using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>Settings saved in %APPDATA%\RavenAgent\config.json. The agent key is encrypted for this Windows user.</summary>
sealed class AgentConfig
{
    public string IngestUrl { get; set; } = "";
    public string AgentKeyProtected { get; set; } = "";
    public string DeviceId { get; set; } = "";
    public int SampleSeconds { get; set; } = 5;
    public int IdleMinutes { get; set; } = 5;
    public int UploadSeconds { get; set; } = 300;
    public int OutlookMinutes { get; set; } = 5;
    public int FirstRunEmailDays { get; set; } = 2;
    public int CalendarDaysAhead { get; set; } = 120; // ~4 months, for planning travel etc.
    public int CalendarDaysBack { get; set; } = 7;    // so this week's earlier days show in week view
    public int MaxEmailBodyChars { get; set; } = 8000;
    public bool ReadOutlook { get; set; } = true;
    public bool PrivateMode { get; set; }
    public List<string> IgnoreApps { get; set; } = new() { "Spotify", "Netflix" };
    public List<string> IgnoreTitleWords { get; set; } = new();
    /// <summary>Mailboxes Raven chose not to have read (set in Clerical Guru, sent with each upload reply).</summary>
    public List<string> IgnoredMailboxes { get; set; } = new();
    /// <summary>Calendars set to Hidden in Clerical Guru: not read from Outlook at all (less work for Outlook).</summary>
    public List<string> HiddenCalendars { get; set; } = new();
    /// <summary>Copy Teams live captions into the meeting while Teams is recording/transcribing it.</summary>
    public bool CaptureCaptions { get; set; } = true;

    // Meeting recordings (on this computer, only after Raven confirms everyone was told)
    /// <summary>Ask "Record this meeting?" when a Teams / Zoom / browser meeting starts.</summary>
    public bool OfferRecording { get; set; } = true;
    /// <summary>Last video choice: window (the meeting window), screen, or none (sound only).</summary>
    public string RecordVideo { get; set; } = "window";
    public string RecordingsFolder { get; set; } = DefaultRecordingsFolder();
    public string RecordingNotice { get; set; } = "Heads-up: I'm recording and transcribing this meeting for my notes. Let me know if you'd rather I didn't.";
    /// <summary>Raven's name on her own lines of a transcript (the captions' version, e.g. "Raven Pick", is used when available).</summary>
    public string MyName { get; set; } = "Raven";

    /// <summary>Documents\Raven Recordings, or a Clerical Guru OneDrive if there is one. Never a client's OneDrive
    /// (a work OneDrive on this computer can belong to a client's company).</summary>
    public static string DefaultRecordingsFolder()
    {
        foreach (var v in new[] { "OneDriveCommercial", "OneDrive" })
        {
            var od = Environment.GetEnvironmentVariable(v);
            if (!string.IsNullOrEmpty(od) && Directory.Exists(od) && IsOwnOneDrive(od)) return Path.Combine(od, "Raven Recordings");
        }
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Raven Recordings");
    }

    /// <summary>A personal OneDrive ("OneDrive", "OneDrive - Personal") or Clerical Guru's own.</summary>
    static bool IsOwnOneDrive(string path)
    {
        var name = Path.GetFileName(path.TrimEnd('\\'));
        return name.Equals("OneDrive", StringComparison.OrdinalIgnoreCase)
            || name.Contains("Personal", StringComparison.OrdinalIgnoreCase)
            || name.Replace(" ", "").Contains("clericalguru", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Recordings went to a client's OneDrive in 0.9.0: move the setting back to the safe default.</summary>
    public bool FixRecordingsFolder()
    {
        var name = RecordingsFolder ?? "";
        var parent = Path.GetFileName(Path.GetDirectoryName(name.TrimEnd('\\')) ?? "");
        if (!parent.StartsWith("OneDrive - ", StringComparison.OrdinalIgnoreCase) || IsOwnOneDrive(parent)) return false;
        RecordingsFolder = DefaultRecordingsFolder();
        return true;
    }

    // Sync bookmarks
    public DateTime? LastMailSyncUtc { get; set; }
    public bool OutlookIdsBackfilled { get; set; }
    public string? LastQboCompanyId { get; set; }
    // Last client seen in the QuickBooks header (kept across restarts, used on full-screen forms)
    public string? LastQboClientId { get; set; }
    public string? LastQboClientName { get; set; }
    public DateTime? LastQboSeenUtc { get; set; }

    public static string Folder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RavenAgent");
    public static string DataFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RavenAgent");
    static string FilePath => Path.Combine(Folder, "config.json");

    static readonly JsonSerializerOptions Opts = new(Json.Options) { WriteIndented = true };

    [JsonIgnore]
    public bool IsConfigured => !string.IsNullOrWhiteSpace(IngestUrl) && !string.IsNullOrWhiteSpace(AgentKey);

    [JsonIgnore] // only the encrypted copy is saved
    public string AgentKey
    {
        get
        {
            if (string.IsNullOrEmpty(AgentKeyProtected)) return "";
            try
            {
                var bytes = ProtectedData.Unprotect(Convert.FromBase64String(AgentKeyProtected), null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(bytes);
            }
            catch { return ""; }
        }
        set => AgentKeyProtected = string.IsNullOrEmpty(value)
            ? ""
            : Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(value), null, DataProtectionScope.CurrentUser));
    }

    public static AgentConfig Load()
    {
        AgentConfig cfg;
        try { cfg = File.Exists(FilePath) ? JsonSerializer.Deserialize<AgentConfig>(File.ReadAllText(FilePath), Opts) ?? new() : new(); }
        catch { cfg = new(); }
        if (cfg.CalendarDaysAhead == 14) cfg.CalendarDaysAhead = 120; // upgrade the old default
        if (cfg.UploadSeconds < 300) cfg.UploadSeconds = 300; // each upload costs Base44 credits; every 5 minutes is plenty
        if (cfg.FixRecordingsFolder()) Log.Info($"Recordings folder moved off a client's OneDrive to {cfg.RecordingsFolder}");
        if (string.IsNullOrEmpty(cfg.DeviceId)) cfg.DeviceId = $"{Environment.MachineName}-{Guid.NewGuid().ToString("N")[..8]}";
        cfg.Save();
        return cfg;
    }

    public void Save()
    {
        Directory.CreateDirectory(Folder);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(this, Opts));
    }
}
