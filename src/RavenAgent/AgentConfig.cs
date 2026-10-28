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

    // Sync bookmarks
    public DateTime? LastMailSyncUtc { get; set; }
    public bool OutlookIdsBackfilled { get; set; }
    public string? LastQboCompanyId { get; set; }

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
