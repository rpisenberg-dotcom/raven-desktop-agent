using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace RavenAgent.Core;

/// <summary>Turns a process name + window title into an activity type, and decides what to skip.</summary>
public static class Classifier
{
    static readonly HashSet<string> Browsers = new(StringComparer.OrdinalIgnoreCase)
        { "msedge", "chrome", "firefox", "brave", "opera", "vivaldi", "arc" };
    static readonly HashSet<string> Documents = new(StringComparer.OrdinalIgnoreCase)
        { "excel", "winword", "powerpnt", "acrobat", "acrord32", "onenote", "notepad", "qbw32", "qbw" };
    static readonly HashSet<string> MeetingApps = new(StringComparer.OrdinalIgnoreCase)
        { "ms-teams", "teams", "zoom", "webex", "ciscocollabhost" };
    static readonly Regex MeetingTitle = new(@"\b(meeting|call|huddle|zoom meeting|webinar)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    static readonly Regex MeetingUrl = new(@"meet\.google\.com|teams\.microsoft\.com/.*/meetup|zoom\.us/(j|wc)/", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Apps that mean "nobody is working" (lock screen etc.).</summary>
    static readonly HashSet<string> Away = new(StringComparer.OrdinalIgnoreCase) { "lockapp", "logonui", "screensaver" };

    public static string Normalize(string process) =>
        Path.GetFileNameWithoutExtension(process ?? "").Trim().ToLowerInvariant();

    public static bool IsAway(string process) => Away.Contains(Normalize(process));

    public static bool IsMeeting(string process, string title) =>
        (MeetingApps.Contains(Normalize(process)) && MeetingTitle.IsMatch(title ?? ""))
        || MeetingUrl.IsMatch(title ?? "");

    public static string ActivityType(string process, string title)
    {
        var p = Normalize(process);
        if (IsMeeting(process, title)) return "meeting";
        if (p == "outlook" || p == "olk") return "email";
        if (Browsers.Contains(p)) return "browser";
        if (Documents.Contains(p)) return "document";
        return "app_usage";
    }

    /// <summary>True when the window should never be recorded (private browsing, ignore list).</summary>
    public static bool IsIgnored(string process, string title, IEnumerable<string> ignoreApps, IEnumerable<string> ignoreTitleWords)
    {
        var p = Normalize(process);
        if (ignoreApps.Any(a => Normalize(a) == p)) return true;
        var t = title ?? "";
        if (t.Contains("InPrivate", StringComparison.OrdinalIgnoreCase) || t.Contains("Incognito", StringComparison.OrdinalIgnoreCase))
            return true;
        return ignoreTitleWords.Any(w => !string.IsNullOrWhiteSpace(w) && t.Contains(w.Trim(), StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Short stable hash used for idempotency keys.</summary>
    public static string Key(params string?[] parts)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("|", parts.Select(p => p ?? ""))));
        return Convert.ToHexString(bytes, 0, 12).ToLowerInvariant();
    }
}
