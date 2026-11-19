using Microsoft.Win32;

namespace RavenAgent;

/// <summary>
/// Knows when Raven is in a call (Teams, Zoom, Google Meet / Webex in a browser, Phone Link), even when the
/// meeting window is behind Excel or Outlook while she shares her screen. Windows records which app is using
/// the microphone right now (the same thing that shows the mic icon in the taskbar); nothing is listened to.
/// </summary>
static class CallDetector
{
    const string MicKey = @"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone";

    static readonly (string match, string label)[] CallApps =
    {
        ("teams", "Teams"), ("zoom", "Zoom"), ("webex", "Webex"), ("ciscocollab", "Webex"), ("gotomeeting", "GoTo"),
        ("yourphone", "Phone Link"), ("phonelink", "Phone Link"), ("phoneexperiencehost", "Phone Link"),
        ("chrome", "Browser"), ("msedge", "Browser"), ("firefox", "Browser"), ("brave", "Browser"), ("opera", "Browser"),
    };

    /// <summary>The call app using the microphone right now ("Teams", "Zoom", "Browser", "Phone Link"), or null.
    /// A meeting app wins over Phone Link, which can hold the microphone open in the background.</summary>
    public static string? ActiveCall()
    {
        var apps = InUse();
        return apps.FirstOrDefault(a => a.label != "Phone Link").label ?? apps.FirstOrDefault().label;
    }

    /// <summary>Every app Windows says is using the microphone now: (registry name, call label or null).</summary>
    public static List<(string key, string? label)> InUse()
    {
        var list = new List<(string, string?)>();
        try
        {
            using var root = Registry.CurrentUser.OpenSubKey(MicKey);
            if (root == null) return list;
            foreach (var name in root.GetSubKeyNames())
            {
                if (name.Equals("NonPackaged", StringComparison.OrdinalIgnoreCase))
                {
                    using var np = root.OpenSubKey(name);
                    if (np == null) continue;
                    foreach (var exe in np.GetSubKeyNames())
                    {
                        using var k = np.OpenSubKey(exe);
                        if (InUse(k)) list.Add((exe, Label(exe)));
                    }
                }
                else
                {
                    using var k = root.OpenSubKey(name);
                    if (InUse(k)) list.Add((name, Label(name)));
                }
            }
        }
        catch { /* no access: treat as not in a call */ }
        return list.Where(x => x.Item2 != null).Concat(list.Where(x => x.Item2 == null)).ToList();
    }

    // In use = started and not yet stopped
    static bool InUse(RegistryKey? k) =>
        k != null && k.GetValue("LastUsedTimeStop") is long stop && stop == 0
        && k.GetValue("LastUsedTimeStart") is long start && start > 0;

    static string? Label(string keyName)
    {
        var n = keyName.ToLowerInvariant();
        foreach (var (match, label) in CallApps) if (n.Contains(match)) return label;
        return null;
    }
}
