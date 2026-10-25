using System.Windows.Automation;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// Reads the address bar of Edge, Chrome, Brave and Firefox through Windows UI Automation,
/// so time logs show the actual site (e.g. qbo.intuit.com) and not just the page title.
/// Only looked up when the window or page changes, then cached.
/// </summary>
static class BrowserUrl
{
    static readonly HashSet<string> Browsers = new(StringComparer.OrdinalIgnoreCase)
        { "msedge", "chrome", "brave", "firefox", "vivaldi", "opera" };

    static readonly Dictionary<string, string?> Cache = new();

    public static bool IsBrowser(string process) => Browsers.Contains(Classifier.Normalize(process));

    public static string? TryGet(IntPtr hwnd, string process, string title)
    {
        if (hwnd == IntPtr.Zero || !IsBrowser(process)) return null;
        var key = $"{hwnd}|{title}";
        if (Cache.TryGetValue(key, out var cached)) return cached;

        string? url = null;
        try
        {
            var root = AutomationElement.FromHandle(hwnd);
            var edits = root.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
            foreach (AutomationElement e in edits)
            {
                var name = e.Current.Name ?? "";
                // Chrome/Edge: "Address and search bar"; Firefox: "Search with ... or enter address"
                var looksLikeAddressBar = name.Contains("address", StringComparison.OrdinalIgnoreCase)
                                          || name.Contains("URL", StringComparison.OrdinalIgnoreCase);
                if (!looksLikeAddressBar && edits.Count > 1) continue;
                if (e.TryGetCurrentPattern(ValuePattern.Pattern, out var p))
                {
                    url = Normalize(((ValuePattern)p).Current.Value);
                    if (url != null) break;
                }
            }
        }
        catch { /* the browser may be busy or the window gone; just skip */ }

        if (Cache.Count > 500) Cache.Clear();
        Cache[key] = url;
        return url;
    }

    static string? Normalize(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var v = value.Trim();
        if (v.Contains(' ') || !v.Contains('.')) return null; // a search phrase, not an address
        if (!v.StartsWith("http", StringComparison.OrdinalIgnoreCase)) v = "https://" + v;
        return v.Length > 500 ? v[..500] : v;
    }
}
