using System.Diagnostics;
using System.Windows.Automation;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// Reads the text visible in browser, Claude, Copilot, Teams and Phone Link windows (via Windows UI Automation)
/// and looks for client names in it, on a background thread so the tray never freezes.
/// The page text itself never leaves this PC; only the matched client is reported, plus a short label so
/// Raven can assign the time: the chat's title (Claude, Copilot) or the contact (Phone Link texts and calls).
/// </summary>
sealed class ScreenContext : IDisposable
{
    public sealed record Result(string? Detail, ClientMatch? Match, DateTime ReadAt);

    static readonly HashSet<string> Apps = new(StringComparer.OrdinalIgnoreCase)
    {
        "msedge", "chrome", "brave", "firefox", "vivaldi", "opera",           // browsers (QBO, NetSuite, claude.ai, web Copilot)
        "claude", "m365copilot", "copilot", "microsoft.copilot", "ms-copilot", // Claude and Copilot desktop apps
        "ms-teams", "teams", "olk",                                            // Teams chats, new Outlook
        "phoneexperiencehost", "phonelink",                                    // Phone Link (texts and calls from the phone)
    };
    static readonly HashSet<string> AiApps = new(StringComparer.OrdinalIgnoreCase) { "claude", "m365copilot", "copilot", "microsoft.copilot", "ms-copilot" };
    static readonly HashSet<string> PhoneApps = new(StringComparer.OrdinalIgnoreCase) { "phoneexperiencehost", "phonelink" };
    static readonly HashSet<string> NavWords = new(StringComparer.OrdinalIgnoreCase)
    { "messages", "calls", "photos", "apps", "notifications", "settings", "contacts", "chats", "all", "unread", "home", "new chat", "chat" };

    const int MaxElements = 4000;
    static readonly TimeSpan Budget = TimeSpan.FromMilliseconds(1500);
    static readonly TimeSpan Refresh = TimeSpan.FromSeconds(30); // pages change without a title change (QBO client switch, chat replies)

    readonly ClientDirectory _clients;
    readonly Thread _thread;
    readonly AutoResetEvent _wake = new(false);
    readonly object _lock = new();
    volatile bool _stop;

    (IntPtr hwnd, string key, string process)? _target;
    string? _resultKey;
    Result? _result;

    public ScreenContext(ClientDirectory clients)
    {
        _clients = clients;
        _thread = new Thread(Loop) { IsBackground = true, Name = "ScreenContext" };
        _thread.SetApartmentState(ApartmentState.MTA);
        _thread.Start();
    }

    public static bool Handles(string process) => Apps.Contains(Classifier.Normalize(process));

    /// <summary>Latest result for this window/page, if one has been read. Also asks the worker to (re)read it.</summary>
    public Result? Get(IntPtr hwnd, string process, string title, string? url)
    {
        if (!Handles(process) || hwnd == IntPtr.Zero) return null;
        var key = $"{hwnd}|{title}|{url}";
        lock (_lock)
        {
            var changed = _target?.key != key;
            _target = (hwnd, key, Classifier.Normalize(process));
            // Phone Link switches conversations without changing its title, so check it more often
            var refresh = PhoneApps.Contains(Classifier.Normalize(process)) ? TimeSpan.FromSeconds(5) : Refresh;
            var stale = _result == null || DateTime.UtcNow - _result.ReadAt > refresh;
            if (changed || stale) _wake.Set();
            return _resultKey == key ? _result : null;
        }
    }

    void Loop()
    {
        while (!_stop)
        {
            _wake.WaitOne(TimeSpan.FromSeconds(5));
            if (_stop) break;
            (IntPtr hwnd, string key, string process)? target;
            lock (_lock) target = _target;
            if (target == null) continue;
            var isPhone = PhoneApps.Contains(target.Value.process);
            lock (_lock)
            {
                var refresh = isPhone ? TimeSpan.FromSeconds(5) : Refresh;
                if (_resultKey == target.Value.key && _result != null && DateTime.UtcNow - _result.ReadAt < refresh) continue;
            }
            ClientMatch? match;
            string? detail;
            if (isPhone)
            {
                // The open conversation or call: the selected item in Phone Link's list
                var contact = SelectedContact(target.Value.hwnd);
                match = contact != null ? _clients.Matcher.MatchText(contact) : null;
                detail = contact != null ? $"Phone Link: {contact}" : null;
            }
            else
            {
                var text = ReadText(target.Value.hwnd);
                match = _clients.Matcher.MatchText(text, requireClearWinner: true);
                detail = match != null ? $"On screen: {match.ClientName}" : null;
                // Claude / Copilot: label the time with the chat's title so it can be assigned later
                if (AiApps.Contains(target.Value.process))
                {
                    var chat = ChatTitle(target.Value.hwnd);
                    if (chat != null) detail = $"Chat: {chat}";
                }
            }
            lock (_lock)
            {
                _resultKey = target.Value.key;
                _result = new Result(detail, match, DateTime.UtcNow);
            }
        }
    }

    /// <summary>Names of on-screen elements, top of the window first (company headers come early).</summary>
    static string ReadText(IntPtr hwnd)
    {
        var sb = new System.Text.StringBuilder();
        var watch = Stopwatch.StartNew();
        try
        {
            var window = AutomationElement.FromHandle(hwnd);
            // Only the page itself: the browser's tab strip, bookmarks bar and address bar list
            // other tabs' titles, which matched clients that weren't on screen
            var root = window.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document)) ?? window;
            var walker = TreeWalker.ControlViewWalker;
            var queue = new Queue<AutomationElement>();
            queue.Enqueue(root);
            int seen = 0;
            while (queue.Count > 0 && seen < MaxElements && watch.Elapsed < Budget && sb.Length < 200_000)
            {
                var el = queue.Dequeue();
                seen++;
                try
                {
                    var name = el.Current.Name;
                    if (!string.IsNullOrWhiteSpace(name) && name.Length < 2000) sb.Append(name).Append('\n');
                    for (var child = walker.GetFirstChild(el); child != null; child = walker.GetNextSibling(child))
                        queue.Enqueue(child);
                }
                catch (ElementNotAvailableException) { }
            }
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or System.Runtime.InteropServices.COMException) { }
        return sb.ToString();
    }

    /// <summary>Contact of the conversation or call open in Phone Link (the selected list item), or null.</summary>
    static string? SelectedContact(IntPtr hwnd)
    {
        try
        {
            var window = AutomationElement.FromHandle(hwnd);
            var selected = window.FindAll(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                new PropertyCondition(SelectionItemPattern.IsSelectedProperty, true)));
            foreach (AutomationElement el in selected)
            {
                var name = (el.Current.Name ?? "").Trim();
                if (name.Length == 0 || NavWords.Contains(name)) continue;
                // "Jane Smith, Can you send..., 10:42 AM" -> "Jane Smith"
                var first = name.Split(new[] { '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                if (first.Length > 0 && first.Length <= 80 && !NavWords.Contains(first)) return first;
            }
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or System.Runtime.InteropServices.COMException) { }
        return null;
    }

    /// <summary>The open chat's title in Claude / Copilot (the web page's title inside the app), or null.</summary>
    static string? ChatTitle(IntPtr hwnd)
    {
        try
        {
            var window = AutomationElement.FromHandle(hwnd);
            var doc = window.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document));
            var name = (doc?.Current.Name ?? "").Trim();
            foreach (var suffix in new[] { " - Claude", " | Microsoft Copilot", " - Microsoft Copilot", " - Copilot", " | Copilot" })
                if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) name = name[..^suffix.Length].Trim();
            if (name.Length < 3 || NavWords.Contains(name) || name.Equals("Claude", StringComparison.OrdinalIgnoreCase)
                || name.Contains("Copilot", StringComparison.OrdinalIgnoreCase)) return null;
            return name.Length > 120 ? name[..120] : name;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or System.Runtime.InteropServices.COMException) { }
        return null;
    }

    public void Dispose()
    {
        _stop = true;
        _wake.Set();
    }
}
