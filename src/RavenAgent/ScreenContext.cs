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
    const int HeaderLines = 100; // QuickBooks: the page header comes first in the reading order
    static readonly TimeSpan Budget = TimeSpan.FromMilliseconds(1500);
    static readonly TimeSpan Refresh = TimeSpan.FromSeconds(30); // pages change without a title change (QBO client switch, chat replies)

    readonly ClientDirectory _clients;
    readonly Thread _thread;
    readonly AutoResetEvent _wake = new(false);
    readonly object _lock = new();
    volatile bool _stop;

    (IntPtr hwnd, string key, string process, string? url)? _target;
    string? _resultKey;
    Result? _result;

    public ScreenContext(ClientDirectory clients)
    {
        _clients = clients;
        _thread = new Thread(Loop) { IsBackground = true, Name = "ScreenContext" };
        _thread.SetApartmentState(ApartmentState.MTA);
        _thread.Start();
    }

    // Copilot and Claude ship under several process names (copilotapp, m365copilot, claude...), so match by name
    static bool IsAi(string normalized) => AiApps.Contains(normalized) || normalized.Contains("copilot") || normalized.Contains("claude");
    public static bool Handles(string process)
    {
        var n = Classifier.Normalize(process);
        return Apps.Contains(n) || IsAi(n);
    }

    /// <summary>Latest result for this window/page, if one has been read. Also asks the worker to (re)read it.</summary>
    public Result? Get(IntPtr hwnd, string process, string title, string? url)
    {
        if (!Handles(process) || hwnd == IntPtr.Zero) return null;
        var key = $"{hwnd}|{title}|{url}";
        lock (_lock)
        {
            var changed = _target?.key != key;
            _target = (hwnd, key, Classifier.Normalize(process), url);
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
            (IntPtr hwnd, string key, string process, string? url)? target;
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
                // Texts are only read for contacts Raven assigned to a client or Clerical Guru
                if (contact != null && _clients.IsBusinessContact(contact)) QueueConversation(target.Value.hwnd, contact);
            }
            else
            {
                var text = ReadText(target.Value.hwnd);
                // QuickBooks: only the top of the page (the company name in the header) says whose books these are.
                // Names further down are payees and customers ("Amazon Prime" in a bank feed), so they never count.
                if (target.Value.url?.IndexOf("qbo.intuit.com", StringComparison.OrdinalIgnoreCase) >= 0)
                    text = string.Join("\n", text.Split('\n').Where(l => l.Trim().Length > 0).Take(HeaderLines));
                match = _clients.Matcher.MatchText(text, requireClearWinner: true);
                detail = match != null ? $"On screen: {match.ClientName}" : null;
                // Claude / Copilot: label the time with the chat's title so it can be assigned later
                if (IsAi(target.Value.process))
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
    static string ReadText(IntPtr hwnd, AutomationElement? rootOverride = null)
    {
        var sb = new System.Text.StringBuilder();
        var watch = Stopwatch.StartNew();
        try
        {
            var window = AutomationElement.FromHandle(hwnd);
            // Only the page itself: the browser's tab strip, bookmarks bar and address bar list
            // other tabs' titles, which matched clients that weren't on screen
            var root = rootOverride ?? window.FindFirst(TreeScope.Descendants,
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

    readonly List<PhoneSnapshotRecord> _phoneOut = new();
    readonly Dictionary<string, string> _lastConversation = new();

    /// <summary>Text conversations waiting to be sent (taken by the upload).</summary>
    public List<PhoneSnapshotRecord> TakePhoneSnapshots() { lock (_phoneOut) { var l = _phoneOut.ToList(); _phoneOut.Clear(); return l; } }
    public void ReturnPhoneSnapshots(List<PhoneSnapshotRecord> l) { lock (_phoneOut) _phoneOut.InsertRange(0, l); }

    void QueueConversation(IntPtr hwnd, string contact)
    {
        var text = ConversationText(hwnd);
        if (string.IsNullOrWhiteSpace(text)) return;
        lock (_phoneOut)
        {
            if (_lastConversation.TryGetValue(contact, out var prev) && prev == text) return;
            _lastConversation[contact] = text;
            _phoneOut.RemoveAll(x => x.Contact == contact); // only the latest view of a conversation is needed
            _phoneOut.Add(new PhoneSnapshotRecord { Contact = contact, Text = text, CapturedAt = Json.Iso(DateTime.UtcNow) });
        }
    }

    /// <summary>
    /// The messages visible in the open Phone Link conversation (the newest ~40 lines), leaving out the
    /// list of other conversations and the app's own buttons.
    /// </summary>
    static string? ConversationText(IntPtr hwnd)
    {
        try
        {
            // Only the open conversation (the right-hand pane under its header); the conversation list and the
            // Notifications panel on the left show other people's messages and are left out
            var window = AutomationElement.FromHandle(hwnd);
            var w = window.Current.BoundingRectangle;
            if (w.IsEmpty) return null;
            var items = new List<(double top, double left, string text)>();
            foreach (AutomationElement el in window.FindAll(TreeScope.Descendants,
                         new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)))
            {
                try
                {
                    var r = el.Current.BoundingRectangle;
                    var t = (el.Current.Name ?? "").Trim();
                    if (r.IsEmpty || t.Length < 2 || NavWords.Contains(t)) continue;
                    if (r.Left <= w.Left + w.Width * 0.45) continue;                      // list / notifications
                    if (r.Top < w.Top + Math.Max(220, w.Height * 0.2)) continue;          // header (name, number)
                    if (t.StartsWith("Send a message", StringComparison.OrdinalIgnoreCase)) continue; // the reply box
                    items.Add((r.Top, r.Left, t));
                }
                catch (ElementNotAvailableException) { }
            }
            items.Sort((a, b) => a.top != b.top ? a.top.CompareTo(b.top) : a.left.CompareTo(b.left));
            var lines = new List<string>();
            foreach (var (_, _, t) in items) if (lines.Count == 0 || lines[^1] != t) lines.Add(t);
            var tail = lines.Count > 40 ? lines.GetRange(lines.Count - 40, 40) : lines;
            var text = string.Join("\n", tail);
            return text.Length > 4000 ? text[^4000..] : text;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or System.Runtime.InteropServices.COMException) { }
        return null;
    }

    static readonly System.Text.RegularExpressions.Regex ContactIds = new(
        "contact|conversation|thread|header|title|displayname|recipient|caller|participant",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    static string? _lastPhoneLog;

    static string? SelectedContact(IntPtr hwnd)
    {
        string? found = null, how = null;
        try
        {
            var window = AutomationElement.FromHandle(hwnd);
            // 0) The open conversation's header: top of the right-hand pane, the name above the phone number
            //    ("Jonathan Egan" / "+1 312-483-4178"). The tabs (Messages, Calls...) and the list are further left.
            var header = HeaderContact(window);
            if (header != null) { found = header; how = "conversation header"; }
            // 1) The selected conversation or call in Phone Link's list (any kind of list item)
            if (found == null)
            foreach (AutomationElement el in window.FindAll(TreeScope.Descendants,
                         new PropertyCondition(SelectionItemPattern.IsSelectedProperty, true)))
            {
                var c = ContactFrom(el.Current.Name);
                if (c != null) { found = c; how = "selected item"; break; }
            }
            // 2) The conversation / call header (an element named like "ContactName" or "ConversationHeader")
            if (found == null)
            {
                var walker = TreeWalker.ControlViewWalker;
                var queue = new Queue<AutomationElement>();
                queue.Enqueue(window);
                int seen = 0;
                var watch = Stopwatch.StartNew();
                while (queue.Count > 0 && seen < 3000 && watch.Elapsed < Budget)
                {
                    var el = queue.Dequeue();
                    seen++;
                    try
                    {
                        var id = el.Current.AutomationId ?? "";
                        if (id.Length > 0 && ContactIds.IsMatch(id))
                        {
                            var c = ContactFrom(el.Current.Name);
                            if (c != null) { found = c; how = $"element {id}"; break; }
                        }
                        for (var child = walker.GetFirstChild(el); child != null; child = walker.GetNextSibling(child)) queue.Enqueue(child);
                    }
                    catch (ElementNotAvailableException) { }
                }
            }
            // 3) A separate call window titled with the caller ("Call with Jane Smith")
            if (found == null)
            {
                var title = (window.Current.Name ?? "").Trim();
                var c = title.Equals("Phone Link", StringComparison.OrdinalIgnoreCase) ? null
                    : ContactFrom(System.Text.RegularExpressions.Regex.Replace(title, @"^(call with|calling|in call with|on call with)\s+|\s*[-|]\s*phone link$", "",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase));
                if (c != null) { found = c; how = "window title"; }
            }
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or System.Runtime.InteropServices.COMException) { }
        var line = found != null ? $"Phone Link contact: {found} (from {how})" : "Phone Link: couldn't find the contact on screen";
        if (line != _lastPhoneLog) { _lastPhoneLog = line; Log.Info(line); }
        return found;
    }

    static readonly System.Text.RegularExpressions.Regex PhoneNumber = new(@"^\+?[\d\s().\-]{7,}$");

    /// <summary>
    /// Phone Link's conversation header: text in the top band of the window, right of the conversation list.
    /// The first line that isn't a phone number is the contact; for unsaved numbers the number itself is used.
    /// </summary>
    static string? HeaderContact(AutomationElement window)
    {
        var w = window.Current.BoundingRectangle;
        if (w.IsEmpty || w.Width < 300) return null;
        var texts = window.FindAll(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
        var inHeader = new List<(double top, double left, string name)>();
        foreach (AutomationElement el in texts)
        {
            try
            {
                var r = el.Current.BoundingRectangle;
                var name = (el.Current.Name ?? "").Trim();
                if (r.IsEmpty || name.Length == 0) continue;
                var right = r.Left > w.Left + w.Width * 0.45;   // the conversation pane, not the list or notifications
                var top = r.Top < w.Top + Math.Max(220, w.Height * 0.2); // header band under the title bar and tabs
                if (right && top) inHeader.Add((r.Top, r.Left, name));
            }
            catch (ElementNotAvailableException) { }
        }
        inHeader.Sort((a, b) => a.top != b.top ? a.top.CompareTo(b.top) : a.left.CompareTo(b.left));
        string? number = null;
        foreach (var (_, _, name) in inHeader)
        {
            if (PhoneNumber.IsMatch(name)) { number ??= name; continue; }
            var c = ContactFrom(name);
            if (c != null && !c.Equals("Phone Link", StringComparison.OrdinalIgnoreCase)) return c;
        }
        return number;
    }

    /// <summary>"Jane Smith, Can you send..., 10:42 AM" -> "Jane Smith"; null for empty / button-like names.</summary>
    static string? ContactFrom(string? name)
    {
        name = (name ?? "").Trim();
        if (name.Length == 0 || NavWords.Contains(name)) return null;
        var first = name.Split(new[] { '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
        foreach (var prefix in new[] { "Conversation with ", "Message from ", "Messages with ", "Call with ", "Chat with " })
            if (first.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) first = first[prefix.Length..].Trim();
        if (first.Length < 2 || first.Length > 80 || NavWords.Contains(first)) return null;
        return first;
    }

    /// <summary>
    /// Writes the layout of a window (element types, ids and short names) to a file, so support can see how an
    /// app shows its information. Only made when Raven asks for it from the tray menu.
    /// </summary>
    public static string DumpLayout(IntPtr hwnd, string process, string folder)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Window layout for {process} at {DateTime.Now:g}");
        try
        {
            var walker = TreeWalker.ControlViewWalker;
            void Walk(AutomationElement el, int depth, ref int count)
            {
                if (depth > 25 || count > 2500) return;
                count++;
                try
                {
                    var c = el.Current;
                    var name = (c.Name ?? "").Replace("\n", " ");
                    if (name.Length > 80) name = name[..80] + "…";
                    bool sel = false;
                    try { sel = (bool)el.GetCurrentPropertyValue(SelectionItemPattern.IsSelectedProperty); } catch { }
                    sb.Append(' ', depth * 2).Append(c.ControlType.ProgrammaticName.Replace("ControlType.", ""))
                      .Append(c.AutomationId is { Length: > 0 } id ? $" #{id}" : "")
                      .Append(name.Length > 0 ? $" \"{name}\"" : "").Append(sel ? " [selected]" : "").AppendLine();
                    for (var child = walker.GetFirstChild(el); child != null; child = walker.GetNextSibling(child)) Walk(child, depth + 1, ref count);
                }
                catch (ElementNotAvailableException) { }
            }
            int n = 0;
            Walk(AutomationElement.FromHandle(hwnd), 0, ref n);
        }
        catch (Exception ex) { sb.AppendLine($"Stopped: {ex.Message}"); }
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, $"layout-{process}-{DateTime.Now:yyyyMMdd-HHmmss}.txt");
        File.WriteAllText(path, sb.ToString());
        return path;
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
