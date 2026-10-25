using System.Diagnostics;
using System.Windows.Automation;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// Reads the text visible in browser, Claude, Copilot and Teams windows (via Windows UI Automation)
/// and looks for client names in it, on a background thread so the tray never freezes.
/// The text itself never leaves this PC; only the matched client is reported.
/// </summary>
sealed class ScreenContext : IDisposable
{
    public sealed record Result(string? Detail, ClientMatch? Match, DateTime ReadAt);

    static readonly HashSet<string> Apps = new(StringComparer.OrdinalIgnoreCase)
    {
        "msedge", "chrome", "brave", "firefox", "vivaldi", "opera",           // browsers (QBO, NetSuite, claude.ai, web Copilot)
        "claude", "m365copilot", "copilot", "microsoft.copilot", "ms-copilot", // Claude and Copilot desktop apps
        "ms-teams", "teams", "olk",                                            // Teams chats, new Outlook
    };

    const int MaxElements = 4000;
    static readonly TimeSpan Budget = TimeSpan.FromMilliseconds(1500);
    static readonly TimeSpan Refresh = TimeSpan.FromSeconds(30); // pages change without a title change (QBO client switch, chat replies)

    readonly ClientDirectory _clients;
    readonly Thread _thread;
    readonly AutoResetEvent _wake = new(false);
    readonly object _lock = new();
    volatile bool _stop;

    (IntPtr hwnd, string key)? _target;
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
            _target = (hwnd, key);
            var stale = _result == null || DateTime.UtcNow - _result.ReadAt > Refresh;
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
            (IntPtr hwnd, string key)? target;
            lock (_lock) target = _target;
            if (target == null) continue;
            lock (_lock)
            {
                if (_resultKey == target.Value.key && _result != null && DateTime.UtcNow - _result.ReadAt < Refresh) continue;
            }
            var text = ReadText(target.Value.hwnd);
            var match = _clients.Matcher.MatchText(text, requireClearWinner: true);
            var detail = match != null ? $"On screen: {match.ClientName}" : null;
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
            var root = AutomationElement.FromHandle(hwnd);
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

    public void Dispose()
    {
        _stop = true;
        _wake.Set();
    }
}
