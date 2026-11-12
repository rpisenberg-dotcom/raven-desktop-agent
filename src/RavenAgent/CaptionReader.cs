using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// Teams live captions -> a running transcript of the meeting, so Raven can summarise it without a Microsoft connection.
/// Works for Teams in a browser tab (Edge/Chrome, any account) and the Teams app, as long as Live captions are showing.
///
/// Only while Teams says the meeting is being recorded or transcribed (everyone in the call has been told), so this
/// never captures a conversation that isn't already being transcribed. Nothing is recorded: it copies the caption text
/// Teams puts on screen.
/// </summary>
sealed class CaptionReader
{
    public sealed record Line(string Speaker, string Text, DateTime At);

    sealed class Session
    {
        public required string Key, Title;
        public DateTime Started, LastSeen;
        public readonly List<Line> Lines = new();
        public int Sent; // lines already sent
        public int Pointer; // last matched line, so "Okay." said twice isn't merged
    }

    Session? _current;
    readonly List<Session> _finished = new();
    readonly object _lock = new();

    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    delegate bool EnumProc(IntPtr h, IntPtr l);
    static readonly string[] Hosts = { "msedge", "chrome", "ms-teams", "teams", "brave" };

    public bool HasSession { get { lock (_lock) return _current != null; } }

    /// <summary>One look (every few seconds during a call). Safe to call on a background thread.</summary>
    public void Tick()
    {
        var win = FindMeetingWindow();
        if (win == null) { EndCurrent(TimeSpan.FromMinutes(2)); return; }
        try
        {
            var root = AutomationElement.FromHandle(win.Value.hwnd);
            var transcribing = root.FindFirst(TreeScope.Descendants, new OrCondition(
                new PropertyCondition(AutomationElement.AutomationIdProperty, "recording-indicator-custom"),
                new PropertyCondition(AutomationElement.AutomationIdProperty, "recording-indicator"),
                new PropertyCondition(AutomationElement.AutomationIdProperty, "transcription-indicator")));
            if (transcribing == null) return;
            var panel = root.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group),
                new PropertyCondition(AutomationElement.NameProperty, "Live Captions")));
            if (panel == null) return;
            var seen = ReadPanel(panel);
            if (seen.Count == 0) return;
            var title = MeetingTitle(win.Value.title);
            lock (_lock)
            {
                if (_current != null && _current.Title != title) EndCurrent(TimeSpan.Zero);
                _current ??= new Session { Key = Guid.NewGuid().ToString("N"), Title = title, Started = DateTime.UtcNow };
                _current.LastSeen = DateTime.UtcNow;
                Merge(_current, seen);
            }
        }
        catch (Exception ex) { Log.Error("Reading captions", ex); }
    }

    /// <summary>Lines not sent yet (whole sessions that ended, plus the new part of the current one).</summary>
    public List<LiveTranscriptRecord> TakeUnsent()
    {
        lock (_lock)
        {
            var list = new List<LiveTranscriptRecord>();
            foreach (var s in _finished.Concat(_current != null ? new[] { _current } : Array.Empty<Session>()))
            {
                // The last line may still be growing: hold it back while the meeting is going
                var upto = s == _current ? Math.Max(s.Sent, s.Lines.Count - 1) : s.Lines.Count;
                if (upto <= s.Sent && s == _current) continue; // nothing new yet (a finished one is still sent, to say it ended)
                list.Add(new LiveTranscriptRecord
                {
                    MeetingKey = s.Key, Title = s.Title, StartedAt = Json.Iso(s.Started), LastSeenAt = Json.Iso(s.LastSeen),
                    Ended = s != _current, FromLine = s.Sent,
                    Lines = s.Lines.Skip(s.Sent).Take(upto - s.Sent).Select(l => new LiveTranscriptLine { Speaker = l.Speaker, Text = l.Text, At = Json.Iso(l.At) }).ToList(),
                });
            }
            return list;
        }
    }

    /// <summary>The server has these; don't send them again.</summary>
    public void MarkSent(List<LiveTranscriptRecord> sent)
    {
        lock (_lock)
        {
            foreach (var r in sent)
            {
                var s = _current?.Key == r.MeetingKey ? _current : _finished.FirstOrDefault(f => f.Key == r.MeetingKey);
                if (s != null) s.Sent = Math.Max(s.Sent, r.FromLine + r.Lines.Count);
            }
            _finished.RemoveAll(f => f.Sent >= f.Lines.Count);
        }
    }

    void EndCurrent(TimeSpan grace)
    {
        lock (_lock)
        {
            if (_current == null || DateTime.UtcNow - _current.LastSeen < grace) return;
            Log.Info($"Captions: meeting '{_current.Title}' ended with {_current.Lines.Count} lines");
            _finished.Add(_current);
            _current = null;
        }
    }

    /// <summary>Speaker / text pairs in the captions panel, oldest first.</summary>
    static List<(string speaker, string text)> ReadPanel(AutomationElement panel)
    {
        var texts = panel.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
        var list = new List<(string, string)>();
        // Each caption is a group holding two Text items: the speaker, then what they said
        foreach (AutomationElement t in texts)
        {
            var parent = TreeWalker.ControlViewWalker.GetParent(t);
            if (parent == null) continue;
            var kids = parent.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
            if (kids.Count < 2 || !Equals(kids[kids.Count - 1], t)) continue; // only handle each caption once, at its last Text
            var speaker = (kids[0].Current.Name ?? "").Trim();
            var said = (t.Current.Name ?? "").Trim();
            if (said.Length > 0) list.Add((speaker, said));
        }
        return list;
    }

    /// <summary>Adds what's on screen now to the transcript: a caption still being spoken grows in place.</summary>
    static void Merge(Session s, List<(string speaker, string text)> seen)
    {
        var p = Math.Max(s.Pointer, Math.Max(0, s.Lines.Count - 60));
        foreach (var (speaker, text) in seen)
        {
            var hit = -1;
            for (var i = p; i < s.Lines.Count; i++)
            {
                var l = s.Lines[i];
                if (l.Speaker == speaker && (l.Text == text || text.StartsWith(l.Text, StringComparison.Ordinal) || l.Text.StartsWith(text, StringComparison.Ordinal))) { hit = i; break; }
            }
            if (hit >= 0)
            {
                if (text.Length > s.Lines[hit].Text.Length && hit >= s.Sent) s.Lines[hit] = s.Lines[hit] with { Text = text };
                p = hit + 1;
            }
            else
            {
                s.Lines.Add(new Line(speaker, text, DateTime.UtcNow));
                p = s.Lines.Count;
            }
        }
        s.Pointer = Math.Max(0, p - 1);
    }

    /// <summary>"Calendar | BBV Team Meeting (External) | Microsoft Teams and 1 more page - Work" -> "BBV Team Meeting".</summary>
    static string MeetingTitle(string window)
    {
        var parts = window.Split(" | ").Select(x => x.Trim()).ToList();
        var name = parts.FirstOrDefault(x => !x.StartsWith("Microsoft Teams") && x is not ("Calendar" or "Chat" or "Teams" or "Activity" or "Calls" or "Meeting") ) ?? "Teams meeting";
        name = System.Text.RegularExpressions.Regex.Replace(name, @"\s*\((External|Guest)\)\s*$", "");
        return name.Length > 200 ? name[..200] : name;
    }

    static (IntPtr hwnd, string title)? FindMeetingWindow()
    {
        (IntPtr, string)? found = null;
        EnumWindows((h, _) =>
        {
            if (!IsWindowVisible(h)) return true;
            var sb = new StringBuilder(512);
            GetWindowText(h, sb, sb.Capacity);
            var title = sb.ToString();
            if (!title.Contains("Microsoft Teams", StringComparison.OrdinalIgnoreCase) && !title.Contains("Meeting", StringComparison.OrdinalIgnoreCase)) return true;
            GetWindowThreadProcessId(h, out var pid);
            string proc;
            try { proc = Process.GetProcessById((int)pid).ProcessName.ToLowerInvariant(); } catch { return true; }
            if (!Hosts.Contains(proc)) return true;
            try
            {
                // The meeting's own window has the Leave button
                var el = AutomationElement.FromHandle(h).FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "hangup-button"));
                if (el != null) { found = (h, title); return false; }
            }
            catch { }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
