using RavenAgent.Core;

namespace RavenAgent;

/// <summary>Looks at the active window every few seconds and turns that into time segments.</summary>
sealed class ActivitySampler
{
    readonly AgentConfig _cfg;
    readonly UploadQueue _queue;
    readonly SegmentBuilder _builder;
    readonly ClientDirectory _clients;
    readonly ScreenContext _screen;
    readonly OutlookContext _outlook;

    public string CurrentWindow { get; private set; } = "";

    public ActivitySampler(AgentConfig cfg, UploadQueue queue, ClientDirectory clients, ScreenContext screen, OutlookContext outlook)
    {
        _cfg = cfg;
        _queue = queue;
        _clients = clients;
        _screen = screen;
        _outlook = outlook;
        _builder = new SegmentBuilder(cfg.DeviceId, TimeSpan.FromSeconds(cfg.SampleSeconds));
    }

    /// <summary>Called on a timer from the UI thread.</summary>
    public void Tick()
    {
        try
        {
            var now = DateTime.UtcNow;
            if (_cfg.PrivateMode)
            {
                Enqueue(_builder.Flush(now));
                CurrentWindow = "Paused";
                return;
            }

            var (hwnd, process, title) = Native.ForegroundWindow();
            // Listening in a meeting isn't idle, even without typing
            var idle = Native.IdleTime() >= TimeSpan.FromMinutes(_cfg.IdleMinutes) && !Classifier.IsMeeting(process, title);
            var ignored = Classifier.IsIgnored(process, title, _cfg.IgnoreApps, _cfg.IgnoreTitleWords);

            CurrentWindow = idle ? "Idle" : ignored ? "Private window" : $"{process}: {title}";
            var url = idle || ignored ? null : BrowserUrl.TryGet(hwnd, process, title);

            // Which client is this about? Selected email in Outlook, names on screen elsewhere, then the title
            string? detail = null, source = null;
            ClientMatch? match = null;
            if (!idle && !ignored)
            {
                if (OutlookContext.Handles(process))
                {
                    (detail, match) = _outlook.Get();
                    source = "email";
                }
                else if (ScreenContext.Handles(process))
                {
                    var r = _screen.Get(hwnd, process, title, url);
                    if (r != null) { detail = r.Detail; match = r.Match; }
                    source = "screen";
                }
                if (match == null)
                {
                    match = _clients.Matcher.MatchText(title);
                    if (match != null) source = "title";
                }
                if (match != null) CurrentWindow += $" [{match.ClientName}]";
            }
            Enqueue(_builder.Add(new Sample(now, process, title, idle, ignored, url,
                detail, match?.ClientId, match?.Term, match != null ? source : null)));
        }
        catch (Exception ex) { Log.Error("Sampling failed", ex); }
    }

    public void Flush() => Enqueue(_builder.Flush(DateTime.UtcNow));

    void Enqueue(IReadOnlyList<TimeLogRecord> logs)
    {
        if (logs.Count > 0) _queue.Add(logs);
    }
}
