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
            // Listening in a meeting or call isn't idle, even without typing (and even with another window in front)
            var call = CallDetector.ActiveCall();
            var idle = Native.IdleTime() >= TimeSpan.FromMinutes(_cfg.IdleMinutes) && !Classifier.IsMeeting(process, title) && call == null;
            var ignored = Classifier.IsIgnored(process, title, _cfg.IgnoreApps, _cfg.IgnoreTitleWords);

            CurrentWindow = idle ? "Idle" : ignored ? "Private window" : $"{process}: {title}";
            var url = idle || ignored ? null : BrowserUrl.TryGet(hwnd, process, title);
            var qbo = QboCompany(url);

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
                detail, match?.ClientId, match?.Term, match != null ? source : null, qbo, call)));
        }
        catch (Exception ex) { Log.Error("Sampling failed", ex); }
    }

    public void Flush() => Enqueue(_builder.Flush(DateTime.UtcNow));

    static readonly System.Text.RegularExpressions.Regex QboIdInUrl =
        new(@"[?&#](?:companyId|realmId|cid)=(\d{6,})", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    /// <summary>
    /// Which QuickBooks Online company is open. QuickBooks only shows the company id in the address when
    /// switching companies (e.g. switchCompany?companyId=...), so the last one seen is remembered for every
    /// QuickBooks page after it, until the next switch.
    /// </summary>
    string? QboCompany(string? url)
    {
        if (string.IsNullOrEmpty(url) || url.IndexOf("intuit.com", StringComparison.OrdinalIgnoreCase) < 0) return null;
        var m = QboIdInUrl.Match(url);
        if (m.Success && m.Groups[1].Value != _cfg.LastQboCompanyId)
        {
            _cfg.LastQboCompanyId = m.Groups[1].Value;
            try { _cfg.Save(); } catch { }
        }
        return url.IndexOf("qbo.intuit.com", StringComparison.OrdinalIgnoreCase) >= 0 ? _cfg.LastQboCompanyId : null;
    }

    void Enqueue(IReadOnlyList<TimeLogRecord> logs)
    {
        if (logs.Count > 0) _queue.Add(logs);
    }
}
