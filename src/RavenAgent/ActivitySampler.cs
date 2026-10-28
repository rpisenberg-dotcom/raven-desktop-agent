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
                    (detail, match) = QuickBooksCompany(hwnd, url, detail, match);
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

    // QuickBooks Online Accountant: the client's name is in the page header, except on full-screen forms
    // (bills, checks, invoices...), where names on screen are vendors/customers instead. You stay in one
    // client's books until you switch, so the last client seen in the header carries over to those forms.
    static readonly System.Text.RegularExpressions.Regex QboForm = new(
        @"qbo\.intuit\.com/app/(bill|expense|check|invoice|salesreceipt|journal|estimate|purchaseorder|deposit|transfer|creditmemo|receivepayment|paybills|vendorcredit|refundreceipt|delayedcharge|timeactivity|statement|creditcardcredit|cccredit|inventoryqtyadjustment)\b",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    static readonly System.Text.RegularExpressions.Regex QboPractice = new(
        @"qbo\.intuit\.com/app/(practice|clients|switchcompany|accountant|firm|team|proadvisor)",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    ClientMatch? _qboClient;
    DateTime _qboSeen;

    (string? detail, ClientMatch? match) QuickBooksCompany(IntPtr hwnd, string? url, string? detail, ClientMatch? match)
    {
        if (string.IsNullOrEmpty(url) || url.IndexOf("qbo.intuit.com", StringComparison.OrdinalIgnoreCase) < 0) return (detail, match);
        // Your own practice pages (client list, switching): not inside any client's books
        if (QboPractice.IsMatch(url)) { _qboClient = null; return (null, null); }
        var fresh = _qboClient != null && DateTime.UtcNow - _qboSeen < TimeSpan.FromHours(2);
        if (QboForm.IsMatch(url))
            return fresh ? ($"On screen: {_qboClient!.ClientName}", _qboClient) : (null, null);
        if (match != null) { _qboClient = match; _qboSeen = DateTime.UtcNow; return (detail, match); }
        return fresh ? ($"On screen: {_qboClient!.ClientName}", _qboClient) : (detail, match);
    }

    static readonly System.Text.RegularExpressions.Regex QboIdInUrl =
        new(@"[?&#](?:companyId|realmId)=(\d{6,})", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    /// <summary>
    /// The QuickBooks Online company id, only when it's in this page's own address. (An id remembered from an
    /// earlier page was wrong whenever the company was switched without the address showing it, so the
    /// company name on screen decides instead.)
    /// </summary>
    static string? QboCompany(string? url)
    {
        if (string.IsNullOrEmpty(url) || url.IndexOf("qbo.intuit.com", StringComparison.OrdinalIgnoreCase) < 0) return null;
        var m = QboIdInUrl.Match(url);
        return m.Success ? m.Groups[1].Value : null;
    }

    void Enqueue(IReadOnlyList<TimeLogRecord> logs)
    {
        if (logs.Count > 0) _queue.Add(logs);
    }
}
