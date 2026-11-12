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
    /// <summary>The client last seen on screen (for pre-filling a quick note), and when.</summary>
    public ClientMatch? LastMatch { get; private set; }
    public DateTime LastMatchUtc { get; private set; }
    /// <summary>What was in front most recently (app and title, or the call), saved with a quick note.</summary>
    public string LastContext { get; private set; } = "";

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
            // Reading an AI answer or a PDF without touching anything is still work: 15 minutes before "away" there
            var idleAfter = TimeSpan.FromMinutes(Classifier.IsReading(process, title) ? Math.Max(15, _cfg.IdleMinutes) : _cfg.IdleMinutes);
            var idle = Native.IdleTime() >= idleAfter && !Classifier.IsMeeting(process, title) && call == null;
            // The desktop, taskbar, Start, search or Alt+Tab: not work in any program, so not recorded
            var ignored = Classifier.IsIgnored(process, title, _cfg.IgnoreApps, _cfg.IgnoreTitleWords) || Native.IsWindowsShell(hwnd, process, title);

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
                    (detail, match) = QuickBooksCompany(hwnd, url, detail, match, read: r != null);
                }
                if (match == null)
                {
                    match = _clients.Matcher.MatchText(title);
                    if (match != null) source = "title";
                }
                if (match != null) { CurrentWindow += $" [{match.ClientName}]"; LastMatch = match; LastMatchUtc = now; }
                if (!process.Contains("RavenAgent", StringComparison.OrdinalIgnoreCase))
                    LastContext = call != null ? $"Call: {call}" : detail is { Length: > 0 } && detail.Length < 80 ? $"{process}: {detail}" : $"{process}: {title}";
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
    bool _qboLoaded;

    void RememberQbo(ClientMatch? m)
    {
        _qboClient = m;
        _qboSeen = DateTime.UtcNow;
        _cfg.LastQboClientId = m?.ClientId;
        _cfg.LastQboClientName = m?.ClientName;
        _cfg.LastQboSeenUtc = m != null ? _qboSeen : null;
        try { _cfg.Save(); } catch { }
    }

    (string? detail, ClientMatch? match) QuickBooksCompany(IntPtr hwnd, string? url, string? detail, ClientMatch? match, bool read)
    {
        if (string.IsNullOrEmpty(url) || url.IndexOf("qbo.intuit.com", StringComparison.OrdinalIgnoreCase) < 0) return (detail, match);
        if (!_qboLoaded)
        {
            // pick up where the agent left off (it restarts for updates and after reboots)
            _qboLoaded = true;
            if (_cfg.LastQboClientId != null && _cfg.LastQboSeenUtc is { } seen)
            {
                _qboClient = new ClientMatch(_cfg.LastQboClientId, _cfg.LastQboClientName ?? "", _cfg.LastQboClientName ?? "", 1);
                _qboSeen = seen;
            }
        }
        // Your own practice pages (client list, switching): not inside any client's books
        if (QboPractice.IsMatch(url)) { if (_qboClient != null) RememberQbo(null); return (null, null); }
        // Full-screen forms hide the header: the client from the header carries over, but only for a while
        var fresh = _qboClient != null && DateTime.UtcNow - _qboSeen < TimeSpan.FromMinutes(45);
        if (QboForm.IsMatch(url))
            return fresh ? ($"On screen: {_qboClient!.ClientName}", _qboClient) : (null, null);
        if (match != null)
        {
            if (_qboClient?.ClientId != match.ClientId || DateTime.UtcNow - _qboSeen > TimeSpan.FromMinutes(5)) RememberQbo(match);
            else _qboSeen = DateTime.UtcNow;
            return (detail, match);
        }
        // An ordinary page whose header was read and names no client you have: you've switched to books Raven
        // doesn't recognise, so the last client is forgotten (it used to carry on for 2 hours, labelling a
        // different client's work). Until the page has been read, the last client stands in briefly.
        if (read) { if (_qboClient != null) { Log.Info($"QuickBooks: header no longer shows {_qboClient.ClientName}"); RememberQbo(null); } return (null, null); }
        return fresh && DateTime.UtcNow - _qboSeen < TimeSpan.FromMinutes(2) ? ($"On screen: {_qboClient!.ClientName}", _qboClient) : (null, null);
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
