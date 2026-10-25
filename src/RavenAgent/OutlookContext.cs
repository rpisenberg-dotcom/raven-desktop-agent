using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// While Outlook is in front, notes which email (or meeting) is selected or open, so the time
/// can be tied to that conversation and its client. Runs on the UI thread (an STA thread, as Outlook requires).
/// </summary>
sealed class OutlookContext
{
    const int olMail = 43, olAppointment = 26, olMeetingRequest = 53;
    const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
    static readonly TimeSpan Every = TimeSpan.FromSeconds(10);

    readonly ClientDirectory _clients;
    DateTime _lastRead = DateTime.MinValue;
    (string? detail, ClientMatch? match) _last;
    object? _app;

    public OutlookContext(ClientDirectory clients) => _clients = clients;

    public static bool Handles(string process) => Classifier.Normalize(process) == "outlook";

    public (string? detail, ClientMatch? match) Get()
    {
        if (DateTime.UtcNow - _lastRead < Every) return _last;
        _lastRead = DateTime.UtcNow;
        try { _last = Read(); }
        catch (Exception ex)
        {
            _app = null; // Outlook may have restarted; reconnect next time
            Log.Error("Couldn't read the selected email", ex);
            _last = (null, null);
        }
        return _last;
    }

    (string?, ClientMatch?) Read()
    {
        if (!OutlookReader.IsOutlookRunning()) return (null, null);
        if (_app == null)
        {
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null) return (null, null);
            _app = Activator.CreateInstance(type);
        }
        dynamic app = _app!;
        object? itemObj = null;
        dynamic? window = app.ActiveWindow();
        if (window != null)
        {
            // An open email window, or the selected item in the main window
            try { itemObj = (object)window.CurrentItem; } catch { }
            if (itemObj == null)
            {
                try
                {
                    dynamic sel = window.Selection;
                    if ((int)sel.Count > 0) itemObj = (object)sel.Item(1);
                }
                catch { }
            }
        }
        if (itemObj == null) return (null, null);
        return Describe(itemObj);
    }

    (string?, ClientMatch?) Describe(object itemObj)
    {
        dynamic item = itemObj;
        int cls = (int)item.Class;
        if (cls != olMail && cls != olAppointment && cls != olMeetingRequest) return (null, null);

        string subject = Safe(() => (string)item.Subject) ?? "";
        if (string.IsNullOrWhiteSpace(subject)) subject = "(no subject)";
        var people = new List<string?>();
        string? senderName = null, sender = null;
        if (cls == olMail || cls == olMeetingRequest)
        {
            senderName = Safe(() => (string)item.SenderName);
            sender = SenderAddress(itemObj);
            people.Add(sender);
        }
        try
        {
            dynamic recips = item.Recipients;
            int n = (int)recips.Count;
            for (int i = 1; i <= n && i <= 30; i++)
            {
                dynamic r = recips.Item(i);
                people.Add(Safe(() => (string)r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)) ?? Safe(() => (string)r.Address));
            }
        }
        catch { }

        var m = _clients.Matcher;
        // Addresses are the strongest signal; fall back to names in the subject or sender name
        var match = m.MatchAddresses(people) ?? m.MatchText($"{subject}\n{senderName}");
        var kind = cls == olAppointment ? "Meeting" : "Email";
        var who = senderName ?? sender;
        var detail = $"{kind}: {Trim(subject, 120)}" + (!string.IsNullOrWhiteSpace(who) && cls != olAppointment ? $" — {Trim(who!, 60)}" : "");
        return (detail, match);
    }

    static string? SenderAddress(object mailObj)
    {
        dynamic m = mailObj;
        try
        {
            if ((string)m.SenderEmailType == "EX")
            {
                dynamic? ex = m.Sender?.GetExchangeUser();
                if (ex != null) return (string)ex.PrimarySmtpAddress;
            }
            return (string)m.SenderEmailAddress;
        }
        catch { return null; }
    }

    static string Trim(string s, int max) => s.Length <= max ? s : s[..max] + "…";
    static string? Safe(Func<string> f) { try { return f(); } catch { return null; } }
}
