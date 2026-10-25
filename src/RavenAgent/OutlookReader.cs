using System.Diagnostics;
using System.Runtime.InteropServices;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// Reads mail and calendars from classic Outlook for Windows through its COM interface.
/// Covers every mailbox loaded in Outlook, including ones on clients' Microsoft or Google systems.
/// Only runs while Outlook is open; it never starts Outlook itself.
/// </summary>
sealed class OutlookReader
{
    const int olFolderSentMail = 5, olFolderInbox = 6, olFolderCalendar = 9;
    const int olMail = 43, olAppointment = 26;
    const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
    const string PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
    const int MaxPerFolder = 200;

    readonly AgentConfig _cfg;
    public string Status { get; private set; } = "Not checked yet";

    public OutlookReader(AgentConfig cfg) => _cfg = cfg;

    public sealed record Result(List<EmailRecord> Emails, List<CalendarRecord> Events, List<CalendarSync> Syncs, DateTime ReadAtUtc);

    public static bool IsOutlookRunning() => Process.GetProcessesByName("OUTLOOK").Length > 0;

    /// <summary>Must be called on an STA thread (see <see cref="RunOnSta{T}"/>).</summary>
    public Result? Read(DateTime mailSinceUtc)
    {
        if (!IsOutlookRunning()) { Status = "Outlook is closed - waiting"; return null; }

        var started = DateTime.UtcNow;
        var emails = new List<EmailRecord>();
        var events = new List<CalendarRecord>();
        var syncs = new List<CalendarSync>();
        dynamic? app = null;
        try
        {
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null) { Status = "Classic Outlook isn't installed"; return null; }
            app = Activator.CreateInstance(type)!; // attaches to the running Outlook
            dynamic ns = app.GetNamespace("MAPI");
            dynamic stores = ns.Stores;
            int storeCount = stores.Count;

            var sinceLocal = mailSinceUtc.ToLocalTime();
            var calStart = DateTime.Today;
            var calEnd = DateTime.Today.AddDays(_cfg.CalendarDaysAhead + 1);
            int mailboxes = 0;

            for (int i = 1; i <= storeCount; i++)
            {
                dynamic store = stores.Item(i);
                string storeName = SafeString(() => (string)store.DisplayName) ?? $"Mailbox {i}";
                bool any = false;

                var inbox = TryFolder((object)store, olFolderInbox);
                if (inbox != null) { ReadMail(inbox, storeName, "received", "ReceivedTime", sinceLocal, emails); any = true; }

                var sent = TryFolder((object)store, olFolderSentMail);
                if (sent != null) { ReadMail(sent, storeName, "sent", "SentOn", sinceLocal, emails); any = true; }

                var cal = TryFolder((object)store, olFolderCalendar);
                if (cal != null)
                {
                    var ids = ReadCalendar(cal, storeName, calStart, calEnd, events);
                    if (ids != null)
                        syncs.Add(new CalendarSync
                        {
                            SourceCalendar = storeName,
                            WindowStart = Json.Iso(calStart.ToUniversalTime()),
                            WindowEnd = Json.Iso(calEnd.ToUniversalTime()),
                            ExternalEventIds = ids,
                        });
                    any = true;
                }
                if (any) mailboxes++;
            }

            Status = $"Read {mailboxes} mailboxes at {DateTime.Now:t}: {emails.Count} new emails, {events.Count} calendar items";
            return new Result(emails, events, syncs, started);
        }
        catch (COMException ex)
        {
            Status = "Outlook was busy - will retry";
            Log.Error("Outlook read failed", ex);
            return null;
        }
        catch (Exception ex)
        {
            Status = "Couldn't read Outlook - see log";
            Log.Error("Outlook read failed", ex);
            return null;
        }
        finally
        {
            if (app != null) try { Marshal.FinalReleaseComObject(app); } catch { }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    static object? TryFolder(object storeObj, int kind)
    {
        dynamic store = storeObj;
        try { return (object)store.GetDefaultFolder(kind); } catch { return null; } // archives / public folders have no inbox
    }

    void ReadMail(object folderObj, string storeName, string direction, string dateProp, DateTime sinceLocal, List<EmailRecord> into)
    {
        dynamic folder = folderObj;
        try
        {
            dynamic items = folder.Items;
            items.Sort($"[{dateProp}]", true);
            dynamic found = items.Restrict(OutlookFilters.Since(dateProp, sinceLocal));
            string folderPath = SafeString(() => (string)folder.FolderPath) ?? "";
            int n = 0;
            for (dynamic? item = found.GetFirst(); item != null && n < MaxPerFolder; item = found.GetNext())
            {
                try
                {
                    if ((int)item.Class != olMail) continue;
                    EmailRecord? rec = ToEmail((object)item, storeName, folderPath, direction);
                    if (rec != null) { into.Add(rec); n++; }
                }
                catch (Exception ex) { Log.Error("Skipped one email", ex); }
            }
        }
        catch (Exception ex) { Log.Error($"Couldn't read {direction} mail in {storeName}", ex); }
    }

    EmailRecord? ToEmail(object mailObj, string storeName, string folderPath, string direction)
    {
        dynamic m = mailObj;
        DateTime when = direction == "sent" ? (DateTime)m.SentOn : (DateTime)m.ReceivedTime;
        string entryId = (string)m.EntryID;
        string? messageId = SafeString(() => (string)m.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID));
        var (to, cc) = Recipients((object)m);
        string body = SafeString(() => (string)m.Body) ?? "";
        if (body.Length > _cfg.MaxEmailBodyChars) body = body[.._cfg.MaxEmailBodyChars];

        return new EmailRecord
        {
            ExternalId = string.IsNullOrWhiteSpace(messageId) ? entryId : messageId!.Trim(),
            ThreadId = SafeString(() => (string)m.ConversationID),
            Subject = SafeString(() => (string)m.Subject) ?? "(No subject)",
            FromEmail = SenderAddress((object)m) ?? "unknown",
            ToEmail = to,
            CcEmail = cc,
            Body = body,
            ReceivedDate = Json.Iso(when.ToUniversalTime()),
            EmailAccountName = storeName,
            FolderPath = folderPath,
            Direction = direction,
            IsRead = direction == "sent" || !(bool)m.UnRead,
        };
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

    static (string? to, string? cc) Recipients(object itemObj)
    {
        dynamic item = itemObj;
        var to = new List<string>();
        var cc = new List<string>();
        try
        {
            dynamic recips = item.Recipients;
            int count = recips.Count;
            for (int i = 1; i <= count && i <= 50; i++)
            {
                dynamic r = recips.Item(i);
                string? addr = SafeString(() => (string)r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)) ?? SafeString(() => (string)r.Address);
                if (string.IsNullOrWhiteSpace(addr)) continue;
                int type = (int)r.Type; // 1 = To / required, 2 = CC / optional
                (type == 2 ? cc : to).Add(addr!);
            }
        }
        catch { }
        return (to.Count > 0 ? string.Join(", ", to) : null, cc.Count > 0 ? string.Join(", ", cc) : null);
    }

    /// <summary>Returns the IDs of all events seen in the window, or null if the calendar couldn't be read.</summary>
    List<string>? ReadCalendar(object folderObj, string storeName, DateTime startLocal, DateTime endLocal, List<CalendarRecord> into)
    {
        dynamic folder = folderObj;
        var ids = new List<string>();
        try
        {
            dynamic items = folder.Items;
            items.IncludeRecurrences = true; // expands repeating meetings into each occurrence
            items.Sort("[Start]");
            dynamic found = items.Restrict(OutlookFilters.Overlapping(startLocal, endLocal));
            int n = 0;
            for (dynamic? a = found.GetFirst(); a != null && n < 1000; a = found.GetNext())
            {
                try
                {
                    if ((int)a.Class != olAppointment) continue;
                    DateTime start = (DateTime)a.Start, end = (DateTime)a.End;
                    if (start >= endLocal) break; // sorted, so nothing later matters
                    var id = OutlookFilters.OccurrenceId((string)a.EntryID, start);
                    int meetingStatus = SafeInt(() => (int)a.MeetingStatus);
                    int busy = SafeInt(() => (int)a.BusyStatus, 2);
                    string location = SafeString(() => (string)a.Location) ?? "";
                    var (req, opt) = Recipients((object)a);

                    into.Add(new CalendarRecord
                    {
                        ExternalEventId = id,
                        Title = SafeString(() => (string)a.Subject) ?? "(No title)",
                        StartTime = Json.Iso(start.ToUniversalTime()),
                        EndTime = Json.Iso(end.ToUniversalTime()),
                        IsAllDay = (bool)a.AllDayEvent,
                        Location = location,
                        Attendees = string.Join(", ", new[] { req, opt }.Where(s => s != null))
                            .Split(", ", StringSplitOptions.RemoveEmptyEntries).ToList(),
                        SourceCalendar = storeName,
                        Metadata = new()
                        {
                            ["show_as"] = busy switch { 0 => "free", 1 => "tentative", 3 => "out_of_office", 4 => "working_elsewhere", _ => "busy" },
                            ["cancelled"] = meetingStatus is 5 or 7,
                            ["online"] = location.Contains("Teams", StringComparison.OrdinalIgnoreCase)
                                         || location.Contains("zoom", StringComparison.OrdinalIgnoreCase),
                            ["organizer"] = SafeString(() => (string)a.Organizer),
                            ["recurring"] = SafeBool(() => (bool)a.IsRecurring),
                        },
                    });
                    ids.Add(id);
                    n++;
                }
                catch (Exception ex) { Log.Error("Skipped one calendar item", ex); }
            }
            return ids;
        }
        catch (Exception ex)
        {
            Log.Error($"Couldn't read calendar in {storeName}", ex);
            return null;
        }
    }

    static string? SafeString(Func<string> f) { try { return f(); } catch { return null; } }
    static int SafeInt(Func<int> f, int fallback = 0) { try { return f(); } catch { return fallback; } }
    static bool SafeBool(Func<bool> f) { try { return f(); } catch { return false; } }

    /// <summary>Outlook's COM interface must be used from a single-threaded apartment.</summary>
    public static Task<T> RunOnSta<T>(Func<T> work)
    {
        var tcs = new TaskCompletionSource<T>();
        var t = new Thread(() =>
        {
            try { tcs.SetResult(work()); }
            catch (Exception ex) { tcs.SetException(ex); }
        }) { IsBackground = true, Name = "OutlookReader" };
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        return tcs.Task;
    }
}
