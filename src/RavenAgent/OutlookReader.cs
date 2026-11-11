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

    public sealed record Result(List<EmailRecord> Emails, List<CalendarRecord> Events, List<CalendarSync> Syncs, DateTime ReadAtUtc, List<string> Mailboxes);

    public static bool IsOutlookRunning() => Process.GetProcessesByName("OUTLOOK").Length > 0;

    /// <summary>Must be called on an STA thread (see <see cref="RunOnSta{T}"/>).</summary>
    /// <param name="fullCalendar">Read the whole calendar window (months ahead); otherwise only yesterday to 3 days ahead,
    /// which is all that changes minute to minute. Reading months of meetings every 5 minutes made Outlook lag.</param>
    public Result? Read(DateTime mailSinceUtc, bool fullCalendar = true)
    {
        if (!IsOutlookRunning()) { Status = "Outlook is closed - waiting"; return null; }

        var started = DateTime.UtcNow;
        if (_attendees.Count > 20000) _attendees.Clear();
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
            var sinceLocal = mailSinceUtc.ToLocalTime();
            var calStart = DateTime.Today.AddDays(fullCalendar ? -_cfg.CalendarDaysBack : -1);
            var calEnd = DateTime.Today.AddDays(fullCalendar ? _cfg.CalendarDaysAhead + 1 : 4);
            int mailboxes = 0;
            var skipped = new List<string>();
            var allMailboxes = new List<string>();

            var seenCalendars = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (storeObj, label) in Mailboxes((object)ns))
            {
                // One broken mailbox (e.g. a .pst on a disconnected drive) must not stop the others
                string storeName = label;
                try
                {
                    dynamic store = storeObj;
                    storeName = SafeString(() => (string)store.DisplayName) ?? label;
                    if (!allMailboxes.Contains(storeName)) allMailboxes.Add(storeName);
                    bool any = false;

                    // Mailboxes Raven marked "don't read" (personal accounts): calendar only, no mail
                    var readMail = !_cfg.IgnoredMailboxes.Contains(storeName, StringComparer.OrdinalIgnoreCase);
                    var inbox = readMail ? TryFolder(storeObj, olFolderInbox) : null;
                    if (inbox != null) { ReadMail(inbox, storeName, "received", "ReceivedTime", sinceLocal, emails); any = true; }

                    var sent = readMail ? TryFolder(storeObj, olFolderSentMail) : null;
                    if (sent != null) { ReadMail(sent, storeName, "sent", "SentOn", sinceLocal, emails); any = true; }

                    foreach (var (calObj, calName) in CalendarFolders(storeObj, storeName))
                    {
                        seenCalendars.Add(SafeString(() => (string)((dynamic)calObj).EntryID) ?? calName);
                        if (_cfg.HiddenCalendars.Contains(calName, StringComparer.OrdinalIgnoreCase)) continue; // Hidden in Clerical Guru
                        var ids = ReadCalendar(calObj, calName, calStart, calEnd, events);
                        if (ids != null)
                            syncs.Add(new CalendarSync
                            {
                                SourceCalendar = calName,
                                WindowStart = Json.Iso(calStart.ToUniversalTime()),
                                WindowEnd = Json.Iso(calEnd.ToUniversalTime()),
                                ExternalEventIds = ids,
                            });
                        any = true;
                    }
                    if (any) mailboxes++;
                }
                catch (Exception ex)
                {
                    skipped.Add(storeName);
                    WarnOnce(storeName, ex);
                }
            }

            // Calendars added in Outlook's Calendar view (Shared Calendars, other people's, other accounts'), which
            // don't live in any mailbox's folder list
            foreach (var (calObj, calName) in NavigationCalendars((object)ns, seenCalendars))
            {
                if (_cfg.HiddenCalendars.Contains(calName, StringComparer.OrdinalIgnoreCase)) continue; // Hidden in Clerical Guru
                try
                {
                    var ids = ReadCalendar(calObj, calName, calStart, calEnd, events);
                    if (ids != null)
                        syncs.Add(new CalendarSync
                        {
                            SourceCalendar = calName,
                            WindowStart = Json.Iso(calStart.ToUniversalTime()),
                            WindowEnd = Json.Iso(calEnd.ToUniversalTime()),
                            ExternalEventIds = ids,
                        });
                }
                catch (Exception ex) { WarnOnce(calName, ex); }
            }

            Status = $"Read {mailboxes} mailboxes at {DateTime.Now:t}: {emails.Count} new emails, {events.Count} calendar items"
                     + (skipped.Count > 0 ? $" ({skipped.Count} unavailable skipped)" : "");
            // Nearest dates first: if an upload is ever cut short, it's far-off meetings that wait, not tomorrow's
            try
            {
                var now = DateTime.UtcNow;
                double Away(CalendarRecord e) => DateTime.TryParse(e.StartTime, null, System.Globalization.DateTimeStyles.RoundtripKind, out var d) ? Math.Abs((d.ToUniversalTime() - now).TotalHours) : double.MaxValue;
                events = events.OrderBy(Away).ToList();
            }
            catch (Exception ex) { Log.Error("Couldn't sort calendar items", ex); }
            return new Result(emails, events, syncs, started, allMailboxes);
        }
        catch (COMException ex)
        {
            Status = "Outlook was busy - will retry";
            Log.Error("Outlook read failed", ex);
            return null;
        }
        catch (Exception ex)
        {
            Status = $"Couldn't read Outlook: {ex.GetType().Name}: {ex.Message} (at {ex.StackTrace?.Split('\n').FirstOrDefault()?.Trim()})";
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

    readonly HashSet<string> _warned = new();
    readonly Dictionary<string, (string?, string?, string?)> _attendees = new();

    static readonly System.Text.RegularExpressions.Regex JoinRe = new(
        @"https://[^\s<>""']*(teams\.microsoft\.com/l/meetup-join|teams\.live\.com/meet|zoom\.us/(j|w|my)/|meet\.google\.com/|webex\.com/(meet|join|[^\s]*j\.php))[^\s<>""']*",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    static string? JoinLink(string body, string location)
    {
        var text = body.Length > 30000 ? body[..30000] : body;
        var m = JoinRe.Match(location + " " + text);
        return m.Success ? m.Value.TrimEnd('>', ')', '.', ',') : null;
    }
    int _calls;
    void WarnOnce(string storeName, Exception ex)
    {
        if (_warned.Add(storeName)) Log.Error($"Skipping mailbox '{storeName}' (unavailable)", ex);
    }

    /// <summary>Every mailbox/data file in the Outlook profile, skipping ones that can't even be opened.</summary>
    IEnumerable<(object store, string label)> Mailboxes(object nsObj)
    {
        dynamic ns = nsObj;
        var list = new List<(object, string)>();
        int count = 0;
        dynamic? stores = null;
        try { stores = ns.Stores; count = (int)stores.Count; }
        catch (Exception ex) { Log.Error("Couldn't list Outlook stores; trying top-level folders", ex); }

        if (stores != null)
        {
            for (int i = 1; i <= count; i++)
            {
                try { list.Add(((object)stores.Item(i), $"Mailbox {i}")); }
                catch (Exception ex) { WarnOnce($"Mailbox {i}", ex); }
            }
            return list;
        }

        // Fallback: top-level folders, one per mailbox
        try
        {
            dynamic folders = ns.Folders;
            int n = (int)folders.Count;
            for (int i = 1; i <= n; i++)
            {
                try { dynamic f = folders.Item(i); list.Add(((object)f.Store, (string)f.Name)); }
                catch (Exception ex) { WarnOnce($"Mailbox {i}", ex); }
            }
        }
        catch (Exception ex) { Log.Error("Couldn't list Outlook mailboxes", ex); }
        return list;
    }

    const int olAppointmentItem = 1; // a folder whose default item type is appointments = a calendar
    const int olModuleCalendar = 1;

    /// <summary>
    /// Calendars listed in Outlook's Calendar view that aren't in a mailbox's folders: "Shared Calendars",
    /// "People's calendars", calendars from other accounts. Named "Group / Calendar".
    /// </summary>
    List<(object folder, string name)> NavigationCalendars(object nsObj, HashSet<string> seen)
    {
        var found = new List<(object, string)>();
        try
        {
            dynamic ns = nsObj;
            dynamic app = ns.Application;
            dynamic? exp = null;
            try { exp = app.ActiveExplorer(); } catch { }
            if (exp == null) { try { if ((int)app.Explorers.Count > 0) exp = app.Explorers.Item(1); } catch { } }
            if (exp == null) return found;
            dynamic module = exp.NavigationPane.Modules.GetNavigationModule(olModuleCalendar);
            dynamic groups = module.NavigationGroups;
            int gn = SafeInt(() => (int)groups.Count);
            for (int g = 1; g <= gn && found.Count < 30; g++)
            {
                try
                {
                    dynamic group = groups.Item(g);
                    string groupName = SafeString(() => (string)group.Name) ?? "Calendars";
                    dynamic navFolders = group.NavigationFolders;
                    int fn = SafeInt(() => (int)navFolders.Count);
                    for (int i = 1; i <= fn && found.Count < 30; i++)
                    {
                        try
                        {
                            dynamic nf = navFolders.Item(i);
                            dynamic folder = nf.Folder;
                            if (SafeInt(() => (int)folder.DefaultItemType, -1) != olAppointmentItem) continue; // e.g. mail "Favorites"
                            string id = SafeString(() => (string)folder.EntryID) ?? "";
                            if (id.Length == 0 || !seen.Add(id)) continue; // already read with its mailbox
                            string name = SafeString(() => (string)nf.DisplayName) ?? SafeString(() => (string)folder.Name) ?? "Calendar";
                            found.Add(((object)folder, $"{groupName} / {name}"));
                        }
                        catch (Exception ex) { WarnOnce($"Shared calendar {g}.{i}", ex); }
                    }
                }
                catch (Exception ex) { WarnOnce($"Calendar group {g}", ex); }
            }
        }
        catch (Exception ex) { WarnOnce("Calendar view", ex); }
        return found;
    }

    /// <summary>
    /// Every calendar in a mailbox: the main one (named after the mailbox) plus extra calendars,
    /// subscribed internet calendars and shared ones stored as folders (named "Mailbox / Folder").
    /// </summary>
    static List<(object folder, string name)> CalendarFolders(object storeObj, string storeName)
    {
        var found = new List<(object, string)>();
        string? defaultId = null;
        var main = TryFolder(storeObj, olFolderCalendar);
        if (main != null)
        {
            found.Add((main, storeName));
            defaultId = SafeString(() => (string)((dynamic)main).EntryID);
        }
        try
        {
            dynamic store = storeObj;
            dynamic root = store.GetRootFolder();
            var queue = new Queue<(object f, int depth)>();
            queue.Enqueue(((object)root, 0));
            while (queue.Count > 0 && found.Count < 25)
            {
                var (fObj, depth) = queue.Dequeue();
                dynamic f = fObj;
                dynamic subs;
                try { subs = f.Folders; } catch { continue; }
                int n = SafeInt(() => (int)subs.Count);
                for (int i = 1; i <= n && found.Count < 25; i++)
                {
                    try
                    {
                        dynamic sub = subs.Item(i);
                        if (SafeInt(() => (int)sub.DefaultItemType, -1) == olAppointmentItem)
                        {
                            string id = SafeString(() => (string)sub.EntryID) ?? "";
                            if (id != defaultId)
                            {
                                string name = SafeString(() => (string)sub.Name) ?? "Calendar";
                                found.Add(((object)sub, $"{storeName} / {name}"));
                            }
                        }
                        if (depth < 3) queue.Enqueue(((object)sub, depth + 1));
                    }
                    catch { }
                }
            }
        }
        catch { }
        return found;
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
                    if ((int)item!.Class != olMail) continue;
                    if (++_calls % 25 == 0) Thread.Sleep(15);
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
            OutlookEntryId = entryId,
            OutlookStoreId = SafeString(() => (string)m.Parent.StoreID),
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
        var seen = new HashSet<string>();
        int n = 0;
        string? storeId = SafeString(() => (string)folder.StoreID);

        void Add(dynamic a, DateTime start, DateTime end)
        {
            var id = OutlookFilters.OccurrenceId((string)a.EntryID, start);
            if (!seen.Add(id)) return;
            int meetingStatus = SafeInt(() => (int)a.MeetingStatus);
            int busy = SafeInt(() => (int)a.BusyStatus, 2);
            string location = SafeString(() => (string)a.Location) ?? "";
            // Looking up every attendee's address is the slowest part: only redo it when the meeting changed
            var modified = SafeString(() => ((DateTime)a.LastModificationTime).ToString("o")) ?? "";
            var key = $"{id}|{modified}";
            if (!_attendees.TryGetValue(key, out var people))
            {
                var (r1, r2) = Recipients((object)a);
                // The Teams / Zoom / Meet link from the invite, for the "Join" button in the app
                people = (r1, r2, JoinLink(SafeString(() => (string)a.Body) ?? "", location));
                _attendees[key] = people;
            }
            var (req, opt, join) = people;
            if (++_calls % 25 == 0) Thread.Sleep(15); // let Outlook handle clicks and typing between reads

            into.Add(new CalendarRecord
            {
                ExternalEventId = id,
                Title = SafeString(() => (string)a.Subject) ?? "(No title)",
                StartTime = Json.Iso(start.ToUniversalTime()),
                EndTime = Json.Iso(end.ToUniversalTime()),
                IsAllDay = SafeBool(() => (bool)a.AllDayEvent),
                Location = location,
                Attendees = string.Join(", ", new[] { req, opt }.Where(x => x != null))
                    .Split(", ", StringSplitOptions.RemoveEmptyEntries).ToList(),
                SourceCalendar = storeName,
                Metadata = new()
                {
                    ["show_as"] = busy switch { 0 => "free", 1 => "tentative", 3 => "out_of_office", 4 => "working_elsewhere", _ => "busy" },
                    ["cancelled"] = meetingStatus is 5 or 7,
                    ["online"] = location.Contains("Teams", StringComparison.OrdinalIgnoreCase)
                                 || location.Contains("zoom", StringComparison.OrdinalIgnoreCase),
                    ["organizer"] = SafeString(() => (string)a.Organizer),
                    // Your own appointment or a meeting you sent (not one you were invited to): the app keeps this copy
                    // when the same meeting is on two of your calendars
                    ["is_organizer"] = meetingStatus is 0 or 1 or 5,
                    ["recurring"] = SafeBool(() => (bool)a.IsRecurring),
                    ["join_url"] = join,
                    ["entry_id"] = SafeString(() => (string)a.EntryID),
                    ["store_id"] = storeId,
                },
            });
            ids.Add(id);
            n++;
        }

        try
        {
            dynamic items = folder.Items;
            // Order matters to Outlook: sort by start FIRST, then expand repeating meetings, then filter.
            // The other way round it can return only a series' first meeting and skip later dates.
            items.Sort("[Start]");
            items.IncludeRecurrences = true;
            dynamic found = items.Restrict(OutlookFilters.Overlapping(startLocal, endLocal));
            for (dynamic? a = found.GetFirst(); a != null && n < 3000; a = found.GetNext())
            {
                try
                {
                    if ((int)a!.Class != olAppointment) continue;
                    DateTime start = (DateTime)a.Start, end = (DateTime)a.End;
                    if (start >= endLocal) break; // sorted, so nothing later matters
                    if (end <= startLocal) continue;
                    Add(a, start, end);
                }
                catch (Exception ex) { Log.Error("Skipped one calendar item", ex); }
            }
        }
        catch (Exception ex)
        {
            Log.Error($"Couldn't read calendar in {storeName}", ex);
            return null;
        }

        // Backup for repeating meetings (some shared and online calendars don't expand them): ask each series
        // directly for its meeting on each day in the window.
        // Only for the next two weeks, to keep it light on Outlook (the main read covers the rest).
        var bStart = startLocal > DateTime.Today.AddDays(-1) ? startLocal : DateTime.Today.AddDays(-1);
        var bEnd = endLocal < DateTime.Today.AddDays(15) ? endLocal : DateTime.Today.AddDays(15);
        if (bEnd > bStart)
            try { AddSeriesOccurrences(folderObj, bStart, bEnd, new Action<dynamic, DateTime, DateTime>(Add)); }
            catch (Exception ex) { WarnOnce($"{storeName} (repeating meetings)", ex); }
        return ids;
    }

    /// <summary>
    /// Support check: every calendar Raven can see and every meeting in the next two weeks whose title contains
    /// <paramref name="word"/>, read exactly the way the regular sync reads them. Returns a plain-text report.
    /// </summary>
    public string CheckCalendars(string word)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Raven calendar check - {DateTime.Now:f} - looking for \"{word}\" from yesterday to 14 days ahead");
        sb.AppendLine($"Agent version {typeof(OutlookReader).Assembly.GetName().Version}");
        sb.AppendLine();
        try
        {
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null) return sb.AppendLine("Classic Outlook isn't installed.").ToString();
            dynamic app = Activator.CreateInstance(type)!;
            dynamic ns = app.GetNamespace("MAPI");
            var from = DateTime.Today.AddDays(-1);
            var to = DateTime.Today.AddDays(15);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var calendars = new List<(object, string, string)>();
            foreach (var (storeObj, label) in Mailboxes((object)ns))
            {
                string storeName = label;
                try
                {
                    storeName = SafeString(() => (string)((dynamic)storeObj).DisplayName) ?? label;
                    foreach (var (calObj, calName) in CalendarFolders(storeObj, storeName))
                    {
                        seen.Add(SafeString(() => (string)((dynamic)calObj).EntryID) ?? calName);
                        calendars.Add((calObj, calName, "mailbox"));
                    }
                }
                catch (Exception ex) { sb.AppendLine($"!! Mailbox {storeName}: {ex.Message}"); }
            }
            foreach (var (calObj, calName) in NavigationCalendars((object)ns, seen)) calendars.Add((calObj, calName, "calendar view"));

            sb.AppendLine($"Calendars Raven reads ({calendars.Count}):");
            foreach (var (calObj, calName, where) in calendars)
            {
                var list = new List<CalendarRecord>();
                var ids = ReadCalendar(calObj, calName, from, to, list);
                var hits = list.Where(e => e.Title.Contains(word, StringComparison.OrdinalIgnoreCase)).ToList();
                sb.AppendLine($"- {calName}  [{where}]  {(ids == null ? "COULDN'T READ" : $"{list.Count} meetings")}{(hits.Count > 0 ? $", {hits.Count} matching" : "")}");
                foreach (var h in hits.OrderBy(h => h.StartTime))
                    sb.AppendLine($"      {DateTime.Parse(h.StartTime).ToLocalTime():ddd MMM d h:mm tt}  {h.Title}{(h.Metadata?.GetValueOrDefault("recurring") is true ? "  (repeating)" : "")}");
                // Repeating series in this calendar whose title matches, and what Outlook says about them
                try
                {
                    dynamic masters = ((dynamic)calObj).Items.Restrict("[IsRecurring] = True");
                    for (dynamic? m = masters.GetFirst(); m != null; m = masters.GetNext())
                    {
                        string subj = SafeString(() => (string)m.Subject) ?? "";
                        if (!subj.Contains(word, StringComparison.OrdinalIgnoreCase)) continue;
                        dynamic p = m.GetRecurrencePattern();
                        sb.AppendLine($"      series: {subj} | state {SafeInt(() => (int)m.RecurrenceState)} | type {SafeInt(() => (int)p.RecurrenceType)} | days {SafeInt(() => (int)p.DayOfWeekMask)} | {SafeString(() => ((DateTime)p.PatternStartDate).ToShortDateString())} to {SafeString(() => ((DateTime)p.PatternEndDate).ToShortDateString())} at {SafeString(() => ((DateTime)p.StartTime).ToShortTimeString())}");
                    }
                }
                catch (Exception ex) { sb.AppendLine($"      (couldn't list repeating series: {ex.Message})"); }
            }
        }
        catch (Exception ex) { sb.AppendLine($"!! {ex.GetType().Name}: {ex.Message}"); }
        sb.AppendLine();
        sb.AppendLine($"Last sync: {Status}");
        return sb.ToString();
    }

    /// <summary>Every calendar Raven reads, by the name the app knows it by.</summary>
    Dictionary<string, object> CalendarsByName(object nsObj)
    {
        var map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (storeObj, label) in Mailboxes(nsObj))
        {
            try
            {
                string storeName = SafeString(() => (string)((dynamic)storeObj).DisplayName) ?? label;
                foreach (var (calObj, calName) in CalendarFolders(storeObj, storeName))
                {
                    seen.Add(SafeString(() => (string)((dynamic)calObj).EntryID) ?? calName);
                    map.TryAdd(calName, calObj);
                }
            }
            catch { }
        }
        foreach (var (calObj, calName) in NavigationCalendars(nsObj, seen)) map.TryAdd(calName, calObj);
        return map;
    }

    const int olAppointmentItemType = 1, olMeeting = 1, olMeetingReceived = 3;

    /// <summary>
    /// Makes calendar changes from the app in Outlook. Never sends invites, updates or cancellations itself: for a
    /// meeting with other people the event is opened in Outlook for Raven to check and click Send.
    /// </summary>
    public List<CalendarChangeResult> ApplyChanges(List<CalendarChange> changes)
    {
        var results = new List<CalendarChangeResult>();
        var type = Type.GetTypeFromProgID("Outlook.Application");
        if (type == null) return changes.Select(c => new CalendarChangeResult(c.Id, "failed", "Classic Outlook isn't installed")).ToList();
        dynamic app = Activator.CreateInstance(type)!;
        dynamic ns = app.GetNamespace("MAPI");
        Dictionary<string, object>? calendars = null;
        DateTime Local(string iso) => DateTime.Parse(iso, null, System.Globalization.DateTimeStyles.RoundtripKind).ToLocalTime();

        dynamic Target(CalendarChange c)
        {
            if (string.IsNullOrEmpty(c.EntryId)) throw new InvalidOperationException("this event hasn't synced from Outlook yet - try again in a few minutes");
            dynamic item = string.IsNullOrEmpty(c.StoreId) ? ns.GetItemFromID(c.EntryId) : ns.GetItemFromID(c.EntryId, c.StoreId);
            if (c.Scope != "series" && !string.IsNullOrEmpty(c.OccurrenceStart) && SafeBool(() => (bool)item.IsRecurring))
                item = item.GetRecurrencePattern().GetOccurrence(Local(c.OccurrenceStart));
            return item;
        }
        void Fill(dynamic item, CalendarChange c)
        {
            if (c.Title != null) item.Subject = c.Title;
            if (c.IsAllDay != null) item.AllDayEvent = c.IsAllDay.Value;
            if (!string.IsNullOrEmpty(c.StartTime)) item.Start = Local(c.StartTime);
            if (!string.IsNullOrEmpty(c.EndTime)) item.End = Local(c.EndTime);
            if (c.Location != null) item.Location = c.Location;
            if (c.Notes != null) item.Body = c.Notes;
        }
        int Others(dynamic item) => Math.Max(0, SafeInt(() => (int)item.Recipients.Count) - 1);

        foreach (var c in changes)
        {
            try
            {
                switch (c.Action)
                {
                    case "open":
                    {
                        dynamic item = Target(c);
                        item.Display();
                        OutlookOpener.BringToFront((object)item);
                        results.Add(new(c.Id, "opened", "Opened in Outlook"));
                        break;
                    }
                    case "create":
                    {
                        calendars ??= CalendarsByName((object)ns);
                        if (!calendars.TryGetValue(c.TargetCalendar, out var folderObj))
                            throw new InvalidOperationException($"Raven couldn't find the calendar \"{c.TargetCalendar}\" in Outlook");
                        dynamic item = ((dynamic)folderObj).Items.Add(olAppointmentItemType);
                        Fill(item, c);
                        var people = c.Attendees.Where(a => !string.IsNullOrWhiteSpace(a)).ToList();
                        if (people.Count > 0)
                        {
                            item.MeetingStatus = olMeeting;
                            foreach (var a in people) item.Recipients.Add(a.Trim());
                            try { item.Recipients.ResolveAll(); } catch { }
                        }
                        item.Save();
                        string? entry = SafeString(() => (string)item.EntryID);
                        if (people.Count > 0 || c.Teams)
                        {
                            item.Display();
                            OutlookOpener.BringToFront((object)item);
                            results.Add(new(c.Id, "opened", people.Count > 0
                                ? (c.Teams ? "Opened in Outlook: click Teams Meeting to add the link, then Send" : "Opened in Outlook: check it and click Send")
                                : "Saved and opened in Outlook: click Teams Meeting to add the link, then Save", entry));
                        }
                        else results.Add(new(c.Id, "done", "Added to Outlook", entry));
                        break;
                    }
                    case "update":
                    {
                        dynamic item = Target(c);
                        Fill(item, c);
                        int status = SafeInt(() => (int)item.MeetingStatus);
                        if (status == olMeeting && Others(item) > 0)
                        {
                            item.Display(); // changes are in the window, not saved or sent: Raven clicks Send Update
                            OutlookOpener.BringToFront((object)item);
                            results.Add(new(c.Id, "opened", "Opened in Outlook with your changes: click Send Update so attendees get them"));
                        }
                        else
                        {
                            item.Save();
                            results.Add(new(c.Id, "done", status == olMeetingReceived ? "Changed in your calendar only (you're not the organizer)" : "Changed in Outlook"));
                        }
                        if (c.Teams) { item.Display(); OutlookOpener.BringToFront((object)item); }
                        break;
                    }
                    case "delete":
                    {
                        dynamic item = Target(c);
                        int status = SafeInt(() => (int)item.MeetingStatus);
                        if ((status == olMeeting && Others(item) > 0) || status == olMeetingReceived)
                        {
                            item.Display();
                            OutlookOpener.BringToFront((object)item);
                            results.Add(new(c.Id, "opened", status == olMeeting
                                ? "Opened in Outlook: click Cancel Meeting to send the cancellation"
                                : "Opened in Outlook: click Decline (or Delete) to remove it and tell the organizer"));
                        }
                        else
                        {
                            item.Delete();
                            results.Add(new(c.Id, "done", "Deleted from Outlook"));
                        }
                        break;
                    }
                    default:
                        results.Add(new(c.Id, "failed", $"Unknown change: {c.Action}"));
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Calendar change {c.Action} failed", ex);
                results.Add(new(c.Id, "failed", ex is InvalidOperationException ? ex.Message : $"Outlook said: {ex.Message}"));
            }
        }
        return results;
    }

    const int olApptMaster = 1;

    void AddSeriesOccurrences(object folderObj, DateTime startLocal, DateTime endLocal, Action<dynamic, DateTime, DateTime> add)
    {
        dynamic folder = folderObj;
        dynamic all = folder.Items;
        dynamic masters = all.Restrict("[IsRecurring] = True");
        int checkedSeries = 0;
        for (dynamic? m = masters.GetFirst(); m != null && checkedSeries < 300; m = masters.GetNext())
        {
            try
            {
                if ((int)m!.Class != olAppointment || SafeInt(() => (int)m.RecurrenceState) != olApptMaster) continue;
                checkedSeries++;
                dynamic pattern = m.GetRecurrencePattern();
                DateTime from = ((DateTime)pattern.PatternStartDate).Date, until = ((DateTime)pattern.PatternEndDate).Date;
                if (until < startLocal.Date || from > endLocal.Date) continue;
                TimeSpan at = ((DateTime)pattern.StartTime).TimeOfDay;
                int type = SafeInt(() => (int)pattern.RecurrenceType); // 0 daily, 1 weekly, 2 monthly, 3 monthly-nth, 5 yearly, 6 yearly-nth
                int mask = SafeInt(() => (int)pattern.DayOfWeekMask);   // Sunday 1, Monday 2, ... Saturday 64
                int dayOfMonth = SafeInt(() => (int)pattern.DayOfMonth);
                for (var day = (startLocal.Date > from ? startLocal.Date : from); day < endLocal && day <= until; day = day.AddDays(1))
                {
                    bool candidate = type switch
                    {
                        0 => true,
                        1 => mask == 0 || (mask & (1 << (int)day.DayOfWeek)) != 0,
                        2 or 5 => dayOfMonth == 0 || day.Day == dayOfMonth,
                        3 or 6 => mask == 0 || (mask & (1 << (int)day.DayOfWeek)) != 0,
                        _ => false,
                    };
                    if (!candidate) continue;
                    try
                    {
                        dynamic occ = pattern.GetOccurrence(day + at); // throws when the series doesn't meet that day
                        DateTime s = (DateTime)occ.Start, e = (DateTime)occ.End;
                        if (e > startLocal && s < endLocal) add(occ, s, e);
                    }
                    catch { /* no meeting that day (or it was moved: the main read has those) */ }
                }
            }
            catch (Exception ex) { Log.Error("Skipped one repeating meeting", ex); }
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
