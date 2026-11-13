using System.Text.RegularExpressions;
using System.Web;

namespace RavenAgent;

/// <summary>
/// Opens an email in classic Outlook for an "Open in Outlook" link from the Raven app:
///   ravenagent://open-email?entry=..&store=..&msgid=..&mailbox=..
/// Uses Outlook's own ids when the app has them (emails read by agent 0.2.6+), otherwise finds
/// the email by its internet message id in that mailbox's Inbox and Sent Items.
/// Only ever displays an email; it never changes or sends anything.
///   ravenagent://reply?draft=<email id>&all=1&entry=..  opens Outlook's Reply (All) window with Raven's draft typed in
///   (fetched from Clerical Guru); Raven reviews it and presses Send herself.
/// </summary>
static class OutlookOpener
{
    const int olFolderSentMail = 5, olFolderInbox = 6;
    const string PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
    static readonly Regex Hex = new("^[0-9A-Fa-f]{16,2000}$");

    public static void Open(string url)
    {
        try
        {
            var q = HttpUtility.ParseQueryString(new Uri(url).Query);
            string? entry = q["entry"], store = q["store"], msgid = q["msgid"], mailbox = q["mailbox"], subject = q["subject"];
            DateTime? when = DateTime.TryParse(q["when"], null, System.Globalization.DateTimeStyles.AdjustToUniversal, out var w) ? w : null;

            if (!OutlookReader.IsOutlookRunning())
            {
                MessageBox.Show("Open Outlook first, then click the link again.", "Raven", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null) { MessageBox.Show("Classic Outlook isn't installed on this PC.", "Raven"); return; }
            dynamic app = Activator.CreateInstance(type)!;
            dynamic ns = app.GetNamespace("MAPI");

            dynamic? item = null;
            if (entry != null && Hex.IsMatch(entry))
            {
                try { item = store != null && Hex.IsMatch(store) ? ns.GetItemFromID(entry, store) : ns.GetItemFromID(entry); }
                catch { item = null; }
            }
            // A calendar event: one meeting of a repeating series opens that meeting, not the whole series
            string? occ = q["occ"];
            if (item != null && !string.IsNullOrEmpty(occ))
            {
                try
                {
                    if ((bool)item.IsRecurring)
                        item = item.GetRecurrencePattern().GetOccurrence(DateTime.Parse(occ, null, System.Globalization.DateTimeStyles.RoundtripKind).ToLocalTime());
                }
                catch { /* open the series instead */ }
            }
            if (item == null && q["kind"] == "event")
            {
                MessageBox.Show("Raven couldn't find this event in Outlook. It may have been moved or deleted.", "Raven", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (item == null) item = Search((object)ns, msgid, subject, when, mailbox);
            if (item == null && !string.IsNullOrEmpty(mailbox)) item = Search((object)ns, msgid, subject, when, null); // any mailbox
            Log.Info(item == null ? $"Open in Outlook: not found (mailbox {mailbox}, subject {subject})" : "Open in Outlook: found");

            if (item == null)
            {
                MessageBox.Show("Raven couldn't find this email in Outlook. It may have been moved to another folder or deleted.", "Raven",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (new Uri(url).Host.Equals("reply", StringComparison.OrdinalIgnoreCase))
            {
                Reply((object)item, q["draft"], q["all"] == "1");
                return;
            }
            item.Display();
            BringToFront((object)item);
        }
        catch (Exception ex)
        {
            Log.Error("Open in Outlook failed", ex);
            MessageBox.Show("Raven couldn't open this email in Outlook.", "Raven", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    /// <summary>Opens the reply window with the draft from Clerical Guru above the quoted email (and your signature).</summary>
    static void Reply(object itemObj, string? draftId, bool all)
    {
        dynamic item = itemObj;
        string draft = "";
        if (!string.IsNullOrWhiteSpace(draftId))
        {
            var cfg = AgentConfig.Load();
            if (cfg.IsConfigured)
            {
                try
                {
                    using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
                    using var req = new HttpRequestMessage(HttpMethod.Post, cfg.IngestUrl)
                    {
                        Content = new StringContent(System.Text.Json.JsonSerializer.Serialize(new { get_draft = draftId, device_id = cfg.DeviceId }),
                            System.Text.Encoding.UTF8, "application/json"),
                    };
                    req.Headers.Add("x-agent-key", cfg.AgentKey);
                    var res = http.Send(req);
                    using var doc = System.Text.Json.JsonDocument.Parse(res.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                    if (doc.RootElement.TryGetProperty("draft", out var d)) draft = d.GetString() ?? "";
                }
                catch (Exception ex) { Log.Error("Couldn't fetch the draft", ex); }
            }
        }
        dynamic reply = all ? item.ReplyAll() : item.Reply();
        reply.Display(); // shown first, so Outlook adds your reply signature
        if (draft.Length > 0)
        {
            var html = string.Join("", draft.Replace("\r", "").Split("\n\n").Select(p =>
                $"<p style=\"margin:0 0 12px 0\">{System.Net.WebUtility.HtmlEncode(p).Replace("\n", "<br>")}</p>"));
            html = $"<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt\">{html}</div>";
            string body = (string)reply.HTMLBody;
            var m = Regex.Match(body, "<body[^>]*>", RegexOptions.IgnoreCase);
            reply.HTMLBody = m.Success ? body.Insert(m.Index + m.Length, html) : html + body;
        }
        else MessageBox.Show("Raven couldn't fetch the draft, so the reply is empty. Use Copy in the app and paste it in.", "Raven");
        BringToFront((object)reply);
        Log.Info($"Reply opened in Outlook (draft {(draft.Length > 0 ? "added" : "missing")})");
    }

    /// <summary>
    /// Looks through Inbox, Sent Items and their subfolders (where emails usually get filed), first by the
    /// email's internet message id, then by subject and time.
    /// </summary>
    static object? Search(object nsObj, string? msgid, string? subject, DateTime? whenUtc, string? mailbox)
    {
        dynamic ns = nsObj;
        string? byId = string.IsNullOrWhiteSpace(msgid) ? null
            : $"@SQL=\"{PR_INTERNET_MESSAGE_ID}\" = '{msgid!.Trim().Replace("'", "''")}'";
        dynamic stores = ns.Stores;
        int count = (int)stores.Count;
        for (int i = 1; i <= count; i++)
        {
            try
            {
                dynamic store = stores.Item(i);
                string name = (string)store.DisplayName;
                if (!string.IsNullOrEmpty(mailbox) && !string.Equals(name, mailbox, StringComparison.OrdinalIgnoreCase)) continue;
                foreach (var folderType in new[] { olFolderInbox, olFolderSentMail })
                {
                    dynamic? top = null;
                    try { top = store.GetDefaultFolder(folderType); } catch { continue; }
                    var found = SearchFolder((object)top!, byId, subject, whenUtc, folderType == olFolderSentMail, 0);
                    if (found != null) return found;
                }
            }
            catch { /* unavailable store */ }
        }
        return null;
    }

    static object? SearchFolder(object folderObj, string? byId, string? subject, DateTime? whenUtc, bool sent, int depth)
    {
        dynamic folder = folderObj;
        try
        {
            dynamic items = folder.Items;
            if (byId != null)
            {
                object? hit = items.Find(byId);
                if (hit != null) return hit;
            }
            if (!string.IsNullOrWhiteSpace(subject) && whenUtc is { } w)
            {
                var local = w.ToLocalTime();
                string field = sent ? "[SentOn]" : "[ReceivedTime]";
                string fmt(DateTime d) => d.ToString("g", System.Globalization.CultureInfo.CurrentCulture);
                dynamic matches = items.Restrict($"{field} >= '{fmt(local.AddMinutes(-3))}' AND {field} <= '{fmt(local.AddMinutes(3))}'");
                int n = (int)matches.Count;
                for (int i = 1; i <= n && i <= 50; i++)
                {
                    dynamic m = matches.Item(i);
                    string s = "";
                    try { s = (string)m.Subject ?? ""; } catch { }
                    if (string.Equals(s.Trim(), subject!.Trim(), StringComparison.OrdinalIgnoreCase)) return (object)m;
                }
            }
        }
        catch { /* this folder can't be searched */ }
        if (depth >= 3) return null;
        try
        {
            dynamic subs = folder.Folders;
            int c = (int)subs.Count;
            for (int i = 1; i <= c; i++)
            {
                var found = SearchFolder((object)subs.Item(i), byId, subject, whenUtc, sent, depth + 1);
                if (found != null) return found;
            }
        }
        catch { /* no subfolders */ }
        return null;
    }

    // Windows doesn't let a program in the background jump in front of the window being used, so the opened
    // email used to appear behind Chrome. Activate Outlook's window for it and bring it forward.
    [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    static extern IntPtr FindWindow(string? cls, string? title);
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h);
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool AllowSetForegroundWindow(int pid);

    internal static void BringToFront(object itemObj)
    {
        try
        {
            dynamic item = itemObj;
            dynamic insp = item.GetInspector;
            try { if ((int)insp.WindowState == 1) insp.WindowState = 2; } catch { } // minimized -> normal
            insp.Activate();
            string caption = (string)insp.Caption;
            AllowSetForegroundWindow(-1);
            for (int i = 0; i < 10; i++)
            {
                var h = FindWindow("rctrl_renwnd32", caption);
                if (h != IntPtr.Zero)
                {
                    // A tap of the Alt key lets this program hand the focus to Outlook's window
                    keybd_event(0x12, 0, 0, UIntPtr.Zero);
                    keybd_event(0x12, 0, 2, UIntPtr.Zero);
                    ShowWindow(h, 9); // restore
                    BringWindowToTop(h);
                    SetForegroundWindow(h);
                    return;
                }
                Thread.Sleep(150);
            }
        }
        catch (Exception ex) { Log.Error("Couldn't bring the email to the front", ex); }
    }
}
