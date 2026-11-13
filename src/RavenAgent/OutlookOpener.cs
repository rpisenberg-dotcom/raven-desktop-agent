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
            if (new Uri(url).Host.Equals("followup", StringComparison.OrdinalIgnoreCase))
            {
                FollowUp((object)item, q["task"]);
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
        if (draft.Length == 0) MessageBox.Show("Raven couldn't fetch the draft, so the reply is empty. Use Copy in the app and paste it in.", "Raven");
        OpenReply(itemObj, draft, all);
    }

    /// <summary>A status follow-up for a task you're waiting on: Raven drafts it, it opens as Reply All in that thread.</summary>
    static void FollowUp(object itemObj, string? taskId)
    {
        var cfg = AgentConfig.Load();
        if (!cfg.IsConfigured || string.IsNullOrWhiteSpace(taskId)) { MessageBox.Show("Raven isn't set up on this PC yet.", "Raven"); return; }
        string draft = "", error = "";
        using (var wait = new Form { Text = "Raven", Size = new System.Drawing.Size(360, 110), StartPosition = FormStartPosition.CenterScreen, TopMost = true,
            FormBorderStyle = FormBorderStyle.FixedDialog, ControlBox = false, Font = new System.Drawing.Font("Segoe UI", 10f) })
        {
            wait.Controls.Add(new Label { Text = "Drafting your follow-up… (about 20 seconds)", AutoSize = true, Location = new System.Drawing.Point(16, 24) });
            wait.Shown += async (_, _) =>
            {
                try
                {
                    using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(180) };
                    using var req = new HttpRequestMessage(HttpMethod.Post, cfg.IngestUrl)
                    {
                        Content = new StringContent(System.Text.Json.JsonSerializer.Serialize(new { followup_for = taskId, device_id = cfg.DeviceId }),
                            System.Text.Encoding.UTF8, "application/json"),
                    };
                    req.Headers.Add("x-agent-key", cfg.AgentKey);
                    var res = await http.SendAsync(req);
                    using var doc = System.Text.Json.JsonDocument.Parse(await res.Content.ReadAsStringAsync());
                    if (doc.RootElement.TryGetProperty("draft", out var d)) draft = d.GetString() ?? "";
                    if (doc.RootElement.TryGetProperty("error", out var e)) error = e.GetString() ?? "";
                }
                catch (Exception ex) { error = ex.Message; Log.Error("Follow-up draft", ex); }
                wait.Close();
            };
            wait.ShowDialog();
        }
        if (draft.Length == 0)
        {
            MessageBox.Show($"Raven couldn't draft the follow-up{(error.Length > 0 ? $": {error}" : ".")}", "Raven", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        OpenReply(itemObj, draft, all: true);
    }

    /// <summary>Outlook's Reply (All) window for the email, with the draft above your signature and the original.</summary>
    internal static void OpenReply(object itemObj, string draft, bool all, string? forwardTo = null)
    {
        dynamic item = itemObj;
        // Forward keeps the attachments; the people you named are filled in (Outlook checks them against your contacts)
        dynamic reply = forwardTo != null ? item.Forward() : all ? item.ReplyAll() : item.Reply();
        if (!string.IsNullOrWhiteSpace(forwardTo))
        {
            try
            {
                foreach (var who in forwardTo.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                    reply.Recipients.Add(who);
                reply.Recipients.ResolveAll();
            }
            catch (Exception ex) { Log.Error("Forward recipients", ex); }
        }
        // Replies to plain-text emails start as plain text, which loses your formatted signature: make it HTML first,
        // so Outlook puts in the formatted signature (with pictures) when it opens
        try { if ((int)reply.BodyFormat != 2) reply.BodyFormat = 2; } catch { }
        reply.Display();
        if (draft.Length > 0)
        {
            string smtp = "", acct = "";
            try { dynamic a = reply.SendUsingAccount; smtp = (string)a.SmtpAddress ?? ""; acct = (string)a.DisplayName ?? ""; } catch { }
            if (smtp.Length == 0) { try { dynamic store = item.Parent.Store; acct = (string)store.DisplayName ?? ""; smtp = acct; } catch { } }
            string body = (string)reply.HTMLBody;
            // Your formatted signature for this mailbox (pictures attached inline), when one is set
            string? sig = null;
            try { sig = Signatures.HtmlFor((object)reply, smtp, acct); } catch (Exception ex) { Log.Error("Signature", ex); }
            var sigText = sig != null ? Signatures.PlainOf(sig) : Signatures.TextFor(smtp, acct);
            // It starts with your name ("Raven"): don't sign the draft twice
            draft = Signatures.DropRepeatedName(draft, sigText);

            var html = DraftHtml(draft);
            var bodyTag = Regex.Match(body, "<body[^>]*>", RegexOptions.IgnoreCase);
            var start = bodyTag.Success ? bodyTag.Index + bodyTag.Length : 0;
            var (cut, reopen) = OriginalStart(body, start);
            if (sig != null && cut > start)
            {
                // Replace whatever Outlook put above the original (its own, sometimes plain-text, copy of the signature)
                reply.HTMLBody = body[..start] + html + $"<div>{sig}</div><p class=MsoNormal>&nbsp;</p>" + reopen + body[cut..];
                Log.Info($"Reply: draft + signature '{smtp}' added, Outlook's top part replaced");
            }
            else
            {
                reply.HTMLBody = body.Insert(start, html + (sig != null ? $"<div>{sig}</div>" : ""));
                Log.Info($"Reply: draft added; signature {(sig != null ? "added" : $"not found for {smtp} / {acct}")}; original found: {cut > start}");
            }
        }
        BringToFront((object)reply);
        Log.Info($"Reply opened in Outlook (draft {(draft.Length > 0 ? "added" : "missing")})");
    }

    const string Font = "font-family:'Segoe UI',sans-serif;font-size:11.0pt";

    /// <summary>
    /// The draft as Outlook paragraphs: Segoe UI 11, single spacing, 0 pt before/after, a blank line where the draft
    /// has one (easy to edit), and real numbered / bulleted lists for lines starting "1. " / "- ".
    /// </summary>
    internal static string DraftHtml(string draft)
    {
        var sb = new System.Text.StringBuilder($"<div style=\"{Font}\">");
        string? list = null; // "ol" or "ul" while inside a list
        void Close() { if (list != null) { sb.Append($"</{list}>"); list = null; } }
        string P(string inner) => $"<p class=MsoNormal style=\"margin:0in;line-height:normal\"><span style=\"{Font}\">{inner}</span></p>";
        foreach (var raw in draft.Replace("\r", "").TrimEnd().Split('\n'))
        {
            var line = raw.TrimEnd();
            var num = Regex.Match(line, @"^\s*\d{1,2}[.)]\s+(.*)$");
            var bul = Regex.Match(line, @"^\s*[-*•]\s+(.*)$");
            if (num.Success || bul.Success)
            {
                var kind = num.Success ? "ol" : "ul";
                if (list != kind) { Close(); sb.Append($"<{kind} style=\"margin-top:0in;margin-bottom:0in\">"); list = kind; }
                var text = System.Net.WebUtility.HtmlEncode((num.Success ? num : bul).Groups[1].Value);
                sb.Append($"<li class=MsoNormal style=\"margin:0in;line-height:normal\"><span style=\"{Font}\">{text}</span></li>");
                continue;
            }
            Close();
            sb.Append(line.Trim().Length == 0 ? P("&nbsp;") : P(System.Net.WebUtility.HtmlEncode(line)));
        }
        Close();
        return sb.Append("</div>").ToString();
    }

    /// <summary>
    /// Where the quoted original email begins in a reply's HTML (its "From: / Sent:" header block), and the tags to
    /// re-open if that point is in the middle of a paragraph (replies to plain-text emails). (-1, "") if not found.
    /// </summary>
    static (int at, string reopen) OriginalStart(string body, int from)
    {
        var best = -1; var reopen = "";
        foreach (var marker in new[] { "id=\"divRplyFwdMsg\"", "id=divRplyFwdMsg", "border-top:solid #E1E1E1", "border-top:solid #B5C4DF" })
        {
            var i = body.IndexOf(marker, from, StringComparison.OrdinalIgnoreCase);
            if (i < 0) continue;
            var div = body.LastIndexOf("<div", i, StringComparison.OrdinalIgnoreCase);
            if (div >= from && (best < 0 || div < best)) { best = div; reopen = ""; }
        }
        var om = body.IndexOf("-----Original Message-----", from, StringComparison.OrdinalIgnoreCase);
        if (om >= 0 && (best < 0 || om < best))
        {
            // Plain-text reply converted to HTML: the header is inside the same paragraph as Outlook's signature
            var p = body.LastIndexOf("<p", om, StringComparison.OrdinalIgnoreCase);
            var tagEnd = p >= 0 ? body.LastIndexOf('>', om) : -1;
            if (p >= from && tagEnd > p)
            {
                var open = Regex.Match(body[p..(tagEnd + 1)], @"^<p[^>]*>(\s*<span[^>]*>)*");
                reopen = open.Success ? open.Value : "<p>";
            }
            best = om;
        }
        return (best, reopen);
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
