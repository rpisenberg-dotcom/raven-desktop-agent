using System.Text.RegularExpressions;
using System.Web;

namespace RavenAgent;

/// <summary>
/// Opens an email in classic Outlook for an "Open in Outlook" link from the Raven app:
///   ravenagent://open-email?entry=..&store=..&msgid=..&mailbox=..
/// Uses Outlook's own ids when the app has them (emails read by agent 0.2.6+), otherwise finds
/// the email by its internet message id in that mailbox's Inbox and Sent Items.
/// Only ever displays an email; it never changes or sends anything.
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
            if (item == null) item = Search((object)ns, msgid, subject, when, mailbox);
            if (item == null && !string.IsNullOrEmpty(mailbox)) item = Search((object)ns, msgid, subject, when, null); // any mailbox
            Log.Info(item == null ? $"Open in Outlook: not found (mailbox {mailbox}, subject {subject})" : "Open in Outlook: found");

            if (item == null)
            {
                MessageBox.Show("Raven couldn't find this email in Outlook. It may have been moved to another folder or deleted.", "Raven",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            item.Display();
        }
        catch (Exception ex)
        {
            Log.Error("Open in Outlook failed", ex);
            MessageBox.Show("Raven couldn't open this email in Outlook.", "Raven", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
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
}
