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
            string? entry = q["entry"], store = q["store"], msgid = q["msgid"], mailbox = q["mailbox"];

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
            if (item == null && !string.IsNullOrWhiteSpace(msgid)) item = FindByMessageId((object)ns, msgid!, mailbox);

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

    static object? FindByMessageId(object nsObj, string msgid, string? mailbox)
    {
        dynamic ns = nsObj;
        var safe = msgid.Trim().Replace("'", "''");
        var filter = $"@SQL=\"{PR_INTERNET_MESSAGE_ID}\" = '{safe}'";
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
                    try
                    {
                        dynamic folder = store.GetDefaultFolder(folderType);
                        object? found = folder.Items.Find(filter);
                        if (found != null) return found;
                    }
                    catch { /* folder missing in this store */ }
                }
            }
            catch { /* unavailable store */ }
        }
        return null;
    }
}
