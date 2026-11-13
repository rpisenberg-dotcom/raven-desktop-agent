using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace RavenAgent;

/// <summary>
/// The Outlook signature set for a mailbox (Outlook → File → Options → Mail → Signatures), so a reply Raven drafts
/// ends with that mailbox's own signature: its HTML, formatting and pictures. Prefers the account's
/// "Replies/forwards" signature, else its "New messages" one. Pictures are attached inline (as Outlook does).
/// </summary>
static class Signatures
{
    const string AccountsKey = "9375CFF0413111d3B88A00104B2A6676";
    const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
    const string PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";
    static string Folder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Signatures");

    /// <summary>The signature's name (file name without .htm) for this mailbox, or null.</summary>
    public static string? NameFor(string smtp, string accountName)
    {
        foreach (var ver in new[] { "16.0", "15.0" })
        {
            using var profiles = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{ver}\Outlook\Profiles");
            if (profiles == null) continue;
            foreach (var p in profiles.GetSubKeyNames())
            {
                using var accounts = profiles.OpenSubKey($@"{p}\{AccountsKey}");
                if (accounts == null) continue;
                foreach (var a in accounts.GetSubKeyNames())
                {
                    using var k = accounts.OpenSubKey(a);
                    if (k == null) continue;
                    var names = new[] { Text(k.GetValue("Account Name")), Text(k.GetValue("Email")), Text(k.GetValue("SMTP Email Address")) };
                    if (!names.Any(n => n.Length > 0 && (n.Equals(smtp, StringComparison.OrdinalIgnoreCase) || n.Equals(accountName, StringComparison.OrdinalIgnoreCase)))) continue;
                    var sig = Text(k.GetValue("Reply-Forward Signature"));
                    if (sig.Length == 0) sig = Text(k.GetValue("New Signature"));
                    if (sig.Length > 0) return sig;
                }
            }
        }
        // Signatures synced from Microsoft 365 are often named after the mailbox: "Raven (rpick@barebonevodka.com)"
        try
        {
            var hit = Directory.GetFiles(Folder, "*.htm").Select(Path.GetFileNameWithoutExtension)
                .FirstOrDefault(n => n != null && n.Contains(smtp, StringComparison.OrdinalIgnoreCase));
            if (hit != null) return hit;
        }
        catch { }
        return null;
    }

    /// <summary>The signature's HTML (body only) with its pictures attached to the message; null if there's none.</summary>
    public static string? HtmlFor(object mailObj, string smtp, string accountName)
    {
        var name = NameFor(smtp, accountName);
        if (name == null) return null;
        var file = Path.Combine(Folder, name + ".htm");
        if (!File.Exists(file)) { Log.Info($"Signature '{name}' not found in {Folder}"); return null; }
        var html = ReadHtml(file);
        var m = Regex.Match(html, @"<body[^>]*>(.*)</body>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        var body = m.Success ? m.Groups[1].Value : html;
        // Pictures sit beside the signature ("Name_files/image001.png"): attach each inline and point at it
        dynamic mail = mailObj;
        var n = 0;
        body = Regex.Replace(body, @"(src\s*=\s*"")([^""]+)("")", match =>
        {
            var src = match.Groups[2].Value;
            if (src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) || src.StartsWith("http", StringComparison.OrdinalIgnoreCase)) return match.Value;
            try
            {
                var rel = Uri.UnescapeDataString(src.Replace("file:///", "")).Replace('/', Path.DirectorySeparatorChar);
                var path = Path.IsPathRooted(rel) ? rel : Path.Combine(Folder, rel);
                if (!File.Exists(path)) return match.Value;
                var cid = $"ravensig{++n}_{Guid.NewGuid():N}".Substring(0, 24);
                dynamic att = mail.Attachments.Add(path, 1, 0, Path.GetFileName(path));
                att.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid);
                try { att.PropertyAccessor.SetProperty(PR_ATTACHMENT_HIDDEN, true); } catch { }
                return $"{match.Groups[1].Value}cid:{cid}{match.Groups[3].Value}";
            }
            catch (Exception ex) { Log.Error("Signature picture", ex); return match.Value; }
        }, RegexOptions.IgnoreCase);
        // VML pictures (Word's "v:imagedata") duplicate the <img>s for old Outlook; the <img> is enough
        body = Regex.Replace(body, @"<!--\[if gte vml 1\]>.*?<!\[endif\]-->", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        body = body.Replace("<![if !vml]>", "").Replace("<![endif]>", "");
        return body;
    }

    /// <summary>Signature files are saved in the PC's old code page (Windows-1252) or with a charset tag.</summary>
    static string ReadHtml(string file)
    {
        var bytes = File.ReadAllBytes(file);
        if (bytes.Length > 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return Encoding.Unicode.GetString(bytes);
        if (bytes.Length > 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        var ascii = Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, 2000));
        if (Regex.IsMatch(ascii, @"charset\s*=\s*[""']?utf-8", RegexOptions.IgnoreCase)) return Encoding.UTF8.GetString(bytes);
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(1252).GetString(bytes);
    }

    static string Text(object? v) => v switch
    {
        string s => s.Trim('\0', ' '),
        byte[] b => Encoding.Unicode.GetString(b).Trim('\0', ' '),
        _ => "",
    };
}
