using System.Diagnostics;
using System.Web;

namespace RavenAgent;

/// <summary>
/// "Open recording" on a meeting page: ravenagent://recording?path=...  Shows the file in File Explorer.
/// Only files inside Raven's recordings folder are ever shown; nothing is run.
/// </summary>
static class RecordingLink
{
    public static void Open(Uri link)
    {
        try
        {
            var path = HttpUtility.ParseQueryString(link.Query)["path"] ?? "";
            var root = Path.GetFullPath(AgentConfig.Load().RecordingsFolder).TrimEnd('\\') + "\\";
            var full = path.Length > 0 ? Path.GetFullPath(path) : "";
            if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Raven only opens recordings from your Raven Recordings folder.", "Raven", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (File.Exists(full)) Process.Start("explorer.exe", $"/select,\"{full}\"");
            else if (Directory.Exists(Path.GetDirectoryName(full))) Process.Start("explorer.exe", $"\"{Path.GetDirectoryName(full)}\"");
            else MessageBox.Show("That recording isn't on this computer (it may have been moved or deleted).", "Raven", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            Log.Error("Open recording", ex);
            MessageBox.Show("Raven couldn't open the recording.", "Raven", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
