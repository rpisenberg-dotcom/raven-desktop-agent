using Microsoft.Win32;

namespace RavenAgent;

/// <summary>"Start with Windows" via the current user's Run key (no admin rights needed).</summary>
static class Startup
{
    const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    const string Name = "RavenAgent";

    public static bool IsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKey);
        return key?.GetValue(Name) is string;
    }

    public static void Set(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKey);
        if (enabled) key.SetValue(Name, $"\"{Environment.ProcessPath}\"");
        else key.DeleteValue(Name, false);
    }

    /// <summary>Lets "Open in Outlook" links in the Raven app hand the email to this agent.</summary>
    public static void RegisterLinkHandler()
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(@"Software\Classes\ravenagent");
            key.SetValue("", "URL:Raven Agent");
            key.SetValue("URL Protocol", "");
            using var cmd = key.CreateSubKey(@"shell\open\command");
            cmd.SetValue("", $"\"{Environment.ProcessPath}\" \"%1\"");
        }
        catch (Exception ex) { Log.Error("Couldn't register Open in Outlook links", ex); }
    }
}
