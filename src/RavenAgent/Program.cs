namespace RavenAgent;

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        // "Open in Outlook" links from the Raven app: ravenagent://open-email?...
        if (args.Length > 0 && args[0].StartsWith("ravenagent:", StringComparison.OrdinalIgnoreCase))
        {
            Application.EnableVisualStyles();
            if (Uri.TryCreate(args[0], UriKind.Absolute, out var link) && link.Host.Equals("recording", StringComparison.OrdinalIgnoreCase))
                RecordingLink.Open(link);
            else
                OutlookOpener.Open(args[0]);
            return;
        }

        // Only one copy per Windows user
        using var mutex = new Mutex(true, @"Local\RavenDesktopAgent", out var first);
        if (!first) return;

        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.ThreadException += (_, e) => Log.Error("Unhandled UI error", e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (_, e) => Log.Error("Unhandled error", e.ExceptionObject as Exception);
        Application.Run(new TrayApp());
    }
}
