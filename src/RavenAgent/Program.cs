namespace RavenAgent;

static class Program
{
    [STAThread]
    static void Main()
    {
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
