namespace RavenAgent;

/// <summary>Plain text log in %LOCALAPPDATA%\RavenAgent\logs, one file per day, kept 7 days.</summary>
static class Log
{
    static readonly object Lock = new();
    public static string Folder => Path.Combine(AgentConfig.DataFolder, "logs");

    public static void Info(string msg) => Write("INFO", msg);
    public static void Error(string msg, Exception? ex = null) => Write("ERROR", ex == null ? msg : $"{msg}: {ex.GetType().Name}: {ex.Message}");

    static void Write(string level, string msg)
    {
        try
        {
            lock (Lock)
            {
                Directory.CreateDirectory(Folder);
                File.AppendAllText(Path.Combine(Folder, $"agent-{DateTime.Now:yyyyMMdd}.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {level} {msg}{Environment.NewLine}");
            }
        }
        catch { /* logging must never crash the agent */ }
    }

    public static void Cleanup()
    {
        try
        {
            foreach (var f in Directory.GetFiles(Folder, "agent-*.log"))
                if (File.GetLastWriteTime(f) < DateTime.Now.AddDays(-7)) File.Delete(f);
        }
        catch { }
    }
}
