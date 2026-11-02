using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace RavenAgent;

/// <summary>Small wrappers around the Windows calls the tracker needs.</summary>
static class Native
{
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LASTINPUTINFO info);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hWnd, StringBuilder name, int count);
    [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern bool QueryFullProcessImageName(IntPtr h, uint flags, StringBuilder name, ref uint size);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);

    // Parts of Windows itself (desktop, taskbar, Start, search, Alt+Tab, notifications): nobody is working "in" these
    static readonly HashSet<string> ShellClasses = new(StringComparer.OrdinalIgnoreCase)
    {
        "Progman", "WorkerW", "Shell_TrayWnd", "Shell_SecondaryTrayWnd", "NotifyIconOverflowWindow", "TopLevelWindowForOverflowXamlIsland",
        "MultitaskingViewFrame", "XamlExplorerHostIslandWindow", "ForegroundStaging", "TaskListThumbnailWnd", "Windows.UI.Core.CoreWindow",
    };
    static readonly HashSet<string> ShellApps = new(StringComparer.OrdinalIgnoreCase)
    {
        "SearchHost", "SearchApp", "StartMenuExperienceHost", "ShellExperienceHost", "ShellHost", "LockApp", "TextInputHost", "Widgets", "consent",
    };

    /// <summary>True when what's in front is Windows itself rather than a program (or nothing at all).</summary>
    public static bool IsWindowsShell(IntPtr hwnd, string process, string title)
    {
        if (hwnd == IntPtr.Zero || string.IsNullOrEmpty(process)) return true;
        if (ShellApps.Contains(process)) return true;
        var sb = new StringBuilder(256);
        GetClassName(hwnd, sb, sb.Capacity);
        var cls = sb.ToString();
        if (ShellClasses.Contains(cls)) return true;
        // explorer.exe with no title is the desktop or taskbar; with a title it's a real folder window
        return process.Equals("explorer", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(title);
    }

    [StructLayout(LayoutKind.Sequential)]
    struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }

    /// <summary>Process name (e.g. "EXCEL") and title of the window in front, or empty strings.</summary>
    public static (IntPtr hwnd, string process, string title) ForegroundWindow()
    {
        var hwnd = GetForegroundWindow();
        if (hwnd == IntPtr.Zero) return (IntPtr.Zero, "", "");
        var sb = new StringBuilder(512);
        GetWindowText(hwnd, sb, sb.Capacity);
        GetWindowThreadProcessId(hwnd, out var pid);
        string process = "";
        try { using var p = Process.GetProcessById((int)pid); process = p.ProcessName; } catch { }
        if (string.IsNullOrEmpty(process) && pid != 0)
        {
            // Programs running as administrator: ask Windows for the program's file name instead
            var h = OpenProcess(0x1000 /* PROCESS_QUERY_LIMITED_INFORMATION */, false, pid);
            if (h != IntPtr.Zero)
            {
                var name = new StringBuilder(1024);
                uint size = (uint)name.Capacity;
                if (QueryFullProcessImageName(h, 0, name, ref size)) process = Path.GetFileNameWithoutExtension(name.ToString());
                CloseHandle(h);
            }
        }
        return (hwnd, process, sb.ToString());
    }

    /// <summary>How long since the last keyboard or mouse input.</summary>
    public static TimeSpan IdleTime()
    {
        var info = new LASTINPUTINFO { cbSize = (uint)Marshal.SizeOf<LASTINPUTINFO>() };
        if (!GetLastInputInfo(ref info)) return TimeSpan.Zero;
        var idleMs = unchecked((uint)Environment.TickCount - info.dwTime);
        return TimeSpan.FromMilliseconds(idleMs);
    }
}
