using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace RavenAgent;

/// <summary>
/// Records a meeting on this computer, only after Raven confirms everyone has been told:
///   me.wav      her microphone (always Raven)
///   others.wav  what the computer plays (everyone else in the call)
///   video-N.mkv the meeting window or a screen, via ffmpeg (it works while the window is behind others,
///               not while it's minimized). Switching the video source starts a new part.
/// Keeping the two sound tracks apart is what lets the transcript tell Raven from everyone else.
/// </summary>
sealed class MeetingRecorder
{
    public sealed record VideoSource(string Kind, IntPtr Handle, string Label); // Kind: window | screen | none
    public sealed record Part(string File, double Start, double End, string Label);
    public sealed record Result(string Folder, string Title, string App, DateTime StartedUtc, DateTime EndedUtc,
        List<Part> Parts, string? MinePath, string? TheirsPath);

    public string Folder { get; }
    public string Title { get; }
    public string App { get; }
    public DateTime StartedUtc { get; } = DateTime.UtcNow;
    public TimeSpan Elapsed => _clock.Elapsed;
    public VideoSource Video { get; private set; } = new("none", IntPtr.Zero, "Sound only");

    readonly Stopwatch _clock = Stopwatch.StartNew();
    readonly AudioTrack? _mine, _theirs;
    readonly List<Part> _parts = new();
    readonly object _videoLock = new();
    Process? _ffmpeg;
    (string file, double start, string label)? _part;

    MeetingRecorder(string folder, string title, string app)
    {
        Folder = folder; Title = title; App = app;
        Directory.CreateDirectory(folder);
        _theirs = AudioTrack.Open(Path.Combine(folder, "others.wav"), _clock, loopback: true);
        _mine = AudioTrack.Open(Path.Combine(folder, "me.wav"), _clock, loopback: false);
    }

    public static MeetingRecorder Start(string root, string title, string app, VideoSource video)
    {
        var folder = Path.Combine(root, $"{DateTime.Now:yyyy-MM-dd HHmm} {SafeName(title)}");
        var r = new MeetingRecorder(folder, title, app);
        Log.Info($"Recording '{title}' ({app}) to {folder}");
        _ = Task.Run(() => r.SwitchVideo(video));
        return r;
    }

    public bool HasSound => _mine != null || _theirs != null;

    /// <summary>Last time anyone could be heard: the meeting's sound or Raven's microphone.</summary>
    public DateTime LastSoundUtc => new[] { _mine?.LastSoundUtc ?? DateTime.MinValue, _theirs?.LastSoundUtc ?? DateTime.MinValue }.Max();

    /// <summary>Changes what the video shows (meeting window / a screen / none). Runs ffmpeg; call off the UI thread.</summary>
    public void SwitchVideo(VideoSource v)
    {
        lock (_videoLock)
        {
            StopVideo();
            Video = v;
            if (v.Kind == "none") return;
            if (!MeetingTools.HasFfmpeg) { Log.Info("Video recorder not downloaded yet: sound only for now"); Video = new("none", IntPtr.Zero, "Sound only"); return; }
            var file = Path.Combine(Folder, $"video-{_parts.Count + 1}.mkv");
            var source = v.Kind == "window" ? $"gfxcapture=hwnd={(ulong)v.Handle}" : $"gfxcapture=hmonitor={(ulong)v.Handle}";
            // The capture hands frames over on the graphics card; the first try copies them back for the encoder
            foreach (var download in new[] { true, false })
            {
                var chain = $"{source}:max_framerate=10:capture_cursor=1{(download ? ",hwdownload,format=bgra" : "")},scale='trunc(min(1920,iw)/2)*2':-2,format=yuv420p";
                var psi = new ProcessStartInfo(MeetingTools.Ffmpeg)
                {
                    UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardError = true,
                };
                foreach (var a in new[] { "-hide_banner", "-loglevel", "warning", "-y", "-filter_complex", chain,
                             "-c:v", "libx264", "-preset", "veryfast", "-crf", "30", "-g", "100", "-f", "matroska", file })
                    psi.ArgumentList.Add(a);
                var errors = new StringBuilder();
                try
                {
                    var p = Process.Start(psi)!;
                    try { p.PriorityClass = ProcessPriorityClass.BelowNormal; } catch { }
                    p.ErrorDataReceived += (_, e) => { if (e.Data != null && errors.Length < 4000) errors.AppendLine(e.Data); };
                    p.BeginErrorReadLine();
                    var start = _clock.Elapsed.TotalSeconds;
                    if (p.WaitForExit(4000)) { Log.Info($"Video ({v.Label}) didn't start{(download ? "" : " (2nd try)")}: {errors}"); continue; }
                    _ffmpeg = p;
                    _part = (file, start, v.Label);
                    Log.Info($"Recording video: {v.Label}");
                    return;
                }
                catch (Exception ex) { Log.Error("Starting the video recorder", ex); break; }
            }
            Video = new("none", IntPtr.Zero, "Sound only");
        }
    }

    void StopVideo()
    {
        var p = _ffmpeg;
        if (p == null || _part == null) return;
        try { p.StandardInput.Write("q"); p.StandardInput.Flush(); } catch { }
        if (!p.WaitForExit(10_000)) try { p.Kill(); } catch { }
        var (file, start, label) = _part.Value;
        if (File.Exists(file) && new FileInfo(file).Length > 0) _parts.Add(new Part(file, start, _clock.Elapsed.TotalSeconds, label));
        _ffmpeg = null;
        _part = null;
    }

    public Result Stop()
    {
        lock (_videoLock) StopVideo();
        _mine?.Stop();
        _theirs?.Stop();
        _clock.Stop();
        Log.Info($"Recording stopped after {_clock.Elapsed:hh\\:mm\\:ss}: {_parts.Count} video part(s)");
        return new Result(Folder, Title, App, StartedUtc, DateTime.UtcNow, _parts.ToList(),
            _mine != null ? Path.Combine(Folder, "me.wav") : null, _theirs != null ? Path.Combine(Folder, "others.wav") : null);
    }

    public static string SafeName(string s)
    {
        var bad = Path.GetInvalidFileNameChars();
        var clean = new string(s.Select(c => bad.Contains(c) ? ' ' : c).ToArray()).Trim();
        clean = System.Text.RegularExpressions.Regex.Replace(clean, @"\s+", " ");
        if (clean.Length == 0) clean = "Meeting";
        return clean.Length > 80 ? clean[..80].Trim() : clean;
    }

    /// <summary>One sound track written to a WAV file. Silence is filled in so both tracks stay in step with the clock
    /// (Windows sends nothing at all while the computer is quiet).</summary>
    sealed class AudioTrack
    {
        readonly IWaveIn _capture;
        readonly WaveFileWriter _writer;
        readonly Stopwatch _clock;
        readonly object _lock = new();
        long _written;
        bool _stopped;
        public DateTime LastSoundUtc { get; private set; } = DateTime.UtcNow;

        AudioTrack(IWaveIn capture, string path, Stopwatch clock)
        {
            _capture = capture; _clock = clock;
            _writer = new WaveFileWriter(path, capture.WaveFormat);
            capture.DataAvailable += OnData;
            capture.RecordingStopped += (_, e) => { if (e.Exception != null) Log.Error("Sound recording stopped", e.Exception); };
            capture.StartRecording();
        }

        public static AudioTrack? Open(string path, Stopwatch clock, bool loopback)
        {
            try
            {
                // The "communications" device is the one Teams and Zoom use unless set otherwise
                var en = new MMDeviceEnumerator();
                var dev = en.GetDefaultAudioEndpoint(loopback ? DataFlow.Render : DataFlow.Capture, Role.Communications);
                Log.Info($"Recording {(loopback ? "meeting sound from" : "microphone")}: {dev.FriendlyName}");
                IWaveIn cap = loopback ? new WasapiLoopbackCapture(dev) : new WasapiCapture(dev);
                return new AudioTrack(cap, path, clock);
            }
            catch (Exception ex)
            {
                Log.Error(loopback ? "Couldn't record the meeting sound" : "Couldn't record the microphone", ex);
                return null;
            }
        }

        long Expected()
        {
            var f = _writer.WaveFormat;
            var bytes = (long)(_clock.Elapsed.TotalSeconds * f.AverageBytesPerSecond);
            return bytes - bytes % f.BlockAlign;
        }

        void OnData(object? sender, WaveInEventArgs a)
        {
            lock (_lock)
            {
                if (_stopped) return;
                var f = _writer.WaveFormat;
                var gap = Expected() - (_written + a.BytesRecorded);
                if (gap > f.AverageBytesPerSecond / 5) Pad(gap - gap % f.BlockAlign);
                _writer.Write(a.Buffer, 0, a.BytesRecorded);
                _written += a.BytesRecorded;
                if (Loud(a.Buffer, a.BytesRecorded, f)) LastSoundUtc = DateTime.UtcNow;
            }
        }

        /// <summary>Anything above a whisper (about -45 dB) in this piece of sound.</summary>
        static bool Loud(byte[] buf, int n, WaveFormat f)
        {
            if (f.Encoding == WaveFormatEncoding.IeeeFloat || (f.Encoding == WaveFormatEncoding.Extensible && f.BitsPerSample == 32))
            {
                for (var i = 0; i + 4 <= n; i += 16) if (Math.Abs(BitConverter.ToSingle(buf, i)) > 0.006f) return true;
                return false;
            }
            if (f.BitsPerSample == 16)
            {
                for (var i = 0; i + 2 <= n; i += 8) if (Math.Abs((int)BitConverter.ToInt16(buf, i)) > 200) return true;
                return false;
            }
            return true; // unknown format: assume sound, so a meeting is never cut short
        }

        void Pad(long n)
        {
            var zeros = new byte[(int)Math.Min(n, 1 << 16)];
            while (n > 0)
            {
                var k = (int)Math.Min(n, zeros.Length);
                _writer.Write(zeros, 0, k);
                _written += k;
                n -= k;
            }
        }

        public void Stop()
        {
            try { _capture.StopRecording(); } catch { }
            lock (_lock)
            {
                if (_stopped) return;
                var f = _writer.WaveFormat;
                var gap = Expected() - _written;
                if (gap > 0) Pad(gap - gap % f.BlockAlign);
                _stopped = true;
                _writer.Dispose();
            }
            try { _capture.Dispose(); } catch { }
        }
    }
}

/// <summary>Finds the meeting's window and the screens, for the "what should the video show" choice.</summary>
static class MeetingWindows
{
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern IntPtr MonitorFromPoint(POINT pt, uint flags);
    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
    delegate bool EnumProc(IntPtr h, IntPtr l);

    static readonly string[] Browsers = { "msedge", "chrome", "firefox", "brave", "opera" };

    /// <summary>The meeting's own window for this call app, or null.</summary>
    public static (IntPtr hwnd, string title)? Find(string? app)
    {
        if (app is null or "Teams" or "Browser")
        {
            var teams = CaptionReader.FindMeetingWindow();
            if (teams != null) return teams;
        }
        (IntPtr, string)? found = null;
        EnumWindows((h, _) =>
        {
            if (!IsWindowVisible(h)) return true;
            var sb = new StringBuilder(512);
            GetWindowText(h, sb, sb.Capacity);
            var title = sb.ToString();
            if (title.Length == 0) return true;
            GetWindowThreadProcessId(h, out var pid);
            string proc;
            try { proc = Process.GetProcessById((int)pid).ProcessName.ToLowerInvariant(); } catch { return true; }
            var hit = app switch
            {
                "Zoom" => proc.Contains("zoom") && (title.Contains("Zoom Meeting", StringComparison.OrdinalIgnoreCase) || title.Contains("Meeting", StringComparison.OrdinalIgnoreCase)),
                "Webex" => proc.Contains("webex") || proc.Contains("ciscocollab"),
                _ => Browsers.Contains(proc) && (title.Contains("Meet -", StringComparison.OrdinalIgnoreCase) || title.StartsWith("Meet ", StringComparison.OrdinalIgnoreCase)
                     || title.Contains("Google Meet", StringComparison.OrdinalIgnoreCase) || title.Contains("Zoom", StringComparison.OrdinalIgnoreCase)
                     || title.Contains("Webex", StringComparison.OrdinalIgnoreCase)),
            };
            if (hit) { found = (h, title); return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static bool IsMinimized(IntPtr h) => IsIconic(h);

    [DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr h, uint flags);
    public static IntPtr MonitorOf(IntPtr hwnd) => MonitorFromWindow(hwnd, 2);

    /// <summary>The browser's little "teams.live.com is sharing your screen" bar, which sits on the shared screen.</summary>
    public static IntPtr SharingBar()
    {
        var found = IntPtr.Zero;
        EnumWindows((h, _) =>
        {
            if (!IsWindowVisible(h)) return true;
            var sb = new StringBuilder(512);
            GetWindowText(h, sb, sb.Capacity);
            var t = sb.ToString();
            if (t.Contains("is sharing your screen", StringComparison.OrdinalIgnoreCase) || t.Contains("is sharing a window", StringComparison.OrdinalIgnoreCase)
                || t.Contains("Sharing control bar", StringComparison.OrdinalIgnoreCase))
            { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    /// <summary>Meeting window first (if found), then each screen, then sound only.</summary>
    public static List<MeetingRecorder.VideoSource> Sources((IntPtr hwnd, string title)? meeting)
    {
        var list = new List<MeetingRecorder.VideoSource>();
        if (meeting != null) list.Add(new("window", meeting.Value.hwnd, "Meeting window"));
        var screens = Screen.AllScreens.OrderBy(s => s.Bounds.Left).ToList();
        for (var i = 0; i < screens.Count; i++)
        {
            var b = screens[i].Bounds;
            var hmon = MonitorFromPoint(new POINT { X = b.Left + b.Width / 2, Y = b.Top + b.Height / 2 }, 2);
            var name = screens.Count == 1 ? "My screen" : $"Screen {i + 1}{(screens[i].Primary ? " (main)" : "")}";
            list.Add(new("screen", hmon, name));
        }
        list.Add(new("none", IntPtr.Zero, "Sound only"));
        return list;
    }

    /// <summary>A name for the recording's folder and the meeting page.</summary>
    public static string TitleFor((IntPtr hwnd, string title)? meeting, string app)
    {
        if (meeting == null) return $"{app} meeting";
        var t = meeting.Value.title;
        if (t.Contains("Microsoft Teams", StringComparison.OrdinalIgnoreCase)) return CaptionReader.MeetingTitle(t);
        t = System.Text.RegularExpressions.Regex.Replace(t, @"\s*[-|]\s*(Google Chrome|Microsoft\u200B? Edge|Mozilla Firefox|Brave).*$", "");
        t = t.Replace("Meet - ", "").Trim();
        return t.Length is > 0 and <= 120 ? t : $"{app} meeting";
    }
}
