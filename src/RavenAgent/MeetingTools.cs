using System.IO.Compression;
using System.Net.Http;

namespace RavenAgent;

/// <summary>
/// The free tools for meeting recordings, downloaded once into %LOCALAPPDATA%\RavenAgent\tools:
///   ffmpeg (records the meeting window or a screen, and joins video and sound), whisper.cpp (speech to text,
///   runs on this computer) and its English speech model. Nothing is sent anywhere to be transcribed.
/// </summary>
static class MeetingTools
{
    const string FfmpegZip = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip";
    const string WhisperZip = "https://github.com/ggml-org/whisper.cpp/releases/download/v1.9.2/whisper-bin-x64.zip";
    const string ModelUrl = "https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-base.en.bin";
    // Voice activity detection: skips silence, which otherwise makes the transcriber invent repeated sentences
    const string VadUrl = "https://huggingface.co/ggml-org/whisper-vad/resolve/main/ggml-silero-v5.1.2.bin";

    public static string Folder => Path.Combine(AgentConfig.DataFolder, "tools");
    public static string Ffmpeg => Path.Combine(Folder, "ffmpeg.exe");
    public static string Whisper => Path.Combine(Folder, "whisper", "whisper-cli.exe");
    public static string Model => Path.Combine(Folder, "ggml-base.en.bin");
    public static string VadModel => Path.Combine(Folder, "ggml-silero-v5.1.2.bin");

    public static bool HasFfmpeg => File.Exists(Ffmpeg);
    public static bool HasWhisper => File.Exists(Whisper) && File.Exists(Model);
    public static bool Ready => HasFfmpeg && HasWhisper && File.Exists(VadModel);

    static readonly SemaphoreSlim Gate = new(1, 1);
    public static string Status { get; private set; } = "";

    /// <summary>Downloads whatever is missing (about 400 MB the first time). Safe to call any time.</summary>
    public static async Task<bool> EnsureAsync()
    {
        if (Ready) return true;
        await Gate.WaitAsync();
        try
        {
            Directory.CreateDirectory(Folder);
            using var http = new HttpClient { Timeout = TimeSpan.FromMinutes(30) };
            http.DefaultRequestHeaders.UserAgent.ParseAdd("RavenAgent");
            if (!HasFfmpeg)
            {
                Status = "Downloading the video recorder (ffmpeg)...";
                Log.Info(Status);
                var zip = Path.Combine(Folder, "ffmpeg.zip");
                await DownloadAsync(http, FfmpegZip, zip);
                using (var z = ZipFile.OpenRead(zip))
                {
                    var exe = z.Entries.FirstOrDefault(e => e.FullName.EndsWith("/bin/ffmpeg.exe", StringComparison.OrdinalIgnoreCase))
                              ?? throw new InvalidOperationException("ffmpeg.exe wasn't in the download");
                    exe.ExtractToFile(Ffmpeg + ".part", true);
                }
                File.Move(Ffmpeg + ".part", Ffmpeg, true);
                File.Delete(zip);
            }
            if (!File.Exists(Whisper))
            {
                Status = "Downloading the transcriber (whisper)...";
                Log.Info(Status);
                var zip = Path.Combine(Folder, "whisper.zip");
                await DownloadAsync(http, WhisperZip, zip);
                var dir = Path.GetDirectoryName(Whisper)!;
                Directory.CreateDirectory(dir);
                using (var z = ZipFile.OpenRead(zip))
                    foreach (var e in z.Entries)
                    {
                        var name = e.Name;
                        if (name.Length == 0) continue;
                        if (name.Equals("whisper-cli.exe", StringComparison.OrdinalIgnoreCase) || name.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
                            e.ExtractToFile(Path.Combine(dir, name), true);
                    }
                File.Delete(zip);
                if (!File.Exists(Whisper)) throw new InvalidOperationException("whisper-cli.exe wasn't in the download");
            }
            if (!File.Exists(Model))
            {
                Status = "Downloading the speech model...";
                Log.Info(Status);
                await DownloadAsync(http, ModelUrl, Model + ".part");
                File.Move(Model + ".part", Model, true);
            }
            if (!File.Exists(VadModel))
            {
                try { await DownloadAsync(http, VadUrl, VadModel + ".part"); File.Move(VadModel + ".part", VadModel, true); }
                catch (Exception ex) { Log.Error("Silence detector download (transcripts still work without it)", ex); }
            }
            Status = "Ready";
            Log.Info("Meeting recording tools ready");
            return true;
        }
        catch (Exception ex)
        {
            Status = $"Couldn't download the recording tools: {ex.Message}";
            Log.Error("Meeting tools download", ex);
            return false;
        }
        finally { Gate.Release(); }
    }

    static async Task DownloadAsync(HttpClient http, string url, string path)
    {
        using var res = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
        res.EnsureSuccessStatusCode();
        await using var src = await res.Content.ReadAsStreamAsync();
        await using var dst = File.Create(path);
        await src.CopyToAsync(dst);
    }
}
