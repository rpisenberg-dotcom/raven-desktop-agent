using System.Diagnostics;
using System.Text;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// After a recorded meeting: joins the video and sound into a normal .mp4 (or .m4a when there's no video),
/// transcribes Raven's track and everyone else's track on this computer with whisper, names the speakers from the
/// meeting's live captions, and saves "Transcript.txt" next to the recording. Runs at low priority so the computer
/// stays responsive.
/// </summary>
static class MeetingTranscriber
{
    public static async Task<LiveTranscriptRecord?> ProcessAsync(MeetingRecorder.Result r, List<CaptionReader.Line> captions, string myName, Action<string> status)
    {
        try
        {
            if (!await MeetingTools.EnsureAsync()) { status(MeetingTools.Status); return null; }
            var safe = MeetingRecorder.SafeName(r.Title);
            var tracks = new[] { r.MinePath, r.TheirsPath }.Where(p => p != null && File.Exists(p) && new FileInfo(p).Length > 1000).Cast<string>().ToList();
            if (tracks.Count == 0) { status("No sound was recorded"); return null; }

            // 1. The recording itself
            status("Saving the recording...");
            var audio = Path.Combine(r.Folder, "audio.m4a");
            var mix = new List<string> { "-hide_banner", "-loglevel", "error", "-y" };
            foreach (var t in tracks) { mix.Add("-i"); mix.Add(t); }
            if (tracks.Count == 2) mix.AddRange(new[] { "-filter_complex", "[0:a][1:a]amix=inputs=2:duration=longest:normalize=0[a]", "-map", "[a]" });
            mix.AddRange(new[] { "-c:a", "aac", "-b:a", "96k", audio });
            await RunAsync(MeetingTools.Ffmpeg, mix, 30);

            string? main = null;
            if (r.Parts.Count == 0)
            {
                main = Path.Combine(r.Folder, $"{safe}.m4a");
                if (File.Exists(audio)) File.Move(audio, main, true);
            }
            else if (r.Parts.Count == 1)
            {
                var part = r.Parts[0];
                var output = Path.Combine(r.Folder, $"{safe}.mp4");
                var ok = await RunAsync(MeetingTools.Ffmpeg, new List<string>
                {
                    "-hide_banner", "-loglevel", "error", "-y", "-i", part.File,
                    "-ss", Secs(part.Start), "-t", Secs(Math.Max(1, part.End - part.Start)), "-i", audio,
                    "-map", "0:v", "-map", "1:a", "-c:v", "copy", "-c:a", "aac", "-b:a", "96k", "-shortest", "-movflags", "+faststart", output,
                }, 30);
                main = ok && File.Exists(output) && new FileInfo(output).Length > 0 ? output : part.File;
                if (main == output) { TryDelete(part.File); TryDelete(audio); }
            }
            else
            {
                // One continuous recording: the parts (meeting window, then the shared screen, ...) joined in order on a
                // common 1920x1080 frame, each held until the next begins so the picture stays in step with the sound
                status("Joining the video into one recording...");
                var output = Path.Combine(r.Folder, $"{safe}.mp4");
                var args = new List<string> { "-hide_banner", "-loglevel", "error", "-y" };
                foreach (var p in r.Parts) { args.Add("-i"); args.Add(p.File); }
                args.Add("-i"); args.Add(audio);
                var chains = new List<string>();
                for (var i = 0; i < r.Parts.Count; i++)
                {
                    var p = r.Parts[i];
                    var span = (i < r.Parts.Count - 1 ? r.Parts[i + 1].Start : p.End) - p.Start;
                    var lead = i == 0 ? p.Start : 0; // the first part starts a few seconds after the sound
                    chains.Add($"[{i}:v]fps=10,scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2:color=black,setsar=1,format=yuv420p,setpts=PTS-STARTPTS,"
                        + $"tpad=stop_mode=clone:stop_duration=60{(lead > 0.05 ? $",tpad=start_mode=clone:start_duration={Secs(lead)}" : "")},"
                        + $"trim=duration={Secs(Math.Max(0.5, span + lead))},setpts=PTS-STARTPTS[v{i}]");
                }
                var graph = string.Join(";", chains) + ";" + string.Concat(Enumerable.Range(0, r.Parts.Count).Select(i => $"[v{i}]")) + $"concat=n={r.Parts.Count}:v=1:a=0[v]";
                args.AddRange(new[] { "-filter_complex", graph, "-map", "[v]", "-map", $"{r.Parts.Count}:a", "-c:v", "libx264", "-preset", "veryfast", "-crf", "28",
                    "-c:a", "aac", "-b:a", "96k", "-movflags", "+faststart", output });
                var ok = await RunAsync(MeetingTools.Ffmpeg, args, 120);
                if (ok && File.Exists(output) && new FileInfo(output).Length > 0)
                {
                    main = output;
                    foreach (var p in r.Parts) TryDelete(p.File);
                    TryDelete(audio);
                }
                else
                {
                    // Joining failed: keep each part as its own video so nothing is lost
                    Log.Info("Couldn't join the video parts; saving them separately");
                    for (var i = 0; i < r.Parts.Count; i++)
                    {
                        var part = r.Parts[i];
                        var name = Path.Combine(r.Folder, $"{safe} - part {i + 1} ({MeetingRecorder.SafeName(part.Label)}).mp4");
                        var ok2 = await RunAsync(MeetingTools.Ffmpeg, new List<string>
                        {
                            "-hide_banner", "-loglevel", "error", "-y", "-i", part.File,
                            "-ss", Secs(part.Start), "-t", Secs(Math.Max(1, part.End - part.Start)), "-i", audio,
                            "-map", "0:v", "-map", "1:a", "-c:v", "copy", "-c:a", "aac", "-b:a", "96k", "-shortest", "-movflags", "+faststart", name,
                        }, 30);
                        if (ok2 && File.Exists(name)) { TryDelete(part.File); main ??= name; } else main ??= part.File;
                    }
                }
            }

            // 2. Speech to text, one track at a time (Raven / everyone else)
            var threads = Math.Clamp(Environment.ProcessorCount / 2, 2, 8).ToString();
            async Task<List<TranscriptBuilder.Seg>> Transcribe(string? wav, string label)
            {
                if (wav == null || !File.Exists(wav)) return new();
                status($"Transcribing ({label})...");
                var small = Path.ChangeExtension(wav, ".16k.wav");
                await RunAsync(MeetingTools.Ffmpeg, new List<string> { "-hide_banner", "-loglevel", "error", "-y", "-i", wav, "-ar", "16000", "-ac", "1", "-c:a", "pcm_s16le", small }, 30);
                var outBase = Path.ChangeExtension(wav, null) + "-text";
                var args = new List<string> { "-m", MeetingTools.Model, "-f", small, "-l", "en", "-oj", "-of", outBase, "-np", "-t", threads };
                if (File.Exists(MeetingTools.VadModel)) args.AddRange(new[] { "--vad", "-vm", MeetingTools.VadModel }); // skip the silences
                await RunAsync(MeetingTools.Whisper, args, 180);
                var json = outBase + ".json";
                var segs = File.Exists(json) ? TranscriptBuilder.DropRepeats(TranscriptBuilder.ParseWhisper(await File.ReadAllTextAsync(json))) : new();
                TryDelete(small); TryDelete(json);
                Log.Info($"Transcribed {label}: {segs.Count} pieces");
                return segs;
            }
            var mine = await Transcribe(r.MinePath, "your microphone");
            var others = await Transcribe(r.TheirsPath, "the other people");
            if (mine.Count + others.Count == 0) { status("Nothing to transcribe (the recording was silent)"); return null; }

            // 3. Names
            var caps = captions.Select(c => new TranscriptBuilder.Caption(c.Speaker, c.Text, (c.At - r.StartedUtc).TotalSeconds)).ToList();
            var lines = TranscriptBuilder.Build(mine, others, caps, myName);
            var header = $"{r.Title}{Environment.NewLine}{r.StartedUtc.ToLocalTime():dddd, MMMM d, yyyy h:mm tt}{Environment.NewLine}{Environment.NewLine}";
            await File.WriteAllTextAsync(Path.Combine(r.Folder, "Transcript.txt"), header + TranscriptBuilder.ToText(lines), Encoding.UTF8);
            foreach (var t in tracks) TryDelete(t);
            status("Transcript ready");

            return new LiveTranscriptRecord
            {
                MeetingKey = "rec-" + Guid.NewGuid().ToString("N"),
                Title = r.Title,
                StartedAt = Json.Iso(r.StartedUtc),
                LastSeenAt = Json.Iso(r.EndedUtc),
                Ended = true,
                FromLine = 0,
                Source = "recording",
                RecordingFile = main,
                Lines = lines.Take(2000).Select(l => new LiveTranscriptLine { Speaker = l.Speaker, Text = l.Text, At = Json.Iso(r.StartedUtc.AddSeconds(l.At)) }).ToList(),
            };
        }
        catch (Exception ex)
        {
            Log.Error("Processing the recording", ex);
            status($"Couldn't finish the transcript: {ex.Message}");
            return null;
        }
    }

    static string Secs(double s) => s.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);

    static void TryDelete(string? path) { try { if (path != null && File.Exists(path)) File.Delete(path); } catch { } }

    static async Task<bool> RunAsync(string exe, List<string> args, int timeoutMinutes)
    {
        var psi = new ProcessStartInfo(exe) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true };
        foreach (var a in args) psi.ArgumentList.Add(a);
        var err = new StringBuilder();
        using var p = Process.Start(psi)!;
        try { p.PriorityClass = ProcessPriorityClass.BelowNormal; } catch { }
        p.ErrorDataReceived += (_, e) => { if (e.Data != null && err.Length < 3000) err.AppendLine(e.Data); };
        p.OutputDataReceived += (_, _) => { };
        p.BeginErrorReadLine();
        p.BeginOutputReadLine();
        using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(timeoutMinutes));
        try { await p.WaitForExitAsync(cts.Token); }
        catch (OperationCanceledException) { try { p.Kill(true); } catch { } Log.Info($"{Path.GetFileName(exe)} took too long and was stopped"); return false; }
        if (p.ExitCode != 0) Log.Info($"{Path.GetFileName(exe)} exit {p.ExitCode}: {err}");
        return p.ExitCode == 0;
    }
}
