using System.Text.Json;
using System.Text.RegularExpressions;

namespace RavenAgent.Core;

/// <summary>
/// Turns a recorded meeting into a transcript with names. Two audio tracks are transcribed separately:
/// Raven's microphone (always Raven) and the computer's sound (everyone else). Names for "everyone else" come
/// from the meeting's own live captions, matched by time and wording.
/// </summary>
public static class TranscriptBuilder
{
    /// <summary>One stretch of speech from the transcriber, in seconds from the start of the recording.</summary>
    public sealed record Seg(double From, double To, string Text);

    /// <summary>A live caption line: who the meeting app said was speaking, and when it appeared.</summary>
    public sealed record Caption(string Speaker, string Text, double At);

    public sealed record Line(string Speaker, string Text, double At);

    static readonly Regex Noise = new(@"^\s*[\[\(\*].*[\]\)\*]\s*$");
    static readonly Regex Word = new(@"[a-z0-9']+");

    /// <summary>Reads whisper.cpp's JSON output (-oj).</summary>
    public static List<Seg> ParseWhisper(string json)
    {
        var list = new List<Seg>();
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("transcription", out var items)) return list;
        foreach (var it in items.EnumerateArray())
        {
            var text = it.TryGetProperty("text", out var t) ? (t.GetString() ?? "").Trim() : "";
            if (text.Length == 0 || Noise.IsMatch(text)) continue; // [BLANK_AUDIO], (music), *typing*
            double from = 0, to = 0;
            if (it.TryGetProperty("offsets", out var o))
            {
                from = o.GetProperty("from").GetDouble() / 1000.0;
                to = o.GetProperty("to").GetDouble() / 1000.0;
            }
            list.Add(new Seg(from, Math.Max(from, to), text));
        }
        return list;
    }

    /// <summary>
    /// The transcriber sometimes repeats one sentence over quiet stretches ("I'm going to go ahead and record this
    /// meeting." again and again). Drops a piece that repeats one of the last few, and repeated sentences inside a piece.
    /// </summary>
    public static List<Seg> DropRepeats(List<Seg> segs)
    {
        static string Norm(string t) => string.Join(' ', Word.Matches(t.ToLowerInvariant()).Select(m => m.Value));
        var kept = new List<Seg>();
        foreach (var s in segs)
        {
            var sentences = Regex.Split(s.Text, @"(?<=[.!?])\s+").Where(x => x.Trim().Length > 0).ToList();
            var unique = new List<string>();
            foreach (var sen in sentences)
                if (unique.Count == 0 || Norm(unique[^1]) != Norm(sen)) unique.Add(sen.Trim());
            var text = string.Join(' ', unique);
            var norm = Norm(text);
            if (norm.Length == 0) continue;
            if (kept.Skip(Math.Max(0, kept.Count - 3)).Any(k => Norm(k.Text) == norm)) continue;
            kept.Add(s with { Text = text });
        }
        return kept;
    }

    static HashSet<string> Words(string s) =>
        Word.Matches(s.ToLowerInvariant()).Select(m => m.Value).Where(w => w.Length >= 3).ToHashSet();

    /// <summary>Share of the shorter text's words that also appear in the other (0 to 1).</summary>
    public static double Similar(string a, string b)
    {
        var wa = Words(a); var wb = Words(b);
        if (wa.Count == 0 || wb.Count == 0) return 0;
        var shared = wa.Count(wb.Contains);
        return (double)shared / Math.Min(wa.Count, wb.Count);
    }

    public static List<Line> Build(List<Seg> mine, List<Seg> others, List<Caption> captions, string myName)
    {
        // 1. Without a headset, Raven's microphone also hears the other people through her speakers. Those echoes
        //    show up on her track at the same moment as on theirs, with the same words: drop them from hers.
        var myOwn = mine.Where(m => !others.Any(o => o.From - 1.5 < m.To && m.From < o.To + 1.5 && Similar(m.Text, o.Text) >= 0.6)).ToList();

        // 2. Raven's own name as the captions show it ("Raven Pick"), from the caption speaker who best matches her track
        var meInCaptions = captions
            .GroupBy(c => c.Speaker)
            .Select(g => (name: g.Key, score: g.Sum(c => myOwn.Any(m => Math.Abs(m.From - c.At) < 15 && Similar(m.Text, c.Text) >= 0.5) ? 1 : 0)))
            .Where(x => x.score >= 1 && !string.IsNullOrWhiteSpace(x.name))
            .OrderByDescending(x => x.score)
            .Select(x => x.name)
            .FirstOrDefault();
        var me = meInCaptions ?? myName;
        var theirCaptions = captions.Where(c => c.Speaker != meInCaptions && !string.IsNullOrWhiteSpace(c.Speaker)).ToList();

        // 3. Everyone else: the caption speaker whose words match best around the same time, then the nearest caption
        string NameFor(Seg s)
        {
            var near = theirCaptions.Where(c => c.At >= s.From - 4 && c.At <= s.To + 6).ToList();
            var best = near.Select(c => (c, score: Similar(s.Text, c.Text))).OrderByDescending(x => x.score).FirstOrDefault();
            if (best.c != null && best.score >= 0.3) return best.c.Speaker;
            var closest = theirCaptions.Where(c => Math.Abs(c.At - s.From) <= 10).OrderBy(c => Math.Abs(c.At - s.From)).FirstOrDefault();
            return closest?.Speaker ?? "Other participant";
        }

        var lines = myOwn.Select(s => new Line(me, s.Text, s.From))
            .Concat(others.Select(s => new Line(NameFor(s), s.Text, s.From)))
            .OrderBy(l => l.At)
            .ToList();

        // 4. One line per turn: join a speaker's back-to-back pieces
        var merged = new List<Line>();
        foreach (var l in lines)
        {
            var last = merged.Count > 0 ? merged[^1] : null;
            if (last != null && last.Speaker == l.Speaker && last.Text.Length + l.Text.Length < 700)
                merged[^1] = last with { Text = $"{last.Text} {l.Text}" };
            else merged.Add(l);
        }
        return merged;
    }

    /// <summary>"[00:12:05] Jonathan Egan: ..." lines for the transcript file saved next to the recording.</summary>
    public static string ToText(IEnumerable<Line> lines) =>
        string.Join(Environment.NewLine, lines.Select(l => $"[{TimeSpan.FromSeconds(Math.Max(0, l.At)):hh\\:mm\\:ss}] {l.Speaker}: {l.Text}"));
}
