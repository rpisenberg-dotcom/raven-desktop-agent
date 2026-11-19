using RavenAgent.Core;
using Xunit;
using static RavenAgent.Core.TranscriptBuilder;

namespace RavenAgent.Core.Tests;

public class TranscriptBuilderTests
{
    [Fact]
    public void Reads_whisper_json_and_skips_blank_audio()
    {
        var json = """
        {"transcription":[
          {"offsets":{"from":0,"to":4000},"text":" Hi Jonathan, thanks for joining."},
          {"offsets":{"from":4000,"to":6000},"text":" [BLANK_AUDIO]"},
          {"offsets":{"from":6500,"to":9000},"text":" (keyboard clicking)"},
          {"offsets":{"from":9000,"to":12500},"text":" Let's start with the AR aging."}
        ]}
        """;
        var segs = ParseWhisper(json);
        Assert.Equal(2, segs.Count);
        Assert.Equal(9.0, segs[1].From);
        Assert.Equal("Let's start with the AR aging.", segs[1].Text);
    }

    [Fact]
    public void Names_other_speakers_from_captions()
    {
        var mine = new List<Seg> { new(0, 3, "Good morning everyone, let's look at the invoices.") };
        var others = new List<Seg>
        {
            new(4, 8, "Sure, I sent the packing list for the Mexico shipment yesterday."),
            new(9, 12, "And I still need the payroll report from ADP."),
        };
        var captions = new List<Caption>
        {
            new("Raven Pick", "Good morning everyone let's look at the invoices", 1),
            new("Jonathan Egan", "Sure I sent the packing list for the Mexico shipment yesterday", 5),
            new("John Smith", "And I still need the payroll report from ADP", 10),
        };
        var lines = Build(mine, others, captions, "Raven");
        Assert.Equal(new[] { "Raven Pick", "Jonathan Egan", "John Smith" }, lines.Select(l => l.Speaker));
    }

    [Fact]
    public void Drops_speaker_echo_from_microphone_track()
    {
        var mine = new List<Seg>
        {
            new(0, 3, "Can you send the bank statements?"),
            new(4.2, 8, "I sent the packing list for the Mexico shipment yesterday"), // heard through the speakers
        };
        var others = new List<Seg> { new(4, 8, "Sure, I sent the packing list for the Mexico shipment yesterday.") };
        var lines = Build(mine, others, new List<Caption>(), "Raven");
        Assert.Equal(2, lines.Count);
        Assert.Equal("Raven", lines[0].Speaker);
        Assert.Equal("Other participant", lines[1].Speaker);
    }

    [Fact]
    public void Joins_back_to_back_pieces_from_the_same_speaker()
    {
        var mine = new List<Seg> { new(0, 2, "First point."), new(2, 4, "Second point.") };
        var lines = Build(mine, new List<Seg>(), new List<Caption>(), "Raven");
        var line = Assert.Single(lines);
        Assert.Equal("First point. Second point.", line.Text);
        Assert.Contains("[00:00:00] Raven: First point.", ToText(lines));
    }

    [Fact]
    public void Drops_invented_repeats()
    {
        var segs = new List<Seg>
        {
            new(0, 4, "Let's look at the invoices."),
            new(10, 12, "I'm going to go ahead and record this meeting."),
            new(14, 16, "I'm going to go ahead and record this meeting."),
            new(20, 24, "I'm going to go ahead and record this meeting. I'm going to go ahead and record this meeting."),
            new(30, 33, "Thanks, talk soon."),
        };
        var kept = DropRepeats(segs);
        Assert.Equal(new[] { "Let's look at the invoices.", "I'm going to go ahead and record this meeting.", "Thanks, talk soon." }, kept.Select(k => k.Text));
    }
}
