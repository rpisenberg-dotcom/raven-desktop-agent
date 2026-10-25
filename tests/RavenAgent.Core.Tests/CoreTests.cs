using RavenAgent.Core;
using Xunit;

namespace RavenAgent.Core.Tests;

public class SegmentBuilderTests
{
    static readonly DateTime T0 = new(2026, 10, 26, 13, 0, 0, DateTimeKind.Utc);
    static readonly TimeSpan Every = TimeSpan.FromSeconds(5);

    static Sample S(int sec, string proc, string title, bool idle = false, bool ignored = false) =>
        new(T0.AddSeconds(sec), proc, title, idle, ignored);

    static List<TimeLogRecord> Run(SegmentBuilder b, IEnumerable<Sample> samples, int flushAt)
    {
        var all = new List<TimeLogRecord>();
        foreach (var s in samples) all.AddRange(b.Add(s));
        all.AddRange(b.Flush(T0.AddSeconds(flushAt)));
        return all;
    }

    [Fact]
    public void Same_window_merges_into_one_segment()
    {
        var b = new SegmentBuilder("pc", Every);
        var samples = Enumerable.Range(0, 13).Select(i => S(i * 5, "EXCEL", "Koplin_EMC_Audit.xlsx - Excel"));
        var logs = Run(b, samples, 65);
        var log = Assert.Single(logs);
        Assert.Equal(1.08, log.DurationMinutes);
        Assert.Equal("document", log.ActivityType);
        Assert.Equal("2026-10-26T13:00:00Z", log.StartTime);
    }

    [Fact]
    public void Switching_windows_closes_the_previous_segment()
    {
        var b = new SegmentBuilder("pc", Every);
        var samples = Enumerable.Range(0, 12).Select(i => S(i * 5, "EXCEL", "A.xlsx - Excel"))
            .Concat(Enumerable.Range(12, 12).Select(i => S(i * 5, "OUTLOOK", "Inbox - raven@clericalguru.com - Outlook")));
        var logs = Run(b, samples, 120);
        Assert.Equal(2, logs.Count);
        Assert.Equal("2026-10-26T13:01:00Z", logs[0].EndTime);
        Assert.Equal("email", logs[1].ActivityType);
    }

    [Fact]
    public void Short_glances_are_dropped()
    {
        var b = new SegmentBuilder("pc", Every);
        var samples = new[] { S(0, "EXCEL", "A"), S(5, "EXCEL", "A"), S(10, "chrome", "x"), S(15, "EXCEL", "A") };
        var logs = Run(b, samples, 20);
        Assert.Empty(logs);
    }

    [Fact]
    public void Idle_and_ignored_time_is_not_recorded()
    {
        var b = new SegmentBuilder("pc", Every);
        var samples = Enumerable.Range(0, 10).Select(i => S(i * 5, "EXCEL", "A"))
            .Concat(Enumerable.Range(10, 100).Select(i => S(i * 5, "EXCEL", "A", idle: true)))
            .Concat(Enumerable.Range(110, 10).Select(i => S(i * 5, "chrome", "Bank", ignored: true)));
        var logs = Run(b, samples, 600);
        var log = Assert.Single(logs);
        Assert.Equal(0.83, log.DurationMinutes); // 50 seconds of real work
    }

    [Fact]
    public void Long_segments_are_split()
    {
        var b = new SegmentBuilder("pc", Every, maxLength: TimeSpan.FromMinutes(15));
        var samples = Enumerable.Range(0, 400).Select(i => S(i * 5, "WINWORD", "Letter.docx - Word")); // ~33 min
        var logs = Run(b, samples, 2000);
        Assert.Equal(3, logs.Count);
        Assert.All(logs, l => Assert.True(l.DurationMinutes <= 15.01));
        Assert.Equal(logs.Count, logs.Select(l => l.AgentKey).Distinct().Count());
    }

    [Fact]
    public void Lock_screen_counts_as_away()
    {
        var b = new SegmentBuilder("pc", Every);
        var samples = Enumerable.Range(0, 100).Select(i => S(i * 5, "LockApp", "Windows Default Lock Screen"));
        Assert.Empty(Run(b, samples, 500));
    }

    [Fact]
    public void Keys_are_stable_across_runs()
    {
        var a = Run(new SegmentBuilder("pc", Every), Enumerable.Range(0, 10).Select(i => S(i * 5, "EXCEL", "A")), 50);
        var c = Run(new SegmentBuilder("pc", Every), Enumerable.Range(0, 10).Select(i => S(i * 5, "EXCEL", "A")), 50);
        Assert.Equal(a[0].AgentKey, c[0].AgentKey);
    }
}

public class ClassifierTests
{
    [Theory]
    [InlineData("ms-teams.exe", "Meeting with Kent | Microsoft Teams", "meeting")]
    [InlineData("Zoom.exe", "Zoom Meeting", "meeting")]
    [InlineData("msedge.exe", "Fabcast - QuickBooks Online - Profile 1 - Microsoft Edge", "browser")]
    [InlineData("chrome.exe", "Meet - abc-defg-hij - meet.google.com", "meeting")]
    [InlineData("OUTLOOK.EXE", "Inbox - Outlook", "email")]
    [InlineData("EXCEL.EXE", "Book1 - Excel", "document")]
    [InlineData("ms-teams.exe", "Chat | Microsoft Teams", "app_usage")]
    public void Activity_types(string proc, string title, string expected) =>
        Assert.Equal(expected, Classifier.ActivityType(proc, title));

    [Fact]
    public void Private_windows_and_ignore_list_are_skipped()
    {
        Assert.True(Classifier.IsIgnored("msedge", "New tab - [InPrivate] - Microsoft Edge", [], []));
        Assert.True(Classifier.IsIgnored("spotify.exe", "x", ["Spotify"], []));
        Assert.True(Classifier.IsIgnored("chrome", "Chase Online - Banking", [], ["chase"]));
        Assert.False(Classifier.IsIgnored("EXCEL", "Koplin.xlsx", ["Spotify"], ["chase"]));
    }
}

public class OutlookFilterTests
{
    [Fact]
    public void Dates_use_the_format_outlook_expects()
    {
        var d = new DateTime(2026, 10, 6, 14, 5, 0);
        Assert.Equal("10/06/2026 02:05 PM", OutlookFilters.Date(d));
        Assert.Equal("[ReceivedTime] >= '10/06/2026 02:05 PM'", OutlookFilters.Since("ReceivedTime", d));
    }

    [Fact]
    public void Recurring_occurrences_get_different_ids()
    {
        var a = OutlookFilters.OccurrenceId("ENTRY1", new DateTime(2026, 10, 26, 9, 0, 0));
        var b = OutlookFilters.OccurrenceId("ENTRY1", new DateTime(2026, 11, 2, 9, 0, 0));
        Assert.NotEqual(a, b);
    }
}

public class UploadQueueTests
{
    [Fact]
    public void Queue_survives_restart_and_removes_sent_items()
    {
        var path = Path.Combine(Path.GetTempPath(), $"raven-q-{Guid.NewGuid():N}.json");
        try
        {
            var q = new UploadQueue(path);
            var log = new TimeLogRecord { AgentKey = "k1", StartTime = "a", EndTime = "b", DurationMinutes = 1, ActivityType = "app_usage" };
            var mail = new EmailRecord { ExternalId = "m1", Subject = "s", FromEmail = "f@x.com", ReceivedDate = "d", Direction = "received" };
            q.Add(new[] { log });
            q.Add(new[] { mail });

            var reopened = new UploadQueue(path);
            Assert.Equal(2, reopened.Count);
            var (logs, emails) = reopened.Peek(100);
            reopened.Remove(logs, emails);
            Assert.Equal(0, new UploadQueue(path).Count);
        }
        finally { File.Delete(path); }
    }
}

public class UrlTests
{
    [Fact]
    public void Segments_keep_the_browser_address()
    {
        var t0 = new DateTime(2026, 10, 26, 13, 0, 0, DateTimeKind.Utc);
        var b = new SegmentBuilder("pc", TimeSpan.FromSeconds(5));
        foreach (var i in Enumerable.Range(0, 10))
            b.Add(new Sample(t0.AddSeconds(i * 5), "msedge", "Fabcast - QuickBooks", false, false, "https://qbo.intuit.com/app/homepage"));
        var log = Assert.Single(b.Flush(t0.AddSeconds(50)));
        Assert.Equal("https://qbo.intuit.com/app/homepage", log.Url);
    }
}

public class ClientMatcherTests
{
    static ClientMatcher M() => new(new[]
    {
        new ClientEntry { Id = "k", Name = "Koplin Construction", Terms = { "Koplin" }, Domains = { "koplinconstruction.com" } },
        new ClientEntry { Id = "f", Name = "Fabcast", Terms = { "Fabcast Inc" }, Domains = { "fabcast.com" }, Emails = { "owner@gmail.com" } },
        new ClientEntry { Id = "a", Name = "ABC", Terms = { } },
    });

    [Fact]
    public void Finds_the_client_named_on_a_qbo_page()
    {
        var page = "QuickBooks Accountant  Koplin Construction  Bank transactions  For review  Categorized";
        var m = M().MatchText(page);
        Assert.Equal("k", m?.ClientId);
        Assert.Equal("Koplin Construction", m?.Term);
    }

    [Fact]
    public void Most_mentioned_client_wins_in_a_chat()
    {
        var chat = "Draft an email to Kent at Koplin about the audit. Koplin's payroll... Also cc Fabcast? Koplin wants it Friday.";
        Assert.Equal("k", M().MatchText(chat)?.ClientId);
    }

    [Fact]
    public void Matches_file_names_with_underscores()
    {
        Assert.Equal("k", M().MatchText("Koplin_EMC_Audit.xlsx - Excel")?.ClientId);
        Assert.Equal("k", M().MatchText("koplin-construction-q3.pdf")?.ClientId);
    }

    [Fact]
    public void Whole_words_only()
    {
        Assert.Null(M().MatchText("ABCD report and Fabcastle pages"));
        Assert.Equal("a", M().MatchText("Notes for ABC, Q3")?.ClientId);
    }

    [Fact]
    public void Email_addresses_match_by_domain_or_exact_address()
    {
        Assert.Equal("k", M().MatchAddresses(new[] { "raven@clericalguru.com", "Jodi@KoplinConstruction.com" })?.ClientId);
        Assert.Equal("f", M().MatchAddresses(new[] { "owner@gmail.com" })?.ClientId);
        Assert.Equal("k", M().MatchAddresses(new[] { "x@mail.koplinconstruction.com" })?.ClientId);
        Assert.Null(M().MatchAddresses(new[] { "someone@gmail.com" }));
    }

    [Fact]
    public void A_late_client_hint_is_adopted_not_split()
    {
        var t0 = new DateTime(2026, 10, 26, 13, 0, 0, DateTimeKind.Utc);
        var b = new SegmentBuilder("pc", TimeSpan.FromSeconds(5));
        var all = new List<TimeLogRecord>();
        for (var i = 0; i < 10; i++)
            all.AddRange(b.Add(new Sample(t0.AddSeconds(i * 5), "claude", "Claude", false, false, null,
                i < 3 ? null : "On screen: Koplin Construction", i < 3 ? null : "k", i < 3 ? null : "Koplin")));
        // switching client mid-window splits the block
        for (var i = 10; i < 20; i++)
            all.AddRange(b.Add(new Sample(t0.AddSeconds(i * 5), "claude", "Claude", false, false, null, "On screen: Fabcast", "f", "Fabcast")));
        all.AddRange(b.Flush(t0.AddSeconds(100)));
        Assert.Equal(2, all.Count);
        Assert.Equal("k", all[0].ClientHintId);
        Assert.Equal("On screen: Koplin Construction", all[0].Detail);
        Assert.Equal("f", all[1].ClientHintId);
    }
}

public class AmbiguityTests
{
    static ClientMatcher M() => new(new[]
    {
        new ClientEntry { Id = "b", Name = "Bare Bone Vodka" },
        new ClientEntry { Id = "f", Name = "Fabcast Solutions", Terms = { "Fabcast" } },
        new ClientEntry { Id = "h", Name = "Hudson Products Corporation", Terms = { "Hudson Products" } },
    });

    [Fact]
    public void Lists_with_several_clients_give_no_screen_guess()
    {
        var orderList = "Orders\nBare Bone Vodka 2026-051\nFabcast Solutions 2026-052\nHudson Products 2026-053";
        Assert.Null(M().MatchText(orderList, requireClearWinner: true));
        Assert.NotNull(M().MatchText(orderList)); // titles/emails still use the plain best match
    }

    [Fact]
    public void A_clearly_dominant_client_still_wins()
    {
        var order = "Order 2026-052 Fabcast Solutions\nBill to: Fabcast Solutions\nShip to: Fabcast\nRecent: Bare Bone Vodka";
        Assert.Equal("f", M().MatchText(order, requireClearWinner: true)?.ClientId);
    }
}
