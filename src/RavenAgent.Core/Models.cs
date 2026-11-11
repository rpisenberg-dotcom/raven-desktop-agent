using System.Text.Json;
using System.Text.Json.Serialization;

namespace RavenAgent.Core;

/// <summary>Shared JSON settings: snake_case names, nulls omitted (matches the Base44 tables).</summary>
public static class Json
{
    public static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false,
    };

    public static string Iso(DateTime utc) =>
        DateTime.SpecifyKind(utc, DateTimeKind.Utc).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'");
}

/// <summary>One continuous stretch of work in one window. Becomes a TimeLog row.</summary>
public sealed record TimeLogRecord
{
    public required string AgentKey { get; init; }
    public required string StartTime { get; init; }
    public required string EndTime { get; init; }
    public required double DurationMinutes { get; init; }
    public required string ActivityType { get; init; }
    public string? App { get; init; }
    public string? WindowTitle { get; init; }
    public string? Url { get; init; }
    public string? FilePath { get; init; }
    /// <summary>What was on screen beyond the title, e.g. "Email: RE: Audit (jodi@koplin.com)".</summary>
    public string? Detail { get; init; }
    /// <summary>Client the agent found on screen or in the selected email (matched on this PC).</summary>
    public string? ClientHintId { get; init; }
    public string? ClientHintTerm { get; init; }
    public string? ClientHintSource { get; init; } // screen | email | title
    /// <summary>QuickBooks Online company open in the browser (its company/realm id), so the app can tell clients apart.</summary>
    public string? QboCompanyId { get; init; }
    /// <summary>In a call at the time (Teams, Zoom, Browser, Phone Link): the app holding the microphone.</summary>
    public string? InCall { get; init; }
}

public sealed record EmailRecord
{
    public required string ExternalId { get; init; }
    public string? ThreadId { get; init; }
    public required string Subject { get; init; }
    public required string FromEmail { get; init; }
    public string? ToEmail { get; init; }
    public string? CcEmail { get; init; }
    public string? Body { get; init; }
    public required string ReceivedDate { get; init; }
    public string? EmailAccountName { get; init; }
    public string? FolderPath { get; init; }
    public required string Direction { get; init; } // received | sent
    public bool IsRead { get; init; }
    /// <summary>Outlook's own ids, so the app can open the original email in Outlook.</summary>
    public string? OutlookEntryId { get; init; }
    public string? OutlookStoreId { get; init; }
}

public sealed record CalendarRecord
{
    public required string ExternalEventId { get; init; }
    public required string Title { get; init; }
    public required string StartTime { get; init; }
    public required string EndTime { get; init; }
    public bool IsAllDay { get; init; }
    public string? Location { get; init; }
    public List<string>? Attendees { get; init; }
    public string? SourceCalendar { get; init; }
    public Dictionary<string, object?>? Metadata { get; init; }
}

/// <summary>
/// Tells the server "these are ALL the events in this calendar between start and end",
/// so meetings deleted in Outlook disappear from the agenda too.
/// </summary>
public sealed record CalendarSync
{
    public required string SourceCalendar { get; init; }
    public required string WindowStart { get; init; }
    public required string WindowEnd { get; init; }
    public required List<string> ExternalEventIds { get; init; }
}

public sealed record IngestPayload
{
    public required string DeviceId { get; init; }
    public required string AgentVersion { get; init; }
    public bool PrivateMode { get; init; }
    public List<TimeLogRecord> TimeLogs { get; init; } = new();
    public List<EmailRecord> Emails { get; init; } = new();
    public List<CalendarRecord> CalendarEvents { get; init; } = new();
    public List<CalendarSync> CalendarSyncs { get; init; } = new();
    /// <summary>Every mailbox loaded in Outlook (including ones Raven chose not to read).</summary>
    public List<string>? Mailboxes { get; init; }
    /// <summary>Outlook ids for emails the app already has (sent once, so older emails can be opened in Outlook).</summary>
    public List<EmailIdRecord>? EmailIds { get; init; }
    /// <summary>Text conversations open in Phone Link, for business contacts only (never personal or unknown ones).</summary>
    public List<PhoneSnapshotRecord>? PhoneMessages { get; init; }
    /// <summary>Raven's answers in the desktop pop-ups (Done / Not yet / Undo, or just that it was shown).</summary>
    public List<NoticeAck>? NoticeAcks { get; init; }
    /// <summary>Notes jotted with Ctrl+Alt+N (questions, to-dos, notes for a client).</summary>
    public List<QuickNoteRecord>? Notes { get; init; }
    /// <summary>A quick check for calendar changes only (nothing else is sent).</summary>
    public bool? PollOnly { get; init; }
    /// <summary>What happened to calendar changes made in the app.</summary>
    public List<CalendarChangeResult>? CalendarChangeResults { get; init; }
}

/// <summary>A change made in the app (new / edited / deleted event, or "open in Outlook") for the agent to make in Outlook.</summary>
public sealed record CalendarChange
{
    public string Id { get; init; } = "";
    public string Action { get; init; } = "";
    public string TargetCalendar { get; init; } = "";
    public string EntryId { get; init; } = "";
    public string StoreId { get; init; } = "";
    public string OccurrenceStart { get; init; } = "";
    public string Scope { get; init; } = "occurrence";
    public string? Title { get; init; }
    public string StartTime { get; init; } = "";
    public string EndTime { get; init; } = "";
    public bool? IsAllDay { get; init; }
    public string? Location { get; init; }
    public string? Notes { get; init; }
    public List<string> Attendees { get; init; } = new();
    public bool Teams { get; init; }
}

public sealed record CalendarChangeResult(string Id, string Status, string Message, string? EntryId = null);

public sealed record QuickNoteRecord
{
    public required string AgentKey { get; init; }
    public required string Text { get; init; }
    public string Kind { get; init; } = "note"; // question | todo | note
    public string? ClientId { get; init; }
    public string? Category { get; init; } // client | business | personal
    public string? Context { get; init; }
}

public sealed record NoticeAck
{
    public required string TaskId { get; init; }
    public required string Action { get; init; } // shown_suggest | shown_done | done | not_yet | undo
}

/// <summary>Something to show Raven: a task Raven marked done from an email, or one it thinks is done.</summary>
public sealed record RavenNotice(string TaskId, string Kind, string Title, string Why);

public sealed record PhoneSnapshotRecord
{
    public required string Contact { get; init; }
    public required string Text { get; init; }
    public required string CapturedAt { get; init; }
}

public sealed record EmailIdRecord
{
    public required string ExternalId { get; init; }
    public required string OutlookEntryId { get; init; }
    public string? OutlookStoreId { get; init; }
}
