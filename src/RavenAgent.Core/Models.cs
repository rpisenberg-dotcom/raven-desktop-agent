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
}

public sealed record EmailIdRecord
{
    public required string ExternalId { get; init; }
    public required string OutlookEntryId { get; init; }
    public string? OutlookStoreId { get; init; }
}
