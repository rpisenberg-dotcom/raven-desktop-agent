using System.Globalization;

namespace RavenAgent.Core;

/// <summary>Builds Outlook "Restrict" filter strings. Outlook wants local time in this exact format.</summary>
public static class OutlookFilters
{
    public static string Date(DateTime local) =>
        local.ToString("MM/dd/yyyy hh:mm tt", CultureInfo.InvariantCulture);

    public static string Since(string property, DateTime local) => $"[{property}] >= '{Date(local)}'";

    /// <summary>Events that overlap [start, end).</summary>
    public static string Overlapping(DateTime startLocal, DateTime endLocal) =>
        $"[Start] < '{Date(endLocal)}' AND [End] > '{Date(startLocal)}'";

    /// <summary>Stable ID for one occurrence of a (possibly recurring) appointment.</summary>
    public static string OccurrenceId(string entryId, DateTime startLocal) =>
        Classifier.Key(entryId, startLocal.ToString("yyyyMMddHHmm", CultureInfo.InvariantCulture));
}
