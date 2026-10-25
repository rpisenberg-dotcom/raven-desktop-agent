namespace RavenAgent.Core;

/// <summary>One look at the screen: which window was in front, and whether Raven was active.</summary>
public readonly record struct Sample(DateTime Utc, string Process, string Title, bool Idle, bool Ignored, string? Url = null);

/// <summary>
/// Merges a stream of samples (taken every few seconds) into time segments:
/// same app + same window title = one segment. Idle, away and ignored time is not recorded.
/// Long segments are split so data reaches the app regularly.
/// </summary>
public sealed class SegmentBuilder
{
    readonly string _deviceId;
    readonly TimeSpan _interval;
    readonly TimeSpan _minLength;
    readonly TimeSpan _maxLength;

    string? _process, _title, _url;
    DateTime _start, _last;

    public SegmentBuilder(string deviceId, TimeSpan sampleInterval, TimeSpan? minLength = null, TimeSpan? maxLength = null)
    {
        _deviceId = deviceId;
        _interval = sampleInterval;
        _minLength = minLength ?? TimeSpan.FromSeconds(20);
        _maxLength = maxLength ?? TimeSpan.FromMinutes(15);
    }

    public bool HasOpenSegment => _process != null;

    /// <summary>Feed one sample; returns any segments that were completed by it.</summary>
    public IReadOnlyList<TimeLogRecord> Add(Sample s)
    {
        var done = new List<TimeLogRecord>();
        var skip = s.Idle || s.Ignored || Classifier.IsAway(s.Process);

        if (_process != null)
        {
            var gap = s.Utc - _last;
            var same = !skip && s.Process == _process && s.Title == _title && gap <= _interval * 3;
            if (same)
            {
                _last = s.Utc;
                _url ??= s.Url;
                if (_last - _start >= _maxLength)
                {
                    Close(done, _last);
                    Open(s);
                }
                return done;
            }
            // The previous window was in front until this sample (or one interval, if there was a gap)
            var end = gap <= _interval * 3 ? s.Utc : _last + _interval;
            Close(done, end);
        }

        if (!skip) Open(s);
        return done;
    }

    /// <summary>Close whatever is open (e.g. on pause or exit).</summary>
    public IReadOnlyList<TimeLogRecord> Flush(DateTime nowUtc)
    {
        var done = new List<TimeLogRecord>();
        if (_process != null)
        {
            var end = nowUtc - _last <= _interval * 3 ? nowUtc : _last + _interval;
            Close(done, end);
        }
        return done;
    }

    void Open(Sample s)
    {
        _process = s.Process;
        _title = s.Title;
        _url = s.Url;
        _start = s.Utc;
        _last = s.Utc;
    }

    void Close(List<TimeLogRecord> done, DateTime end)
    {
        var length = end - _start;
        if (length >= _minLength && _process != null)
        {
            done.Add(new TimeLogRecord
            {
                AgentKey = Classifier.Key(_deviceId, _start.Ticks.ToString(), _process),
                StartTime = Json.Iso(_start),
                EndTime = Json.Iso(end),
                DurationMinutes = Math.Round(length.TotalMinutes, 2),
                ActivityType = Classifier.ActivityType(_process, _title ?? ""),
                App = _process,
                WindowTitle = Truncate(_title, 400),
                Url = _url,
            });
        }
        _process = null;
        _title = null;
        _url = null;
    }

    static string? Truncate(string? s, int max) => s == null || s.Length <= max ? s : s[..max];
}
