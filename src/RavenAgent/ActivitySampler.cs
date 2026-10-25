using RavenAgent.Core;

namespace RavenAgent;

/// <summary>Looks at the active window every few seconds and turns that into time segments.</summary>
sealed class ActivitySampler
{
    readonly AgentConfig _cfg;
    readonly UploadQueue _queue;
    readonly SegmentBuilder _builder;

    public string CurrentWindow { get; private set; } = "";

    public ActivitySampler(AgentConfig cfg, UploadQueue queue)
    {
        _cfg = cfg;
        _queue = queue;
        _builder = new SegmentBuilder(cfg.DeviceId, TimeSpan.FromSeconds(cfg.SampleSeconds));
    }

    /// <summary>Called on a timer from the UI thread.</summary>
    public void Tick()
    {
        try
        {
            var now = DateTime.UtcNow;
            if (_cfg.PrivateMode)
            {
                Enqueue(_builder.Flush(now));
                CurrentWindow = "Paused";
                return;
            }

            var (process, title) = Native.ForegroundWindow();
            // Listening in a meeting isn't idle, even without typing
            var idle = Native.IdleTime() >= TimeSpan.FromMinutes(_cfg.IdleMinutes) && !Classifier.IsMeeting(process, title);
            var ignored = Classifier.IsIgnored(process, title, _cfg.IgnoreApps, _cfg.IgnoreTitleWords);

            CurrentWindow = idle ? "Idle" : ignored ? "Private window" : $"{process}: {title}";
            Enqueue(_builder.Add(new Sample(now, process, title, idle, ignored)));
        }
        catch (Exception ex) { Log.Error("Sampling failed", ex); }
    }

    public void Flush() => Enqueue(_builder.Flush(DateTime.UtcNow));

    void Enqueue(IReadOnlyList<TimeLogRecord> logs)
    {
        if (logs.Count > 0) _queue.Add(logs);
    }
}
