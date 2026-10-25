using System.Text.Json;

namespace RavenAgent.Core;

/// <summary>
/// Holds time logs and emails until the server confirms it has them.
/// Saved to disk after every change, so nothing is lost if the PC restarts or Wi-Fi drops.
/// </summary>
public sealed class UploadQueue
{
    sealed class State
    {
        public List<TimeLogRecord> TimeLogs { get; set; } = new();
        public List<EmailRecord> Emails { get; set; } = new();
    }

    readonly string _path;
    readonly object _lock = new();
    readonly int _maxItems;
    State _state = new();

    public UploadQueue(string path, int maxItems = 20000)
    {
        _path = path;
        _maxItems = maxItems;
        Load();
    }

    public int Count { get { lock (_lock) return _state.TimeLogs.Count + _state.Emails.Count; } }

    public void Add(IEnumerable<TimeLogRecord> logs) { lock (_lock) { _state.TimeLogs.AddRange(logs); Trim(); Save(); } }
    public void Add(IEnumerable<EmailRecord> emails) { lock (_lock) { _state.Emails.AddRange(emails); Trim(); Save(); } }

    /// <summary>Takes up to <paramref name="max"/> of each kind without removing them.</summary>
    public (List<TimeLogRecord> logs, List<EmailRecord> emails) Peek(int max)
    {
        lock (_lock) return (_state.TimeLogs.Take(max).ToList(), _state.Emails.Take(max).ToList());
    }

    /// <summary>Removes items the server accepted.</summary>
    public void Remove(IReadOnlyCollection<TimeLogRecord> logs, IReadOnlyCollection<EmailRecord> emails)
    {
        lock (_lock)
        {
            var logKeys = logs.Select(l => l.AgentKey).ToHashSet();
            var mailKeys = emails.Select(e => e.ExternalId).ToHashSet();
            _state.TimeLogs.RemoveAll(l => logKeys.Contains(l.AgentKey));
            _state.Emails.RemoveAll(e => mailKeys.Contains(e.ExternalId));
            Save();
        }
    }

    void Trim()
    {
        // Never grow without limit if the server is unreachable for weeks: drop the oldest emails first
        while (_state.TimeLogs.Count + _state.Emails.Count > _maxItems && _state.Emails.Count > 0) _state.Emails.RemoveAt(0);
        while (_state.TimeLogs.Count > _maxItems) _state.TimeLogs.RemoveAt(0);
    }

    void Load()
    {
        try
        {
            if (File.Exists(_path))
                _state = JsonSerializer.Deserialize<State>(File.ReadAllText(_path), Json.Options) ?? new State();
        }
        catch
        {
            // A damaged queue file shouldn't stop the agent; keep a copy and start fresh
            try { File.Copy(_path, _path + ".bad", true); } catch { }
            _state = new State();
        }
    }

    void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        var tmp = _path + ".tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(_state, Json.Options));
        File.Move(tmp, _path, true);
    }
}
