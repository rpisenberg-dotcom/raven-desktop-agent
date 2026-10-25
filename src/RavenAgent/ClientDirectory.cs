using System.Text.Json;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// Client names, aliases and email domains, sent by Clerical Guru with every upload reply
/// and cached on this PC so matching works offline.
/// </summary>
sealed class ClientDirectory
{
    static string FilePath => Path.Combine(AgentConfig.DataFolder, "clients.json");
    readonly object _lock = new();
    ClientMatcher _matcher = new(Array.Empty<ClientEntry>());
    string _signature = "";

    public ClientDirectory()
    {
        try
        {
            if (File.Exists(FilePath))
                Update(JsonSerializer.Deserialize<List<ClientEntry>>(File.ReadAllText(FilePath), Json.Options) ?? new());
        }
        catch (Exception ex) { Log.Error("Couldn't load cached client list", ex); }
    }

    public ClientMatcher Matcher { get { lock (_lock) return _matcher; } }
    public int Count => Matcher.ClientCount;

    public void Update(List<ClientEntry> clients)
    {
        var json = JsonSerializer.Serialize(clients, Json.Options);
        lock (_lock)
        {
            if (json == _signature) return;
            _signature = json;
            _matcher = new ClientMatcher(clients);
        }
        try
        {
            Directory.CreateDirectory(AgentConfig.DataFolder);
            File.WriteAllText(FilePath, json);
        }
        catch { }
        Log.Info($"Client list updated: {clients.Count} clients");
    }
}
