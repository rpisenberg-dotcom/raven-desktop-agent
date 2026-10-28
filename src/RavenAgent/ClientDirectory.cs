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
        try
        {
            if (File.Exists(ContactsPath))
                _business = (JsonSerializer.Deserialize<List<string>>(File.ReadAllText(ContactsPath)) ?? new()).Select(ContactKey).ToHashSet();
        }
        catch { }
    }

    public ClientMatcher Matcher { get { lock (_lock) return _matcher; } }

    // Phone Link contacts assigned to a client or Clerical Guru (names and phone numbers), sent by the app
    HashSet<string> _business = new();
    static string ContactsPath => Path.Combine(AgentConfig.DataFolder, "business-contacts.json");
    public static string ContactKey(string s)
    {
        var digits = new string(s.Where(char.IsDigit).ToArray());
        if (digits.Length >= 7) return digits.Length > 10 ? digits[^10..] : digits;
        return System.Text.RegularExpressions.Regex.Replace(s.ToLowerInvariant(), "[^a-z0-9]+", " ").Trim();
    }
    public bool IsBusinessContact(string contact) { lock (_lock) return _business.Contains(ContactKey(contact)); }
    public void UpdateBusinessContacts(List<string> contacts)
    {
        var set = contacts.Select(ContactKey).Where(k => k.Length > 0).ToHashSet();
        lock (_lock) { if (set.SetEquals(_business)) return; _business = set; }
        try { File.WriteAllText(ContactsPath, JsonSerializer.Serialize(contacts)); } catch { }
    }
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
