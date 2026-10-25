using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>Sends batches to the Clerical Guru agentIngest function.</summary>
sealed class IngestClient : IDisposable
{
    readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(60) };

    public sealed record SendResult(bool Ok, string Message);

    public async Task<SendResult> SendAsync(string url, string key, IngestPayload payload, CancellationToken ct = default)
    {
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Post, url);
            req.Headers.Add("x-agent-key", key);
            req.Headers.UserAgent.Add(new ProductInfoHeaderValue("RavenAgent", payload.AgentVersion));
            req.Content = new StringContent(JsonSerializer.Serialize(payload, Json.Options), Encoding.UTF8, "application/json");
            using var res = await _http.SendAsync(req, ct);
            var body = await res.Content.ReadAsStringAsync(ct);
            if (res.IsSuccessStatusCode) return new(true, body);
            var reason = (int)res.StatusCode switch
            {
                401 or 403 => "The agent key was rejected - check AGENT_INGEST_KEY in Base44 and in Settings",
                404 => "The ingest URL wasn't found - check the address in Settings",
                _ => $"Server returned {(int)res.StatusCode}",
            };
            return new(false, $"{reason}. {Trim(body)}");
        }
        catch (TaskCanceledException) { return new(false, "Timed out reaching Clerical Guru"); }
        catch (HttpRequestException ex) { return new(false, $"Couldn't reach Clerical Guru: {ex.Message}"); }
    }

    /// <summary>The server's reply includes the current client list (names, aliases, domains) for on-PC matching.</summary>
    public static List<ClientEntry>? ParseClientDirectory(string body)
    {
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("client_directory", out var dir) || dir.ValueKind != JsonValueKind.Array) return null;
            return dir.Deserialize<List<ClientEntry>>(Json.Options);
        }
        catch { return null; }
    }

    static string Trim(string s) => s.Length > 200 ? s[..200] : s;

    public void Dispose() => _http.Dispose();
}
