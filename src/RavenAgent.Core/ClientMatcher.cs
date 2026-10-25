using System.Text.RegularExpressions;

namespace RavenAgent.Core;

/// <summary>One client as sent by the server: names/aliases to look for, plus email domains and addresses.</summary>
public sealed record ClientEntry
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public List<string> Terms { get; init; } = new();
    public List<string> Domains { get; init; } = new();
    public List<string> Emails { get; init; } = new();
}

/// <summary>Best client found in some text or addresses.</summary>
public sealed record ClientMatch(string ClientId, string ClientName, string Term, int Count);

/// <summary>
/// Finds which client a piece of on-screen text or an email is about, entirely on this PC.
/// Only the result (client id + the word that matched) is ever sent to the server.
/// </summary>
public sealed class ClientMatcher
{
    sealed record Pattern(ClientEntry Client, string Term, Regex Regex);

    readonly List<Pattern> _patterns = new();
    readonly Dictionary<string, ClientEntry> _byDomain = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, ClientEntry> _byEmail = new(StringComparer.OrdinalIgnoreCase);

    public int ClientCount { get; }

    public ClientMatcher(IEnumerable<ClientEntry> clients)
    {
        var list = clients.ToList();
        ClientCount = list.Count;
        foreach (var c in list)
        {
            var terms = new[] { c.Name }.Concat(c.Terms)
                .Select(t => (t ?? "").Trim())
                .Where(t => t.Length >= 3)
                .Distinct(StringComparer.OrdinalIgnoreCase);
            foreach (var t in terms)
            {
                // Whole words only, so "ABC" doesn't match "ABCD"; spaces/underscores/dashes are interchangeable
                // (file names like Koplin_EMC_Audit.xlsx)
                var body = string.Join(@"[\s_\-\.]*", t.Split(new[] { ' ', '_', '-' }, StringSplitOptions.RemoveEmptyEntries).Select(Regex.Escape));
                _patterns.Add(new Pattern(c, t, new Regex($@"(?<![\p{{L}}\p{{N}}]){body}(?![\p{{L}}\p{{N}}])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)));
            }
            foreach (var d in c.Domains.Where(d => !string.IsNullOrWhiteSpace(d))) _byDomain[d.Trim().TrimStart('@')] = c;
            foreach (var e in c.Emails.Where(e => !string.IsNullOrWhiteSpace(e))) _byEmail[e.Trim()] = c;
        }
        // Longer names first so "Koplin Construction" wins over "Koplin" for the reported term
        _patterns.Sort((a, b) => b.Term.Length.CompareTo(a.Term.Length));
    }

    /// <summary>
    /// The client mentioned most often in the text (ties go to the longer, more specific name).
    /// With <paramref name="requireClearWinner"/>, pages that show several clients (order lists,
    /// customer dropdowns, reports) return no match unless one client clearly dominates, so the
    /// guess doesn't flip back and forth on every re-read.
    /// </summary>
    public ClientMatch? MatchText(string? text, bool requireClearWinner = false)
    {
        if (string.IsNullOrWhiteSpace(text) || _patterns.Count == 0) return null;
        var counts = new Dictionary<string, (ClientEntry c, string term, int n)>();
        // Patterns run longest-first; a shorter alias inside an already-counted longer name
        // ("Fabcast" inside "Fabcast Solutions") is the same mention and isn't counted again
        var taken = new Dictionary<string, List<(int start, int end)>>();
        foreach (var p in _patterns)
        {
            if (!taken.TryGetValue(p.Client.Id, out var spans)) taken[p.Client.Id] = spans = new();
            var n = 0;
            foreach (Match m in p.Regex.Matches(text))
            {
                if (spans.Any(sp => m.Index < sp.end && m.Index + m.Length > sp.start)) continue;
                spans.Add((m.Index, m.Index + m.Length));
                n++;
            }
            if (n == 0) continue;
            if (counts.TryGetValue(p.Client.Id, out var cur))
                counts[p.Client.Id] = (cur.c, cur.term, cur.n + n); // keep the longest term (seen first)
            else
                counts[p.Client.Id] = (p.Client, p.Term, n);
        }
        if (counts.Count == 0) return null;
        var ranked = counts.Values.OrderByDescending(v => v.n).ThenByDescending(v => v.term.Length).ToList();
        var best = ranked[0];
        if (requireClearWinner && ranked.Count > 1 && best.n < 2 * ranked[1].n) return null;
        return new ClientMatch(best.c.Id, best.c.Name, best.term, best.n);
    }

    /// <summary>A client whose email address or domain appears among the given addresses.</summary>
    public ClientMatch? MatchAddresses(IEnumerable<string?> addresses)
    {
        var hits = new Dictionary<string, (ClientEntry c, string term, int n)>();
        foreach (var raw in addresses)
        {
            var a = (raw ?? "").Trim().Trim('<', '>');
            if (a.Length == 0) continue;
            ClientEntry? c = null;
            string term = a;
            if (_byEmail.TryGetValue(a, out var byEmail)) c = byEmail;
            else
            {
                var at = a.LastIndexOf('@');
                if (at > 0)
                {
                    var domain = a[(at + 1)..];
                    // also match subdomains, e.g. mail.koplin.com -> koplin.com
                    for (var d = domain; c == null && d.Contains('.'); d = d[(d.IndexOf('.') + 1)..])
                        if (_byDomain.TryGetValue(d, out var byDomain)) { c = byDomain; term = "@" + d; }
                }
            }
            if (c == null) continue;
            hits[c.Id] = hits.TryGetValue(c.Id, out var cur) ? (c, cur.term, cur.n + 1) : (c, term, 1);
        }
        if (hits.Count == 0) return null;
        var best = hits.Values.OrderByDescending(v => v.n).First();
        return new ClientMatch(best.c.Id, best.c.Name, best.term, best.n);
    }
}
