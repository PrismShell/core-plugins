using System.Collections.Concurrent;
using System.Text.Json;
using Plugboard.Contracts;

namespace Plugboard.Plugins.MarketCache;

// In-memory cache for vendor market data. Ported from DeskHub-Gateway's MarketCacheHandler.
//
// Why memory and not the sqlite store: that store is an append-only changelog that
// sqlite/publish exports WHOLESALE to a file share for other desks to sync - it does not
// filter by dataset. Market data that never enters it therefore cannot be redistributed by
// accident, which is a structural guarantee rather than a rule someone has to remember. It
// also does not survive a restart, which is the right lifetime for a price.
//
// One TTL for everything (PLUGBOARD_CACHE_TTL_SECONDS, default 300). An earlier version
// tried to tier fields (prices short, reference data long) and that was a bad idea: BBG has
// tens of thousands of fields, so any hand-written list covers a rounding error of them
// while looking authoritative, and a misclassified field quietly serves hours-old data. A
// caller that genuinely knows better can pass ttlSeconds per request.
//
// Timestamps are kept per FIELD rather than per security, which is not tiering - it is just
// correct. Fetching PX_BID at 10:04 must not make a NAME cached at 10:00 look freshly
// retrieved, and a request for both should be able to reuse the fresh half.
//
// The intended use is cache-through composition: a service req.Calls cache/get first, goes
// live (e.g. bbg/bdp) only for the fields that came back missing, then cache/puts the fresh
// values.
public sealed class MarketCachePlugin : IPlugin
{
    public string Name => "cache";

    private sealed record Cell(string? Value, DateTime StoredUtc);

    // dataset -> id -> field -> cell
    private static readonly ConcurrentDictionary<string,
        ConcurrentDictionary<string, ConcurrentDictionary<string, Cell>>> _sets = new();

    // Short by default: the cache exists to stop reruns and accidental refreshes firing the
    // same request repeatedly while building a sheet, not to keep data for the day.
    private static readonly int TtlSeconds =
        int.TryParse(Environment.GetEnvironmentVariable("PLUGBOARD_CACHE_TTL_SECONDS"), out var t) && t > 0 ? t : 300;
    private static Timer? _sweeper;
    private static long _hits, _misses, _stores, _evicted;

    public void Register(IEndpointRegistry r)
    {
        // A sweeper, not just lazy expiry on read: a security nobody asks about again would
        // otherwise sit in memory until the process ends.
        var everyMs = Math.Max(30, TtlSeconds / 2) * 1000;
        _sweeper ??= new Timer(_ => Sweep(), null, everyMs, everyMs);

        r.Map("POST", "svc/cache/get", req => Run(() => Get(Parse(req.Body))), new RouteInfo(
            "Read fresh fields",
            "Returns only the requested fields that are still fresh, each with its own age. Whatever is missing from the result is the caller's cue to go live for those.",
            new { dataset = "bbg", id = "AAPL US Equity", fields = new[] { "PX_LAST", "NAME" } },
            new[]
            {
                new ParamInfo("dataset", "string", true, "Cache partition, e.g. \"bbg\"."),
                new ParamInfo("id", "string", true, "Row id, e.g. a security identifier."),
                new ParamInfo("fields", "array", true, "Fields to look up.", Items: "string"),
                new ParamInfo("ttlSeconds", "integer", false, "Per-request freshness override."),
            }));

        r.Map("POST", "svc/cache/put", req => Run(() => Put(Parse(req.Body))), new RouteInfo(
            "Store field values",
            "Each field gets its own timestamp; a later put of one field never refreshes its neighbours.",
            new { dataset = "bbg", id = "AAPL US Equity", fields = new { PX_LAST = "227.16", NAME = "APPLE INC" } },
            new[]
            {
                new ParamInfo("dataset", "string", true, "Cache partition."),
                new ParamInfo("id", "string", true, "Row id."),
                new ParamInfo("fields", "object", true, "Field name -> value (values stored as strings; null allowed)."),
            }));

        r.Map("GET", "svc/cache/stats", _ => Run(Stats), new RouteInfo(
            "Cache statistics",
            "Datasets with row/value counts plus hit/miss/store/evict counters."));

        r.Map("POST", "svc/cache/clear", req => Run(() => Clear(Parse(req.Body))), new RouteInfo(
            "Clear the cache",
            "Clears one dataset, or everything if dataset is omitted.",
            new { dataset = "bbg" },
            new[] { new ParamInfo("dataset", "string", false, "Dataset to clear; omit for all.") }));
    }

    private static Task<object?> Run(Func<object?> f) => Task.FromResult(f());

    // ---------------------------------------------------------------- endpoints

    private static object? Get(JsonElement b)
    {
        var dataset = ReqStr(b, "dataset");
        var id = ReqStr(b, "id");
        if (!b.TryGetProperty("fields", out var fs) || fs.ValueKind != JsonValueKind.Array)
            throw new Exception("fields array is required");
        var fields = fs.EnumerateArray().Where(x => x.ValueKind == JsonValueKind.String)
                       .Select(x => x.GetString()!).ToList();
        int? ttlOverride = b.TryGetProperty("ttlSeconds", out var tv) && tv.TryGetInt32(out var t) && t > 0 ? t : null;

        var ttl = ttlOverride ?? TtlSeconds;
        var found = new Dictionary<string, object>();
        if (!_sets.TryGetValue(dataset, out var set) || !set.TryGetValue(id, out var row))
        {
            Interlocked.Increment(ref _misses);
            return new { dataset, id, found, missing = fields };
        }
        var now = DateTime.UtcNow;
        var missing = new List<string>();
        foreach (var f in fields)
        {
            if (!row.TryGetValue(f, out var cell)) { Interlocked.Increment(ref _misses); missing.Add(f); continue; }
            var age = (int)(now - cell.StoredUtc).TotalSeconds;
            if (age > ttl)
            {
                row.TryRemove(f, out _);
                Interlocked.Increment(ref _evicted);
                Interlocked.Increment(ref _misses);
                missing.Add(f);
                continue;
            }
            found[f] = new { value = cell.Value, ageSeconds = age };
            Interlocked.Increment(ref _hits);
        }
        return new { dataset, id, found, missing };
    }

    private static object? Put(JsonElement b)
    {
        var dataset = ReqStr(b, "dataset");
        var id = ReqStr(b, "id");
        if (!b.TryGetProperty("fields", out var f) || f.ValueKind != JsonValueKind.Object)
            throw new Exception("fields object is required");

        var fields = new Dictionary<string, string?>();
        foreach (var p in f.EnumerateObject())
            fields[p.Name] = p.Value.ValueKind switch
            {
                JsonValueKind.String => p.Value.GetString(),
                JsonValueKind.Null => null,
                _ => p.Value.GetRawText(),
            };
        if (fields.Count == 0) throw new Exception("fields object is empty");

        var set = _sets.GetOrAdd(dataset, _ => new ConcurrentDictionary<string, ConcurrentDictionary<string, Cell>>());
        var row = set.GetOrAdd(id, _ => new ConcurrentDictionary<string, Cell>());
        var now = DateTime.UtcNow;
        foreach (var kv in fields) row[kv.Key] = new Cell(kv.Value, now);
        Interlocked.Increment(ref _stores);
        return new { dataset, id, stored = fields.Count };
    }

    private static object? Stats() => new
    {
        ttlSeconds = TtlSeconds,
        datasets = _sets.ToDictionary(kv => kv.Key, kv => new
        {
            rows = kv.Value.Count,
            values = kv.Value.Sum(r => r.Value.Count),
        }),
        hits = Interlocked.Read(ref _hits),
        misses = Interlocked.Read(ref _misses),
        stores = Interlocked.Read(ref _stores),
        evicted = Interlocked.Read(ref _evicted),
    };

    private static object? Clear(JsonElement b)
    {
        var dataset = Str(b, "dataset");
        int removed;
        if (string.IsNullOrWhiteSpace(dataset))
        {
            removed = _sets.Sum(kv => kv.Value.Sum(r => r.Value.Count));
            _sets.Clear();
        }
        else
        {
            removed = _sets.TryRemove(dataset!, out var set) ? set.Sum(r => r.Value.Count) : 0;
        }
        return new { clearedValues = removed, dataset = dataset ?? "(all)" };
    }

    private static void Sweep()
    {
        var now = DateTime.UtcNow;
        foreach (var (_, set) in _sets)
        {
            foreach (var (id, row) in set)
            {
                foreach (var (field, cell) in row)
                    if ((now - cell.StoredUtc).TotalSeconds > TtlSeconds && row.TryRemove(field, out _))
                        Interlocked.Increment(ref _evicted);
                if (row.IsEmpty) set.TryRemove(id, out _);
            }
        }
    }

    // ---------------------------------------------------------------- helpers

    private static JsonElement Parse(string body) =>
        JsonSerializer.Deserialize<JsonElement>(string.IsNullOrWhiteSpace(body) ? "{}" : body);

    private static string ReqStr(JsonElement b, string name) =>
        Str(b, name) is { Length: > 0 } v ? v : throw new Exception($"'{name}' is required.");

    private static string? Str(JsonElement b, string name) =>
        b.ValueKind == JsonValueKind.Object && b.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString() : null;
}
