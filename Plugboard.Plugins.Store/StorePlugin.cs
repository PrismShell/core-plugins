using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Nodes;
using Plugboard.Contracts;

namespace Plugboard.Plugins.Store;

// In-memory shared state store - think a Redux/Zustand store, but it lives in the gateway
// and is shared by every local page/tool via HTTP. State is held in RAM (fast, volatile),
// organized as NAMESPACES (store "slices") -> KEY -> arbitrary JSON value.
//
// Ported from DeskHub-Gateway's StoreHandler, minus two gateway-only features:
//   - SSE subscribe: the plugin contract returns one value per request, no streaming.
//     Pages that need change-push should poll snapshot, or that feature stays gateway-side.
//   - SQLite write-through persistence: rehydration ran at startup with no request context;
//     a service that wants durable state should put it through the sqlite connector instead.
public sealed class StorePlugin : IPlugin
{
    public string Name => "store";

    // ns -> (key -> raw JSON text). Raw text so any JSON shape round-trips untouched.
    private static readonly ConcurrentDictionary<string, ConcurrentDictionary<string, string>> _store = new();

    public void Register(IEndpointRegistry r)
    {
        r.Map("POST", "svc/store/set", req => Run(() => Set(Parse(req.Body))), new RouteInfo(
            "Set a value",
            "Stores any JSON value (object/array/scalar) under ns/key, replacing what was there.",
            new { ns = "blotter", key = "filters", value = new { desk = "Agency MBS" } },
            new[]
            {
                new ParamInfo("ns", "string", true, "Namespace (store slice)."),
                new ParamInfo("key", "string", true, "Key within the namespace."),
                new ParamInfo("value", "object", true, "Any JSON value."),
            }));

        r.Map("POST", "svc/store/patch", req => Run(() => Patch(Parse(req.Body))), new RouteInfo(
            "Shallow-merge into a value",
            "Merges an object's top-level properties into the existing object (or a new one).",
            new { ns = "blotter", key = "filters", value = new { book = "ALL" } },
            new[]
            {
                new ParamInfo("ns", "string", true, "Namespace."),
                new ParamInfo("key", "string", true, "Key."),
                new ParamInfo("value", "object", true, "Object whose properties are merged in."),
            }));

        r.Map("POST", "svc/store/delete", req => Run(() => Delete(Parse(req.Body))), new RouteInfo(
            "Delete a key or namespace",
            "Removes one key, or the whole namespace if key is omitted.",
            new { ns = "blotter", key = "filters" },
            new[]
            {
                new ParamInfo("ns", "string", true, "Namespace."),
                new ParamInfo("key", "string", false, "Key; omit to clear the namespace."),
            }));

        r.Map("POST", "svc/store/get", req => Run(() => Get(Parse(req.Body))), new RouteInfo(
            "Get one value",
            null,
            new { ns = "blotter", key = "filters" },
            new[]
            {
                new ParamInfo("ns", "string", true, "Namespace."),
                new ParamInfo("key", "string", true, "Key."),
            }));

        r.Map("POST", "svc/store/select", req => Run(() => Select(Parse(req.Body))), new RouteInfo(
            "Read a subset of a large value",
            "Drill into a stored value without shipping the whole thing: path resolves a node (JSON-Pointer-ish, \"/rows/3/name\"); where filters an array of objects by exact field match; fields projects objects to just those keys; offset/limit window an array (response carries total for pagination).",
            new { ns = "ref", key = "pools", path = "rows", where = new { Desk = "Agency MBS" }, fields = new[] { "Cusip", "Coupon" }, limit = 50 },
            new[]
            {
                new ParamInfo("ns", "string", true, "Namespace."),
                new ParamInfo("key", "string", true, "Key."),
                new ParamInfo("path", "string", false, "Slash-separated path into the value (~1 => /, ~0 => ~)."),
                new ParamInfo("where", "object", false, "Exact-match filter on array elements: { field: value }."),
                new ParamInfo("fields", "array", false, "Project objects down to these keys.", Items: "string"),
                new ParamInfo("offset", "integer", false, "Array window start (after where)."),
                new ParamInfo("limit", "integer", false, "Array window size."),
            }));

        r.Map("POST", "svc/store/snapshot", req => Run(() => Snapshot(Parse(req.Body))), new RouteInfo(
            "Read a whole namespace",
            "The entire store slice as one { key: value } object.",
            new { ns = "blotter" },
            new[] { new ParamInfo("ns", "string", true, "Namespace.") }));

        r.Map("POST", "svc/store/keys", req => Run(() => Keys(Parse(req.Body))), new RouteInfo(
            "List keys in a namespace",
            null,
            new { ns = "blotter" },
            new[] { new ParamInfo("ns", "string", true, "Namespace.") }));

        r.Map("GET", "svc/store/stats", _ => Run(Stats), new RouteInfo(
            "Store statistics",
            "Namespaces with key counts."));
    }

    private static Task<object?> Run(Func<object?> f) => Task.FromResult(f());

    // ---------------------------------------------------------------- writes

    private static object? Set(JsonElement body)
    {
        var ns = ReqStr(body, "ns"); var key = ReqStr(body, "key");
        if (!body.TryGetProperty("value", out var val)) throw new Exception("value is required");

        _store.GetOrAdd(ns, _ => new())[key] = val.GetRawText();
        return new { ns, key };
    }

    private static object? Patch(JsonElement body)
    {
        var ns = ReqStr(body, "ns"); var key = ReqStr(body, "key");
        if (!body.TryGetProperty("value", out var val) || val.ValueKind != JsonValueKind.Object)
            throw new Exception("value must be an object to patch");

        var slice = _store.GetOrAdd(ns, _ => new());
        JsonObject merged;
        if (slice.TryGetValue(key, out var existing) &&
            JsonNode.Parse(existing) is JsonObject cur) merged = cur;
        else merged = new JsonObject();

        foreach (var p in val.EnumerateObject())
            merged[p.Name] = JsonNode.Parse(p.Value.GetRawText());

        var json = merged.ToJsonString();
        slice[key] = json;
        return new { ns, key, value = JsonSerializer.Deserialize<JsonElement>(json) };
    }

    private static object? Delete(JsonElement body)
    {
        var ns = ReqStr(body, "ns"); var key = Str(body, "key");

        if (string.IsNullOrEmpty(key))
        {
            var cleared = _store.TryRemove(ns, out var removed) ? removed.Count : 0;
            return new { ns, cleared };
        }

        var deleted = _store.TryGetValue(ns, out var slice) && slice.TryRemove(key!, out _);
        return new { ns, key, deleted };
    }

    // ---------------------------------------------------------------- reads

    private static object? Get(JsonElement body)
    {
        var ns = ReqStr(body, "ns"); var key = ReqStr(body, "key");
        if (_store.TryGetValue(ns, out var slice) && slice.TryGetValue(key, out var json))
            return new { ns, key, value = JsonSerializer.Deserialize<JsonElement>(json) };
        throw new Exception("not found");
    }

    // Pull a SUBSET of a (possibly large) stored value WITHOUT shipping the whole thing.
    // The whole value is parsed in RAM here, but only the resulting subset is serialized back.
    private static object? Select(JsonElement body)
    {
        var ns = ReqStr(body, "ns"); var key = ReqStr(body, "key");
        if (!(_store.TryGetValue(ns, out var slice) && slice.TryGetValue(key, out var json)))
            throw new Exception("not found");

        JsonNode? node;
        try { node = JsonNode.Parse(json); }
        catch (Exception ex) { throw new Exception("stored value is not valid JSON: " + ex.Message); }

        var path = Str(body, "path") ?? "";
        if (path != "")
        {
            node = Navigate(node, path);
            if (node == null) throw new Exception($"path '{path}' not found");
        }

        var fields = StrArray(body, "fields");
        Dictionary<string, string?>? where = null;
        if (body.TryGetProperty("where", out var w) && w.ValueKind == JsonValueKind.Object)
        {
            where = new();
            foreach (var p in w.EnumerateObject())
                where[p.Name] = p.Value.ValueKind == JsonValueKind.String ? p.Value.GetString() : p.Value.GetRawText();
        }
        int? offset = IntOrNull(body, "offset");
        int? limit = IntOrNull(body, "limit");

        // Array: where -> project -> window (with total = pre-window match count).
        if (node is JsonArray arr)
        {
            IEnumerable<JsonNode?> seq = arr;
            if (where is { Count: > 0 })
                seq = seq.Where(el => el is JsonObject o && where.All(kv =>
                {
                    var v = o[kv.Key];
                    var s = v == null ? null : (v is JsonValue jv && jv.TryGetValue<string>(out var sv) ? sv : v.ToJsonString());
                    return s == kv.Value;
                }));
            var list = seq.ToList();
            int total = list.Count;
            int off = Math.Max(0, offset ?? 0);
            if (off > 0) list = list.Skip(off).ToList();
            if (limit is >= 0) list = list.Take(limit.Value).ToList();

            var outArr = new JsonArray();
            foreach (var el in list) outArr.Add(Project(el, fields));
            return new { ns, key, path, total, offset = off, limit,
                count = outArr.Count, value = JsonSerializer.Deserialize<JsonElement>(outArr.ToJsonString()) };
        }

        // Object / scalar: optional field projection.
        var outNode = Project(node, fields);
        return new { ns, key, path,
            value = outNode == null ? (object?)null : JsonSerializer.Deserialize<JsonElement>(outNode.ToJsonString()) };
    }

    private static object? Snapshot(JsonElement body)
    {
        var ns = ReqStr(body, "ns");
        var obj = new JsonObject();
        if (_store.TryGetValue(ns, out var slice))
            foreach (var kv in slice) obj[kv.Key] = JsonNode.Parse(kv.Value);
        return new { ns, count = obj.Count, state = JsonSerializer.Deserialize<JsonElement>(obj.ToJsonString()) };
    }

    private static object? Keys(JsonElement body)
    {
        var ns = ReqStr(body, "ns");
        var keys = _store.TryGetValue(ns, out var slice)
            ? slice.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray()
            : Array.Empty<string>();
        return new { ns, count = keys.Length, keys };
    }

    private static object? Stats() => new
    {
        namespaces = _store.ToDictionary(kv => kv.Key, kv => new { keys = kv.Value.Count }),
        totalKeys = _store.Sum(kv => kv.Value.Count),
    };

    // ---------------------------------------------------------------- helpers

    // Walk a JSON-Pointer-ish path (slash-separated; ~1=>/ ~0=>~) into object keys / array indices.
    private static JsonNode? Navigate(JsonNode? node, string path)
    {
        foreach (var raw in path.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            var tok = raw.Replace("~1", "/").Replace("~0", "~");
            if (node is JsonObject o) { if (!o.TryGetPropertyValue(tok, out node)) return null; }
            else if (node is JsonArray a) { if (!int.TryParse(tok, out var i) || i < 0 || i >= a.Count) return null; node = a[i]; }
            else return null;
        }
        return node;
    }

    // Deep-clone a node (so it can be re-parented), optionally projecting an object to `fields`.
    private static JsonNode? Project(JsonNode? node, string[]? fields)
    {
        if (node == null) return null;
        if (fields == null || fields.Length == 0) return node.DeepClone();
        if (node is JsonObject o)
        {
            var res = new JsonObject();
            foreach (var f in fields) if (o.TryGetPropertyValue(f, out var v)) res[f] = v?.DeepClone();
            return res;
        }
        return node.DeepClone();
    }

    private static string[]? StrArray(JsonElement b, string name)
    {
        if (b.ValueKind != JsonValueKind.Object || !b.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.Array)
            return null;
        return v.EnumerateArray().Where(x => x.ValueKind == JsonValueKind.String).Select(x => x.GetString()!).ToArray();
    }

    private static int? IntOrNull(JsonElement b, string name)
    {
        if (b.ValueKind != JsonValueKind.Object || !b.TryGetProperty(name, out var v)) return null;
        if (v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var n)) return n;
        if (v.ValueKind == JsonValueKind.String && int.TryParse(v.GetString(), out var m)) return m;
        return null;
    }

    private static JsonElement Parse(string body) =>
        JsonSerializer.Deserialize<JsonElement>(string.IsNullOrWhiteSpace(body) ? "{}" : body);

    private static string ReqStr(JsonElement b, string name) =>
        Str(b, name) is { Length: > 0 } v ? v : throw new Exception($"'{name}' is required.");

    private static string? Str(JsonElement b, string name) =>
        b.ValueKind == JsonValueKind.Object && b.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString() : null;
}
