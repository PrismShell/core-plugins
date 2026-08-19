using Microsoft.Data.Sqlite;
using System.Text.Json;
using Plugboard.Contracts;

namespace Plugboard.Plugins.Sqlite;

// Local SQLite "store" connector. Ported from DeskHub-Gateway's SqliteHandler.
//
// Caches ingested data PLUS user-written records into a LOCAL SQLite DB, and syncs with a
// file share using an append-only change-log with newest-wins conflict resolution.
//
// Purpose-built endpoints ONLY. There is deliberately NO arbitrary/raw SQL surface - that is
// what the sql connector is for (SQL Server), and this store must not grow one.
//
// Storage model (WAL-mode SQLite at %LOCALAPPDATA%\Plugboard\data\sqlite.db, overridable via
// PLUGBOARD_SQLITE_DB):
//   changelog(seq, dataset, record_id, field, value, usr, ts) - append-only; THIS machine's
//             own writes.
//   state(dataset, record_id, field, value, usr, ts, PK(dataset,record_id,field)) - the
//             materialized, merged "current" view.
//
// Sync model: each machine PUBLISHES its own changelog to the share as <usr>.changelog.db
// (per-user file => no cross-writer contention, built locally then copied). SYNC reads EVERY
// user's changelog file plus the local changelog and rebuilds `state` by keeping, per
// (dataset,record_id,field), the row with the max (ts, usr) - NEWEST WINS, usr breaks ties.
//
// put/batch/get/query accept an optional "db": an absolute path to a .db file. Without it
// everything targets the default database. A caller-supplied database gets the same two-table
// schema but is DELIBERATELY OUTSIDE publish/sync - a project-local file that a solution owns
// has no business in the desk-wide newest-wins merge.
public sealed class SqlitePlugin : IPlugin
{
    public string Name => "sqlite";

    private static readonly object _writeLock = new();   // SQLite allows one writer; serialize ours
    private static volatile string _lastPublish = "";
    private static volatile string _lastSync = "";
    private static readonly string Usr = Environment.UserName;
    // Reserved field used as a delete tombstone ("1" = deleted). Written with ts=now so the
    // deletion wins under newest-wins publish/sync; a put writes "0" to un-delete. Hidden from
    // query/get output.
    private const string DEL = "__deleted";

    private static readonly string DbPath = ResolveDbPath();
    private static readonly string ConnStr = new SqliteConnectionStringBuilder { DataSource = DbPath }.ToString();
    // Optional FALLBACK only: no invented default share - a placeholder folder that shares with
    // nobody while reporting success is worse than refusing. Unset means local-only.
    private static readonly string ShareRootFallback =
        Environment.GetEnvironmentVariable("PLUGBOARD_SQLITE_SHAREROOT") ?? "";

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _extConn = new();
    private static readonly string[] DbExt = { ".db", ".sqlite", ".sqlite3" };

    private static string ResolveDbPath()
    {
        var overridePath = Environment.GetEnvironmentVariable("PLUGBOARD_SQLITE_DB");
        if (!string.IsNullOrWhiteSpace(overridePath))
            return Path.GetFullPath(Environment.ExpandEnvironmentVariables(overridePath));
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                               "Plugboard", "data");
        return Path.Combine(dir, "sqlite.db");
    }

    public void Register(IEndpointRegistry r)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(DbPath)!);
        InitSchema(ConnStr);

        r.Map("POST", "sqlite/put", req => Run(() => Put(Parse(req.Body))), new RouteInfo(
            "Upsert one record",
            "Writes each field through the changelog and into state. Clears any prior delete tombstone.",
            new { dataset = "lists", id = "batch-42", fields = new { status = "open", owner = "pat" } },
            new[]
            {
                new ParamInfo("dataset", "string", true, "Logical table name."),
                new ParamInfo("id", "string", true, "Record id (unique within the dataset)."),
                new ParamInfo("fields", "object", true, "Field name -> value. Objects/arrays stored as raw JSON text."),
                new ParamInfo("db", "string", false, "Absolute path to an alternate .db file (outside publish/sync)."),
            }));

        r.Map("POST", "sqlite/batch", req => Run(() => Batch(Parse(req.Body))), new RouteInfo(
            "Apply puts/deletes in one transaction",
            "ops: [{op:\"put\", id, fields:{...}}, {op:\"delete\", id}]. Default op is put. Deletes write a tombstone that publishes/syncs like any other write; query/get hide tombstoned records.",
            new { dataset = "lists", ops = new object[] { new { op = "put", id = "a", fields = new { x = "1" } }, new { op = "delete", id = "b" } } },
            new[]
            {
                new ParamInfo("dataset", "string", true, "Logical table name."),
                new ParamInfo("ops", "array", true, "Operations to apply atomically.", Items: "object"),
                new ParamInfo("db", "string", false, "Absolute path to an alternate .db file."),
            }));

        r.Map("POST", "sqlite/get", req => Run(() => Get(Parse(req.Body))), new RouteInfo(
            "Read one record",
            "Returns { record, meta } where meta carries per-field { usr, ts }. Tombstoned records read as not found.",
            new { dataset = "lists", id = "batch-42" },
            new[]
            {
                new ParamInfo("dataset", "string", true, "Logical table name."),
                new ParamInfo("id", "string", true, "Record id."),
                new ParamInfo("db", "string", false, "Absolute path to an alternate .db file."),
            }));

        r.Map("POST", "sqlite/query", req => Run(() => Query(Parse(req.Body))), new RouteInfo(
            "List records in a dataset",
            "Optional exact-match where on field values. Returns { count, records } with each record's id inlined.",
            new { dataset = "lists", where = new { status = "open" } },
            new[]
            {
                new ParamInfo("dataset", "string", true, "Logical table name."),
                new ParamInfo("where", "object", false, "Exact-match filter: { field: value }."),
                new ParamInfo("db", "string", false, "Absolute path to an alternate .db file."),
            }));

        r.Map("POST", "sqlite/publish", req => Run(() =>
        {
            var b = Parse(req.Body);
            return Publish(ShareFor(b), ConnFor(b));
        }), new RouteInfo(
            "Publish this machine's changelog to the share",
            "Exports the changelog to <shareRoot>\\<user>.changelog.db (built in local temp, then copied - never opened over SMB). shareRoot from the body wins, else PLUGBOARD_SQLITE_SHAREROOT.",
            new { shareRoot = @"\\server\share\desk-sync" },
            new[]
            {
                new ParamInfo("shareRoot", "string", false, "UNC/absolute folder to publish into."),
                new ParamInfo("db", "string", false, "Absolute path to an alternate .db file."),
            }));

        r.Map("POST", "sqlite/sync", req => Run(() =>
        {
            var b = Parse(req.Body);
            return Sync(ShareFor(b), ConnFor(b));
        }), new RouteInfo(
            "Rebuild state from every published changelog",
            "Reads every <shareRoot>\\*.changelog.db plus the local changelog and rebuilds state with newest-wins per (dataset,record,field). Reports conflicts (multiple users, differing values).",
            new { shareRoot = @"\\server\share\desk-sync" },
            new[]
            {
                new ParamInfo("shareRoot", "string", false, "UNC/absolute folder to sync from."),
                new ParamInfo("db", "string", false, "Absolute path to an alternate .db file."),
            }));

        r.Map("POST", "sqlite/purge", req => Run(() => Purge(Parse(req.Body))), new RouteInfo(
            "Purge datasets by prefix",
            "Deletes every changelog + state row whose dataset starts with the prefix, then VACUUMs. Prefix-scoped so records outside the prefix are never in range.",
            new { prefix = "bbg." },
            new[] { new ParamInfo("prefix", "string", true, "Dataset prefix, e.g. \"bbg.\".") }));

        r.Map("GET", "sqlite/status", _ => Run(Status), new RouteInfo(
            "Store status",
            "Datasets with record counts, changelog size, last publish/sync, share reachability."));
    }

    private static Task<object?> Run(Func<object?> f) => Task.FromResult(f());

    // ---------------------------------------------------------------- schema / connections

    private static void InitSchema(string connStr)
    {
        lock (_writeLock)
        {
            using var con = new SqliteConnection(connStr);
            con.Open();
            Exec(con, "PRAGMA journal_mode=WAL;");
            Exec(con, "PRAGMA synchronous=NORMAL;");
            Exec(con, @"CREATE TABLE IF NOT EXISTS changelog(
                          seq INTEGER PRIMARY KEY AUTOINCREMENT,
                          dataset TEXT, record_id TEXT, field TEXT, value TEXT, usr TEXT, ts TEXT);");
            Exec(con, @"CREATE TABLE IF NOT EXISTS state(
                          dataset TEXT, record_id TEXT, field TEXT, value TEXT, usr TEXT, ts TEXT,
                          PRIMARY KEY(dataset,record_id,field));");
        }
    }

    private static string ConnFor(JsonElement body)
    {
        var raw = Str(body, "db");
        if (string.IsNullOrWhiteSpace(raw)) return ConnStr;
        return ConnForPath(raw!);
    }

    // Every response says which file answered it. "I thought I was looking at the other
    // database" is the failure this prevents, and it costs one string.
    private static string DbLabel(string connStr) =>
        connStr == ConnStr ? "default" : new SqliteConnectionStringBuilder(connStr).DataSource;

    // Where to publish to / sync from, for THIS call. Body "shareRoot" wins; otherwise the
    // env fallback; otherwise refuse plainly rather than inventing a local folder that
    // shares with nobody.
    private static string ShareFor(JsonElement body)
    {
        var raw = Str(body, "shareRoot");
        if (string.IsNullOrWhiteSpace(raw)) raw = ShareRootFallback;
        if (string.IsNullOrWhiteSpace(raw))
            throw new Exception(
                "no shareRoot: pass \"shareRoot\":\"\\\\\\\\server\\\\share\\\\folder\" with the call, "
                + "or set PLUGBOARD_SQLITE_SHAREROOT. Without one this store is local-only and there "
                + "is nowhere to publish to.");
        var expanded = Environment.ExpandEnvironmentVariables(raw!.Trim());
        // Rootedness is tested on the INPUT, not on GetFullPath's result. GetFullPath resolves
        // a relative path against the working directory and hands back something absolute, so
        // checking after the fact always passes and "some/folder" quietly becomes a folder
        // inside the install directory.
        if (!Path.IsPathRooted(expanded))
            throw new Exception($"shareRoot must be an absolute or UNC path (got '{raw!.Trim()}')");
        return Path.GetFullPath(expanded);
    }

    private static string ConnForPath(string raw)
    {
        var expanded = Environment.ExpandEnvironmentVariables(raw.Trim());
        if (!Path.IsPathRooted(expanded))
            throw new Exception($"db must be an absolute path (got '{raw.Trim()}')");

        var full = Path.GetFullPath(expanded);
        if (!DbExt.Contains(Path.GetExtension(full).ToLowerInvariant()))
            throw new Exception($"db must end in {string.Join(" / ", DbExt)}");

        // Same file as the default, spelled differently: treat it as the default rather than
        // opening a second connection to it (and thereby skipping publish/sync semantics).
        if (string.Equals(full, DbPath, StringComparison.OrdinalIgnoreCase)) return ConnStr;

        return _extConn.GetOrAdd(full, p =>
        {
            var dir = Path.GetDirectoryName(p);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            var cs = new SqliteConnectionStringBuilder { DataSource = p }.ToString();
            InitSchema(cs);              // create-if-missing, so a fresh path just works
            return cs;
        });
    }

    // ---------------------------------------------------------------- endpoints

    private static object? Put(JsonElement body)
    {
        var dataset = ReqStr(body, "dataset");
        var id = ReqStr(body, "id");
        if (!body.TryGetProperty("fields", out var f) || f.ValueKind != JsonValueKind.Object)
            throw new Exception("fields object is required");

        var fields = new Dictionary<string, string?>();
        foreach (var p in f.EnumerateObject()) fields[p.Name] = JsonToStr(p.Value);
        if (fields.Count == 0) throw new Exception("fields object is empty");

        var conn = ConnFor(body);
        var n = fields.Count;
        fields[DEL] = "0";   // a put clears any prior delete tombstone for this id
        WriteMany(dataset, new() { (id, fields) }, Usr, conn);
        return new { dataset, id, fields = n, db = DbLabel(conn) };
    }

    private static object? Batch(JsonElement body)
    {
        var dataset = ReqStr(body, "dataset");
        if (!body.TryGetProperty("ops", out var ops) || ops.ValueKind != JsonValueKind.Array)
            throw new Exception("ops array is required");

        var records = new List<(string id, IDictionary<string, string?> fields)>();
        int puts = 0, dels = 0, fieldCount = 0;
        foreach (var op in ops.EnumerateArray())
        {
            if (op.ValueKind != JsonValueKind.Object) throw new Exception("each op must be an object");
            var kind = (op.TryGetProperty("op", out var k) && k.ValueKind == JsonValueKind.String
                        ? k.GetString() : "put")?.Trim().ToLowerInvariant() ?? "put";
            var id = op.TryGetProperty("id", out var idv) && idv.ValueKind == JsonValueKind.String ? idv.GetString() ?? "" : "";
            if (string.IsNullOrWhiteSpace(id)) throw new Exception("each op needs a non-empty id");

            if (kind == "delete")
            {
                records.Add((id, new Dictionary<string, string?> { [DEL] = "1" }));
                dels++;
            }
            else if (kind == "put")
            {
                if (!op.TryGetProperty("fields", out var f) || f.ValueKind != JsonValueKind.Object)
                    throw new Exception($"put op for id '{id}' needs a fields object");
                var fields = new Dictionary<string, string?>();
                foreach (var p in f.EnumerateObject()) fields[p.Name] = JsonToStr(p.Value);
                if (fields.Count == 0) throw new Exception($"put op for id '{id}' has empty fields");
                fieldCount += fields.Count;
                fields[DEL] = "0";   // clear any prior tombstone
                records.Add((id, fields));
                puts++;
            }
            else throw new Exception($"unknown op '{kind}' for id '{id}' (use \"put\" or \"delete\")");
        }
        if (records.Count == 0) throw new Exception("ops is empty");

        var conn = ConnFor(body);
        WriteMany(dataset, records, Usr, conn);   // one transaction for the whole batch
        return new { dataset, put = puts, deleted = dels, fields = fieldCount, db = DbLabel(conn) };
    }

    private static object? Get(JsonElement body)
    {
        var dataset = ReqStr(body, "dataset");
        var id = ReqStr(body, "id");
        var connStr = ConnFor(body);

        var record = new Dictionary<string, string?>();
        var meta = new Dictionary<string, object>();
        using var con = new SqliteConnection(connStr);
        con.Open();
        using var c = con.CreateCommand();
        c.CommandText = "SELECT field,value,usr,ts FROM state WHERE dataset=@d AND record_id=@r";
        c.Parameters.AddWithValue("@d", dataset);
        c.Parameters.AddWithValue("@r", id);
        using var rd = c.ExecuteReader();
        while (rd.Read())
        {
            var field = rd.GetString(0);
            record[field] = rd.IsDBNull(1) ? null : rd.GetString(1);
            meta[field] = new { usr = rd.GetString(2), ts = rd.GetString(3) };
        }
        if (record.Count == 0) throw new Exception("not found");
        if (record.TryGetValue(DEL, out var del) && del == "1") throw new Exception("not found");   // tombstoned
        record.Remove(DEL); meta.Remove(DEL);
        return new { record, meta, db = DbLabel(connStr) };
    }

    private static object? Query(JsonElement body)
    {
        var dataset = ReqStr(body, "dataset");

        var where = new Dictionary<string, string?>();
        if (body.TryGetProperty("where", out var w) && w.ValueKind == JsonValueKind.Object)
            foreach (var p in w.EnumerateObject()) where[p.Name] = JsonToStr(p.Value);

        // Pull every field of every record in the dataset, then pivot in memory.
        var conn = ConnFor(body);
        var byId = new Dictionary<string, Dictionary<string, string?>>();
        using (var con = new SqliteConnection(conn))
        {
            con.Open();
            using var c = con.CreateCommand();
            c.CommandText = "SELECT record_id,field,value FROM state WHERE dataset=@d";
            c.Parameters.AddWithValue("@d", dataset);
            using var rd = c.ExecuteReader();
            while (rd.Read())
            {
                var rid = rd.GetString(0);
                if (!byId.TryGetValue(rid, out var m)) byId[rid] = m = new();
                m[rd.GetString(1)] = rd.IsDBNull(2) ? null : rd.GetString(2);
            }
        }

        var records = new List<Dictionary<string, object?>>();
        foreach (var kv in byId.OrderBy(k => k.Key, StringComparer.Ordinal))
        {
            if (kv.Value.TryGetValue(DEL, out var del) && del == "1") continue;   // tombstoned - hidden
            if (!where.All(f => kv.Value.TryGetValue(f.Key, out var v) && v == f.Value)) continue;
            var rec = new Dictionary<string, object?> { ["id"] = kv.Key };
            foreach (var f in kv.Value) if (f.Key != "id" && f.Key != DEL) rec[f.Key] = f.Value;
            records.Add(rec);
        }
        return new { count = records.Count, records, db = DbLabel(conn) };
    }

    private static object? Publish(string shareRoot, string connStr)
    {
        try { if (!Directory.Exists(shareRoot)) Directory.CreateDirectory(shareRoot); }
        catch (Exception ex) { throw new Exception($"share not reachable: {ex.Message}"); }

        var target = Path.Combine(shareRoot, $"{Usr}.changelog.db");
        // Build it in LOCAL temp, never in the share: the whole design is "build here, copy
        // there, never open a .db over SMB", and nothing extra ever lands in the shared
        // folder alongside the published file.
        var tmp = Path.Combine(Path.GetTempPath(), $"plugboard-{Usr}-changelog-{Guid.NewGuid():N}.db");
        int rows;

        lock (_writeLock)
        {
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            var tmpConn = new SqliteConnectionStringBuilder { DataSource = tmp }.ToString();
            using (var dst = new SqliteConnection(tmpConn))
            {
                dst.Open();
                Exec(dst, @"CREATE TABLE changelog(
                              seq INTEGER, dataset TEXT, record_id TEXT, field TEXT,
                              value TEXT, usr TEXT, ts TEXT);");
                using var tx = dst.BeginTransaction();
                using var ins = dst.CreateCommand();
                ins.CommandText = @"INSERT INTO changelog(seq,dataset,record_id,field,value,usr,ts)
                                    VALUES(@s,@d,@r,@f,@v,@u,@t)";
                var pS = ins.Parameters.Add("@s", SqliteType.Integer);
                var pD = ins.Parameters.Add("@d", SqliteType.Text);
                var pR = ins.Parameters.Add("@r", SqliteType.Text);
                var pF = ins.Parameters.Add("@f", SqliteType.Text);
                var pV = ins.Parameters.Add("@v", SqliteType.Text);
                var pU = ins.Parameters.Add("@u", SqliteType.Text);
                var pT = ins.Parameters.Add("@t", SqliteType.Text);

                rows = 0;
                using (var src = new SqliteConnection(connStr))
                {
                    src.Open();
                    using var read = src.CreateCommand();
                    read.CommandText = "SELECT seq,dataset,record_id,field,value,usr,ts FROM changelog ORDER BY seq";
                    using var rd = read.ExecuteReader();
                    while (rd.Read())
                    {
                        pS.Value = rd.GetInt64(0);
                        pD.Value = rd.GetString(1);
                        pR.Value = rd.GetString(2);
                        pF.Value = rd.GetString(3);
                        pV.Value = rd.IsDBNull(4) ? DBNull.Value : rd.GetString(4);
                        pU.Value = rd.GetString(5);
                        pT.Value = rd.GetString(6);
                        ins.ExecuteNonQuery();
                        rows++;
                    }
                }
                tx.Commit();
            }
            // Release the pooled handle on the staged file so it can be copied and deleted.
            SqliteConnection.ClearAllPools();
        }

        // Copy local -> share. Not a rename: the staged file is on a different volume now, so
        // a Move would be a copy-and-delete anyway. Retry for transient share locks, and
        // delete the local staging file whatever happens, so nothing is left behind here either.
        Exception? last = null;
        try
        {
            for (int i = 0; i < 5; i++)
            {
                try { File.Copy(tmp, target, overwrite: true); last = null; break; }
                catch (Exception ex) { last = ex; Thread.Sleep(150 * (i + 1)); }
            }
        }
        finally { try { if (File.Exists(tmp)) File.Delete(tmp); } catch { /* temp dir: harmless */ } }
        if (last != null) throw new Exception($"publish copy failed after retries: {last.Message}");

        _lastPublish = NowTs();
        return new { path = target, rows };
    }

    private static object? Sync(string shareRoot, string connStr)
    {
        var rows = new List<(string dataset, string id, string field, string? value, string usr, string ts)>();

        // 1) Local changelog - so unpublished local writes survive the rebuild.
        ReadChangelog(connStr, rows);

        // 2) Every user's published changelog on the share (includes our own file too).
        var files = Directory.Exists(shareRoot)
            ? Directory.GetFiles(shareRoot, "*.changelog.db")
            : Array.Empty<string>();
        var skipped = new List<string>();
        foreach (var file in files)
        {
            try
            {
                var ro = new SqliteConnectionStringBuilder { DataSource = file, Mode = SqliteOpenMode.ReadOnly }.ToString();
                ReadChangelog(ro, rows);
            }
            catch { skipped.Add(Path.GetFileName(file)); }
        }
        SqliteConnection.ClearAllPools();

        var usersSeen = rows.Select(r => r.usr).Distinct().OrderBy(u => u, StringComparer.Ordinal).ToList();
        var conflicts = new List<object>();
        int merged = 0;

        lock (_writeLock)
        {
            using var con = new SqliteConnection(connStr);
            con.Open();
            using var tx = con.BeginTransaction();
            Exec(con, "DELETE FROM state;");

            using var ins = con.CreateCommand();
            ins.CommandText = @"INSERT INTO state(dataset,record_id,field,value,usr,ts)
                                VALUES(@d,@r,@f,@v,@u,@t)";
            var pD = ins.Parameters.Add("@d", SqliteType.Text);
            var pR = ins.Parameters.Add("@r", SqliteType.Text);
            var pF = ins.Parameters.Add("@f", SqliteType.Text);
            var pV = ins.Parameters.Add("@v", SqliteType.Text);
            var pU = ins.Parameters.Add("@u", SqliteType.Text);
            var pT = ins.Parameters.Add("@t", SqliteType.Text);

            foreach (var g in rows.GroupBy(r => (r.dataset, r.id, r.field)))
            {
                // Winner = max by (ts, usr). ts is fixed-format ISO8601 => ordinal-sortable.
                var winner = g.OrderBy(x => x.ts, StringComparer.Ordinal)
                              .ThenBy(x => x.usr, StringComparer.Ordinal)
                              .Last();

                pD.Value = g.Key.dataset;
                pR.Value = g.Key.id;
                pF.Value = g.Key.field;
                pV.Value = (object?)winner.value ?? DBNull.Value;
                pU.Value = winner.usr;
                pT.Value = winner.ts;
                ins.ExecuteNonQuery();
                merged++;

                // One contribution per distinct user (their latest). Conflict = more than one
                // distinct user AND more than one distinct value for this field.
                var perUsr = g.GroupBy(x => x.usr)
                              .Select(u => u.OrderBy(x => x.ts, StringComparer.Ordinal).Last())
                              .ToList();
                int distinctValues = g.Select(x => x.value).Distinct().Count();
                if (perUsr.Count > 1 && distinctValues > 1)
                {
                    var losers = perUsr.Where(x => x.usr != winner.usr)
                                       .Select(x => new { usr = x.usr, ts = x.ts, value = x.value })
                                       .ToList();
                    conflicts.Add(new
                    {
                        dataset = g.Key.dataset,
                        id = g.Key.id,
                        field = g.Key.field,
                        winner = new { usr = winner.usr, ts = winner.ts, value = winner.value },
                        losers
                    });
                }
            }
            tx.Commit();
        }

        _lastSync = NowTs();
        return new { usersSeen, rowsMerged = merged, conflicts, skippedFiles = skipped };
    }

    // Remove every row whose dataset starts with `prefix`, from both the append-only
    // changelog and the materialised state, then VACUUM so the file actually shrinks.
    // Prefix-scoped rather than "clear everything" so records outside the prefix are
    // never in range.
    private static object? Purge(JsonElement body)
    {
        var prefix = ReqStr(body, "prefix");
        long before, changelog, state;
        lock (_writeLock)
        {
            using var con = new SqliteConnection(ConnStr);
            con.Open();
            using (var c = con.CreateCommand())
            {
                c.CommandText = "SELECT COUNT(*) FROM changelog WHERE dataset LIKE @p";
                c.Parameters.AddWithValue("@p", prefix + "%");
                before = Convert.ToInt64(c.ExecuteScalar() ?? 0L);
            }
            using (var c = con.CreateCommand())
            {
                c.CommandText = "DELETE FROM changelog WHERE dataset LIKE @p";
                c.Parameters.AddWithValue("@p", prefix + "%");
                changelog = c.ExecuteNonQuery();
            }
            using (var c = con.CreateCommand())
            {
                c.CommandText = "DELETE FROM state WHERE dataset LIKE @p";
                c.Parameters.AddWithValue("@p", prefix + "%");
                state = c.ExecuteNonQuery();
            }
            // Without VACUUM the pages are freed but the file stays the same size.
            Exec(con, "VACUUM;");
        }
        return new { prefix, matched = before, changelogDeleted = changelog, stateDeleted = state };
    }

    private static object? Status()
    {
        var datasets = new Dictionary<string, int>();
        long changelogRows;
        using (var con = new SqliteConnection(ConnStr))
        {
            con.Open();
            using (var c = con.CreateCommand())
            {
                c.CommandText = "SELECT dataset, COUNT(DISTINCT record_id) FROM state GROUP BY dataset ORDER BY dataset";
                using var rd = c.ExecuteReader();
                while (rd.Read()) datasets[rd.GetString(0)] = rd.GetInt32(1);
            }
            using (var c = con.CreateCommand())
            {
                c.CommandText = "SELECT COUNT(*) FROM changelog";
                changelogRows = Convert.ToInt64(c.ExecuteScalar());
            }
        }
        bool reachable = false;
        try { reachable = !string.IsNullOrEmpty(ShareRootFallback) && Directory.Exists(ShareRootFallback); } catch { }
        return new
        {
            dbPath = DbPath,
            datasets,
            changelogRows,
            lastPublish = _lastPublish,
            lastSync = _lastSync,
            shareRoot = ShareRootFallback,
            shareReachable = reachable
        };
    }

    // ---------------------------------------------------------------- internals

    // Append changelog rows + upsert state for a batch of records, in one transaction.
    // Own writes carry ts=now, so they are authoritative and overwrite state unconditionally.
    private static void WriteMany(string dataset, List<(string id, IDictionary<string, string?> fields)> records,
                                  string usr, string connStr)
    {
        var ts = NowTs();
        lock (_writeLock)
        {
            using var con = new SqliteConnection(connStr);
            con.Open();
            using var tx = con.BeginTransaction();

            using var clog = con.CreateCommand();
            clog.CommandText = @"INSERT INTO changelog(dataset,record_id,field,value,usr,ts)
                                 VALUES(@d,@r,@f,@v,@u,@t)";
            var cD = clog.Parameters.Add("@d", SqliteType.Text);
            var cR = clog.Parameters.Add("@r", SqliteType.Text);
            var cF = clog.Parameters.Add("@f", SqliteType.Text);
            var cV = clog.Parameters.Add("@v", SqliteType.Text);
            var cU = clog.Parameters.Add("@u", SqliteType.Text);
            var cT = clog.Parameters.Add("@t", SqliteType.Text);

            using var st = con.CreateCommand();
            st.CommandText = @"INSERT INTO state(dataset,record_id,field,value,usr,ts)
                               VALUES(@d,@r,@f,@v,@u,@t)
                               ON CONFLICT(dataset,record_id,field)
                               DO UPDATE SET value=excluded.value, usr=excluded.usr, ts=excluded.ts";
            var sD = st.Parameters.Add("@d", SqliteType.Text);
            var sR = st.Parameters.Add("@r", SqliteType.Text);
            var sF = st.Parameters.Add("@f", SqliteType.Text);
            var sV = st.Parameters.Add("@v", SqliteType.Text);
            var sU = st.Parameters.Add("@u", SqliteType.Text);
            var sT = st.Parameters.Add("@t", SqliteType.Text);

            cU.Value = usr; cT.Value = ts; sU.Value = usr; sT.Value = ts;

            foreach (var (id, fields) in records)
            {
                cR.Value = id; sR.Value = id; cD.Value = dataset; sD.Value = dataset;
                foreach (var kv in fields)
                {
                    object v = (object?)kv.Value ?? DBNull.Value;
                    cF.Value = kv.Key; cV.Value = v; clog.ExecuteNonQuery();
                    sF.Value = kv.Key; sV.Value = v; st.ExecuteNonQuery();
                }
            }
            tx.Commit();
        }
    }

    private static void ReadChangelog(string connStr, List<(string, string, string, string?, string, string)> sink)
    {
        using var con = new SqliteConnection(connStr);
        con.Open();
        using var c = con.CreateCommand();
        c.CommandText = "SELECT dataset,record_id,field,value,usr,ts FROM changelog";
        using var rd = c.ExecuteReader();
        while (rd.Read())
            sink.Add((rd.GetString(0), rd.GetString(1), rd.GetString(2),
                      rd.IsDBNull(3) ? null : rd.GetString(3), rd.GetString(4), rd.GetString(5)));
    }

    private static void Exec(SqliteConnection con, string sql)
    {
        using var c = con.CreateCommand();
        c.CommandText = sql;
        c.ExecuteNonQuery();
    }

    private static string NowTs() => DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

    private static JsonElement Parse(string body) =>
        JsonSerializer.Deserialize<JsonElement>(string.IsNullOrWhiteSpace(body) ? "{}" : body);

    private static string ReqStr(JsonElement b, string name) =>
        Str(b, name) is { Length: > 0 } v ? v : throw new Exception($"'{name}' is required.");

    private static string? Str(JsonElement b, string name) =>
        b.ValueKind == JsonValueKind.Object && b.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString() : null;

    // Normalize a JSON value to the string we store. Objects/arrays keep their raw JSON text.
    private static string? JsonToStr(JsonElement v) => v.ValueKind switch
    {
        JsonValueKind.String => v.GetString(),
        JsonValueKind.Null => null,
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        JsonValueKind.Number => v.GetRawText(),
        _ => v.GetRawText(),
    };
}
