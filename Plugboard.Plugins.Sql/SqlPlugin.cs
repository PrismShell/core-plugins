using System.Text.Json;
using Microsoft.Data.SqlClient;
using Plugboard.Contracts;

namespace Plugboard.Plugins.Sql;

// SQL Server access. Ported from local-gateway's SqlHandler. Connect with a full
// `connectionString`, or with `server`/`database` (Windows integrated auth by default;
// pass `user`+`password` for SQL auth). Loopback-only, runs as the logged-in user - same
// trust as SSMS on this box. Execute (non-query) is gated by the env var
// PLUGBOARD_SQL_ALLOW_EXECUTE (default allowed; set to "false" to disable writes).
public sealed class SqlPlugin : IPlugin
{
    public string Name => "sql";

    private static readonly bool AllowExecute =
        !string.Equals(Environment.GetEnvironmentVariable("PLUGBOARD_SQL_ALLOW_EXECUTE"), "false", StringComparison.OrdinalIgnoreCase);

    private static readonly ParamInfo[] ConnParams =
    {
        new("server", "string", false, "Server/instance. Default localhost."),
        new("database", "string", false, "Initial catalog."),
        new("connectionString", "string", false, "Full connection string (overrides server/database)."),
        new("user", "string", false, "SQL auth user (integrated auth if omitted)."),
        new("password", "string", false, "SQL auth password."),
    };

    public void Register(IEndpointRegistry r)
    {
        r.Map("POST", "sql/ping", req => Ping(Parse(req.Body)),
            new RouteInfo("Test a connection", "Opens a connection and reports server/database/version.",
                new { server = "localhost", database = "master" }, ConnParams));

        r.Map("POST", "sql/query", req => { var b = Parse(req.Body); return Query(b, Req(b, "sql"), Max(b)); },
            new RouteInfo("Run a SELECT", "Returns { columns, rowCount, truncated, rows }. Supports named @params.",
                new { server = "localhost", database = "master", sql = "SELECT TOP 10 name FROM sys.objects", maxRows = 1000 },
                new[]
                {
                    new ParamInfo("sql", "string", true, "The query text."),
                    new ParamInfo("maxRows", "integer", false, "Row cap.", Default: 1000),
                    new ParamInfo("params", "object", false, "Named parameters mapped to @name."),
                }.Concat(ConnParams).ToArray()));

        r.Map("POST", "sql/execute", req =>
        {
            if (!AllowExecute) throw new Exception("SQL execute is disabled (PLUGBOARD_SQL_ALLOW_EXECUTE=false).");
            var b = Parse(req.Body); return Execute(b, Req(b, "sql"));
        }, new RouteInfo("Run a non-query (INSERT/UPDATE/DELETE/DDL)", "Returns { rowsAffected }. Gated by PLUGBOARD_SQL_ALLOW_EXECUTE.",
                new { server = "localhost", database = "mydb", sql = "UPDATE t SET x=1 WHERE id=@id", @params = new { id = 1 } }));

        r.Map("POST", "sql/databases", req => Query(Parse(req.Body),
                "SELECT name, database_id, create_date FROM sys.databases ORDER BY name", 5000),
            new RouteInfo("List databases on the server", null, new { server = "localhost" }, ConnParams));

        r.Map("POST", "sql/tables", req => Query(Parse(req.Body),
                "SELECT TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_SCHEMA, TABLE_NAME", 5000),
            new RouteInfo("List tables/views in the connected database", null, new { server = "localhost", database = "mydb" }, ConnParams));
    }

    // ---- helpers ----
    private static JsonElement Parse(string body) =>
        JsonSerializer.Deserialize<JsonElement>(string.IsNullOrWhiteSpace(body) ? "{}" : body);

    private static string Req(JsonElement b, string name) =>
        Str(b, name) ?? throw new Exception($"'{name}' is required.");

    private static int Max(JsonElement b) =>
        b.TryGetProperty("maxRows", out var mr) && mr.TryGetInt32(out var m) ? m : 1000;

    private static string ConnStr(JsonElement b)
    {
        var explicitCs = Str(b, "connectionString");
        if (!string.IsNullOrWhiteSpace(explicitCs)) return explicitCs!;
        var sb = new SqlConnectionStringBuilder
        {
            DataSource = Str(b, "server") ?? Str(b, "dataSource") ?? "localhost",
            ConnectTimeout = 15,
        };
        var db = Str(b, "database"); if (db != null) sb.InitialCatalog = db;
        var user = Str(b, "user");
        if (!string.IsNullOrEmpty(user)) { sb.UserID = user; sb.Password = Str(b, "password") ?? ""; }
        else sb.IntegratedSecurity = true;
        sb.TrustServerCertificate = !(b.TryGetProperty("trustServerCertificate", out var t) && t.ValueKind == JsonValueKind.False);
        return sb.ConnectionString;
    }

    private static string? Str(JsonElement b, string name) =>
        b.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static void AddParams(SqlCommand cmd, JsonElement b)
    {
        if (b.TryGetProperty("params", out var pr) && pr.ValueKind == JsonValueKind.Object)
            foreach (var p in pr.EnumerateObject())
                cmd.Parameters.AddWithValue("@" + p.Name.TrimStart('@'), JsonToValue(p.Value) ?? DBNull.Value);
    }

    private static object? JsonToValue(JsonElement v) => v.ValueKind switch
    {
        JsonValueKind.String => v.GetString(),
        JsonValueKind.Number => v.TryGetInt64(out var l) ? l : v.GetDouble(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null => null,
        _ => v.GetRawText(),
    };

    private static object? Normalize(object? v) => v switch
    {
        null => null,
        DateTime dt => dt.ToString("o"),
        DateTimeOffset dto => dto.ToString("o"),
        byte[] bytes => Convert.ToBase64String(bytes),
        decimal m => (double)m,
        Guid g => g.ToString(),
        _ => v,
    };

    private static async Task<object?> Ping(JsonElement b)
    {
        using var con = new SqlConnection(ConnStr(b));
        await con.OpenAsync();
        return new { server = con.DataSource, database = con.Database, version = con.ServerVersion };
    }

    private static async Task<object?> Query(JsonElement b, string sql, int max)
    {
        using var con = new SqlConnection(ConnStr(b));
        await con.OpenAsync();
        using var cmd = new SqlCommand(sql, con) { CommandTimeout = 60 };
        AddParams(cmd, b);
        using var rd = await cmd.ExecuteReaderAsync();
        var cols = new List<string>();
        for (int i = 0; i < rd.FieldCount; i++) cols.Add(rd.GetName(i));
        var rows = new List<Dictionary<string, object?>>();
        bool truncated = false;
        while (await rd.ReadAsync())
        {
            if (rows.Count >= max) { truncated = true; break; }
            var row = new Dictionary<string, object?>(cols.Count);
            for (int i = 0; i < rd.FieldCount; i++) row[cols[i]] = Normalize(rd.IsDBNull(i) ? null : rd.GetValue(i));
            rows.Add(row);
        }
        return new { columns = cols, rowCount = rows.Count, truncated, rows };
    }

    private static async Task<object?> Execute(JsonElement b, string sql)
    {
        using var con = new SqlConnection(ConnStr(b));
        await con.OpenAsync();
        using var cmd = new SqlCommand(sql, con) { CommandTimeout = 120 };
        AddParams(cmd, b);
        return new { rowsAffected = await cmd.ExecuteNonQueryAsync() };
    }
}
