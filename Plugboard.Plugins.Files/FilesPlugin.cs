using System.Diagnostics;
using System.Text.Json;
using Plugboard.Contracts;

namespace Plugboard.Plugins.Files;

// Ported from the local-gateway FileHandler (/api/file, /api/file/open). Reads a
// local file (binary-safe: text as utf8, anything else as base64), or opens one
// with the OS shell. Note: like the original, the path is not clamped to a root -
// this is a local trusted tool.
public sealed class FilesPlugin : IPlugin
{
    public string Name => "files";

    public void Register(IEndpointRegistry r)
    {
        // GET files/read?path=...  -> { path, mime, encoding, content }
        // Text types return content as utf8; binary types (xlsx, pdf, …) return
        // content as base64 so the bytes survive the JSON envelope intact.
        r.Map("GET", "files/read", req =>
        {
            if (!req.Query.TryGetValue("path", out var path) || string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("missing 'path' query parameter");
            path = Environment.ExpandEnvironmentVariables(path);
            if (!File.Exists(path)) throw new FileNotFoundException("not found: " + path);
            var mime  = Mime(Path.GetExtension(path));
            var bytes = File.ReadAllBytes(path);
            return Task.FromResult<object?>(IsText(mime)
                ? new { path, mime, encoding = "utf8",   content = System.Text.Encoding.UTF8.GetString(bytes) }
                : new { path, mime, encoding = "base64", content = Convert.ToBase64String(bytes) });
        });

        // POST files/open { "path": "..." }  -> opens the file with the OS shell
        r.Map("POST", "files/open", req =>
        {
            using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(req.Body) ? "{}" : req.Body);
            var path = Environment.ExpandEnvironmentVariables(doc.RootElement.GetProperty("path").GetString() ?? "");
            if (!File.Exists(path)) throw new FileNotFoundException("not found: " + path);
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            return Task.FromResult<object?>(new { opened = path });
        });

        // POST files/write { path, content, encoding: "utf8"|"base64", append: false }
        // Writes a file (text by default; base64 for binary). Creates parent dirs.
        // Gated by env var PLUGBOARD_FILES_ALLOW_WRITE (default allowed; "false" disables).
        r.Map("POST", "files/write", req =>
        {
            if (string.Equals(Environment.GetEnvironmentVariable("PLUGBOARD_FILES_ALLOW_WRITE"), "false", StringComparison.OrdinalIgnoreCase))
                throw new Exception("file write is disabled (PLUGBOARD_FILES_ALLOW_WRITE=false).");
            using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(req.Body) ? "{}" : req.Body);
            var root = doc.RootElement;
            var path = Environment.ExpandEnvironmentVariables(
                root.TryGetProperty("path", out var p) ? (p.GetString() ?? "") : "");
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("'path' is required.");
            var content  = root.TryGetProperty("content", out var c) ? (c.GetString() ?? "") : "";
            var encoding = root.TryGetProperty("encoding", out var e) ? (e.GetString() ?? "utf8") : "utf8";
            var append   = root.TryGetProperty("append", out var a) && a.ValueKind == JsonValueKind.True;

            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            if (encoding.Equals("base64", StringComparison.OrdinalIgnoreCase))
            {
                var bytes = Convert.FromBase64String(content);
                if (append) { using var fs = new FileStream(path, FileMode.Append, FileAccess.Write); fs.Write(bytes, 0, bytes.Length); }
                else File.WriteAllBytes(path, bytes);
            }
            else
            {
                if (append) File.AppendAllText(path, content);
                else File.WriteAllText(path, content);
            }
            return Task.FromResult<object?>(new { path, bytes = new FileInfo(path).Length });
        },
        new RouteInfo("Write a file",
            "Text by default; encoding=base64 for binary. Creates parent dirs. append=true to append.",
            new { path = "%TEMP%\\note.txt", content = "hello", encoding = "utf8", append = false },
            new[]
            {
                new ParamInfo("path", "string", true, "Destination path (env vars expanded)."),
                new ParamInfo("content", "string", true, "File content (utf8 text or base64)."),
                new ParamInfo("encoding", "string", false, "utf8 or base64.", Enum: new[] { "utf8", "base64" }, Default: "utf8"),
                new ParamInfo("append", "boolean", false, "Append instead of overwrite.", Default: false),
            }));
    }

    private static string Mime(string ext) => ext.ToLowerInvariant() switch
    {
        ".css"  => "text/css",
        ".js"   => "application/javascript",
        ".html" or ".htm" => "text/html",
        ".json" => "application/json",
        ".csv"  => "text/csv",
        ".txt"  => "text/plain",
        ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
        ".xls"  => "application/vnd.ms-excel",
        ".pdf"  => "application/pdf",
        _       => "application/octet-stream"
    };

    private static bool IsText(string mime) =>
        mime.StartsWith("text/") || mime == "application/json" || mime == "application/javascript";
}
