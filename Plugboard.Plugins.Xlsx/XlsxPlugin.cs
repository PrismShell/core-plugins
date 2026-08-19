using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml;
using Plugboard.Contracts;

namespace Plugboard.Plugins.Xlsx;

// Read .xlsx files straight off disk or a share, with Excel CLOSED and Excel not even
// installed. Ported from DeskHub-Gateway's XlsxHandler.
//
// This is deliberately NOT the excel connector. That one drives a live Excel COM instance -
// it attaches to a workbook someone has open, which is right for "read the cell the trader
// is looking at" and wrong for "load a reference dataset". A service endpoint must never
// open Excel windows or fight with a user's open workbook.
//
// An .xlsx is a zip of XML, so this reads it with System.IO.Compression + XmlReader: no COM,
// no interop, no Access Database Engine, nothing to install.
//
// Three details that are easy to get wrong and produce silently wrong data rather than an error:
//
//  1. SHARED STRINGS. A text cell stores an INDEX into xl/sharedStrings.xml, not the text.
//     Miss the table and every string comes back as a number. Each <si> is read as a subtree
//     because ReadElementContentAsString() consumes through </t> and would skip </si>, and
//     because rich text splits one string across several <r><t> runs.
//  2. OMITTED CELLS. Empty cells are absent from the XML entirely, so the column has to come
//     from the cell reference ("B7" -> column 1). Taking cells in document order shifts every
//     value right of the first blank.
//  3. SHEET ORDER. xl/worksheets/sheet1.xml is not necessarily the first tab. The name -> part
//     mapping goes through r:id in xl/_rels/workbook.xml.rels.
//
// Parsed workbooks are cached on (path, last-write, length), so a 40MB reference file is read
// once and a monthly refresh invalidates itself with no restart. Optional allow-list of
// directory prefixes via PLUGBOARD_XLSX_ROOTS (semicolon-separated); empty = any readable path,
// the same reach the files connector already has.
public sealed class XlsxPlugin : IPlugin
{
    public string Name => "xlsx";

    // Bounded so a stray loop over a directory cannot pin the whole share in memory.
    private const int MaxCachedWorkbooks = 8;
    private static readonly object _gate = new();
    private static readonly Dictionary<string, CachedBook> _cache = new(StringComparer.OrdinalIgnoreCase);

    private static readonly string[] Roots =
        (Environment.GetEnvironmentVariable("PLUGBOARD_XLSX_ROOTS") ?? "")
        .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    private sealed record Sheet(string Name, List<string> Columns, List<Dictionary<string, object?>> Rows);
    private sealed record CachedBook(string Key, DateTime Modified, long Length, DateTime LoadedUtc,
                                     Dictionary<string, Sheet> Sheets, List<string> Order);

    public void Register(IEndpointRegistry r)
    {
        r.Map("POST", "xlsx/sheets", req => Task.FromResult(Sheets(Parse(req.Body))), new RouteInfo(
            "List sheets in a workbook",
            "Sheet names with row and column counts. Cheap way to see what a workbook holds before asking for data.",
            new { path = @"\\share\ref\pools.xlsx" },
            new[] { new ParamInfo("path", "string", true, "Path to the .xlsx file (env vars expand).") }));

        r.Map("POST", "xlsx/read", req => Task.FromResult(Read(Parse(req.Body))), new RouteInfo(
            "Read rows from a sheet",
            "Filtering happens HERE, not in the caller: a 40MB reference file is 50k+ rows and nobody wants that over the wire to find one row. First row of the sheet is the header.",
            new { path = @"\\share\ref\pools.xlsx", sheet = "Pools", where = new { Cusip = "3142GXH34" }, maxRows = 100 },
            new[]
            {
                new ParamInfo("path", "string", true, "Path to the .xlsx file (env vars expand)."),
                new ParamInfo("sheet", "string", false, "Sheet name. Defaults to the first sheet."),
                new ParamInfo("columns", "array", false, "Project to just these columns.", Items: "string"),
                new ParamInfo("where", "object", false, "Exact-match filter: { column: value } or { column: [values] }. Case-insensitive on text."),
                new ParamInfo("maxRows", "integer", false, "Row cap after filtering.", Default: 1000),
                new ParamInfo("offset", "integer", false, "Rows to skip after filtering.", Default: 0),
            }));

        r.Map("POST", "xlsx/forget", req => Task.FromResult(Forget(Parse(req.Body))), new RouteInfo(
            "Drop the workbook cache",
            "Forget one cached workbook ({path}) or all of them (empty body), without a restart.",
            new { path = @"\\share\ref\pools.xlsx" },
            new[] { new ParamInfo("path", "string", false, "Workbook to forget; omit to clear everything.") }));
    }

    // ---- endpoints ----

    private static object? Sheets(JsonElement b)
    {
        var full = CheckPath(Req(b, "path"));
        var book = Load(full);
        var fi = new FileInfo(full);
        return new
        {
            file = new { path = fi.FullName, modified = fi.LastWriteTime.ToString("s"), bytes = fi.Length },
            cachedUtc = book.LoadedUtc.ToString("s"),
            sheets = book.Order.Select(n => new
            {
                name = n,
                rows = book.Sheets[n].Rows.Count,
                columns = book.Sheets[n].Columns.Count,
            }),
        };
    }

    private static object? Read(JsonElement b)
    {
        var full = CheckPath(Req(b, "path"));
        var book = Load(full);

        var sheetName = Str(b, "sheet");
        if (string.IsNullOrWhiteSpace(sheetName)) sheetName = book.Order.FirstOrDefault() ?? "";
        if (!book.Sheets.TryGetValue(sheetName!, out var sheet))
            throw new Exception($"sheet '{sheetName}' not found (sheets: {string.Join(", ", book.Order)})");

        IEnumerable<Dictionary<string, object?>> rows = sheet.Rows;

        // where: exact match, case-insensitive on text, or membership in a list
        if (b.TryGetProperty("where", out var w) && w.ValueKind == JsonValueKind.Object)
        {
            foreach (var f in w.EnumerateObject())
            {
                var col = sheet.Columns.FirstOrDefault(c => c.Equals(f.Name, StringComparison.OrdinalIgnoreCase))
                    ?? throw new Exception($"unknown column in where: {f.Name} (columns: {string.Join(", ", sheet.Columns)})");
                var wanted = new List<string>();
                if (f.Value.ValueKind == JsonValueKind.Array)
                    foreach (var el in f.Value.EnumerateArray()) wanted.Add(Norm(el.ToString()));
                else wanted.Add(Norm(f.Value.ToString()));
                rows = rows.Where(r => r.TryGetValue(col, out var v) && wanted.Contains(Norm(v?.ToString())));
            }
        }

        var total = rows is ICollection<Dictionary<string, object?>> c0 ? c0.Count : rows.Count();
        var offset = Math.Max(0, Int(b, "offset", 0));
        var maxRows = Int(b, "maxRows", 1000);
        if (maxRows <= 0) maxRows = int.MaxValue;
        var page = rows.Skip(offset).Take(maxRows).ToList();

        // columns: narrow the projection
        List<string> outCols = sheet.Columns;
        if (b.TryGetProperty("columns", out var cs) && cs.ValueKind == JsonValueKind.Array)
        {
            var want = new List<string>();
            foreach (var el in cs.EnumerateArray())
            {
                var nm = el.GetString() ?? "";
                var col = sheet.Columns.FirstOrDefault(x => x.Equals(nm, StringComparison.OrdinalIgnoreCase))
                    ?? throw new Exception($"unknown column: {nm} (columns: {string.Join(", ", sheet.Columns)})");
                want.Add(col);
            }
            if (want.Count > 0)
            {
                outCols = want;
                page = page.Select(r =>
                {
                    var o = new Dictionary<string, object?>();
                    foreach (var k in want) o[k] = r.TryGetValue(k, out var v) ? v : null;
                    return o;
                }).ToList();
            }
        }

        var fi = new FileInfo(full);
        return new
        {
            sheet = sheetName, matched = total, returned = page.Count, offset,
            file = new { path = fi.FullName, modified = fi.LastWriteTime.ToString("s") },
            columns = outCols, rows = page,
        };
    }

    private static object? Forget(JsonElement b)
    {
        var path = Str(b, "path");
        lock (_gate)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                var n = _cache.Count; _cache.Clear();
                return new { cleared = n };
            }
            var key = Path.GetFullPath(Environment.ExpandEnvironmentVariables(path!));
            return new { cleared = _cache.Remove(key) ? 1 : 0 };
        }
    }

    // ---- path guard / cache ----

    private static string CheckPath(string path)
    {
        var full = Path.GetFullPath(Environment.ExpandEnvironmentVariables(path));
        if (Roots.Length > 0 &&
            !Roots.Any(r => full.StartsWith(Path.GetFullPath(r), StringComparison.OrdinalIgnoreCase)))
            throw new Exception("path is outside PLUGBOARD_XLSX_ROOTS");
        return full;
    }

    private static CachedBook Load(string path)
    {
        var fi = new FileInfo(path);
        if (!fi.Exists) throw new FileNotFoundException("file not found: " + path);
        var key = fi.FullName;
        lock (_gate)
        {
            if (_cache.TryGetValue(key, out var hit) && hit.Modified == fi.LastWriteTimeUtc && hit.Length == fi.Length)
                return hit;
        }

        var book = ParseBook(path, fi);
        lock (_gate)
        {
            if (_cache.Count >= MaxCachedWorkbooks)
            {
                var oldest = _cache.Values.OrderBy(x => x.LoadedUtc).First().Key;
                _cache.Remove(oldest);
            }
            _cache[key] = book;
        }
        return book;
    }

    // ---- xlsx parsing ----

    private static CachedBook ParseBook(string path, FileInfo fi)
    {
        // FileShare.ReadWrite is the point: Excel holds a write lock while the workbook is open,
        // and ZipFile.OpenRead (FileShare.Read) throws "in use". This way a reader never blocks
        // the person editing the file, and they never block the reader.
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);

        ZipArchiveEntry? E(string name) =>
            zip.Entries.FirstOrDefault(e => e.FullName.Equals(name, StringComparison.OrdinalIgnoreCase));

        // 1. shared strings
        var shared = new List<string>();
        var ssEntry = E("xl/sharedStrings.xml");
        if (ssEntry != null)
        {
            using var s = ssEntry.Open();
            using var rd = XmlReader.Create(s);
            while (rd.Read())
            {
                if (rd.NodeType != XmlNodeType.Element || rd.Name != "si") continue;
                using var sub = rd.ReadSubtree();
                var sb = new StringBuilder();
                while (sub.Read())
                    if (sub.NodeType == XmlNodeType.Element && sub.Name == "t")
                        sb.Append(sub.ReadElementContentAsString());
                shared.Add(sb.ToString());
            }
        }

        // 2. which numeric formats are dates (so a serial does not surface as 45874)
        var dateStyles = ReadDateStyles(E("xl/styles.xml"));

        // 3. sheet name -> r:id -> part path
        var order = new List<string>();
        var nameToRid = new Dictionary<string, string>();
        var wbEntry = E("xl/workbook.xml") ?? throw new InvalidDataException("xl/workbook.xml missing - not an xlsx?");
        using (var s = wbEntry.Open())
        using (var rd = XmlReader.Create(s))
            while (rd.Read())
                if (rd.NodeType == XmlNodeType.Element && rd.Name == "sheet")
                {
                    var nm = rd.GetAttribute("name") ?? "";
                    var rid = rd.GetAttribute("r:id") ?? rd.GetAttribute("id") ?? "";
                    if (nm.Length > 0) { order.Add(nm); nameToRid[nm] = rid; }
                }

        var ridToTarget = new Dictionary<string, string>();
        var relEntry = E("xl/_rels/workbook.xml.rels");
        if (relEntry != null)
        {
            using var s = relEntry.Open();
            using var rd = XmlReader.Create(s);
            while (rd.Read())
                if (rd.NodeType == XmlNodeType.Element && rd.Name == "Relationship")
                {
                    var id = rd.GetAttribute("Id") ?? "";
                    var tgt = rd.GetAttribute("Target") ?? "";
                    if (id.Length > 0) ridToTarget[id] = tgt;
                }
        }

        var sheets = new Dictionary<string, Sheet>(StringComparer.OrdinalIgnoreCase);
        foreach (var nm in order)
        {
            string part = "";
            if (nameToRid.TryGetValue(nm, out var rid) && ridToTarget.TryGetValue(rid, out var tgt))
                part = "xl/" + tgt.TrimStart('/').Replace("xl/", "", StringComparison.OrdinalIgnoreCase);
            var entry = part.Length > 0 ? E(part) : null;
            if (entry == null) { sheets[nm] = new Sheet(nm, new(), new()); continue; }
            sheets[nm] = ReadSheet(nm, entry, shared, dateStyles);
        }

        return new CachedBook(fi.FullName, fi.LastWriteTimeUtc, fi.Length, DateTime.UtcNow, sheets, order);
    }

    // numFmtId -> is-a-date. Built-ins 14-22 and 45-47 are date/time; custom formats are detected
    // by looking for y/m/d/h in the format string while ignoring anything in quotes.
    private static HashSet<int> ReadDateStyles(ZipArchiveEntry? styles)
    {
        var dateFmtIds = new HashSet<int> { 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47 };
        var result = new HashSet<int>();
        if (styles == null) return result;
        var custom = new Dictionary<int, string>();
        var xfs = new List<int>();
        using (var s = styles.Open())
        using (var rd = XmlReader.Create(s))
        {
            bool inCellXfs = false;
            while (rd.Read())
            {
                if (rd.NodeType == XmlNodeType.Element && rd.Name == "numFmt")
                {
                    if (int.TryParse(rd.GetAttribute("numFmtId"), out var id))
                        custom[id] = rd.GetAttribute("formatCode") ?? "";
                }
                else if (rd.NodeType == XmlNodeType.Element && rd.Name == "cellXfs") inCellXfs = true;
                else if (rd.NodeType == XmlNodeType.EndElement && rd.Name == "cellXfs") inCellXfs = false;
                else if (inCellXfs && rd.NodeType == XmlNodeType.Element && rd.Name == "xf")
                    xfs.Add(int.TryParse(rd.GetAttribute("numFmtId"), out var n) ? n : 0);
            }
        }
        for (int i = 0; i < xfs.Count; i++)
        {
            var id = xfs[i];
            bool isDate = dateFmtIds.Contains(id);
            if (!isDate && custom.TryGetValue(id, out var code))
            {
                var bare = System.Text.RegularExpressions.Regex.Replace(code, "\"[^\"]*\"", "");
                isDate = bare.IndexOfAny(new[] { 'y', 'd' }) >= 0
                      || bare.Contains("mm") || bare.Contains("hh") || bare.Contains("ss");
            }
            if (isDate) result.Add(i);   // index into cellXfs = the cell's s= attribute
        }
        return result;
    }

    private static Sheet ReadSheet(string name, ZipArchiveEntry entry, List<string> shared, HashSet<int> dateStyles)
    {
        var header = new List<string>();
        var rows = new List<Dictionary<string, object?>>();
        using var s = entry.Open();
        using var rd = XmlReader.Create(s);

        var cells = new SortedDictionary<int, object?>();
        bool inRow = false;
        int col = 0, styleIdx = -1;
        string type = "";

        void Flush()
        {
            if (header.Count == 0)
            {
                // First row is the header. Blank headers become col1/col2/... so the position is
                // still addressable rather than collapsing into one empty key.
                var max = cells.Count > 0 ? cells.Keys.Max() : -1;
                for (int i = 0; i <= max; i++)
                {
                    var h = cells.TryGetValue(i, out var v) ? v?.ToString()?.Trim() : null;
                    header.Add(string.IsNullOrEmpty(h) ? $"col{i + 1}" : h!);
                }
                // Duplicate header names would silently overwrite each other in the row dictionary.
                for (int i = 0; i < header.Count; i++)
                {
                    var n = header[i];
                    if (header.FindIndex(x => x.Equals(n, StringComparison.OrdinalIgnoreCase)) != i)
                        header[i] = n + "_" + (i + 1);
                }
                return;
            }
            var row = new Dictionary<string, object?>();
            for (int i = 0; i < header.Count; i++) row[header[i]] = cells.TryGetValue(i, out var v) ? v : null;
            // Cells beyond the header width are kept rather than dropped.
            foreach (var kv in cells) if (kv.Key >= header.Count) row["col" + (kv.Key + 1)] = kv.Value;
            rows.Add(row);
        }

        while (rd.Read())
        {
            if (rd.NodeType == XmlNodeType.Element && rd.Name == "row") { cells.Clear(); inRow = true; }
            else if (rd.NodeType == XmlNodeType.Element && rd.Name == "c" && inRow)
            {
                col = ColumnIndex(rd.GetAttribute("r"));
                type = rd.GetAttribute("t") ?? "";
                styleIdx = int.TryParse(rd.GetAttribute("s"), out var si) ? si : -1;
                if (rd.IsEmptyElement) cells[col] = null;
            }
            else if (rd.NodeType == XmlNodeType.Element && rd.Name == "v" && inRow)
            {
                var raw = rd.ReadElementContentAsString();
                cells[col] = ConvertCell(raw, type, styleIdx, shared, dateStyles);
            }
            else if (rd.NodeType == XmlNodeType.Element && rd.Name == "t" && inRow && type == "inlineStr")
                cells[col] = rd.ReadElementContentAsString();
            else if (rd.NodeType == XmlNodeType.EndElement && rd.Name == "row") { Flush(); inRow = false; }
        }
        if (inRow) Flush();   // a final row with no closing tag
        return new Sheet(name, header, rows);
    }

    private static object? ConvertCell(string raw, string type, int styleIdx, List<string> shared, HashSet<int> dateStyles)
    {
        if (type == "s")
            return int.TryParse(raw, out var ix) && ix >= 0 && ix < shared.Count ? shared[ix] : raw;
        if (type == "str" || type == "inlineStr") return raw;
        if (type == "b") return raw == "1";
        if (type == "e") return raw;                       // #N/A, #REF! etc: keep the error text
        if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var d))
        {
            if (styleIdx >= 0 && dateStyles.Contains(styleIdx) && d > 0 && d < 2958466)
            {
                // Excel's serial epoch is 1899-12-30, absorbing the fictional 1900-02-29.
                var dt = new DateTime(1899, 12, 30).AddDays(d);
                return dt.TimeOfDay == TimeSpan.Zero ? dt.ToString("yyyy-MM-dd") : dt.ToString("s");
            }
            return d;
        }
        return raw.Length == 0 ? null : raw;
    }

    private static int ColumnIndex(string? cellRef)
    {
        if (string.IsNullOrEmpty(cellRef)) return 0;
        int n = 0;
        foreach (var ch in cellRef)
        {
            var c = char.ToUpperInvariant(ch);
            if (c < 'A' || c > 'Z') break;
            n = n * 26 + (c - 'A' + 1);
        }
        return Math.Max(0, n - 1);
    }

    // ---- body helpers ----

    private static string Norm(string? s) => (s ?? "").Trim().ToUpperInvariant();

    private static JsonElement Parse(string body) =>
        JsonSerializer.Deserialize<JsonElement>(string.IsNullOrWhiteSpace(body) ? "{}" : body);

    private static string Req(JsonElement b, string name) =>
        Str(b, name) is { Length: > 0 } v ? v : throw new Exception($"'{name}' is required.");

    private static string? Str(JsonElement b, string name) =>
        b.ValueKind == JsonValueKind.Object && b.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString() : null;

    private static int Int(JsonElement b, string name, int dflt) =>
        b.ValueKind == JsonValueKind.Object && b.TryGetProperty(name, out var v) && v.TryGetInt32(out var i) ? i : dflt;
}
