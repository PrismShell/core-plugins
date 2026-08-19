using System.Runtime.InteropServices;
using System.Text.Json;
using Plugboard.Contracts;

namespace Plugboard.Plugins.Excel;

// Attaches to the ALREADY-RUNNING Excel instance (GetActiveObject), same as the
// local-gateway ExcelHandler.
internal static class ComHelper
{
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject([MarshalAs(UnmanagedType.LPStruct)] Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    public static object GetActiveObject(string progId)
    {
        var clsid = Type.GetTypeFromProgID(progId, true)!.GUID;
        GetActiveObject(clsid, IntPtr.Zero, out var obj);
        return obj;
    }
}

// Ported from the local-gateway ExcelHandler (COM automation via dynamic). Each
// route returns its data object (host wraps in { ok, data }); errors throw (host
// wraps in { ok:false, error }). Detect is the exception: "Excel not running" is
// data, not an error.
public sealed class ExcelPlugin : IPlugin
{
    public string Name => "excel";

    public void Register(IEndpointRegistry r)
    {
        r.Map("GET",  "excel/detect",            _ => Task.FromResult<object?>(Detect()));
        r.Map("GET",  "excel/sheets",            q => Task.FromResult<object?>(GetSheets(Req(q, "workbook"))));
        r.Map("GET",  "excel/app-info",          _ => Task.FromResult<object?>(GetAppInfo()));
        r.Map("GET",  "excel/list-named-ranges", q => Task.FromResult<object?>(ListNamedRanges(Req(q, "workbook"))));
        r.Map("POST", "excel/get-range",         q => Json(q, GetRange));
        r.Map("POST", "excel/set-range",         q => Json(q, SetRange));
        r.Map("POST", "excel/get-formula",       q => Json(q, GetFormula));
        r.Map("POST", "excel/set-formula",       q => Json(q, SetFormula));
        r.Map("POST", "excel/get-named-range",   q => Json(q, GetNamedRange));
        r.Map("POST", "excel/get-used-range",    q => Json(q, GetUsedRange));
        r.Map("POST", "excel/run-macro",         q => Json(q, RunMacro));
        r.Map("POST", "excel/save",              q => Json(q, Save));
        r.Map("POST", "excel/refresh",           q => Json(q, Refresh));
        r.Map("POST", "excel/create-sheet",      q => Json(q, CreateSheet));
        r.Map("POST", "excel/rename-sheet",      q => Json(q, RenameSheet));
        r.Map("POST", "excel/delete-sheet",      q => Json(q, DeleteSheet));
        r.Map("POST", "excel/copy-sheet",        q => Json(q, CopySheet));
        r.Map("POST", "excel/hide-sheet",        q => Json(q, HideSheet));
        r.Map("POST", "excel/unhide-sheet",      q => Json(q, UnhideSheet));
        r.Map("POST", "excel/protect-sheet",     q => Json(q, ProtectSheet));
        r.Map("POST", "excel/unprotect-sheet",   q => Json(q, UnprotectSheet));
        r.Map("POST", "excel/clear-range",       q => Json(q, ClearRange));
        r.Map("POST", "excel/delete-range",      q => Json(q, DeleteRange));
        r.Map("POST", "excel/insert-range",      q => Json(q, InsertRange));
        r.Map("POST", "excel/get-cell-format",   q => Json(q, GetCellFormat));
        r.Map("POST", "excel/set-cell-format",   q => Json(q, SetCellFormat));
        r.Map("POST", "excel/add-named-range",   q => Json(q, AddNamedRange));
        r.Map("POST", "excel/delete-named-range",q => Json(q, DeleteNamedRange));
        r.Map("POST", "excel/find-replace",      q => Json(q, FindReplace));
        r.Map("POST", "excel/sort-range",        q => Json(q, SortRange));
        r.Map("POST", "excel/autofit",           q => Json(q, Autofit));
        r.Map("POST", "excel/calculate",         q => Json(q, Calculate));
        r.Map("POST", "excel/set-calc-mode",     q => Json(q, SetCalcMode));
        r.Map("POST", "excel/save-as",           q => Json(q, SaveAs));
        r.Map("POST", "excel/protect-workbook",  q => Json(q, ProtectWorkbook));
        r.Map("POST", "excel/unprotect-workbook",q => Json(q, UnprotectWorkbook));
        r.Map("POST", "excel/get-doc-properties",q => Json(q, GetDocProperties));
        r.Map("POST", "excel/export-pdf",        q => Json(q, ExportPdf));
        r.Map("POST", "excel/set-screen-updating",q => Json(q, SetScreenUpdating));
        r.Map("POST", "excel/set-status-bar",    q => Json(q, SetStatusBar));
        r.Map("GET",  "excel/vba/list-components",q => Task.FromResult<object?>(VbaListComponents(Req(q, "workbook"))));
        r.Map("POST", "excel/vba/get-module",    q => Json(q, VbaGetModule));
        r.Map("POST", "excel/vba/set-module",    q => Json(q, VbaSetModule));
        r.Map("POST", "excel/vba/add-module",    q => Json(q, VbaAddModule));
        r.Map("POST", "excel/vba/add-class",     q => Json(q, VbaAddClass));
        r.Map("POST", "excel/vba/add-form",      q => Json(q, VbaAddForm));
        r.Map("POST", "excel/vba/delete-component",q => Json(q, VbaDeleteComponent));
        r.Map("POST", "excel/vba/export-component",q => Json(q, VbaExportComponent));
        r.Map("POST", "excel/vba/import-component",q => Json(q, VbaImportComponent));
    }

    private static Task<object?> Json(PluginRequest req, Func<JsonElement, object?> handler)
    {
        using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(req.Body) ? "{}" : req.Body);
        return Task.FromResult(handler(doc.RootElement));
    }

    private static string Req(PluginRequest q, string key) =>
        q.Query.TryGetValue(key, out var v) && !string.IsNullOrEmpty(v) ? v : throw new ArgumentException($"missing '{key}' query parameter");

    private static dynamic GetExcel() => ComHelper.GetActiveObject("Excel.Application");

    private static dynamic GetWorkbook(dynamic excel, string name)
    {
        foreach (dynamic wb in excel.Workbooks)
            if ((string)wb.Name == name) return wb;
        throw new Exception($"Workbook '{name}' not found");
    }

    private static object Detect()
    {
        dynamic excel;
        try { excel = GetExcel(); }
        catch { return new { excelRunning = false, workbooks = Array.Empty<object>() }; }

        var workbooks = new List<object>();
        foreach (dynamic wb in excel.Workbooks)
        {
            var sheets = new List<object>();
            foreach (dynamic ws in wb.Worksheets) sheets.Add(new { name = (string)ws.Name, index = (int)ws.Index });
            workbooks.Add(new { name = (string)wb.Name, fullName = (string)wb.FullName, saved = (bool)wb.Saved, sheets });
        }
        return new { excelRunning = true, workbooks };
    }

    private static object? GetSheets(string workbook)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, workbook);
        var sheets = new List<object>();
        foreach (dynamic ws in wb.Worksheets) sheets.Add(new { name = (string)ws.Name, index = (int)ws.Index, visible = (int)ws.Visible == -1 });
        return new { sheets };
    }

    private static object? GetAppInfo()
    {
        dynamic excel = GetExcel();
        var calcMap = new Dictionary<int, string> { [-4105] = "Automatic", [-4135] = "Manual", [2] = "Semiautomatic" };
        return new
        {
            version = (string)excel.Version,
            build = (string)excel.Build.ToString(),
            path = (string)excel.Path,
            calcMode = calcMap.GetValueOrDefault((int)excel.Calculation, "Unknown"),
            screenUpdating = (bool)excel.ScreenUpdating,
            workbookCount = (int)excel.Workbooks.Count
        };
    }

    private static object? ListNamedRanges(string workbook)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, workbook);
        var names = new List<object>();
        foreach (dynamic n in wb.Names) names.Add(new { name = (string)n.Name, refersTo = (string)n.RefersTo, address = (string)n.RefersToRange.Address });
        return new { names };
    }

    private static object? GetRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        dynamic rng = ws.Range[body.GetProperty("range").GetString()!];
        int rows = rng.Rows.Count, cols = rng.Columns.Count;
        var data = new List<List<string?>>();
        for (int r = 1; r <= rows; r++)
        {
            var row = new List<string?>();
            for (int c = 1; c <= cols; c++) row.Add(rng.Cells[r, c].Value2?.ToString());
            data.Add(row);
        }
        return new { rows, cols, data };
    }

    private static object? SetRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        var values = body.GetProperty("values");
        excel.ScreenUpdating = false;
        for (int r = 0; r < values.GetArrayLength(); r++)
        {
            var row = values[r];
            for (int c = 0; c < row.GetArrayLength(); c++)
            {
                var cell = row[c];
                object? val = cell.ValueKind switch
                {
                    JsonValueKind.Null => null,
                    JsonValueKind.Number => cell.GetDouble(),
                    JsonValueKind.True => true,
                    JsonValueKind.False => false,
                    _ => cell.GetString()
                };
                ws.Cells[r + 1, c + 1].Value2 = val;
            }
        }
        excel.ScreenUpdating = true;
        return new { message = "Range written" };
    }

    private static object? GetFormula(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        return new { formula = (string)ws.Range[body.GetProperty("range").GetString()!].Formula };
    }

    private static object? SetFormula(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        ws.Range[body.GetProperty("range").GetString()!].Formula = body.GetProperty("formula").GetString()!;
        return new { message = "Formula set" };
    }

    private static object? GetNamedRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic rng = wb.Names.Item(body.GetProperty("name").GetString()!).RefersToRange;
        int rows = rng.Rows.Count, cols = rng.Columns.Count;
        var data = new List<List<string?>>();
        for (int r = 1; r <= rows; r++)
        {
            var row = new List<string?>();
            for (int c = 1; c <= cols; c++) row.Add(rng.Cells[r, c].Value2?.ToString());
            data.Add(row);
        }
        return new { data, address = (string)rng.Address };
    }

    private static object? GetUsedRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        dynamic used = ws.UsedRange;
        int rows = used.Rows.Count, cols = used.Columns.Count;
        var data = new List<List<string?>>();
        for (int r = 1; r <= rows; r++)
        {
            var row = new List<string?>();
            for (int c = 1; c <= cols; c++) row.Add(used.Cells[r, c].Value2?.ToString());
            data.Add(row);
        }
        return new { address = (string)used.Address[true, true], rows, cols, data };
    }

    private static object? RunMacro(JsonElement body)
    {
        dynamic excel = GetExcel();
        excel.Run(body.GetProperty("macro").GetString()!);
        return new { message = "Macro executed" };
    }

    private static object? Save(JsonElement body)
    {
        dynamic excel = GetExcel();
        GetWorkbook(excel, body.GetProperty("workbook").GetString()!).Save();
        return new { message = "Saved" };
    }

    private static object? Refresh(JsonElement body)
    {
        dynamic excel = GetExcel();
        GetWorkbook(excel, body.GetProperty("workbook").GetString()!).RefreshAll();
        return new { message = "Refresh triggered" };
    }

    private static object? CreateSheet(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets.Add();
        if (body.TryGetProperty("name", out var n)) ws.Name = n.GetString()!;
        return new { message = $"Sheet '{ws.Name}' created", name = (string)ws.Name };
    }

    private static object? RenameSheet(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        wb.Worksheets[body.GetProperty("sheetName").GetString()!].Name = body.GetProperty("newName").GetString()!;
        return new { message = "Renamed" };
    }

    private static object? DeleteSheet(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        excel.DisplayAlerts = false;
        wb.Worksheets[body.GetProperty("sheetName").GetString()!].Delete();
        excel.DisplayAlerts = true;
        return new { message = "Sheet deleted" };
    }

    private static object? CopySheet(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        ws.Copy(wb.Worksheets[1]);
        dynamic newWs = wb.Worksheets[1];
        if (body.TryGetProperty("newName", out var nn)) newWs.Name = nn.GetString()!;
        return new { message = "Copied", name = (string)newWs.Name };
    }

    private static object? HideSheet(JsonElement body)
    {
        dynamic excel = GetExcel();
        GetWorkbook(excel, body.GetProperty("workbook").GetString()!).Worksheets[body.GetProperty("sheetName").GetString()!].Visible = 0;
        return new { message = "Sheet hidden" };
    }

    private static object? UnhideSheet(JsonElement body)
    {
        dynamic excel = GetExcel();
        GetWorkbook(excel, body.GetProperty("workbook").GetString()!).Worksheets[body.GetProperty("sheetName").GetString()!].Visible = -1;
        return new { message = "Sheet visible" };
    }

    private static object? ProtectSheet(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        var pwd = body.TryGetProperty("password", out var p) ? p.GetString() : null;
        wb.Worksheets[body.GetProperty("sheetName").GetString()!].Protect(pwd);
        return new { message = "Sheet protected" };
    }

    private static object? UnprotectSheet(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        var pwd = body.TryGetProperty("password", out var p) ? p.GetString() : null;
        wb.Worksheets[body.GetProperty("sheetName").GetString()!].Unprotect(pwd);
        return new { message = "Sheet unprotected" };
    }

    private static object? ClearRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic rng = wb.Worksheets[body.GetProperty("sheetName").GetString()!].Range[body.GetProperty("range").GetString()!];
        if (body.TryGetProperty("contentsOnly", out var co) && co.GetBoolean()) rng.ClearContents();
        else rng.Clear();
        return new { message = "Range cleared" };
    }

    private static object? DeleteRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        var shiftMap = new Dictionary<string, int> { ["Up"] = -4162, ["Left"] = -4159, ["Down"] = 4121, ["Right"] = 4161 };
        var shift = body.TryGetProperty("shift", out var s) && shiftMap.ContainsKey(s.GetString()!) ? shiftMap[s.GetString()!] : -4162;
        ws.Range[body.GetProperty("range").GetString()!].Delete(shift);
        return new { message = "Range deleted" };
    }

    private static object? InsertRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        var shiftMap = new Dictionary<string, int> { ["Down"] = -4121, ["Right"] = -4161 };
        var shift = body.TryGetProperty("shift", out var s) && shiftMap.ContainsKey(s.GetString()!) ? shiftMap[s.GetString()!] : -4121;
        ws.Range[body.GetProperty("range").GetString()!].Insert(shift);
        return new { message = "Range inserted" };
    }

    private static object? GetCellFormat(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic cell = wb.Worksheets[body.GetProperty("sheetName").GetString()!].Range[body.GetProperty("range").GetString()!];
        return new
        {
            numberFormat = (string)cell.NumberFormat,
            bold = (bool)cell.Font.Bold,
            italic = (bool)cell.Font.Italic,
            fontSize = (double)cell.Font.Size,
            fontName = (string)cell.Font.Name,
            fontColor = (double)cell.Font.Color,
            bgColor = (double)cell.Interior.Color,
            hAlign = (int)cell.HorizontalAlignment,
            vAlign = (int)cell.VerticalAlignment,
            wrapText = (bool)cell.WrapText
        };
    }

    private static object? SetCellFormat(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic cell = wb.Worksheets[body.GetProperty("sheetName").GetString()!].Range[body.GetProperty("range").GetString()!];
        if (body.TryGetProperty("numberFormat", out var nf)) cell.NumberFormat = nf.GetString()!;
        if (body.TryGetProperty("bold", out var b)) cell.Font.Bold = b.GetBoolean();
        if (body.TryGetProperty("italic", out var it)) cell.Font.Italic = it.GetBoolean();
        if (body.TryGetProperty("fontSize", out var fs)) cell.Font.Size = fs.GetDouble();
        if (body.TryGetProperty("fontName", out var fn)) cell.Font.Name = fn.GetString()!;
        if (body.TryGetProperty("fontColor", out var fc)) cell.Font.Color = fc.GetDouble();
        if (body.TryGetProperty("bgColor", out var bg)) cell.Interior.Color = bg.GetDouble();
        if (body.TryGetProperty("wrapText", out var wt)) cell.WrapText = wt.GetBoolean();
        return new { message = "Format applied" };
    }

    private static object? AddNamedRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        wb.Names.Add(body.GetProperty("name").GetString()!, ws.Range[body.GetProperty("range").GetString()!]);
        return new { message = "Named range added" };
    }

    private static object? DeleteNamedRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        wb.Names.Item(body.GetProperty("name").GetString()!).Delete();
        return new { message = "Named range deleted" };
    }

    private static object? FindReplace(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        dynamic rng = body.TryGetProperty("range", out var r) ? ws.Range[r.GetString()!] : ws.UsedRange;
        int lookAt = body.TryGetProperty("partial", out var p) && p.GetBoolean() ? 2 : 1;
        rng.Replace(body.GetProperty("find").GetString()!, body.GetProperty("replace").GetString()!, lookAt);
        return new { message = "Find/replace complete" };
    }

    private static object? SortRange(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        dynamic rng = ws.Range[body.GetProperty("range").GetString()!];
        dynamic key = ws.Range[body.GetProperty("keyColumn").GetString()!];
        int order = body.TryGetProperty("descending", out var d) && d.GetBoolean() ? 2 : 1;
        int header = body.TryGetProperty("hasHeader", out var h) && h.GetBoolean() ? 1 : 2;
        rng.Sort(key, order, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, header);
        return new { message = "Range sorted" };
    }

    private static object? Autofit(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic ws = wb.Worksheets[body.GetProperty("sheetName").GetString()!];
        dynamic rng = body.TryGetProperty("range", out var r) ? ws.Range[r.GetString()!].EntireColumn : ws.UsedRange.EntireColumn;
        rng.AutoFit();
        if (body.TryGetProperty("rows", out var rows) && rows.GetBoolean()) ws.UsedRange.EntireRow.AutoFit();
        return new { message = "AutoFit applied" };
    }

    private static object? Calculate(JsonElement body)
    {
        dynamic excel = GetExcel();
        if (body.TryGetProperty("workbook", out var wbName) && body.TryGetProperty("sheetName", out var sn))
            GetWorkbook(excel, wbName.GetString()!).Worksheets[sn.GetString()!].Calculate();
        else if (body.TryGetProperty("workbook", out var wbn))
            GetWorkbook(excel, wbn.GetString()!).Calculate();
        else
            excel.Calculate();
        return new { message = "Calculated" };
    }

    private static object? SetCalcMode(JsonElement body)
    {
        dynamic excel = GetExcel();
        var modeMap = new Dictionary<string, int> { ["Automatic"] = -4105, ["Manual"] = -4135, ["Semiautomatic"] = 2 };
        var mode = body.GetProperty("mode").GetString()!;
        excel.Calculation = modeMap[mode];
        return new { message = $"Calc mode set to {mode}" };
    }

    private static object? SaveAs(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        var fmtMap = new Dictionary<string, int> { ["xlsx"] = 51, ["xlsm"] = 52, ["xls"] = 56, ["pdf"] = 57, ["csv"] = 6 };
        var fmt = body.TryGetProperty("format", out var f) && fmtMap.ContainsKey(f.GetString()!) ? fmtMap[f.GetString()!] : 51;
        excel.DisplayAlerts = false;
        wb.SaveAs(body.GetProperty("filePath").GetString()!, fmt);
        excel.DisplayAlerts = true;
        return new { message = "Saved" };
    }

    private static object? ProtectWorkbook(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        var pwd = body.TryGetProperty("password", out var p) ? p.GetString() : null;
        wb.Protect(pwd);
        return new { message = "Workbook protected" };
    }

    private static object? UnprotectWorkbook(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        var pwd = body.TryGetProperty("password", out var p) ? p.GetString() : null;
        wb.Unprotect(pwd);
        return new { message = "Workbook unprotected" };
    }

    private static object? GetDocProperties(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic props = wb.BuiltinDocumentProperties;
        return new
        {
            title = (string)props["Title"].Value,
            subject = (string)props["Subject"].Value,
            author = (string)props["Author"].Value,
            company = (string)props["Company"].Value,
            created = ((DateTime)props["Creation Date"].Value).ToString("yyyy-MM-dd HH:mm:ss"),
            modified = ((DateTime)props["Last Save Time"].Value).ToString("yyyy-MM-dd HH:mm:ss")
        };
    }

    private static object? ExportPdf(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic target = body.TryGetProperty("sheetName", out var sn) ? wb.Worksheets[sn.GetString()!] : wb;
        target.ExportAsFixedFormat(0, body.GetProperty("filePath").GetString()!);
        return new { message = "Exported to PDF" };
    }

    private static object? SetScreenUpdating(JsonElement body)
    {
        dynamic excel = GetExcel();
        excel.ScreenUpdating = body.GetProperty("enabled").GetBoolean();
        return new { message = "ScreenUpdating updated" };
    }

    private static object? SetStatusBar(JsonElement body)
    {
        dynamic excel = GetExcel();
        excel.StatusBar = body.TryGetProperty("text", out var t) ? t.GetString()! : (object)false;
        return new { message = "Status bar updated" };
    }

    // ── VBA ──
    private static object? VbaListComponents(string workbook)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, workbook);
        var typeMap = new Dictionary<int, string> { [1] = "Module", [2] = "Class", [3] = "Form", [100] = "Document" };
        var comps = new List<object>();
        foreach (dynamic c in wb.VBProject.VBComponents)
            comps.Add(new { name = (string)c.Name, type = typeMap.GetValueOrDefault((int)c.Type, "Unknown"), lines = (int)c.CodeModule.CountOfLines });
        return new { components = comps };
    }

    private static object? VbaGetModule(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic cm = wb.VBProject.VBComponents.Item(body.GetProperty("component").GetString()!).CodeModule;
        int lines = cm.CountOfLines;
        string code = lines > 0 ? cm.Lines[1, lines] : "";
        return new { component = body.GetProperty("component").GetString()!, lines, code };
    }

    private static object? VbaSetModule(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic cm = wb.VBProject.VBComponents.Item(body.GetProperty("component").GetString()!).CodeModule;
        if (cm.CountOfLines > 0) cm.DeleteLines(1, cm.CountOfLines);
        cm.AddFromString(body.GetProperty("code").GetString()!);
        return new { message = "Module updated" };
    }

    private static object? VbaAddModule(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic comp = wb.VBProject.VBComponents.Add(1);
        if (body.TryGetProperty("name", out var n)) comp.Name = n.GetString()!;
        if (body.TryGetProperty("code", out var c)) comp.CodeModule.AddFromString(c.GetString()!);
        return new { message = "Module added", name = (string)comp.Name };
    }

    private static object? VbaAddClass(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic comp = wb.VBProject.VBComponents.Add(2);
        if (body.TryGetProperty("name", out var n)) comp.Name = n.GetString()!;
        if (body.TryGetProperty("code", out var c)) comp.CodeModule.AddFromString(c.GetString()!);
        return new { message = "Class added", name = (string)comp.Name };
    }

    private static object? VbaAddForm(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic comp = wb.VBProject.VBComponents.Add(3);
        if (body.TryGetProperty("name", out var n)) comp.Name = n.GetString()!;
        return new { message = "UserForm added", name = (string)comp.Name };
    }

    private static object? VbaDeleteComponent(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic comp = wb.VBProject.VBComponents.Item(body.GetProperty("component").GetString()!);
        wb.VBProject.VBComponents.Remove(comp);
        return new { message = "Component deleted" };
    }

    private static object? VbaExportComponent(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        wb.VBProject.VBComponents.Item(body.GetProperty("component").GetString()!).Export(body.GetProperty("filePath").GetString()!);
        return new { message = "Exported" };
    }

    private static object? VbaImportComponent(JsonElement body)
    {
        dynamic excel = GetExcel();
        dynamic wb = GetWorkbook(excel, body.GetProperty("workbook").GetString()!);
        dynamic comp = wb.VBProject.VBComponents.Import(body.GetProperty("filePath").GetString()!);
        return new { message = "Imported", name = (string)comp.Name };
    }
}
