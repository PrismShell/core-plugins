using System.Diagnostics;
using System.Text;
using System.Text.Json;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Utils;
using Plugboard.Contracts;

namespace Plugboard.Plugins.Pdf;

// Ported from the local-gateway PdfHandler (itext7 + headless Chrome/Edge for
// HTML->PDF). Each route returns its data object (host wraps in { ok, data });
// errors throw (host wraps in { ok:false, error }).
public sealed class PdfPlugin : IPlugin
{
    public string Name => "pdf";

    public void Register(IEndpointRegistry r)
    {
        r.Map("POST", "pdf/generate",       GenerateAsync);
        r.Map("POST", "pdf/read",           q => Json(q, Read));
        r.Map("POST", "pdf/extract-tables", q => Json(q, ExtractTables));
        r.Map("POST", "pdf/metadata",       q => Json(q, GetMetadata));
        r.Map("POST", "pdf/merge",          q => Json(q, Merge));
        r.Map("POST", "pdf/split",          q => Json(q, Split));
        r.Map("POST", "pdf/rotate",         q => Json(q, Rotate));
        r.Map("POST", "pdf/watermark",      q => Json(q, Watermark));
        r.Map("POST", "pdf/protect",        q => Json(q, Protect));
        r.Map("POST", "pdf/unlock",         q => Json(q, Unlock));
        r.Map("POST", "pdf/extract-pages",  q => Json(q, ExtractPages));
        r.Map("POST", "pdf/compress",       q => Json(q, Compress));
    }

    private static Task<object?> Json(PluginRequest req, Func<JsonElement, object?> handler)
    {
        using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(req.Body) ? "{}" : req.Body);
        return Task.FromResult(handler(doc.RootElement));
    }

    private static async Task<object?> GenerateAsync(PluginRequest req)
    {
        using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(req.Body) ? "{}" : req.Body);
        return await Generate(doc.RootElement);
    }

    // HTML/URL -> PDF via system-installed Chrome or Edge (no bundled browser).
    static async Task<object?> Generate(JsonElement body)
    {
        var url        = body.TryGetProperty("url",  out var u) ? u.GetString() : null;
        var html       = body.TryGetProperty("html", out var h) ? h.GetString() : null;
        if (url == null && html == null) throw new Exception("Provide either 'url' or 'html'");
        var outputPath = Environment.ExpandEnvironmentVariables(body.GetProperty("outputPath").GetString()!);
        var landscape  = body.TryGetProperty("landscape", out var l) && l.GetBoolean();

        var candidates = new[]
        {
            @"C:\Program Files\Google\Chrome\Application\chrome.exe",
            @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Chrome\Application\chrome.exe"),
            @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            @"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        };
        var execPath = candidates.FirstOrDefault(File.Exists)
            ?? throw new Exception("Chrome or Edge not found. Install either browser to use PDF generation.");

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        string? tempHtmlPath = null;
        string target;
        if (html != null)
        {
            tempHtmlPath = Path.Combine(Path.GetTempPath(), $"plugboard_{Guid.NewGuid():N}.html");
            await File.WriteAllTextAsync(tempHtmlPath, html, Encoding.UTF8);
            target = tempHtmlPath;
        }
        else target = url!;

        try
        {
            var args = string.Join(" ",
                "--headless=new", "--disable-gpu", "--no-sandbox", "--disable-extensions",
                "--disable-software-rasterizer", $"--print-to-pdf=\"{outputPath}\"",
                "--print-to-pdf-no-header", $"\"{target}\"");

            var psi = new ProcessStartInfo(execPath, args)
            { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };

            using var proc = Process.Start(psi) ?? throw new Exception("Failed to start browser process");
            var stderr = await proc.StandardError.ReadToEndAsync();
            await proc.WaitForExitAsync();

            if (!File.Exists(outputPath))
                throw new Exception($"PDF was not created. Browser output: {stderr.Trim()}");

            return new { outputPath, size = new FileInfo(outputPath).Length };
        }
        finally
        {
            if (tempHtmlPath != null && File.Exists(tempHtmlPath)) File.Delete(tempHtmlPath);
        }
    }

    static object? Read(JsonElement body)
    {
        var path = body.GetProperty("path").GetString()!;
        int[]? pageFilter = body.TryGetProperty("pages", out var pg) ? pg.EnumerateArray().Select(e => e.GetInt32()).ToArray() : null;

        using var reader = new PdfReader(path);
        using var doc    = new PdfDocument(reader);
        var pages = new List<object>();
        for (int i = 1; i <= doc.GetNumberOfPages(); i++)
        {
            if (pageFilter != null && !pageFilter.Contains(i)) continue;
            var text  = PdfTextExtractor.GetTextFromPage(doc.GetPage(i), new LocationTextExtractionStrategy());
            var words = text.Split([' ', '\n', '\r', '\t'], StringSplitOptions.RemoveEmptyEntries).ToList();
            pages.Add(new { page = i, text, words, wordCount = words.Count });
        }
        return new { pageCount = doc.GetNumberOfPages(), pages };
    }

    static object? ExtractTables(JsonElement body)
    {
        var path = body.GetProperty("path").GetString()!;
        using var reader = new PdfReader(path);
        using var doc    = new PdfDocument(reader);
        var result = new List<object>();
        for (int i = 1; i <= doc.GetNumberOfPages(); i++)
        {
            var text = PdfTextExtractor.GetTextFromPage(doc.GetPage(i), new LocationTextExtractionStrategy());
            var rows = text.Split('\n', StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries).ToList())
                .Where(rw => rw.Count > 0).ToList();
            result.Add(new { page = i, rows });
        }
        return new { pages = result };
    }

    static object? GetMetadata(JsonElement body)
    {
        var path = body.GetProperty("path").GetString()!;
        using var reader = new PdfReader(path);
        using var doc    = new PdfDocument(reader);
        var info = doc.GetDocumentInfo();
        return new
        {
            pageCount  = doc.GetNumberOfPages(),
            title      = info.GetTitle(),
            author     = info.GetAuthor(),
            subject    = info.GetSubject(),
            keywords   = info.GetKeywords(),
            creator    = info.GetCreator(),
            producer   = info.GetProducer(),
            pdfVersion = doc.GetPdfVersion().ToString()
        };
    }

    static object? Merge(JsonElement body)
    {
        var inputs = body.GetProperty("inputPaths").EnumerateArray().Select(e => e.GetString()!).ToList();
        var output = body.GetProperty("outputPath").GetString()!;
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);

        using var writer = new PdfWriter(output);
        using var outDoc = new PdfDocument(writer);
        var merger = new PdfMerger(outDoc);
        foreach (var input in inputs)
        {
            using var rd = new PdfReader(input);
            using var d  = new PdfDocument(rd);
            merger.Merge(d, 1, d.GetNumberOfPages());
        }
        return new { outputPath = output, pageCount = outDoc.GetNumberOfPages() };
    }

    static object? Split(JsonElement body)
    {
        var input  = body.GetProperty("inputPath").GetString()!;
        var outDir = body.GetProperty("outputDir").GetString()!;
        var ranges = body.GetProperty("ranges").EnumerateArray().Select(rg => rg.EnumerateArray().Select(e => e.GetInt32()).ToArray()).ToList();
        Directory.CreateDirectory(outDir);

        var outputs = new List<string>();
        using var reader = new PdfReader(input);
        using var inDoc  = new PdfDocument(reader);
        for (int ri = 0; ri < ranges.Count; ri++)
        {
            var outPath = Path.Combine(outDir, $"part_{ri + 1}.pdf");
            using var writer = new PdfWriter(outPath);
            using var outDoc = new PdfDocument(writer);
            inDoc.CopyPagesTo(ranges[ri][0], ranges[ri][1], outDoc);
            outputs.Add(outPath);
        }
        return new { outputs };
    }

    static object? Rotate(JsonElement body)
    {
        var input   = body.GetProperty("inputPath").GetString()!;
        var output  = body.GetProperty("outputPath").GetString()!;
        var degrees = body.TryGetProperty("degrees", out var d) ? d.GetInt32() : 90;
        int[]? pages = body.TryGetProperty("pages", out var pg) ? pg.EnumerateArray().Select(e => e.GetInt32()).ToArray() : null;
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);

        using var reader = new PdfReader(input);
        using var writer = new PdfWriter(output);
        using var doc    = new PdfDocument(reader, writer);
        for (int i = 1; i <= doc.GetNumberOfPages(); i++)
        {
            if (pages != null && !pages.Contains(i)) continue;
            var page = doc.GetPage(i);
            page.SetRotation((page.GetRotation() + degrees) % 360);
        }
        return new { outputPath = output };
    }

    static object? Watermark(JsonElement body)
    {
        var input    = body.GetProperty("inputPath").GetString()!;
        var output   = body.GetProperty("outputPath").GetString()!;
        var text     = body.TryGetProperty("text",     out var t)  ? t.GetString()!        : "CONFIDENTIAL";
        var opacity  = body.TryGetProperty("opacity",  out var op) ? (float)op.GetDouble() : 0.15f;
        var fontSize = body.TryGetProperty("fontSize", out var fs) ? fs.GetInt32()         : 48;
        int[]? pages = body.TryGetProperty("pages", out var pg) ? pg.EnumerateArray().Select(e => e.GetInt32()).ToArray() : null;
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);

        using var reader = new PdfReader(input);
        using var writer = new PdfWriter(output);
        using var doc    = new PdfDocument(reader, writer);
        var font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
        for (int i = 1; i <= doc.GetNumberOfPages(); i++)
        {
            if (pages != null && !pages.Contains(i)) continue;
            var page   = doc.GetPage(i);
            var size   = page.GetPageSize();
            var canvas = new PdfCanvas(page);
            canvas.SaveState();
            var gs = new iText.Kernel.Pdf.Extgstate.PdfExtGState().SetFillOpacity(opacity);
            canvas.SetExtGState(gs)
                .BeginText()
                .SetFontAndSize(font, fontSize)
                .SetTextMatrix(
                    (float)Math.Cos(Math.PI / 4), (float)Math.Sin(Math.PI / 4),
                    -(float)Math.Sin(Math.PI / 4), (float)Math.Cos(Math.PI / 4),
                    size.GetWidth() / 2 - fontSize * text.Length / 4f,
                    size.GetHeight() / 2)
                .ShowText(text)
                .EndText();
            canvas.RestoreState();
        }
        return new { outputPath = output };
    }

    static object? Protect(JsonElement body)
    {
        var input   = body.GetProperty("inputPath").GetString()!;
        var output  = body.GetProperty("outputPath").GetString()!;
        var userPw  = body.TryGetProperty("userPassword",  out var up) ? up.GetString() ?? "" : "";
        var ownPw   = body.TryGetProperty("ownerPassword", out var op) ? op.GetString() ?? Guid.NewGuid().ToString() : Guid.NewGuid().ToString();
        var allowPrint = !body.TryGetProperty("allowPrinting", out var ap) || ap.GetBoolean();
        var allowCopy  =  body.TryGetProperty("allowCopying",  out var ac) && ac.GetBoolean();

        int perms = 0;
        if (allowPrint) perms |= EncryptionConstants.ALLOW_PRINTING;
        if (allowCopy)  perms |= EncryptionConstants.ALLOW_COPY;
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);

        using var reader = new PdfReader(input);
        using var writer = new PdfWriter(output, new WriterProperties()
            .SetStandardEncryption(Encoding.UTF8.GetBytes(userPw), Encoding.UTF8.GetBytes(ownPw), perms, EncryptionConstants.ENCRYPTION_AES_256));
        using var doc = new PdfDocument(reader, writer);
        return new { outputPath = output };
    }

    static object? Unlock(JsonElement body)
    {
        var input  = body.GetProperty("inputPath").GetString()!;
        var output = body.GetProperty("outputPath").GetString()!;
        var pw     = body.TryGetProperty("password", out var p) ? p.GetString() ?? "" : "";
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);

        using var reader = new PdfReader(input, new ReaderProperties().SetPassword(Encoding.UTF8.GetBytes(pw)));
        using var writer = new PdfWriter(output);
        using var doc    = new PdfDocument(reader, writer);
        return new { outputPath = output };
    }

    static object? ExtractPages(JsonElement body)
    {
        var input  = body.GetProperty("inputPath").GetString()!;
        var output = body.GetProperty("outputPath").GetString()!;
        var pages  = body.GetProperty("pages").EnumerateArray().Select(e => e.GetInt32()).ToList();
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);

        using var reader = new PdfReader(input);
        using var writer = new PdfWriter(output);
        using var inDoc  = new PdfDocument(reader);
        using var outDoc = new PdfDocument(writer);
        inDoc.CopyPagesTo(pages, outDoc);
        return new { outputPath = output, pageCount = pages.Count };
    }

    static object? Compress(JsonElement body)
    {
        var input  = body.GetProperty("inputPath").GetString()!;
        var output = body.GetProperty("outputPath").GetString()!;
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);

        using (var reader = new PdfReader(input))
        using (var writer = new PdfWriter(output, new WriterProperties().SetCompressionLevel(9)))
        using (var doc    = new PdfDocument(reader, writer))
        {
            doc.SetFlushUnusedObjects(true);
        }
        var inSize  = new FileInfo(input).Length;
        var outSize = new FileInfo(output).Length;
        return new { outputPath = output, inputBytes = inSize, outputBytes = outSize, savedBytes = inSize - outSize };
    }
}
