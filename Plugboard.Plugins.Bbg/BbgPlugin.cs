using Bloomberglp.Blpapi;
using System.Text.Json;
using Plugboard.Contracts;
using Plugboard.Blpapi;

namespace Plugboard.Plugins.Bbg;

// Reference/market data (BDP/BDS/BDH), intraday, field/security search, and bbg:// launch.
// Ported from the local-gateway BloombergHandler. Uses the shared BLPAPI core
// (Plugboard.Blpapi.Blp) - one process-wide session, opened services cached - so it
// coexists with the cmp connector (CMP is another service on the same session). Each route
// returns its data (the host wraps it in { ok, data }); errors throw (host wraps { ok:false }).
public sealed class BbgPlugin : IPlugin
{
    public string Name => "bbg";

    public void Register(IEndpointRegistry r)
    {
        r.Map("GET",  "bbg/status",          _   => Task.FromResult<object?>(GetStatus()));
        r.Map("POST", "bbg/bdp", req => Json(req, RunBdp), new RouteInfo(
            "Reference data point",
            "Current field values (BDP) for one or more securities.",
            new { securities = new[] { "AAPL US Equity" }, fields = new[] { "PX_LAST", "NAME" } },
            new[]
            {
                new ParamInfo("securities", "array", true, "BBG security identifiers.", Items: "string"),
                new ParamInfo("fields", "array", true, "Field mnemonics to fetch.", Items: "string"),
                new ParamInfo("overrides", "object", false, "Optional field overrides, e.g. { \"SETTLE_DT\": \"2026/05/15\" }.")
            }));
        r.Map("POST", "bbg/bds", req => Json(req, RunBds), new RouteInfo(
            "Bulk data set",
            "Multi-row bulk field data (BDS) for one or more securities.",
            new { securities = new[] { "3142GXH34 Mtge" }, fields = new[] { "MTG_CASH_FLOW" } },
            new[]
            {
                new ParamInfo("securities", "array", true, "BBG security identifiers.", Items: "string"),
                new ParamInfo("fields", "array", true, "Bulk field mnemonics.", Items: "string"),
                new ParamInfo("overrides", "object", false, "Optional field overrides.")
            }));
        r.Map("POST", "bbg/bdh", req => Json(req, RunBdh), new RouteInfo(
            "Historical data",
            "Historical time series (BDH) for one or more securities.",
            new { securities = new[] { "AAPL US Equity" }, fields = new[] { "PX_LAST" }, startDate = "2025-01-01", endDate = "2025-12-31" },
            new[]
            {
                new ParamInfo("securities", "array", true, "BBG security identifiers.", Items: "string"),
                new ParamInfo("fields", "array", true, "Field mnemonics.", Items: "string"),
                new ParamInfo("startDate", "string", false, "Start date (YYYY-MM-DD or YYYYMMDD). Defaults to one year ago."),
                new ParamInfo("endDate", "string", false, "End date (YYYY-MM-DD or YYYYMMDD). Defaults to today."),
                new ParamInfo("periodicitySelection", "string", false, "Sampling frequency.", Enum: new[] { "DAILY", "WEEKLY", "MONTHLY", "QUARTERLY", "YEARLY" }, Default: "DAILY"),
                new ParamInfo("overrides", "object", false, "Optional field overrides.")
            }));
        r.Map("POST", "bbg/intraday-bar",    req => Json(req, RunIntradayBar));
        r.Map("POST", "bbg/intraday-tick",   req => Json(req, RunIntradayTick));
        r.Map("POST", "bbg/field-search",    req => Json(req, RunFieldSearch));
        r.Map("POST", "bbg/security-lookup", req => Json(req, RunSecurityLookup));
        // CMP moved to its own connector (Plugboard.Plugins.Cmp, /con/cmp/*) - it shares
        // this session but is a separate plugin with its own //blp/cmp service.
        r.Map("POST", "bbg/open",            req => Json(req, RunOpen));
    }

    private static Task<object?> Json(PluginRequest req, Func<JsonElement, object?> handler)
    {
        using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(req.Body) ? "{}" : req.Body);
        return Task.FromResult(handler(doc.RootElement));
    }

    // ── endpoints ──

    private static object GetStatus()
    {
        var p = Blp.ProbeService("//blp/refdata");
        return new
        {
            connected = p.Connected,
            sessionReady = p.ServiceReady,
            lastKnownReady = p.LastKnownReady,
            host = Blp.Host, port = Blp.Port,
            error = p.Error
        };
    }

    // Launch a BBG Terminal deep link (bbg:// scheme) via the local OS protocol
    // handler. Pass either a raw {url} or {security, mnemonic} to build
    // bbg://securities/<security>/<mnemonic> (mnemonic defaults to DES). Restricted to the
    // bbg:// scheme so the shell can't launch other protocols or local files. No BLPAPI.
    private static object? RunOpen(JsonElement body)
    {
        string url;
        if (body.TryGetProperty("url", out var u) && u.ValueKind == JsonValueKind.String)
        {
            url = u.GetString()!.Trim();
        }
        else
        {
            var security = body.GetProperty("security").GetString()!.Trim();
            var mnemonic = body.TryGetProperty("mnemonic", out var m) && m.ValueKind == JsonValueKind.String
                ? m.GetString()!.Trim() : "DES";
            url = $"bbg://securities/{Uri.EscapeDataString(security)}/{Uri.EscapeDataString(mnemonic)}";
        }
        if (!url.StartsWith("bbg://", StringComparison.OrdinalIgnoreCase))
            throw new Exception("only bbg:// URLs may be launched");

        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true   // route through the registered bbg:// handler
        });
        return new { success = true, url };
    }

    private static object? RunBdp(JsonElement body)
    {
        var securities = body.GetProperty("securities").EnumerateArray().Select(e => e.GetString()!).ToList();
        var fields     = body.GetProperty("fields").EnumerateArray().Select(e => e.GetString()!).ToList();
        return Blp.WithService("//blp/refdata", (session, svc) =>
        {
            var req = svc.CreateRequest("ReferenceDataRequest");
            foreach (var sec in securities) req.Append("securities", sec);
            foreach (var fld in fields)     req.Append("fields", fld);
            if (body.TryGetProperty("overrides", out var ovProp) && ovProp.ValueKind == JsonValueKind.Object)
            {
                var ov = req.GetElement("overrides");
                foreach (var o in ovProp.EnumerateObject())
                { var e = ov.AppendElement(); e.SetElement("fieldId", o.Name); e.SetElement("value", Blp.NormalizeOverrideValue(o.Value)); }
            }
            var rows = new Dictionary<string, Dictionary<string, string?>>();
            foreach (var msg in Blp.SendAndReceive(session, req))
            {
                if (!msg.HasElement("securityData")) continue;
                var arr = msg.GetElement("securityData");
                for (int i = 0; i < arr.NumValues; i++)
                {
                    var se = arr.GetValueAsElement(i);
                    var fd = se.GetElement("fieldData");
                    var row = new Dictionary<string, string?>();
                    foreach (var fld in fields)
                        row[fld] = fd.HasElement(fld) ? fd.GetElement(fld).GetValueAsString() : null;
                    rows[se.GetElementAsString("security")] = row;
                }
            }
            return (object?)rows;
        });
    }

    private static object? RunBds(JsonElement body)
    {
        var securities = body.GetProperty("securities").EnumerateArray().Select(e => e.GetString()!).ToList();
        var fields     = body.GetProperty("fields").EnumerateArray().Select(e => e.GetString()!).ToList();
        return Blp.WithService("//blp/refdata", (session, svc) =>
        {
            var req = svc.CreateRequest("ReferenceDataRequest");
            foreach (var sec in securities) req.Append("securities", sec);
            foreach (var fld in fields)     req.Append("fields", fld);
            if (body.TryGetProperty("overrides", out var ovProp) && ovProp.ValueKind == JsonValueKind.Object)
            {
                var ov = req.GetElement("overrides");
                foreach (var o in ovProp.EnumerateObject())
                { var e = ov.AppendElement(); e.SetElement("fieldId", o.Name); e.SetElement("value", Blp.NormalizeOverrideValue(o.Value)); }
            }
            var data = new Dictionary<string, Dictionary<string, object?>>();
            foreach (var msg in Blp.SendAndReceive(session, req))
            {
                if (!msg.HasElement("securityData")) continue;
                var arr = msg.GetElement("securityData");
                for (int i = 0; i < arr.NumValues; i++)
                {
                    var se = arr.GetValueAsElement(i);
                    var fd = se.GetElement("fieldData");
                    var row = new Dictionary<string, object?>();
                    foreach (var fld in fields)
                    {
                        if (!fd.HasElement(fld)) { row[fld] = null; continue; }
                        var fe = fd.GetElement(fld);
                        if (fe.IsArray)
                        {
                            var rows = new List<Dictionary<string, string>>();
                            for (int ri = 0; ri < fe.NumValues; ri++)
                            {
                                var re = fe.GetValueAsElement(ri);
                                var rr = new Dictionary<string, string>();
                                for (int ci = 0; ci < re.NumElements; ci++) { var c = re.GetElement(ci); rr[c.Name.ToString()] = c.GetValueAsString(); }
                                rows.Add(rr);
                            }
                            row[fld] = rows;
                        }
                        else row[fld] = fe.GetValueAsString();
                    }
                    data[se.GetElementAsString("security")] = row;
                }
            }
            return (object?)data;
        });
    }

    private static object? RunBdh(JsonElement body)
    {
        var securities  = body.GetProperty("securities").EnumerateArray().Select(e => e.GetString()!).ToList();
        var fields      = body.GetProperty("fields").EnumerateArray().Select(e => e.GetString()!).ToList();
        var startDate   = body.TryGetProperty("startDate",   out var sd) ? sd.GetString()!.Replace("-","") : DateTime.Now.AddYears(-1).ToString("yyyyMMdd");
        var endDate     = body.TryGetProperty("endDate",     out var ed) ? ed.GetString()!.Replace("-","") : DateTime.Now.ToString("yyyyMMdd");
        var periodicity = body.TryGetProperty("periodicitySelection", out var ps) ? ps.GetString()! : "DAILY";
        return Blp.WithService("//blp/refdata", (session, svc) =>
        {
            var req = svc.CreateRequest("HistoricalDataRequest");
            foreach (var sec in securities) req.Append("securities", sec);
            foreach (var fld in fields)     req.Append("fields", fld);
            req.Set("startDate", startDate); req.Set("endDate", endDate); req.Set("periodicitySelection", periodicity);
            if (body.TryGetProperty("overrides", out var ovProp) && ovProp.ValueKind == JsonValueKind.Object)
            {
                var ov = req.GetElement("overrides");
                foreach (var o in ovProp.EnumerateObject())
                { var e = ov.AppendElement(); e.SetElement("fieldId", o.Name); e.SetElement("value", Blp.NormalizeOverrideValue(o.Value)); }
            }
            var data = new Dictionary<string, List<Dictionary<string, string?>>>();
            foreach (var msg in Blp.SendAndReceive(session, req))
            {
                if (!msg.HasElement("securityData")) continue;
                var se   = msg.GetElement("securityData");
                var fd   = se.GetElement("fieldData");
                var rows = new List<Dictionary<string, string?>>();
                for (int ri = 0; ri < fd.NumValues; ri++)
                {
                    var row = fd.GetValueAsElement(ri);
                    var obj = new Dictionary<string, string?> { ["date"] = row.GetElementAsString("date") };
                    foreach (var fld in fields) obj[fld] = row.HasElement(fld) ? row.GetElement(fld).GetValueAsString() : null;
                    rows.Add(obj);
                }
                data[se.GetElementAsString("security")] = rows;
            }
            return (object?)data;
        });
    }

    private static object? RunIntradayBar(JsonElement body)
    {
        var security  = body.GetProperty("security").GetString()!;
        var eventType = body.TryGetProperty("eventType", out var et) ? et.GetString()! : "TRADE";
        var interval  = body.TryGetProperty("interval",  out var iv) ? iv.GetInt32() : 60;
        var startDt   = DateTime.Parse(body.GetProperty("startDateTime").GetString()!).ToUniversalTime();
        var endDt     = DateTime.Parse(body.GetProperty("endDateTime").GetString()!).ToUniversalTime();
        var bars = Blp.WithService("//blp/refdata", (session, svc) =>
        {
            var req = svc.CreateRequest("IntradayBarRequest");
            req.Set("security", security); req.Set("eventType", eventType); req.Set("interval", interval);
            req.Set("startDateTime", new Datetime(startDt));
            req.Set("endDateTime",   new Datetime(endDt));
            var list = new List<object>();
            foreach (var msg in Blp.SendAndReceive(session, req))
            {
                if (!msg.HasElement("barData")) continue;
                var btd = msg.GetElement("barData").GetElement("barTickData");
                for (int i = 0; i < btd.NumValues; i++)
                {
                    var b = btd.GetValueAsElement(i);
                    list.Add(new { time = b.GetElementAsString("time"), open = b.GetElementAsFloat64("open"), high = b.GetElementAsFloat64("high"), low = b.GetElementAsFloat64("low"), close = b.GetElementAsFloat64("close"), volume = b.GetElementAsInt64("volume"), numEvents = b.GetElementAsInt32("numEvents") });
                }
            }
            return list;
        });
        return new { security, bars };
    }

    private static object? RunIntradayTick(JsonElement body)
    {
        var security   = body.GetProperty("security").GetString()!;
        var startDt    = body.GetProperty("startDateTime").GetString()!;
        var endDt      = body.GetProperty("endDateTime").GetString()!;
        var eventTypes = body.TryGetProperty("eventTypes", out var ets) ? ets.EnumerateArray().Select(e => e.GetString()!).ToList() : new List<string> { "TRADE" };
        var includeCc  = body.TryGetProperty("includeConditionCodes", out var cc) && cc.GetBoolean();
        var ticks = Blp.WithService("//blp/refdata", (session, svc) =>
        {
            var req    = svc.CreateRequest("IntradayTickRequest");
            req.Set("security", security); req.Set("startDateTime", startDt); req.Set("endDateTime", endDt);
            var etElem = req.GetElement("eventTypes");
            foreach (var e in eventTypes) etElem.AppendValue(e);
            if (includeCc) req.Set("includeConditionCodes", true);
            var list = new List<object>();
            foreach (var msg in Blp.SendAndReceive(session, req))
            {
                if (!msg.HasElement("tickData")) continue;
                var td = msg.GetElement("tickData").GetElement("tickData");
                for (int i = 0; i < td.NumValues; i++)
                {
                    var t = td.GetValueAsElement(i);
                    list.Add(new { time = t.GetElementAsString("time"), type = t.GetElementAsString("type"), value = t.GetElementAsFloat64("value"), size = t.GetElementAsInt32("size"), conditionCodes = t.HasElement("conditionCodes") ? t.GetElementAsString("conditionCodes") : null });
                }
            }
            return list;
        });
        return new { security, ticks };
    }

    private static object? RunFieldSearch(JsonElement body)
    {
        var query      = body.GetProperty("query").GetString()!;
        var returnDocs = body.TryGetProperty("returnFieldDocumentation", out var rd) && rd.GetBoolean();
        var fields = Blp.WithService("//blp/apiflds", (session, svc) =>
        {
            var req = svc.CreateRequest("FieldSearchRequest");
            req.Set("searchSpec", query);
            if (returnDocs) req.Set("returnFieldDocumentation", true);
            var list = new List<object>();
            foreach (var msg in Blp.SendAndReceive(session, req))
            {
                if (!msg.HasElement("fieldData")) continue;
                var fd = msg.GetElement("fieldData");
                for (int i = 0; i < fd.NumValues; i++)
                {
                    var f = fd.GetValueAsElement(i);
                    var fi = f.HasElement("fieldInfo") ? f.GetElement("fieldInfo") : null;
                    string g(string nm) => (fi != null && fi.HasElement(nm)) ? fi.GetElementAsString(nm) : "";
                    list.Add(new {
                        id           = f.HasElement("id") ? f.GetElementAsString("id") : "",
                        mnemonic     = g("mnemonic"),
                        description  = g("description"),
                        datatype     = g("datatype"),
                        categoryName = g("categoryName")
                    });
                }
            }
            return list;
        });
        return new { fields, count = fields.Count };
    }

    private static object? RunSecurityLookup(JsonElement body)
    {
        var query      = body.GetProperty("query").GetString()!;
        var maxResults = body.TryGetProperty("maxResults", out var mr) ? mr.GetInt32() : 20;
        var securities = Blp.WithService("//blp/instruments", (session, svc) =>
        {
            var req = svc.CreateRequest("instrumentListRequest");
            req.Set("query", query); req.Set("maxResults", maxResults);
            if (body.TryGetProperty("yellowKey", out var yk) && yk.GetString() is string ykVal && ykVal.Length > 0)
                req.Set("yellowKeyFilter", ykVal);
            var list = new List<object>();
            foreach (var msg in Blp.SendAndReceive(session, req))
            {
                if (!msg.HasElement("results")) continue;
                var res = msg.GetElement("results");
                for (int i = 0; i < res.NumValues; i++)
                { var rr = res.GetValueAsElement(i); list.Add(new { security = rr.GetElementAsString("security"), description = rr.GetElementAsString("description") }); }
            }
            return list;
        });
        return new { securities, count = securities.Count };
    }
}
