using Bloomberglp.Blpapi;
using System.Text.Json;
using Plugboard.Contracts;
using Plugboard.Blpapi;

namespace Plugboard.Plugins.Cmp;

// CMP (BBG's CMBS / structured-finance analytics add-in) as its OWN connector.
//
// It shares the process-wide BLPAPI session (Plugboard.Blpapi.Blp) with the bbg
// connector - BLPAPI's native library is a process singleton, so there can only be one
// managed session. "Its own connector" means its own plugin, its own /con/cmp/* routes,
// and its own service (//blp/cmp) opened on that shared session - exactly how CMP is a
// separate service in BBG itself.
//
// Surface: one TYPED endpoint per CMP request type at /con/cmp/<requestType> (the
// //blp/cmp `request` CHOICE has ~61 members), plus:
//   - /con/cmp/status   session/entitlement probe
//   - /con/cmp/raw      the proven JSON-envelope escape hatch (cmpJsonRequest)
// The two envelope members (cmpExcelRequest, cmpJsonRequest) route through the envelope;
// every other member is built directly from the request JSON. Write-type requests are
// gated (env PLUGBOARD_CMP_ALLOW_WRITES + a confirm token), fail-closed.
public sealed class CmpPlugin : IPlugin
{
    public string Name => "cmp";

    private const string SERVICE = "//blp/cmp";
    private const int    CMP_TIMEOUT_MS = 60000;   // analytics can be slow

    public void Register(IEndpointRegistry r)
    {
        r.Map("GET", "cmp/status", _ => Task.FromResult<object?>(GetStatus()), new RouteInfo(
            "CMP session status",
            "Whether the BBG Terminal is reachable and //blp/cmp opens (i.e. the session is CMP-entitled)."));

        r.Map("POST", "cmp/raw", req => Json(req, RunRaw), new RouteInfo(
            "CMP raw JSON envelope",
            "Escape hatch: send a requestData JSON blob through the cmpJsonRequest envelope and get the parsed cmpJsonResponse back. Prefer a typed /con/cmp/<requestType> endpoint when one fits.",
            new { requestData = new { cmpExcelRequest = new { parameters = new object[0] } } },
            new[] { new ParamInfo("requestData", "object", true, "The CMP request document (e.g. { cmpExcelRequest: {...} }); a JSON string is also accepted.") }));

        foreach (var rt in CMP_REQUEST_TYPES)
        {
            var reqType = rt;
            bool write = IsWriteOp(reqType);
            var summary = "CMP " + reqType + (write ? " (write)" : "");
            var desc = write
                ? $"CMP {reqType} - a WRITE request. Disabled unless PLUGBOARD_CMP_ALLOW_WRITES=true and body.confirm == \"{SERVICE} {reqType}\"."
                : $"CMP {reqType}. POST the request fields as JSON (top level, or under \"request\": {{...}}).";
            r.Map("POST", $"cmp/{reqType}", req => Json(req, b => RunTyped(reqType, b)),
                new RouteInfo(summary, desc,
                    new { request = new { } },
                    new[] { new ParamInfo("request", "object", false, "The typed request body for this CMP request type. May also be posted at the top level.") }));
        }
    }

    private static Task<object?> Json(PluginRequest req, Func<JsonElement, object?> handler)
    {
        using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(req.Body) ? "{}" : req.Body);
        return Task.FromResult(handler(doc.RootElement));
    }

    // Write-request gate: opt in via env, then require an explicit confirm token per call.
    private static readonly bool _allowWrites =
        (Environment.GetEnvironmentVariable("PLUGBOARD_CMP_ALLOW_WRITES") ?? "")
            .Trim().ToLowerInvariant() is "1" or "true" or "yes";

    private static object GetStatus()
    {
        var p = Blp.ProbeService(SERVICE);
        return new
        {
            connected = p.Connected,
            serviceReady = p.ServiceReady,           // //blp/cmp opened -> session is CMP-entitled
            lastKnownReady = p.LastKnownReady,
            service = SERVICE, host = Blp.Host, port = Blp.Port,
            writesEnabled = _allowWrites,
            error = p.Error
        };
    }

    private static object? RunRaw(JsonElement body)
    {
        string reqJson = body.TryGetProperty("requestData", out var rd)
            ? (rd.ValueKind == JsonValueKind.String ? rd.GetString()! : rd.GetRawText())
            : body.GetRawText();
        return SendEnvelope(reqJson);
    }

    private static object? RunTyped(string reqType, JsonElement body)
    {
        JsonElement inner = body.TryGetProperty("request", out var r) && r.ValueKind == JsonValueKind.Object ? r : body;

        if (reqType == "cmpExcelRequest") return SendEnvelope("{\"cmpExcelRequest\":" + inner.GetRawText() + "}");
        if (reqType == "cmpJsonRequest")  return SendEnvelope(inner.GetRawText());

        if (IsWriteOp(reqType))
        {
            if (!_allowWrites)
                throw new Exception($"'{reqType}' is a write request; disabled. Set PLUGBOARD_CMP_ALLOW_WRITES=true to enable.");
            var expected = SERVICE + " " + reqType;
            var confirm = body.TryGetProperty("confirm", out var c) ? c.GetString() : null;
            if (confirm != expected)
                throw new Exception($"write request requires body.confirm == \"{expected}\"");
        }

        return Blp.WithService(SERVICE, (session, svc) =>
        {
            var req = svc.CreateRequest("request");
            var choice = req.AsElement.GetElement(reqType);           // select the CHOICE member
            if (inner.ValueKind == JsonValueKind.Object) Blp.ApplyJsonToElement(choice, inner);
            var outMsgs = new List<object?>();
            foreach (var msg in Blp.SendAndReceive(session, req, CMP_TIMEOUT_MS))
            {
                if (msg.HasElement("responseError"))
                    throw new Exception("responseError: " + msg.GetElement("responseError").ToString());
                outMsgs.Add(new { messageType = msg.MessageType.ToString(), data = Blp.ElementToJson(msg.AsElement) });
            }
            return (object?)new { requestType = reqType, service = SERVICE, messages = outMsgs };
        });
    }

    // Send a requestData blob through the cmpJsonRequest envelope; return parsed cmpJsonResponse.
    private static object? SendEnvelope(string reqJson)
    {
        return Blp.WithService(SERVICE, (session, svc) =>
        {
            var req = svc.CreateRequest("request");
            req.AsElement.GetElement("cmpJsonRequest").SetElement("requestData", reqJson);
            string? p = null; var seen = new List<string>();
            foreach (var msg in Blp.SendAndReceive(session, req, CMP_TIMEOUT_MS))
            {
                seen.Add(msg.MessageType.ToString());
                if (msg.HasElement("responseError"))
                    throw new Exception("responseError: " + msg.GetElement("responseError").ToString());
                if (msg.HasElement("cmpJsonResponse"))
                {
                    var e = msg.GetElement("cmpJsonResponse");
                    p = e.HasElement("responseData") ? e.GetElement("responseData").GetValueAsString()
                                                     : e.GetValueAsString();
                }
            }
            if (string.IsNullOrEmpty(p))
                throw new Exception("no cmpJsonResponse payload returned; messageTypes: " + string.Join(",", seen));
            try { return (object?)JsonSerializer.Deserialize<JsonElement>(p); }
            catch { return (object?)new { raw = p }; }
        });
    }

    // ── write-op gate ──

    private static readonly string[] _writeVerbs =
    {
        "create", "modify", "delete", "route", "cancel", "upload", "submit",
        "contribut", "insert", "update", "amend", "replace", "addtag", "settings",
        "remove", "post", "write", "activate"
    };

    private static bool IsWriteOp(string op)
    {
        var o = op.ToLowerInvariant();
        if (o.StartsWith("get") || o.StartsWith("list") || o.StartsWith("send")) return false;
        return _writeVerbs.Any(v => o.Contains(v));
    }

    // The request types the //blp/cmp `request` CHOICE exposes; each gets a typed endpoint.
    private static readonly string[] CMP_REQUEST_TYPES =
    {
        "bondHistoricDataRequest","getAllCurves","mtgeAnalyticsRequest","dealRequest","cashFlowsAggregate",
        "getForwardIndexRates","dealAnalyticsRequest","modelMetadataRequest","activateDealRequest","getActiveDealsRequest",
        "loanStatsRequest","assetsRequest","requiredIndexRatesRequest","getAllCurvesDelayed","getForwardIndexRatesDelayed",
        "getSummaryDetails","getExtLoanDetails","getPropertyDetails","getLeaseDetails","list_canned_flows",
        "getHistLoanDetails","getPaymentHistory","getAbLoanDetails","getLoanDetails","getHopeNotesRelatives",
        "getYearlyFinancials","triggerHistoricalRequestJson","getSinglePropertyDetails","getPropertyDetailedFinancials","mtgeTotalReturnRequest",
        "groupDistributionRequestJson","securityIdRequest","dealRequestUri","dealRequestUriMsgpack","getBval3pmSnapshotId",
        "analyticsRequest","riskRequest","bootstrapRequest","getMarketSourceRequest","listCannedFlowsRequest",
        "calcAnalyticsRequest","modelProjRequest","settleDateRequest","loanStatsRequestJson","getExtPropertyDetails",
        "ShockXmkt","histCreditSupportRequest","listFlowsRequest","cloMarketValueRequest","getCfsvcScenariosRequest",
        "writeCfsvcScenariosRequest","getTicketInfoRequest","cmpExcelRequest","ratesRequest","getModelErrorShifts",
        "getSpotRateShifts","getModelErrors","getMarketId","cmpJsonRequest","shockXmkt","cloTransactionHistoryRequest"
    };
}
