// DocForge — the AI document workshop.
// Four modules: Fill (template + messy text → filled docx/xlsx), Sheet (prompt → formula-driven xlsx),
// Doc (brief → structured docx), Deck (notes → themed pptx). Zero NuGet dependencies.
using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocForge;

var builder = WebApplication.CreateBuilder(args);
var port = Environment.GetEnvironmentVariable("PORT") ?? "8080";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");
var app = builder.Build();
app.UseDefaultFiles();
app.UseStaticFiles();

// ============================== config ==============================
bool mockAi = Environment.GetEnvironmentVariable("MOCK_AI") == "1";
string? geminiKey = Environment.GetEnvironmentVariable("GEMINI_API_KEY");
string? supaUrl = Environment.GetEnvironmentVariable("SUPABASE_URL")?.TrimEnd('/');
if (supaUrl is not null && supaUrl.EndsWith("/rest/v1")) supaUrl = supaUrl[..^"/rest/v1".Length];
string? supaKey = Environment.GetEnvironmentVariable("SUPABASE_SERVICE_KEY");
string dataDir = Environment.GetEnvironmentVariable("DATA_DIR") ?? "/tmp/docforge";
Directory.CreateDirectory(dataDir);

IStore store = (supaUrl, supaKey) is (not null, not null) && !mockAi
    ? new SupabaseStore(supaUrl!, supaKey!)
    : new FileStore(dataDir);
var http = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };

// ============================== rate limiting ==============================
var buckets = new ConcurrentDictionary<string, List<DateTime>>();
bool RateLimited(HttpContext ctx, int perMinute = 10)
{
    var ip = ctx.Connection.RemoteIpAddress?.ToString() ?? "?";
    var now = DateTime.UtcNow;
    var list = buckets.GetOrAdd(ip, _ => new());
    lock (list)
    {
        list.RemoveAll(t => (now - t).TotalSeconds > 60);
        if (list.Count >= perMinute) return true;
        list.Add(now);
        return false;
    }
}

// ============================== Gemini ==============================
async Task<JsonNode> AskJson(string system, string user, string mockKey)
{
    if (mockAi) return JsonNode.Parse(Mock.Responses[mockKey])!;
    if (geminiKey is null) throw new AppError(503, "AI is not configured — set GEMINI_API_KEY.");
    for (int attempt = 0; attempt < 2; attempt++)
    {
        var req = new JsonObject
        {
            ["contents"] = new JsonArray(new JsonObject
            {
                ["parts"] = new JsonArray(new JsonObject { ["text"] = user + (attempt == 1 ? "\n\nREMINDER: reply with ONLY the JSON object. No prose, no markdown fences." : "") })
            }),
            ["systemInstruction"] = new JsonObject { ["parts"] = new JsonArray(new JsonObject { ["text"] = system }) },
            ["generationConfig"] = new JsonObject { ["temperature"] = 0.25, ["maxOutputTokens"] = 8192, ["responseMimeType"] = "application/json" }
        };
        var resp = await http.PostAsync(
            $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={geminiKey}",
            new StringContent(req.ToJsonString(), Encoding.UTF8, "application/json"));
        var body = await resp.Content.ReadAsStringAsync();
        if ((int)resp.StatusCode == 429) throw new AppError(429, "AI quota briefly exceeded — try again in a minute.");
        if (!resp.IsSuccessStatusCode) throw new AppError(502, $"AI error {(int)resp.StatusCode}");
        try
        {
            var text = JsonNode.Parse(body)?["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.GetValue<string>() ?? "";
            text = text.Trim();
            if (text.StartsWith("```")) text = text[(text.IndexOf('\n') + 1)..].TrimEnd('`', '\n', ' ');
            return JsonNode.Parse(text)!;
        }
        catch { if (attempt == 1) throw new AppError(502, "AI returned malformed JSON twice — try rephrasing."); }
    }
    throw new AppError(502, "AI unavailable");
}

// ============================== helpers ==============================
static IResult Err(int code, string msg) => Results.Json(new { error = msg }, statusCode: code);
static string NewId() => Guid.NewGuid().ToString("N")[..12];
static string Sanitize(string s, int max) => (s ?? "").Length > max ? s![..max] : s ?? "";

async Task Record(string module, string title, JsonObject meta)
{
    try
    {
        var h = new JsonObject { ["id"] = NewId(), ["module"] = module, ["title"] = title,
            ["createdAt"] = DateTime.UtcNow.ToString("o"), ["meta"] = meta };
        await store.Upsert("df_history", h["id"]!.GetValue<string>(), h);
    }
    catch { /* history is best-effort */ }
}

// ============================== endpoints ==============================
app.MapGet("/api/health", () => Results.Json(new
{
    ok = true, app = "docforge",
    provider = mockAi ? "mock" : (geminiKey is null ? "none" : "gemini"),
    store = store is SupabaseStore ? "supabase" : "file",
    modules = new[] { "fill", "sheet", "doc", "deck" }
}));

// -------- template analyze + save --------
app.MapPost("/api/template/analyze", async (HttpContext ctx) =>
{
    if (RateLimited(ctx)) return Err(429, "Too many requests — slow down a little.");
    var body = await JsonNode.ParseAsync(ctx.Request.Body);
    var name = Sanitize(body?["filename"]?.GetValue<string>() ?? "template", 120);
    var b64 = body?["fileBase64"]?.GetValue<string>();
    if (b64 is null) return Err(400, "fileBase64 is required.");
    byte[] bytes;
    try { bytes = Convert.FromBase64String(b64); } catch { return Err(400, "fileBase64 is not valid base64."); }
    if (bytes.Length > 2_500_000) return Err(400, "Template too large — keep it under 2.5 MB.");
    var kind = name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? "xlsx"
             : name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ? "docx" : null;
    if (kind is null) return Err(400, "Only .docx and .xlsx templates are supported in v1.");
    List<string> placeholders;
    try { placeholders = Ooxml.FindPlaceholders(bytes, kind); }
    catch { return Err(400, "That file couldn't be opened as a valid Office document."); }
    if (placeholders.Count == 0)
        return Err(422, "No {{placeholders}} found. Add fields like {{client_name}} to the template and re-upload.");
    var id = NewId();
    var doc = new JsonObject
    {
        ["id"] = id, ["name"] = name, ["kind"] = kind,
        ["placeholders"] = new JsonArray(placeholders.Select(p => (JsonNode)p!).ToArray()),
        ["createdAt"] = DateTime.UtcNow.ToString("o"), ["fileBase64"] = b64
    };
    try { await store.Upsert("df_templates", id, doc); }
    catch (Exception ex) { return Err(502, $"storage error — check Supabase setup (df_templates table + env vars): {ex.Message}"); }
    return Results.Json(new { id, name, kind, placeholders });
});

app.MapGet("/api/templates", async () =>
{
    var all = await store.List("df_templates");
    var list = all.Select(t => new
    {
        id = t["id"]!.GetValue<string>(), name = t["name"]!.GetValue<string>(),
        kind = t["kind"]!.GetValue<string>(),
        placeholders = t["placeholders"]!.AsArray().Select(p => p!.GetValue<string>()).ToArray(),
        createdAt = t["createdAt"]!.GetValue<string>()
    }).OrderByDescending(t => t.createdAt).ToArray();
    return Results.Json(list);
});

app.MapDelete("/api/templates/{id}", async (string id) =>
{
    await store.Delete("df_templates", id);
    return Results.Json(new { deleted = id });
});

// -------- fill --------
app.MapPost("/api/template/fill", async (HttpContext ctx) =>
{
    if (RateLimited(ctx)) return Err(429, "Too many requests — slow down a little.");
    var body = await JsonNode.ParseAsync(ctx.Request.Body);
    var tid = body?["templateId"]?.GetValue<string>();
    var source = Sanitize(body?["sourceText"]?.GetValue<string>() ?? "", 20000);
    if (tid is null || source.Length < 10) return Err(400, "templateId and sourceText (a decent chunk of it) are required.");
    var tpl = await store.Get("df_templates", tid);
    if (tpl is null) return Err(404, "Template not found — upload it again.");
    var kind = tpl["kind"]!.GetValue<string>();
    var placeholders = tpl["placeholders"]!.AsArray().Select(p => p!.GetValue<string>()).ToList();

    JsonNode mapped;
    try
    {
        mapped = await AskJson(
            """
            You fill document templates. Given a list of placeholder field names and a source text
            (emails, notes, CRM data), produce values for each field FROM THE SOURCE ONLY.
            Reply with ONLY JSON: {"values": {"field": "value", ...}, "confidence": 0-100, "missing": ["field", ...]}
            Rules: never invent facts; a field not clearly present in the source goes in "missing" (and NOT in values);
            format dates like "1 September 2026"; format money with its currency symbol/code as given;
            expand obvious shorthand; confidence reflects overall mapping certainty.
            """,
            $"PLACEHOLDER FIELDS:\n{string.Join("\n", placeholders)}\n\nSOURCE TEXT:\n{source}",
            "fill");
    }
    catch (AppError e) { return Err(e.Code, e.Message); }

    var values = new Dictionary<string, string>();
    foreach (var kv in mapped["values"]?.AsObject() ?? new JsonObject())
        if (placeholders.Contains(kv.Key) && kv.Value is not null)
            values[kv.Key] = kv.Value.GetValue<string>();
    var missing = placeholders.Where(p => !values.ContainsKey(p)).ToArray();

    byte[] outBytes;
    try { outBytes = Ooxml.FillTemplate(Convert.FromBase64String(tpl["fileBase64"]!.GetValue<string>()), kind, values); }
    catch (Exception ex) { return Err(500, "Filling failed: " + ex.Message); }

    var outName = Path.GetFileNameWithoutExtension(tpl["name"]!.GetValue<string>()) + "-filled." + kind;
    // privacy: history stores field NAMES only, never the filled values or source text
    await Record("fill", outName, new JsonObject
    {
        ["template"] = tpl["name"]!.GetValue<string>(),
        ["filled"] = values.Count, ["missing"] = missing.Length,
        ["confidence"] = mapped["confidence"]?.GetValue<int>() ?? 0
    });
    return Results.Json(new
    {
        fileBase64 = Convert.ToBase64String(outBytes), filename = outName,
        filled = values, missing, confidence = mapped["confidence"]?.GetValue<int>() ?? 0
    });
});

// -------- generate: sheet --------
app.MapPost("/api/generate/sheet", async (HttpContext ctx) =>
{
    if (RateLimited(ctx)) return Err(429, "Too many requests — slow down a little.");
    var body = await JsonNode.ParseAsync(ctx.Request.Body);
    var prompt = Sanitize(body?["prompt"]?.GetValue<string>() ?? "", 4000);
    if (prompt.Length < 8) return Err(400, "Describe the spreadsheet you want.");
    JsonNode spec;
    try
    {
        spec = await AskJson(
            """
            You design Excel workbooks. Reply with ONLY JSON:
            {"filename":"snake_case.xlsx","summary":"one sentence",
             "sheets":[{"name":"...", "columns":[{"header":"...","width":14}],
                        "rows":[[cell,...]], "boldLastRow":true|false}]}
            Cell rules: numbers as JSON numbers; formulas as strings starting with "=" using A1 references
            (headers are row 1, data starts row 2); text as strings. USE REAL FORMULAS for totals, margins,
            growth, running balances — never precompute what a formula should do. Max 5 sheets, 300 rows/sheet.
            Populate realistic example data when the user gives none. Sheet names ≤ 28 chars, no []:*?/\.
            """,
            prompt, "sheet");
    }
    catch (AppError e) { return Err(e.Code, e.Message); }
    byte[] bytes;
    try { bytes = Ooxml.BuildXlsx(spec); }
    catch (Exception ex) { return Err(500, "Workbook build failed: " + ex.Message); }
    var fname = Sanitize(spec["filename"]?.GetValue<string>() ?? "workbook.xlsx", 80);
    if (!fname.EndsWith(".xlsx")) fname += ".xlsx";
    await Record("sheet", fname, new JsonObject { ["prompt"] = prompt, ["sheets"] = spec["sheets"]!.AsArray().Count });
    return Results.Json(new { fileBase64 = Convert.ToBase64String(bytes), filename = fname,
        summary = spec["summary"]?.GetValue<string>() ?? "", sheets = spec["sheets"]!.AsArray().Select(s => s!["name"]!.GetValue<string>()).ToArray() });
});

// -------- generate: doc --------
app.MapPost("/api/generate/doc", async (HttpContext ctx) =>
{
    if (RateLimited(ctx)) return Err(429, "Too many requests — slow down a little.");
    var body = await JsonNode.ParseAsync(ctx.Request.Body);
    var prompt = Sanitize(body?["prompt"]?.GetValue<string>() ?? "", 6000);
    if (prompt.Length < 8) return Err(400, "Describe the document you want.");
    JsonNode spec;
    try
    {
        spec = await AskJson(
            """
            You write business documents. Reply with ONLY JSON:
            {"filename":"snake_case.docx","title":"Document Title","summary":"one sentence",
             "blocks":[{"type":"h1|h2|p|bullets|table", "text":"(h1/h2/p)", "items":["(bullets)"],
                        "headers":["(table)"], "rows":[["(table)"]]}]}
            Write complete, professional prose — real paragraphs, not lorem ipsum. Use h1/h2 to structure,
            bullets for lists, tables for anything tabular (pricing, timelines, RACI). 300–900 words unless asked otherwise.
            """,
            prompt, "doc");
    }
    catch (AppError e) { return Err(e.Code, e.Message); }
    byte[] bytes;
    try { bytes = Ooxml.BuildDocx(spec["title"]?.GetValue<string>() ?? "Document", spec["blocks"]!.AsArray()); }
    catch (Exception ex) { return Err(500, "Document build failed: " + ex.Message); }
    var fname = Sanitize(spec["filename"]?.GetValue<string>() ?? "document.docx", 80);
    if (!fname.EndsWith(".docx")) fname += ".docx";
    await Record("doc", spec["title"]?.GetValue<string>() ?? fname, new JsonObject { ["prompt"] = prompt });
    return Results.Json(new { fileBase64 = Convert.ToBase64String(bytes), filename = fname,
        summary = spec["summary"]?.GetValue<string>() ?? "", title = spec["title"]?.GetValue<string>() ?? "" });
});

// -------- generate: deck --------
app.MapPost("/api/generate/deck", async (HttpContext ctx) =>
{
    if (RateLimited(ctx)) return Err(429, "Too many requests — slow down a little.");
    var body = await JsonNode.ParseAsync(ctx.Request.Body);
    var prompt = Sanitize(body?["prompt"]?.GetValue<string>() ?? "", 6000);
    var accent = Sanitize(body?["accent"]?.GetValue<string>() ?? "1F6E54", 7);
    if (prompt.Length < 8) return Err(400, "Describe the deck you want.");
    JsonNode spec;
    try
    {
        spec = await AskJson(
            """
            You design presentation outlines. Reply with ONLY JSON:
            {"filename":"snake_case.pptx","title":"Deck Title","subtitle":"one line",
             "summary":"one sentence",
             "slides":[{"title":"...","bullets":["...","..."]}]}
            First slide = title slide (its bullets are ignored). 5–12 slides total.
            Bullets: max 6 per slide, each ≤ 12 words, concrete and punchy — no sub-bullets.
            """,
            prompt, "deck");
    }
    catch (AppError e) { return Err(e.Code, e.Message); }
    byte[] bytes;
    try
    {
        bytes = Ooxml.BuildPptx(spec["title"]?.GetValue<string>() ?? "Deck",
            spec["subtitle"]?.GetValue<string>() ?? "", accent, spec["slides"]!.AsArray());
    }
    catch (Exception ex) { return Err(500, "Deck build failed: " + ex.Message); }
    var fname = Sanitize(spec["filename"]?.GetValue<string>() ?? "deck.pptx", 80);
    if (!fname.EndsWith(".pptx")) fname += ".pptx";
    await Record("deck", spec["title"]?.GetValue<string>() ?? fname,
        new JsonObject { ["prompt"] = prompt, ["slides"] = spec["slides"]!.AsArray().Count, ["accent"] = accent });
    return Results.Json(new { fileBase64 = Convert.ToBase64String(bytes), filename = fname,
        summary = spec["summary"]?.GetValue<string>() ?? "",
        slides = spec["slides"]!.AsArray().Select(s => s!["title"]!.GetValue<string>()).ToArray() });
});

// -------- history --------
app.MapGet("/api/history", async () =>
{
    var all = await store.List("df_history");
    return Results.Json(all.OrderByDescending(h => h["createdAt"]!.GetValue<string>()).Take(60));
});

app.Run();

// ============================== plumbing ==============================
class AppError(int code, string message) : Exception(message) { public int Code { get; } = code; }

interface IStore
{
    Task Upsert(string table, string id, JsonObject doc);
    Task<JsonObject?> Get(string table, string id);
    Task<List<JsonObject>> List(string table);
    Task Delete(string table, string id);
}

class FileStore(string dir) : IStore
{
    string PathFor(string table, string id) => System.IO.Path.Combine(dir, $"{table}__{id}.json");
    public Task Upsert(string table, string id, JsonObject doc)
    { File.WriteAllText(PathFor(table, id), doc.ToJsonString()); return Task.CompletedTask; }
    public Task<JsonObject?> Get(string table, string id)
    {
        var p = PathFor(table, id);
        return Task.FromResult(File.Exists(p) ? JsonNode.Parse(File.ReadAllText(p))!.AsObject() : null);
    }
    public Task<List<JsonObject>> List(string table) =>
        Task.FromResult(Directory.GetFiles(dir, $"{table}__*.json")
            .Select(f => JsonNode.Parse(File.ReadAllText(f))!.AsObject()).ToList());
    public Task Delete(string table, string id)
    { var p = PathFor(table, id); if (File.Exists(p)) File.Delete(p); return Task.CompletedTask; }
}

class SupabaseStore(string url, string key) : IStore
{
    readonly HttpClient c = new();
    HttpRequestMessage Req(HttpMethod m, string path)
    {
        var r = new HttpRequestMessage(m, $"{url}/rest/v1/{path}");
        r.Headers.Add("apikey", key);
        r.Headers.Add("Authorization", $"Bearer {key}");
        return r;
    }
    async Task<string> Send(HttpRequestMessage r)
    {
        var resp = await c.SendAsync(r);
        var body = await resp.Content.ReadAsStringAsync();
        if (!resp.IsSuccessStatusCode) throw new Exception($"store error {(int)resp.StatusCode}: {body}");
        return body;
    }
    public async Task Upsert(string table, string id, JsonObject doc)
    {
        var r = Req(HttpMethod.Post, table);
        r.Headers.Add("Prefer", "resolution=merge-duplicates");
        r.Content = new StringContent(new JsonArray(new JsonObject { ["id"] = id, ["doc"] = doc.DeepClone() }).ToJsonString(),
            Encoding.UTF8, "application/json");
        await Send(r);
    }
    public async Task<JsonObject?> Get(string table, string id)
    {
        var body = await Send(Req(HttpMethod.Get, $"{table}?id=eq.{id}&select=doc"));
        var arr = JsonNode.Parse(body)!.AsArray();
        return arr.Count == 0 ? null : arr[0]!["doc"]!.AsObject();
    }
    public async Task<List<JsonObject>> List(string table)
    {
        var body = await Send(Req(HttpMethod.Get, $"{table}?select=doc"));
        return JsonNode.Parse(body)!.AsArray().Select(x => x!["doc"]!.AsObject()).ToList();
    }
    public async Task Delete(string table, string id) => await Send(Req(HttpMethod.Delete, $"{table}?id=eq.{id}"));
}

static class Mock
{
    public static readonly Dictionary<string, string> Responses = new()
    {
        ["fill"] = """
        {"values":{"candidate_name":"Nadeesha Perera","job_title":"Senior Software Engineer","company":"Meridian Labs","salary":"LKR 480,000","start_date":"1 September 2026","manager":"R. Fernando"},"confidence":92,"missing":["probation_months"]}
        """,
        ["sheet"] = """
        {"filename":"bakery_cash_flow.xlsx","summary":"12-month bakery cash flow with assumptions.",
         "sheets":[{"name":"Cash Flow","columns":[{"header":"Month","width":12},{"header":"Revenue","width":13},{"header":"Costs","width":13},{"header":"Net","width":13}],
         "rows":[["Jan",120000,80000,"=B2-C2"],["Feb",132000,82400,"=B3-C3"],["Total","=SUM(B2:B3)","=SUM(C2:C3)","=SUM(D2:D3)"]],"boldLastRow":true},
         {"name":"Assumptions","columns":[{"header":"Item","width":24},{"header":"Value","width":12}],"rows":[["Monthly growth",0.1],["Cost inflation",0.03]]}]}
        """,
        ["doc"] = """
        {"filename":"meridian_proposal.docx","title":"Meridian Project Proposal","summary":"A concise project proposal.",
         "blocks":[{"type":"p","text":"This proposal outlines the recommended approach for the Meridian initiative."},
         {"type":"h1","text":"1. Background"},{"type":"p","text":"The client operates a legacy platform with rising costs."},
         {"type":"h2","text":"1.1 Objectives"},{"type":"bullets","items":["Reduce operating cost by 30%","Cut release cycle to two weeks","Improve customer satisfaction"]},
         {"type":"h1","text":"2. Commercials"},{"type":"table","headers":["Item","Qty","Price"],"rows":[["Discovery","1","$8,000"],["Build (per sprint)","6","$12,000"]]}]}
        """,
        ["deck"] = """
        {"filename":"meridian_kickoff.pptx","title":"Meridian Kick-off","subtitle":"Project briefing — August 2026","summary":"Kick-off deck.",
         "slides":[{"title":"Meridian Kick-off","bullets":[]},
         {"title":"Why now","bullets":["Market shift toward platforms","Legacy cost is rising","Team capacity is available"]},
         {"title":"The plan","bullets":["Phase 1 — discovery","Phase 2 — build","Phase 3 — rollout"]},
         {"title":"Next steps","bullets":["Confirm scope by Friday","Kick-off workshop next week"]}]}
        """
    };
}
