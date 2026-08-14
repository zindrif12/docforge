// DocForge v2 template engine — loops, conditionals, filters over docx/xlsx.
// Syntax: {{field}}, {{field | filter:arg | filter2}}, {{#each coll}}...{{/each}}, {{#if flag}}...{{/if}}
using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DocForge;

public class AppError(int code, string message) : Exception(message) { public int Code { get; } = code; }

public static class Tpl
{
    static readonly Regex Tag = new(@"\{\{\s*(#each|#if|/each|/if)?\s*([A-Za-z0-9_.\- ]*?)\s*((?:\|[^}]*?)?)\s*\}\}", RegexOptions.Compiled);
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ===================== schema =====================
    // { scalars:[], collections:[{name, fields:[]}], conditionals:[] }
    public static JsonObject ExtractSchema(byte[] fileBytes, string kind)
    {
        var texts = new List<string>();
        using (var ms = new MemoryStream(fileBytes))
        using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read))
            foreach (var e in z.Entries)
                if (IsTarget(e.FullName, kind))
                { using var r = new StreamReader(e.Open()); texts.Add(kind == "docx" ? JoinWt(r.ReadToEnd()) : r.ReadToEnd()); }

        var scalars = new List<string>(); var conds = new List<string>();
        var colls = new Dictionary<string, List<string>>();
        foreach (var text in texts)
        {
            string? inLoop = null; int depth = 0;
            foreach (Match m in Tag.Matches(text))
            {
                var op = m.Groups[1].Value; var name = m.Groups[2].Value.Trim();
                if (op == "#each") { if (inLoop is not null) throw new AppError(422, "Nested {{#each}} loops are not supported in v2."); inLoop = name; colls.TryAdd(name, new()); }
                else if (op == "/each") inLoop = null;
                else if (op == "#if") { depth++; if (!conds.Contains(name)) conds.Add(name); }
                else if (op == "/if") depth--;
                else if (name.Length > 0)
                {
                    if (inLoop is not null) { if (name != "this" && !colls[inLoop].Contains(name)) colls[inLoop].Add(name); }
                    else if (!scalars.Contains(name) && !conds.Contains(name)) scalars.Add(name);
                }
            }
            if (inLoop is not null) throw new AppError(422, $"{{{{#each {inLoop}}}}} has no matching {{{{/each}}}}.");
            if (depth != 0) throw new AppError(422, "Unbalanced {{#if}} / {{/if}} markers.");
        }
        return new JsonObject
        {
            ["scalars"] = new JsonArray(scalars.Select(s => (JsonNode)s!).ToArray()),
            ["collections"] = new JsonArray(colls.Select(kv => (JsonNode)new JsonObject
            { ["name"] = kv.Key, ["fields"] = new JsonArray(kv.Value.Select(f => (JsonNode)f!).ToArray()) }).ToArray()),
            ["conditionals"] = new JsonArray(conds.Select(c => (JsonNode)c!).ToArray())
        };
    }

    static bool IsTarget(string p, string kind) =>
        kind == "docx" ? p is "word/document.xml" or "word/header1.xml" or "word/header2.xml" or "word/footer1.xml" or "word/footer2.xml"
                       : (p == "xl/sharedStrings.xml" || p.StartsWith("xl/worksheets/"));
    static string JoinWt(string xml)
    {
        var sb = new StringBuilder();
        foreach (Match m in Regex.Matches(xml, @"<w:t[^>]*>([^<]*)</w:t>"))
            sb.Append(System.Net.WebUtility.HtmlDecode(m.Groups[1].Value));
        return sb.ToString();
    }

    // ===================== filters =====================
    public static string ApplyFilters(string value, string filterChain)
    {
        if (string.IsNullOrWhiteSpace(filterChain)) return value;
        foreach (var raw in filterChain.Split('|', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = raw.Trim().Split(':', 2);
            var f = parts[0].Trim().ToLowerInvariant();
            var arg = parts.Length > 1 ? parts[1].Trim() : "";
            value = f switch
            {
                "upper" => value.ToUpperInvariant(),
                "lower" => value.ToLowerInvariant(),
                "currency" => double.TryParse(value.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var c)
                              ? (arg.Length > 0 ? arg + " " : "") + c.ToString("#,##0.00", CultureInfo.InvariantCulture) : value,
                "number" => double.TryParse(value.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var n)
                              ? n.ToString("#,##0.##", CultureInfo.InvariantCulture) : value,
                "round" => double.TryParse(value.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out var r)
                              ? Math.Round(r, int.TryParse(arg, out var dp) ? dp : 0).ToString(CultureInfo.InvariantCulture) : value,
                "date" => DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
                              ? d.ToString(arg == "short" ? "dd/MM/yyyy" : "d MMMM yyyy", CultureInfo.InvariantCulture) : value,
                _ => value
            };
        }
        return value;
    }

    // ===================== payload access =====================
    static string Resolve(JsonObject payload, JsonObject? scope, string name)
    {
        if (name == "this" && scope?["this"] is JsonNode t) return t.GetValue<string>();
        if (scope is not null && scope.ContainsKey(name)) return ValStr(scope[name]);
        var sc = payload["scalars"]?.AsObject();
        if (sc is not null && sc.ContainsKey(name)) return ValStr(sc[name]);
        return "";
    }
    static string ValStr(JsonNode? n) => n is null ? "" :
        n.GetValueKind() switch
        {
            System.Text.Json.JsonValueKind.Number => n.GetValue<double>().ToString(CultureInfo.InvariantCulture),
            System.Text.Json.JsonValueKind.True => "true",
            System.Text.Json.JsonValueKind.False => "false",
            _ => n.GetValue<string>()
        };
    static bool Truthy(JsonObject payload, string name)
    {
        var c = payload["conditionals"]?.AsObject()?[name] ?? payload["scalars"]?.AsObject()?[name];
        if (c is null) return false;
        return c.GetValueKind() switch
        {
            System.Text.Json.JsonValueKind.True => true,
            System.Text.Json.JsonValueKind.False => false,
            System.Text.Json.JsonValueKind.Number => c.GetValue<double>() != 0,
            _ => !string.IsNullOrWhiteSpace(c.GetValue<string>()) && c.GetValue<string>().ToLowerInvariant() is not ("false" or "no" or "0")
        };
    }
    static JsonArray Items(JsonObject payload, string coll) =>
        payload["collections"]?.AsObject()?[coll]?.AsArray() ?? new JsonArray();
    static JsonObject ItemScope(JsonNode item) =>
        item is JsonObject o ? o : new JsonObject { ["this"] = ValStr(item) };

    // substitute scalar/filtered tags in a plain string within an optional loop scope
    static string Substitute(string text, JsonObject payload, JsonObject? scope, bool xmlEscape)
    {
        return Tag.Replace(text, m =>
        {
            if (m.Groups[1].Value.Length > 0) return m.Value; // block markers handled elsewhere
            var name = m.Groups[2].Value.Trim();
            if (name.Length == 0) return m.Value;
            var v = ApplyFilters(Resolve(payload, scope, name), m.Groups[3].Value.TrimStart('|'));
            return xmlEscape ? Ooxml.XmlEsc(v) : v;
        });
    }

    // ===================== DOCX render =====================
    public static byte[] RenderDocx(byte[] fileBytes, JsonObject payload)
    {
        using var src = new MemoryStream(fileBytes);
        using var zin = new System.IO.Compression.ZipArchive(src, System.IO.Compression.ZipArchiveMode.Read);
        using var outMs = new MemoryStream();
        using (var zout = new System.IO.Compression.ZipArchive(outMs, System.IO.Compression.ZipArchiveMode.Create, true))
            foreach (var e in zin.Entries)
            {
                var ne = zout.CreateEntry(e.FullName, System.IO.Compression.CompressionLevel.Optimal);
                using var inS = e.Open(); using var outS = ne.Open();
                if (IsTarget(e.FullName, "docx"))
                {
                    using var r = new StreamReader(inS);
                    var xml = RenderDocxXml(r.ReadToEnd(), payload);
                    var b = new UTF8Encoding(false).GetBytes(xml);
                    outS.Write(b, 0, b.Length);
                }
                else inS.CopyTo(outS);
            }
        return outMs.ToArray();
    }

    static string RenderDocxXml(string xml, JsonObject payload)
    {
        var doc = XDocument.Parse(xml);
        // 1) merge runs in any paragraph containing markers so tags are contiguous
        foreach (var p in doc.Descendants(W + "p"))
        {
            var ts = p.Descendants(W + "t").ToList();
            if (ts.Count < 2) continue;
            var joined = string.Concat(ts.Select(t => t.Value));
            if (!joined.Contains("{{")) continue;
            ts[0].Value = joined;
            ts[0].SetAttributeValue(XNamespace.Xml + "space", "preserve");
            for (int i = 1; i < ts.Count; i++) ts[i].Value = "";
        }
        // 2) expand blocks over TABLE ROWS, then over body PARAGRAPHS
        foreach (var tbl in doc.Descendants(W + "tbl").ToList())
            ExpandBlocks(tbl, W + "tr", payload);
        var body = doc.Root!.Element(W + "body");
        if (body is not null) ExpandBlocks(body, W + "p", payload);
        // 3) scalar substitution everywhere left
        foreach (var t in doc.Descendants(W + "t"))
            if (t.Value.Contains("{{")) t.Value = Substitute(t.Value, payload, null, xmlEscape: false);
        return doc.Declaration is null ? doc.ToString(SaveOptions.DisableFormatting)
             : doc.Declaration + doc.ToString(SaveOptions.DisableFormatting);
    }

    // generic block expansion: find unit (row/para) with a start marker, matching unit with end marker,
    // then repeat (each) or keep/drop (if) the unit block. Handles multiple sibling blocks.
    static void ExpandBlocks(XContainer container, XName unitName, JsonObject payload)
    {
        for (var guard = 0; guard < 64; guard++)
        {
            var units = container.Elements(unitName).ToList();
            int s = -1; string op = "", name = "";
            for (int i = 0; i < units.Count; i++)
            {
                var m = Tag.Match(UnitText(units[i]));
                if (m.Success && m.Groups[1].Value is "#each" or "#if")
                { s = i; op = m.Groups[1].Value; name = m.Groups[2].Value.Trim(); break; }
            }
            if (s < 0) return;
            var endTag = op == "#each" ? "/each" : "/if";
            int e = -1;
            for (int i = s; i < units.Count; i++)
                if (Tag.Matches(UnitText(units[i])).Any(m => m.Groups[1].Value == endTag)) { e = i; break; }
            if (e < 0) throw new AppError(422, $"{{{{{op} {name}}}}} has no matching {{{{{endTag}}}}} in the same table/section.");

            var block = units.Skip(s).Take(e - s + 1).ToList();
            var anchor = block[0].PreviousNode;
            foreach (var u in block) StripMarkers(u, op, endTag);

            var rendered = new List<XElement>();
            if (op == "#if")
            {
                if (Truthy(payload, name))
                    foreach (var u in block) { if (!IsEmptyUnit(u)) rendered.Add(u); }
            }
            else
            {
                foreach (var item in Items(payload, name))
                {
                    var scope = ItemScope(item!);
                    foreach (var u in block)
                    {
                        var clone = new XElement(u);
                        foreach (var t in clone.Descendants(W + "t"))
                            if (t.Value.Contains("{{")) t.Value = Substitute(t.Value, payload, scope, xmlEscape: false);
                        if (!IsEmptyUnit(clone)) rendered.Add(clone);
                    }
                }
            }
            foreach (var u in block) u.Remove();
            if (anchor is null) foreach (var rEl in Enumerable.Reverse(rendered)) container.AddFirst(rEl);
            else foreach (var rEl in rendered) { anchor.AddAfterSelf(rEl); anchor = rEl; }
        }
        throw new AppError(422, "Too many template blocks (64+) — simplify the template.");
    }
    static string UnitText(XElement u) => string.Concat(u.Descendants(W + "t").Select(t => t.Value));
    static void StripMarkers(XElement u, string op, string endTag)
    {
        foreach (var t in u.Descendants(W + "t"))
            t.Value = Regex.Replace(t.Value, @"\{\{\s*(#each|#if|/each|/if)[^}]*\}\}", "");
    }
    // a marker-only paragraph/row left empty after stripping should vanish rather than leave a blank line
    static bool IsEmptyUnit(XElement u)
    {
        if (u.Name != W + "p") return false;
        return string.Concat(u.Descendants(W + "t").Select(t => t.Value)).Trim().Length == 0
            && !u.Descendants(W + "drawing").Any() && !u.Descendants(W + "tbl").Any();
    }

    // ===================== XLSX render =====================
    public static byte[] RenderXlsx(byte[] fileBytes, JsonObject payload)
    {
        using var src = new MemoryStream(fileBytes);
        using var zin = new System.IO.Compression.ZipArchive(src, System.IO.Compression.ZipArchiveMode.Read);
        using var outMs = new MemoryStream();
        using (var zout = new System.IO.Compression.ZipArchive(outMs, System.IO.Compression.ZipArchiveMode.Create, true))
            foreach (var e in zin.Entries)
            {
                var ne = zout.CreateEntry(e.FullName, System.IO.Compression.CompressionLevel.Optimal);
                using var inS = e.Open(); using var outS = ne.Open();
                if (IsTarget(e.FullName, "xlsx"))
                {
                    using var r = new StreamReader(inS);
                    var xml = r.ReadToEnd();
                    xml = e.FullName.StartsWith("xl/worksheets/") ? RenderSheetXml(xml, payload)
                        : Substitute(xml, payload, null, xmlEscape: true); // sharedStrings scalar fill
                    var b = new UTF8Encoding(false).GetBytes(xml);
                    outS.Write(b, 0, b.Length);
                }
                else inS.CopyTo(outS);
            }
        return outMs.ToArray();
    }

    static string RenderSheetXml(string xml, JsonObject payload)
    {
        var doc = XDocument.Parse(xml);
        XNamespace S = doc.Root!.Name.Namespace;
        var sheetData = doc.Descendants(S + "sheetData").FirstOrDefault();
        if (sheetData is null) return xml;
        var outRows = new List<XElement>();
        int cur = 1;
        var rows = sheetData.Elements(S + "row").ToList();
        for (int i = 0; i < rows.Count; i++)
        {
            var rowText = string.Concat(rows[i].Descendants().Where(x => x.Name.LocalName is "t" or "f").Select(x => x.Value));
            var m = Tag.Match(rowText);
            if (m.Success && m.Groups[1].Value == "#each")
            {
                var coll = m.Groups[2].Value.Trim();
                // template block: this row up to the row containing {{/each}} (usually the same row)
                int e = i;
                while (e < rows.Count && !Tag.Matches(string.Concat(rows[e].Descendants().Where(x => x.Name.LocalName is "t" or "f").Select(x => x.Value))).Any(mm => mm.Groups[1].Value == "/each")) e++;
                if (e >= rows.Count) throw new AppError(422, $"{{{{#each {coll}}}}} has no matching {{{{/each}}}} in the sheet.");
                var block = rows.Skip(i).Take(e - i + 1).ToList();
                foreach (var item in Items(payload, coll))
                {
                    var scope = ItemScope(item!);
                    foreach (var tmpl in block)
                    {
                        var clone = new XElement(tmpl);
                        foreach (var t in clone.Descendants())
                            if (t.Name.LocalName is "t" or "f" && t.Value.Contains("{{"))
                                t.Value = Regex.Replace(Substitute(t.Value, payload, scope, xmlEscape: false),
                                                        @"\{\{\s*(#each|/each)[^}]*\}\}", "");
                        outRows.Add(Renumber(clone, cur, S)); cur++;
                    }
                }
                i = e;
            }
            else
            {
                var clone = new XElement(rows[i]);
                foreach (var t in clone.Descendants())
                    if (t.Name.LocalName is "t" or "f" && t.Value.Contains("{{"))
                        t.Value = Substitute(t.Value, payload, null, xmlEscape: false);
                outRows.Add(Renumber(clone, cur, S)); cur++;
            }
        }
        sheetData.RemoveNodes();
        foreach (var r in outRows) sheetData.Add(r);
        // dimension can go stale after expansion; removing it is legal and Excel recomputes
        doc.Descendants(S + "dimension").FirstOrDefault()?.Remove();
        return doc.Declaration is null ? doc.ToString(SaveOptions.DisableFormatting)
             : doc.Declaration + doc.ToString(SaveOptions.DisableFormatting);
    }
    // rewrite row @r, each cell @r (A5→A{new}) and row digits inside formulas that referenced the template row
    static XElement Renumber(XElement row, int newIndex, XNamespace S)
    {
        var oldIdx = row.Attribute("r")?.Value;
        row.SetAttributeValue("r", newIndex);
        foreach (var c in row.Elements(S + "c"))
        {
            var rf = c.Attribute("r")?.Value;
            if (rf is not null)
                c.SetAttributeValue("r", Regex.Replace(rf, @"\d+$", newIndex.ToString()));
        }
        if (oldIdx is not null)
            foreach (var f in row.Descendants(S + "f"))
                f.Value = Regex.Replace(f.Value, @"(?<=[A-Z])" + oldIdx + @"\b", newIndex.ToString());
        return row;
    }

    // ===================== verification =====================
    // status per field: verified (value or its evidence found in source), derived (AI transformed), missing
    public static JsonArray Verify(JsonObject payload, JsonObject evidence, string source, JsonArray missing)
    {
        var report = new JsonArray();
        string Norm(string s) => Regex.Replace(s.ToLowerInvariant(), @"[\s,]+", "");
        var normSrc = Norm(source);
        void Check(string field, string value)
        {
            var ev = evidence[field]?.GetValue<string>() ?? "";
            string status =
                value.Length == 0 ? "missing"
                : normSrc.Contains(Norm(value)) ? "verified"
                : ev.Length > 0 && normSrc.Contains(Norm(ev)) ? "derived"
                : "unverified";
            report.Add(new JsonObject { ["field"] = field, ["value"] = value, ["evidence"] = ev, ["status"] = status });
        }
        foreach (var kv in payload["scalars"]?.AsObject() ?? new JsonObject()) Check(kv.Key, ValStr(kv.Value));
        foreach (var kv in payload["collections"]?.AsObject() ?? new JsonObject())
        {
            int idx = 0;
            foreach (var item in kv.Value!.AsArray())
            {
                if (item is JsonObject o)
                    foreach (var f in o) Check($"{kv.Key}[{idx}].{f.Key}", ValStr(f.Value));
                else Check($"{kv.Key}[{idx}]", ValStr(item));
                idx++;
            }
        }
        foreach (var miss in missing) report.Add(new JsonObject
        { ["field"] = miss!.GetValue<string>(), ["value"] = "", ["evidence"] = "", ["status"] = "missing" });
        return report;
    }
}
