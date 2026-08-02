// DocForge OpenXML writers — zero-dependency Office file generation via ZipArchive.
// Every XML part is minimal but spec-complete: files must open with NO repair dialogs.
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;

namespace DocForge;

public static class Ooxml
{
    public static string XmlEsc(string s) => System.Security.SecurityElement.Escape(s ?? "") ?? "";

    static void Add(ZipArchive z, string path, string content)
    {
        var e = z.CreateEntry(path, CompressionLevel.Optimal);
        using var w = new StreamWriter(e.Open(), new UTF8Encoding(false));
        w.Write(content);
    }

    // ============================================================ XLSX ============
    // spec: { sheets: [ { name, columns:[{header,width}], rows:[[cell,...]], } ] }
    // cell: number → numeric; string starting '=' → formula; else inline string.
    public static byte[] BuildXlsx(JsonNode spec)
    {
        var sheets = spec["sheets"]!.AsArray();
        using var ms = new MemoryStream();
        using (var z = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            var ct = new StringBuilder();
            ct.Append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>""");
            for (int i = 1; i <= sheets.Count; i++)
                ct.Append($"""<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>""");
            ct.Append("</Types>");
            Add(z, "[Content_Types].xml", ct.ToString());

            Add(z, "_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");

            var wbRels = new StringBuilder("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""");
            var wbSheets = new StringBuilder();
            for (int i = 1; i <= sheets.Count; i++)
            {
                // CRITICAL: sheetId allocated as index — unique by construction (the Max+1 lesson, avoided by design)
                var nm = XmlEsc(sheets[i - 1]!["name"]?.GetValue<string>() ?? $"Sheet{i}");
                wbSheets.Append($"""<sheet name="{nm}" sheetId="{i}" r:id="rIdS{i}"/>""");
                wbRels.Append($"""<Relationship Id="rIdS{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>""");
            }
            wbRels.Append($"""<Relationship Id="rIdSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>""");
            wbRels.Append("</Relationships>");
            Add(z, "xl/_rels/workbook.xml.rels", wbRels.ToString());

            Add(z, "xl/workbook.xml", $"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{wbSheets}</sheets></workbook>""");

            // styles: 0 default, 1 bold header w/ fill, 2 number 2dp, 3 bold total, 4 currency-ish #,##0.00
            Add(z, "xl/styles.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F6E54"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>""");

            for (int si = 0; si < sheets.Count; si++)
            {
                var sh = sheets[si]!;
                var cols = sh["columns"]?.AsArray();
                var rows = sh["rows"]?.AsArray() ?? new JsonArray();
                var sb = new StringBuilder("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""");
                if (cols is { Count: > 0 })
                {
                    sb.Append("<cols>");
                    for (int c = 0; c < cols.Count; c++)
                    {
                        var wd = cols[c]!["width"]?.GetValue<double>() ?? 16;
                        sb.Append($"""<col min="{c + 1}" max="{c + 1}" width="{wd}" customWidth="1"/>""");
                    }
                    sb.Append("</cols>");
                }
                sb.Append("<sheetData>");
                int r = 1;
                if (cols is { Count: > 0 })
                {
                    sb.Append($"""<row r="{r}">""");
                    for (int c = 0; c < cols.Count; c++)
                        sb.Append(InlineCell(Col(c) + r, cols[c]!["header"]?.GetValue<string>() ?? "", 1));
                    sb.Append("</row>");
                    r++;
                }
                foreach (var row in rows)
                {
                    sb.Append($"""<row r="{r}">""");
                    var cells = row!.AsArray();
                    for (int c = 0; c < cells.Count; c++)
                    {
                        var cellRef = Col(c) + r;
                        var v = cells[c];
                        bool lastRowBold = sh["boldLastRow"]?.GetValue<bool>() == true && ReferenceEquals(row, rows[rows.Count - 1]);
                        if (v is null) continue;
                        if (v.GetValueKind() is System.Text.Json.JsonValueKind.Number)
                        {
                            var num = v.GetValue<double>();
                            sb.Append($"""<c r="{cellRef}" s="{(lastRowBold ? 4 : 2)}"><v>{num}</v></c>""");
                        }
                        else
                        {
                            var s = v.GetValue<string>();
                            if (s.StartsWith("="))
                                sb.Append($"""<c r="{cellRef}" s="{(lastRowBold ? 4 : 2)}"><f>{XmlEsc(s[1..])}</f></c>""");
                            else
                                sb.Append(InlineCell(cellRef, s, lastRowBold ? 3 : 0));
                        }
                    }
                    sb.Append("</row>");
                    r++;
                }
                sb.Append("</sheetData></worksheet>");
                Add(z, $"xl/worksheets/sheet{si + 1}.xml", sb.ToString());
            }
        }
        return ms.ToArray();
    }
    static string InlineCell(string r, string text, int style) =>
        $"""<c r="{r}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{XmlEsc(text)}</t></is></c>""";
    public static string Col(int i) { var s = ""; i++; while (i > 0) { s = (char)('A' + (i - 1) % 26) + s; i = (i - 1) / 26; } return s; }

    // ============================================================ DOCX ============
    // blocks: [ {type:"h1"|"h2"|"p"|"bullets"|"table", text?, items?, headers?, rows?} ]
    public static byte[] BuildDocx(string title, JsonArray blocks)
    {
        using var ms = new MemoryStream();
        using (var z = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            Add(z, "[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>""");
            Add(z, "_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""");
            Add(z, "word/_rels/document.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>""");
            Add(z, "word/styles.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="52"/><w:color w:val="1F6E54"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="280" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F6E54"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="2E8B6F"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style></w:styles>""");
            Add(z, "word/numbering.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>""");

            var body = new StringBuilder();
            body.Append(Para("Title", title));
            foreach (var b in blocks)
            {
                var type = b!["type"]?.GetValue<string>() ?? "p";
                switch (type)
                {
                    case "h1": body.Append(Para("Heading1", b["text"]?.GetValue<string>() ?? "")); break;
                    case "h2": body.Append(Para("Heading2", b["text"]?.GetValue<string>() ?? "")); break;
                    case "bullets":
                        foreach (var it in b["items"]?.AsArray() ?? new JsonArray())
                            body.Append($"""<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>{Run(it!.GetValue<string>())}</w:p>""");
                        break;
                    case "table":
                        body.Append("""<w:tbl><w:tblPr><w:tblStyle w:val="Normal"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="CFCFCF"/><w:left w:val="single" w:sz="4" w:color="CFCFCF"/><w:bottom w:val="single" w:sz="4" w:color="CFCFCF"/><w:right w:val="single" w:sz="4" w:color="CFCFCF"/><w:insideH w:val="single" w:sz="4" w:color="CFCFCF"/><w:insideV w:val="single" w:sz="4" w:color="CFCFCF"/></w:tblBorders></w:tblPr>""");
                        var headers = b["headers"]?.AsArray();
                        int nCols = headers?.Count ?? b["rows"]?.AsArray().FirstOrDefault()?.AsArray().Count ?? 1;
                        body.Append("<w:tblGrid>");
                        for (int gc = 0; gc < nCols; gc++) body.Append($"""<w:gridCol w:w="{9638 / Math.Max(nCols, 1)}"/>""");
                        body.Append("</w:tblGrid>");
                        if (headers is not null)
                        {
                            body.Append("<w:tr>");
                            foreach (var h in headers) body.Append($"""<w:tc><w:tcPr><w:shd w:val="clear" w:fill="1F6E54"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">{XmlEsc(h!.GetValue<string>())}</w:t></w:r></w:p></w:tc>""");
                            body.Append("</w:tr>");
                        }
                        foreach (var row in b["rows"]?.AsArray() ?? new JsonArray())
                        {
                            body.Append("<w:tr>");
                            foreach (var cell in row!.AsArray())
                                body.Append($"<w:tc><w:p>{Run(cell?.ToString() ?? "")}</w:p></w:tc>");
                            body.Append("</w:tr>");
                        }
                        body.Append("</w:tbl><w:p/>");
                        break;
                    default: body.Append(Para("Normal", b["text"]?.GetValue<string>() ?? "")); break;
                }
            }
            body.Append("""<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr>""");
            Add(z, "word/document.xml", $"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>""");
        }
        return ms.ToArray();
        static string Para(string style, string text) => $"""<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{Run(text)}</w:p>""";
        static string Run(string text) => $"""<w:r><w:t xml:space="preserve">{XmlEsc(text)}</w:t></w:r>""";
    }

    // ============================================================ PPTX ============
    // slides: [ {title, bullets:[..]} ] ; slide 0 rendered as title slide with subtitle
    public static byte[] BuildPptx(string deckTitle, string subtitle, string accentHex, JsonArray slides)
    {
        accentHex = (accentHex ?? "1F6E54").TrimStart('#').ToUpperInvariant();
        if (accentHex.Length != 6) accentHex = "1F6E54";
        int n = slides.Count;
        using var ms = new MemoryStream();
        using (var z = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            var ct = new StringBuilder("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>""");
            for (int i = 1; i <= n; i++)
                ct.Append($"""<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>""");
            ct.Append("</Types>");
            Add(z, "[Content_Types].xml", ct.ToString());

            Add(z, "_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>""");

            var pRels = new StringBuilder("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>""");
            var sldIds = new StringBuilder();
            for (int i = 1; i <= n; i++)
            {
                pRels.Append($"""<Relationship Id="rIdSl{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>""");
                sldIds.Append($"""<p:sldId id="{255 + i}" r:id="rIdSl{i}"/>""");
            }
            pRels.Append("</Relationships>");
            Add(z, "ppt/_rels/presentation.xml.rels", pRels.ToString());
            Add(z, "ppt/presentation.xml", $"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst><p:sldIdLst>{sldIds}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>""");

            Add(z, "ppt/theme/theme1.xml", $"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="DocForge"><a:themeElements><a:clrScheme name="DocForge"><a:dk1><a:srgbClr val="10221C"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="223A32"/></a:dk2><a:lt2><a:srgbClr val="F2F6F4"/></a:lt2><a:accent1><a:srgbClr val="{accentHex}"/></a:accent1><a:accent2><a:srgbClr val="4FB08C"/></a:accent2><a:accent3><a:srgbClr val="B9C9C2"/></a:accent3><a:accent4><a:srgbClr val="8AA79B"/></a:accent4><a:accent5><a:srgbClr val="5C7A6E"/></a:accent5><a:accent6><a:srgbClr val="32574A"/></a:accent6><a:hlink><a:srgbClr val="{accentHex}"/></a:hlink><a:folHlink><a:srgbClr val="5C7A6E"/></a:folHlink></a:clrScheme><a:fontScheme name="DocForge"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>""");

            Add(z, "ppt/slideMasters/_rels/slideMaster1.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdL1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rIdTh" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>""");
            Add(z, "ppt/slideMasters/slideMaster1.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdL1"/></p:sldLayoutIdLst></p:sldMaster>""");

            Add(z, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>""");
            Add(z, "ppt/slideLayouts/slideLayout1.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sldLayout>""");

            for (int i = 0; i < n; i++)
            {
                var s = slides[i]!;
                bool isTitle = i == 0;
                var titleTxt = s["title"]?.GetValue<string>() ?? "";
                var bullets = s["bullets"]?.AsArray() ?? new JsonArray();
                var shapes = new StringBuilder();
                uint id = 2;
                if (isTitle)
                {
                    shapes.Append(Shape(id++, "accentbar", 0, 5486400, 12192000, 118872, $"""<a:solidFill><a:srgbClr val="{accentHex}"/></a:solidFill>"""));
                    shapes.Append(TextShape(id++, "Title", 838200, 2286000, 10515600, 1600200, titleTxt, 4400, true, "10221C"));
                    var sub = subtitle ?? "";
                    if (isTitle && sub.Length > 0)
                        shapes.Append(TextShape(id++, "Subtitle", 838200, 3962400, 10515600, 800100, sub, 2000, false, "5C7A6E"));
                }
                else
                {
                    shapes.Append(Shape(id++, "accentbar", 0, 0, 12192000, 91440, $"""<a:solidFill><a:srgbClr val="{accentHex}"/></a:solidFill>"""));
                    shapes.Append(TextShape(id++, "Title", 838200, 365760, 10515600, 990600, titleTxt, 3200, true, "10221C"));
                    if (bullets.Count > 0)
                    {
                        var paras = new StringBuilder();
                        foreach (var b in bullets)
                            paras.Append($"""<a:p><a:pPr marL="342900" indent="-342900"><a:buClr><a:srgbClr val="{accentHex}"/></a:buClr><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="en-US" sz="2000" dirty="0"><a:solidFill><a:srgbClr val="223A32"/></a:solidFill></a:rPr><a:t>{XmlEsc(b!.GetValue<string>())}</a:t></a:r></a:p>""");
                        shapes.Append($"""<p:sp><p:nvSpPr><p:cNvPr id="{id++}" name="Body"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="1554480"/><a:ext cx="10515600" cy="4800600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>{paras}</p:txBody></p:sp>""");
                    }
                }
                Add(z, $"ppt/slides/_rels/slide{i + 1}.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>""");
                Add(z, $"ppt/slides/slide{i + 1}.xml", $"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>""");
            }
        }
        return ms.ToArray();

        static string Shape(uint id, string name, long x, long y, long cx, long cy, string fill) =>
            $"""<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>""";
        static string TextShape(uint id, string name, long x, long y, long cx, long cy, string text, int sz, bool bold, string color) =>
            $"""<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="{sz}" {(bold ? "b=\"1\"" : "")} dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr><a:t>{XmlEsc(text)}</a:t></a:r></a:p></p:txBody></p:sp>""";
    }

    // ============================================================ TEMPLATE FILL ============
    // Fill {{placeholders}} in docx/xlsx even when Word splits them across runs.
    public static List<string> FindPlaceholders(byte[] fileBytes, string kind)
    {
        var found = new List<string>();
        using var ms = new MemoryStream(fileBytes);
        using var z = new ZipArchive(ms, ZipArchiveMode.Read);
        foreach (var entry in z.Entries)
        {
            if (!TargetPart(entry.FullName, kind)) continue;
            using var r = new StreamReader(entry.Open());
            var xml = r.ReadToEnd();
            var text = kind == "docx" ? JoinDocxText(xml) : xml;
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(text, @"\{\{\s*([A-Za-z0-9_ .\-]+?)\s*\}\}"))
            {
                var name = m.Groups[1].Value.Trim();
                if (!found.Contains(name)) found.Add(name);
            }
        }
        return found;
    }
    static bool TargetPart(string p, string kind) =>
        kind == "docx" ? (p is "word/document.xml" or "word/footer1.xml" or "word/footer2.xml" or "word/header1.xml" or "word/header2.xml")
                       : (p == "xl/sharedStrings.xml" || p.StartsWith("xl/worksheets/"));
    static string JoinDocxText(string xml)
    {
        var sb = new StringBuilder();
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(xml, @"<w:t[^>]*>([^<]*)</w:t>"))
            sb.Append(System.Net.WebUtility.HtmlDecode(m.Groups[1].Value));
        return sb.ToString();
    }

    public static byte[] FillTemplate(byte[] fileBytes, string kind, Dictionary<string, string> values)
    {
        using var src = new MemoryStream(fileBytes);
        using var zin = new ZipArchive(src, ZipArchiveMode.Read);
        using var outMs = new MemoryStream();
        using (var zout = new ZipArchive(outMs, ZipArchiveMode.Create, true))
        {
            foreach (var entry in zin.Entries)
            {
                var ne = zout.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var inS = entry.Open();
                using var outS = ne.Open();
                if (TargetPart(entry.FullName, kind))
                {
                    using var r = new StreamReader(inS);
                    var xml = r.ReadToEnd();
                    xml = kind == "docx" ? FillDocxXml(xml, values) : FillPlainXml(xml, values);
                    var bytes = new UTF8Encoding(false).GetBytes(xml);
                    outS.Write(bytes, 0, bytes.Length);
                }
                else inS.CopyTo(outS);
            }
        }
        return outMs.ToArray();
    }
    static string FillPlainXml(string xml, Dictionary<string, string> values)
    {
        foreach (var (k, v) in values)
            xml = System.Text.RegularExpressions.Regex.Replace(xml, @"\{\{\s*" + System.Text.RegularExpressions.Regex.Escape(k) + @"\s*\}\}", XmlEsc(v));
        return xml;
    }
    // Word splits {{name}} into multiple <w:t> runs. Strategy: per paragraph, join run texts,
    // do replacements on the joined string, then write the new text back into the first run
    // and blank the others — preserving the paragraph's first-run formatting.
    static string FillDocxXml(string xml, Dictionary<string, string> values)
    {
        var doc = XDocument.Parse(xml);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach (var p in doc.Descendants(w + "p"))
        {
            var ts = p.Descendants(w + "t").ToList();
            if (ts.Count == 0) continue;
            var joined = string.Concat(ts.Select(t => t.Value));
            if (!joined.Contains("{{")) continue;
            var replaced = joined;
            foreach (var (k, v) in values)
                replaced = System.Text.RegularExpressions.Regex.Replace(replaced, @"\{\{\s*" + System.Text.RegularExpressions.Regex.Escape(k) + @"\s*\}\}", v.Replace("$", "$$"));
            if (replaced == joined) continue;
            ts[0].Value = replaced;
            ts[0].SetAttributeValue(XNamespace.Xml + "space", "preserve");
            for (int i = 1; i < ts.Count; i++) ts[i].Value = "";
        }
        return doc.Declaration is null ? doc.ToString(SaveOptions.DisableFormatting)
             : doc.Declaration + doc.ToString(SaveOptions.DisableFormatting);
    }
}
