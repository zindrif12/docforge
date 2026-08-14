# DocForge v2 — the AI document workshop

One app, four forges — now with a real template engine, batch generation, and verification.

## v2 template language
```
{{client_name}}                     scalar field
{{amount | currency:LKR}}           filters: currency:CODE, number, round:n, date:long|short, upper, lower
{{#each line_items}} ... {{/each}}  loops — repeat table rows (docx/xlsx) or paragraphs per item
{{#if discount_applies}} ... {{/if}} conditionals — keep or drop whole blocks
```
Loops expand Word table rows and Excel rows (cell references and per-row formulas are renumbered
automatically — `=B2*C2` in the template row becomes `=B3*C3`, `=B4*C4`, … in the clones; totals
below a loop should use column ranges like `=SUM(D:D)`). Nested loops are rejected with a clear
error in v2.

## Batch
Select a template, drop a CSV: headers map to scalar fields, one document per row, optional
`filename` column names each file, capped at 200 rows, delivered as a zip. No AI involved —
deterministic and fast.

## Verification
Every AI fill returns a per-field report: **verified** (the value appears in your source),
**derived** (the AI transformed something that does appear — the supporting quote is included),
**unverified** (take a look), or **missing** (flagged, never invented). Nothing goes to a client
on the AI's word alone.

- **Fill** — upload your own `.docx`/`.xlsx` template containing `{{placeholders}}`, paste any messy source
  (email threads, meeting notes, CRM text) and get the finished document back with your formatting intact.
  Handles placeholders that Word has split across runs. Shows AI confidence and flags fields it could not find
  rather than inventing them.
- **Sheet** — describe a workbook in plain English and get a real `.xlsx` with **live formulas**
  (`=SUM`, margins, running balances), styled headers and multiple sheets.
- **Doc** — a brief becomes a structured `.docx` with title, headings, bullets and tables.
- **Deck** — notes become a themed `.pptx` with a title slide, accent colour of your choice and clean bullet slides.

All Office files are generated from scratch over `ZipArchive` + hand-written OpenXML — no NuGet packages,
no Office installation, no repair dialogs.

## Privacy
Source text and filled values are processed in memory and returned to you — they are never stored.
History records only module, title, field names and counts.

## Stack
ASP.NET Core 8 (zero NuGet) · Gemini 2.5 Flash (JSON mode) · Supabase (template library + history) · vanilla JS frontend.

## Run locally
```
MOCK_AI=1 dotnet run          # no keys needed, canned AI responses
GEMINI_API_KEY=... dotnet run # real AI, file-based storage
```

## Deploy (Render)
1. Run `supabase-setup.sql` in your Supabase project's SQL Editor.
2. Push this repo to GitHub, create a Render Blueprint from it (`render.yaml` is picked up).
3. Set env vars: `GEMINI_API_KEY`, `SUPABASE_URL` (project URL only, no `/rest/v1`), `SUPABASE_SERVICE_KEY`.
4. Smoke test `/api/health` — expect `"provider":"gemini","store":"supabase"`.

## Tests
`tests/suite.sh` — 35 checks. Every generated file is independently opened and verified by
openpyxl / python-docx / python-pptx, including a split-run placeholder fill round-trip.

## API additions in v2
`POST /api/template/fill-manual` — render a JSON payload directly, no AI (integrations, pipelines).
`POST /api/template/batch` — templateId + csvBase64 → zip.
`GET /api/health` — now reports `version: 2` and the engine feature list.

## Roadmap
Scheduled recurring runs · data-source connectors (SQL / OData / Sheets) · PDF export · pptx templates · nested loops.
