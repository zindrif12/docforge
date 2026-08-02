# DocForge — the AI document workshop

One app, four forges. Real Office files, zero dependencies.

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
`tests/suite.sh` — 25 checks. Every generated file is independently opened and verified by
openpyxl / python-docx / python-pptx, including a split-run placeholder fill round-trip.

## Roadmap
PDF export · scheduled recurring reports (ReportRun) · pptx templates for Fill · bulk fill from CSV.
