---
trigger: glob
globs: "research/**/*"
description: Directory structure, naming conventions, and institutional formatting standards for all equity research deliverables stored under research/.
---

# Research Deliverables & Workspace Organization

1. **Directory layout.** Place every output in the matching location under `research/`:
   - `research/companies/<TICKER>/thesis.md` — Living investment thesis (maintained via `thesis-tracker`)
   - `research/companies/<TICKER>/models/` — `.xlsx` models (`<TICKER>_3Statement_YYYY-MM-DD.xlsx`, `<TICKER>_DCF_YYYY-MM-DD.xlsx`, `<TICKER>_Comps_YYYY-MM-DD.xlsx`) and generator scripts
   - `research/companies/<TICKER>/notes/` — Earnings previews, earnings updates, channel checks, meeting notes (`<TICKER>_Q3FY26_Earnings_YYYY-MM-DD.md`)
   - `research/companies/<TICKER>/reports/` — Initiating coverage reports (`.md`, `.docx`, `.pdf`), charts, and pitch decks (`.pptx`)
   - `research/companies/<TICKER>/data/` — Raw filings, extracted CSVs, and source logs
   - `research/sectors/<sector-slug>/` — Sector primers (`overview.md`), competitive landscapes (`competitive-landscape.md`), peer comps (`comps/`), and decks (`decks/`)
   - `research/screens/` — Quantitative and thematic stock screens (`YYYY-MM-DD_<screen-name>.md` / `.xlsx`)
   - `research/morning-notes/YYYY-MM-DD.md` — Daily morning notes
   - `research/catalysts.md` — Master coverage catalyst calendar
2. **Executive summary first.** Every research note, primer, and report must open with a concise headline takeaway (rating/stance, target/valuation range or key thesis hook, and 3–5 bullet takeaways) before diving into supporting analysis.
3. **Consistent units and precision.** State currency and unit scale clearly in every table header (e.g., `USD Millions, except per-share data`). Use consistent decimals: `$M`/`$B` (1 decimal or integer), EPS (`$0.00`), multiples (`12.4x`), percentages (`14.2%`), margin changes (`+120 bps`).
4. **Source footnotes on every table and chart.** End every table, chart, and memo with explicit source citations and as-of dates.
