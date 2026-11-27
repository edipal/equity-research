<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/initiating-coverage/references/task5-report-assembly.md), Apache-2.0; modified for Google Antigravity. -->

# Task 5: Report Assembly - Detailed Workflow

Step-by-step instructions for executing Task 5 (Report Assembly) of the `initiating-coverage` skill. The output is staged as an internal institutional draft: **`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`**.

## Task Overview

**Purpose:** Programmatically assemble `research/companies/<TICKER>/reports/<TICKER>_Initiation_Report_YYYY-MM-DD.docx` using Python (`python-docx` + `openpyxl`), integrating Task 1 prose, Task 2/3 Excel tables, Task 4 PNG charts, and mandatory Disclosures & Important Disclaimers.

**Prerequisites (MANDATORY — stop if any Task 1–4 deliverable is missing):**
1. **Task 1:** `research/companies/<TICKER>/reports/<TICKER>_Research_Document_<Date>.md`
2. **Task 2 & Task 3:** `research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx` with canonical tabs (`Revenue Model`, `Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios`, `WACC`, `DCF`, `Comparable Companies`, `Valuation Summary` — plus inline 5×5 sensitivity on `DCF` or optional `Sensitivity` tab) verified by `recalc_xlsx.py` (exit `0`) and `validate_dcf.py` (exit `0`), plus `<TICKER>_Valuation_Analysis_<Date>.md`
3. **Task 4:** `research/companies/<TICKER>/reports/charts/` containing 300 DPI PNG charts + `chart_index.txt`

> **Compliance & Integrity Rules (per `.agents/rules/research-integrity.md`):**
> - **Drafts only:** Label Page 1 header and document footer `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)`. Never describe the output as client-ready for external distribution without human analyst sign-off.
> - **No fabricated identities:** Use `[ANALYST NAME — REVIEW REQUIRED]` and `[DATE]` placeholders — NEVER invent analyst names, CFA/Ph.D. credentials, emails, phone numbers, or firm names.
> - **No fabrication to hit quotas:** Target ~30–50 pages, ~10,000–15,000 words, ~20–30+ charts, and ~12–20 tables where primary filings support that depth. Never fabricate segment rows, geographies, or charts to hit numerical targets; tag estimates `[ASSUMPTION]` and unverified items `[UNSOURCED]`.
> - **No placeholders in prose:** Write every included analytical section in full — never write `"This section would include..."` or `"Charts would be inserted here..."`.

---

## Report Structure & Target Allocation

| Section | Target Length | Content & Source |
|---|---|---|
| **Page 1: Initiating Coverage** | 1 page (~500–700 words) | Draft banner, **`INITIATING COVERAGE`** header (NEVER `"Company Update"`), standardized **`Buy / Hold / Sell`** rating box, `[ANALYST NAME — REVIEW REQUIRED]`, Figure 1 stock chart, 3–4 quantified `■` bullets, financial summary table (`[FY-2]A–[FY+2]E`) — **write Page 1 LAST** |
| **Page 2: Table of Contents** | 1 page | Word `TOC` field (updated on LibreOffice conversion) |
| **Investment Thesis & Risks** | ~3–5 pages (~1,400–2,100 words) | 3–5 quantified thesis pillars + 8–12 balanced risks across 4 categories (Company, Industry, Financial, Macro) |
| **Company 101** | ~10–12 pages (~6,000–8,000 words) | Reuse Task 1 `.md` prose verbatim (converted to DOCX styles) with Task 4 charts embedded inline |
| **Financial Analysis, Projections & Scenarios** | ~10–13 pages (~5,000–7,000 words) | Historical performance (`[FY-4]A–[FY]A` + LTM stub), Projection Assumptions (product/geography/margin/OpEx/CapEx/NWC `[ASSUMPTION]` detail), Bear/Base/Bull Scenario Analysis, Growth Drivers, extracted Excel tables + charts |
| **Valuation Analysis** | ~8–10 pages (~1,500–2,500 words) | DCF (`WACC` + `DCF` tabs + sensitivity heatmap `chart_28`), Comps (`Comparable Companies` tab with mandatory Max/75th/Median/25th/Min summary), Precedents (cross-check only if applicable), Valuation Summary (`Valuation Summary` tab + football field `chart_32`), 12M Price Target, `Buy / Hold / Sell` rating, Catalysts & Risks |
| **Appendices & Disclosures (MANDATORY)** | ~4–8 pages | Appendix A: Extended Model Tables; Appendix B: Data Sources & References (dated clickable hyperlinks); **Appendix C: Disclosures & Important Disclaimers** |

---

## Programmatic Assembly Workflow (`python-docx` + `openpyxl`)

### Step 1: Verify & Load Inputs
1. Run `python3 scripts/recalc_xlsx.py research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx` (non-destructive default / `--check-only`) to verify zero formula errors and zero uncached values on the Task 3 recalculated workbook before reading `openpyxl.load_workbook(..., data_only=True)` (if any `openpyxl` edit was made since Task 3, first repopulate cached values with `--in-place` or `--output <recalc.xlsx>`).
2. Read Task 1 and Task 3 markdown files with `Path.read_text(encoding="utf-8")`.
3. Parse `research/companies/<TICKER>/reports/charts/chart_index.txt` to map every generated PNG (and confirm any disclosure-driven omissions).

### Step 2: Essential `python-docx` Helpers (Sizing, Hyperlinks, Page Numbers, TOC)

- **Image width rule:** Default `python-docx` Letter pages (`8.5"` wide with `1.0"`–`1.25"` left/right margins) have `6.0"`–`6.5"` of usable text width. **Cap all embedded chart widths at `Inches(6.0)`** (`Inches(5.5)` for square/org visuals) so images NEVER clip margins.
- **Clickable hyperlinks, page numbers & TOC XML helpers:**

```python
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_hyperlink(paragraph, text: str, url: str) -> None:
    """Insert a clickable external hyperlink into a python-docx paragraph."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK,
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(color)
    r_pr.append(underline)
    new_run.append(r_pr)

    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def add_field(run, field_instruction: str) -> None:
    """Insert a Word XML field (e.g., 'PAGE', 'NUMPAGES', or 'TOC \\\\o \"1-3\" \\\\h \\\\z \\\\u')."""
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_instruction} "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_sep, fld_end])

def add_chart(doc, img_path, caption: str, source: str, width_inches: float = 6.0) -> None:
    """Embed a 300 DPI PNG sized <= 6.0 inches with sequential caption and source line."""
    cap_p = doc.add_paragraph()
    cap_run = cap_p.add_run(caption)
    cap_run.bold = True
    cap_run.font.size = Pt(10)
    doc.add_picture(str(img_path), width=Inches(min(width_inches, 6.0)))
    src_p = doc.add_paragraph()
    src_run = src_p.add_run(f"Source: {source}")
    src_run.italic = True
    src_run.font.size = Pt(8.5)
    src_run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
```

Add `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)` + `Page [PAGE] of [NUMPAGES]` to the section footer, and insert `add_field(toc_run, 'TOC \\o "1-3" \\h \\z \\u')` on Page 2.

### Step 3: Extract Canonical Tables from Excel (`openpyxl`)

Read cached values directly from `<TICKER>_Financial_Model_<Date>.xlsx`:
- Page 1 Summary Financials (`Income Statement` + `Valuation Summary`)
- Full Income Statement (`Income Statement`)
- Revenue by Product/Segment and Revenue by Geography (`Revenue Model` — matching filing disclosures)
- Cash Flow Statement & Balance Sheet highlights (`Cash Flow Statement`, `Balance Sheet`)
- Scenario Comparison (`Scenarios`)
- WACC & DCF Key Assumptions (`WACC`, `DCF`)
- DCF Sensitivity Matrix (`DCF` inline sensitivity or optional `Sensitivity` tab)
- Comparable Companies Table including **mandatory** Max / 75th / Median / 25th / Min summary (`Comparable Companies`)
- Valuation Summary Table (`Valuation Summary`)

### Step 4: Assemble Sections & Mandatory Disclosures Appendix

Follow [assets/report-template.md](../assets/report-template.md) for layout and Calibri typography:
1. Convert Task 1 markdown headings/bullets/tables into native Word styles (zero raw `#` or `**` syntax left in text).
2. Intersperse Task 4 charts inline immediately after the paragraph discussing each metric (avoid chart-dump pages or text-only walls).
3. Write **Appendix C: Disclosures & Important Disclaimers (MANDATORY)** containing:
   - Draft status notice (`DRAFT — FOR INTERNAL ANALYST REVIEW ONLY`)
   - Not investment advice / no solicitation disclaimer
   - `[ASSUMPTION]` and `[UNSOURCED]` methodology disclosure
   - Review sign-off block: `[ANALYST NAME — REVIEW REQUIRED]` | `[SUPERVISORY REVIEW — PENDING]` | `[DATE]`
4. Write **Page 1 (`INITIATING COVERAGE`) LAST** so every headline metric ties cleanly to the completed tables.

### Step 5: Render PDF & Verify Page Count / Layout (MANDATORY QA GATE)

Convert the `.docx` to PDF via LibreOffice (which also updates Word TOC and `PAGE`/`NUMPAGES` fields) and verify exact page count with `pypdf`:

```bash
soffice --headless --convert-to pdf \
  --outdir /tmp/docx_qa \
  research/companies/<TICKER>/reports/<TICKER>_Initiation_Report_<Date>.docx

python3 -c '
from pypdf import PdfReader
reader = PdfReader("/tmp/docx_qa/<TICKER>_Initiation_Report_<Date>.pdf")
print("Rendered PDF page count:", len(reader.pages))
'
```

*(Optionally convert sampled PDF pages to PNG via `pdftoppm -png -r 150` to inspect Page 1, chart margins, and tables visually. Clean up `/tmp/docx_qa` afterward.)*

---

## Final Quality Verification

Complete [assets/quality-checklist.md](../assets/quality-checklist.md) before handing off the draft. Confirm:
- [ ] File saved at `research/companies/<TICKER>/reports/<TICKER>_Initiation_Report_YYYY-MM-DD.docx`
- [ ] Prominent `DRAFT — FOR INTERNAL ANALYST REVIEW ONLY (NOT INVESTMENT ADVICE)` banner and footer
- [ ] Page 1 header reads `INITIATING COVERAGE`; rating uses `Buy / Hold / Sell`; analyst block uses `[ANALYST NAME — REVIEW REQUIRED]`
- [ ] Disclosures & Important Disclaimers appendix + Data Sources & References appendix (with clickable hyperlinks) present
- [ ] All embedded charts `<= 6.0"` wide with sequential figure numbers and dated source lines
- [ ] 10–15 key metrics spot-checked against `<TICKER>_Financial_Model_<Date>.xlsx` — exact match
- [ ] Rendered PDF verified via LibreOffice + `pypdf` with zero broken layout or raw markdown tokens
