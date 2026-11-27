---
name: pptx-author
description: Build institutional-quality PowerPoint (.pptx) equity research decks, sector primers, and pitch presentations programmatically using Python (python-pptx) with model-linked charts and visual verification. Use when creating presentation slides in Antigravity.
---

# pptx-author

Build clean, institutional-grade PowerPoint (`.pptx`) presentations programmatically in Antigravity using Python (`python-pptx`) and high-DPI rendered charts (`matplotlib`).

## Output Contract

- Default output location:
  - Single-company pitch / coverage decks: `research/companies/<TICKER>/reports/<name>.pptx`
  - Sector / competitive landscape decks: `research/sectors/<sector-slug>/decks/<name>.pptx`
  - Thematic / cross-sector pitch decks: `research/screens/<name>.pptx` (create parent directories if needed)
- Return the workspace file path and a slide-by-slide summary in your final response.

## How to Build the Deck

Write a Python script using `python-pptx` and run it in the terminal:

```python
from pptx import Presentation
from pptx.util import Inches, Pt

# Use a user-provided template if present, otherwise standard 16:9 widescreen
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title-only
slide.shapes.title.text = "Valuation Summary — DCF & Peer Multiples"
# Add tables, high-DPI chart PNGs, and source footnotes...

prs.save("research/sectors/cloud-saas/decks/primer.pptx")
```

## Institutional Conventions

- **Action titles (one takeaway per slide)**: The slide title states the analytical conclusion, not just a topic label.
- **Traceability**: Every number traces to an underlying model or filing; include a concise source & date footnote at the bottom of every slide.
- **Template fidelity**: Use a user-provided `.pptx` template when supplied; otherwise use a clean, restrained institutional palette (dark navy `#1F4E79`, slate grey, crisp whitespace).
- **High-DPI charts**: Render financial charts at 300 DPI via `matplotlib` and embed PNGs rather than relying on cramped default chart layouts.
- **Drafts only**: Write files locally in the workspace; never email, publish, or distribute externally without explicit user direction.

## Visual Verification

When LibreOffice and `pdftoppm` are available, convert the generated `.pptx` to PDF and slide PNGs to visually inspect layout, text overlap, and chart legibility before finalizing:

```bash
soffice --headless --convert-to pdf <deck.pptx> --outdir /tmp/pptx_qa
pdftoppm -png -r 150 /tmp/pptx_qa/<deck>.pdf /tmp/pptx_qa/slide
```
