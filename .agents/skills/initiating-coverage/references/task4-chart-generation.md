<!-- Derived from anthropics/financial-services@3e32c22 (plugins/vertical-plugins/equity-research/skills/initiating-coverage/references/task4-chart-generation.md), Apache-2.0; modified for Google Antigravity. -->

# Task 4: Chart Generation - Detailed Workflow

Step-by-step instructions for executing Task 4 (Chart Generation) of the `initiating-coverage` skill.

## Task Overview

**Purpose:** Generate institutional-quality 300 DPI PNG charts saved directly in `research/companies/<TICKER>/reports/charts/` alongside `chart_index.txt` for embedding in Task 5's `.docx` draft report.

**Prerequisites (MANDATORY — stop if Tasks 1–3 are incomplete):**
- **Task 1:** `research/companies/<TICKER>/reports/<TICKER>_Research_Document_<Date>.md`
- **Task 2 & Task 3:** `research/companies/<TICKER>/models/<TICKER>_Financial_Model_<Date>.xlsx` containing canonical tabs (`Revenue Model`, `Income Statement`, `Cash Flow Statement`, `Balance Sheet`, `Scenarios`, `WACC`, `DCF`, `Comparable Companies`, `Valuation Summary` — plus inline 5×5 sensitivity on `DCF` or optional `Sensitivity` tab) verified via `recalc_xlsx.py` and `validate_dcf.py`, plus `<TICKER>_Valuation_Analysis_<Date>.md`
- **External Market Data:** Verified historical stock price series (and historical multiples where available) with provider + as-of date.

> **No-Fabrication Rule:** Target **20–30+ charts** across the report where supported by primary filings, model tabs, and verified market data. **NEVER invent data to hit a chart count.** If a company reports only one operating segment, lacks geographic disaggregation, or does not disclose customer cohorts, omit that chart and log the omission (and reason) cleanly in `chart_index.txt`. Tag chart series driven by forward projections `[ASSUMPTION]` / `"E"` suffix.

---

## Priority Chart Catalog (Target 20–30+ Where Disclosed)

### 4 Core Priority Charts (Always Build Where Disclosed) ⭐
1. **`chart_03_revenue_by_product_stacked_area.png`** — Revenue by Product/Segment (stacked area; omit only for true single-product companies, noting in `chart_index.txt`)
2. **`chart_04_revenue_by_geography_stacked_bar.png`** — Revenue by Geography (stacked bar; omit only when 10-K discloses single-country revenue)
3. **`chart_28_dcf_sensitivity_heatmap.png`** — DCF Sensitivity 2-Way Heatmap (5×5 centered on Base WACC & Terminal `g`)
4. **`chart_32_valuation_football_field.png`** — Valuation Football Field (horizontal range bars + current price & target price lines)

### Target Core Set (~20–25 Charts Mapped to Source Files)

| Section | Charts & Canonical Data Source |
|---|---|
| **Page 1 Summary** | `chart_01` Stock Price Performance (12–24M vs. benchmark; external market data) |
| **Financial Performance (`Task 2` tabs)** | `chart_02` Revenue Growth (`Income Statement`), `chart_03`⭐ Revenue by Product (`Revenue Model`), `chart_04`⭐ Revenue by Geography (`Revenue Model`), `chart_10` Gross Margin (`Income Statement`), `chart_11` EBITDA Margin (`Income Statement`), `chart_12` Free Cash Flow (`Cash Flow Statement`), `chart_13` Operating Metrics Dashboard, `chart_14` Bear/Base/Bull Scenarios (`Scenarios`) |
| **Company 101 & Market (`Task 1`)** | `chart_05` Company Overview/Scale, `chart_06` Milestones Timeline, `chart_07` Leadership/Org Overview, `chart_08` Product Portfolio Matrix, `chart_09` Customer Segmentation *(where disclosed)*, `chart_15` TAM Evolution, `chart_16` Competitive Positioning Matrix, `chart_17` Market Share *(where disclosed)*, `chart_18` Competitive Benchmarking |
| **Valuation (`Task 3` tabs)** | `chart_28`⭐ DCF Sensitivity Heatmap (`DCF` / `Sensitivity`), `chart_29` DCF Value Bridge Waterfall (`DCF`), `chart_30` Comps Growth vs. Multiple Scatter (`Comparable Companies`), `chart_31` Peer Multiples Bar (`Comparable Companies`), `chart_32`⭐ Valuation Football Field (`Valuation Summary`), `chart_33` Bear/Base/Bull Implied Prices (`Valuation Summary`), `chart_34` Historical Valuation Multiples *(where available)* |

### Optional Supplemental Charts (Where Disclosed in Filings)
`chart_19` Customer Acquisition Trends, `chart_20` Unit Economics Evolution, `chart_21` Product Roadmap, `chart_22` Geographic Footprint, `chart_23` R&D Intensity, `chart_24` S&M Efficiency, `chart_25` Working Capital Cycle (DSO/DIO/DPO), `chart_26` Debt Maturity Schedule (10-K debt footnote), `chart_27` Ownership Structure (DEF 14A), `chart_35` Street Target Distribution *(only if verified feed available)*.

---

## Step-by-Step Python Chart Generation Workflow

### Step 1: Environment & Shared Output Directory

Install libraries from root `requirements.txt` (`matplotlib`, `seaborn`, `Pillow`, `pandas`, `numpy`, `openpyxl`). Always write charts into `output_dir = Path("research/companies/<TICKER>/reports/charts")` with `mkdir(parents=True, exist_ok=True)`:

```python
from pathlib import Path
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np

TICKER = "TICKER"  # replace with actual ticker
OUTPUT_DIR = Path(f"research/companies/{TICKER}/reports/charts")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

plt.style.use("seaborn-v0_8-whitegrid")
plt.rcParams["font.sans-serif"] = ["Calibri", "DejaVu Sans"]
DPI = 300
```

### Step 2: Load Verified Numbers from Excel (`data_only=True`)

Always read reconciled cached values from `<TICKER>_Financial_Model_<Date>.xlsx` after running `scripts/recalc_xlsx.py` — never re-type numbers manually when they exist in Excel. Use dynamic fiscal year labels pulled directly from the workbook headers (e.g., `[FY-4]A`..`[FY]A`, `LTM`, `[FY+1]E`..`[FY+5]E`).

### Step 3: Core Chart Implementation Snippets

#### 1. Revenue by Product — Stacked Area (`chart_03`) ⭐

```python
def create_revenue_by_product_chart(
    output_dir: Path,
    year_labels: list[str],
    last_actual_idx: int,
    series_dict: dict[str, list[float]],
    source_text: str,
) -> Path:
    fig, ax = plt.subplots(figsize=(9, 5.2))
    x = np.arange(len(year_labels))
    palette = ["#1F4E79", "#2E75B6", "#70AD47", "#ED7D31", "#A5A5A5"]

    ax.stackplot(
        x,
        *series_dict.values(),
        labels=list(series_dict.keys()),
        colors=palette[: len(series_dict)],
        alpha=0.88,
    )
    ax.axvline(x=last_actual_idx, color="#595959", linestyle="--", linewidth=1.1)
    ax.text(last_actual_idx + 0.15, ax.get_ylim()[1] * 0.94, "Projected ->", fontsize=9, color="#595959")

    ax.set_xticks(x)
    ax.set_xticklabels(year_labels, rotation=30, ha="right")
    ax.set_ylabel("Revenue ($M)", fontsize=11, fontweight="bold")
    ax.set_title("Figure 3 - Revenue by Product / Segment", fontsize=13, fontweight="bold", pad=14)
    ax.legend(loc="upper left", frameon=False, fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.text(0.08, 0.01, f"Source: {source_text}", fontsize=8.5, style="italic", color="#595959")

    out_path = output_dir / "chart_03_revenue_by_product_stacked_area.png"
    plt.tight_layout(rect=(0, 0.04, 1, 1))
    plt.savefig(out_path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close()
    return out_path
```

#### 2. Revenue by Geography — Stacked Bar (`chart_04`) ⭐

```python
def create_revenue_by_geography_chart(
    output_dir: Path,
    year_labels: list[str],
    region_dict: dict[str, list[float]],
    source_text: str,
) -> Path:
    fig, ax = plt.subplots(figsize=(9, 5.2))
    x = np.arange(len(year_labels))
    width = 0.62
    palette = ["#1F4E79", "#2E75B6", "#5B9BD5", "#A5A5A5", "#ED7D31"]
    bottom = np.zeros(len(year_labels), dtype=float)

    for idx, (region, vals) in enumerate(region_dict.items()):
        arr = np.asarray(vals, dtype=float)
        ax.bar(x, arr, width, bottom=bottom, label=region, color=palette[idx % len(palette)])
        bottom += arr

    ax.set_xticks(x)
    ax.set_xticklabels(year_labels, rotation=30, ha="right")
    ax.set_ylabel("Revenue ($M)", fontsize=11, fontweight="bold")
    ax.set_title("Figure 4 - Revenue by Geography", fontsize=13, fontweight="bold", pad=14)
    ax.legend(loc="upper left", frameon=False, fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.text(0.08, 0.01, f"Source: {source_text}", fontsize=8.5, style="italic", color="#595959")

    out_path = output_dir / "chart_04_revenue_by_geography_stacked_bar.png"
    plt.tight_layout(rect=(0, 0.04, 1, 1))
    plt.savefig(out_path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close()
    return out_path
```

#### 3. DCF Sensitivity 5×5 Heatmap (`chart_28`) ⭐ — Float-Safe Formatting!

```python
def create_dcf_sensitivity_heatmap(
    output_dir: Path,
    wacc_values: list[float],
    terminal_growth: list[float],
    price_matrix: np.ndarray,
    source_text: str,
) -> Path:
    """Use float formatting (fmt='.2f' or '.1f' — NEVER fmt='d' on float arrays!)."""
    df = pd.DataFrame(
        price_matrix.astype(float),
        index=[f"{w:.2f}%" for w in wacc_values],
        columns=[f"{g:.2f}%" for g in terminal_growth],
    )
    fig, ax = plt.subplots(figsize=(7.5, 5.5))
    sns.heatmap(
        df,
        annot=True,
        fmt=".2f",  # float-safe dollar price; use '.1%' for percentage heatmaps
        cmap="RdYlGn",
        cbar_kws={"label": "Implied Share Price ($)"},
        linewidths=0.6,
        linecolor="white",
        ax=ax,
    )
    ax.set_xlabel("Terminal Growth Rate (g)", fontsize=11, fontweight="bold")
    ax.set_ylabel("WACC", fontsize=11, fontweight="bold")
    ax.set_title("Figure 28 - DCF Sensitivity Analysis ($/share)", fontsize=13, fontweight="bold", pad=14)
    plt.yticks(rotation=0)
    fig.text(0.08, 0.01, f"Source: {source_text}", fontsize=8.5, style="italic", color="#595959")

    out_path = output_dir / "chart_28_dcf_sensitivity_heatmap.png"
    plt.tight_layout(rect=(0, 0.04, 1, 1))
    plt.savefig(out_path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close()
    return out_path
```

#### 4. Valuation Football Field (`chart_32`) ⭐

```python
def create_valuation_football_field(
    output_dir: Path,
    methods: list[str],
    lows: list[float],
    highs: list[float],
    current_price: float,
    target_price: float,
    source_text: str,
) -> Path:
    fig, ax = plt.subplots(figsize=(9, 4.8))
    y_pos = np.arange(len(methods))
    colors = ["#1F4E79", "#2E75B6", "#70AD47", "#A5A5A5"]

    for i, (method, lo, hi) in enumerate(zip(methods, lows, highs)):
        ax.barh(i, hi - lo, left=lo, height=0.55, color=colors[i % len(colors)], alpha=0.82)
        ax.text(lo - 0.8, i, f"${lo:.2f}", va="center", ha="right", fontsize=9.5)
        ax.text(hi + 0.8, i, f"${hi:.2f}", va="center", ha="left", fontsize=9.5)

    ax.axvline(x=current_price, color="#C00000", linestyle="--", linewidth=1.8, label=f"Current: ${current_price:.2f}")
    ax.axvline(x=target_price, color="#000000", linestyle="-", linewidth=1.8, label=f"12M Target: ${target_price:.2f}")

    x_min = min(min(lows), current_price, target_price) * 0.88
    x_max = max(max(highs), current_price, target_price) * 1.12
    ax.set_xlim(x_min, x_max)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(methods, fontsize=10.5)
    ax.set_xlabel("Implied Price Per Share ($)", fontsize=11, fontweight="bold")
    ax.set_title("Figure 32 - Valuation Football Field", fontsize=13, fontweight="bold", pad=14)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.grid(axis="x", alpha=0.3, linestyle="--")
    ax.legend(loc="upper right", frameon=False, fontsize=9.5)
    fig.text(0.08, 0.01, f"Source: {source_text}", fontsize=8.5, style="italic", color="#595959")

    out_path = output_dir / "chart_32_valuation_football_field.png"
    plt.tight_layout(rect=(0, 0.04, 1, 1))
    plt.savefig(out_path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close()
    return out_path
```

### Step 4: Write `chart_index.txt` & Verify All Generated PNGs

Write `research/companies/<TICKER>/reports/charts/chart_index.txt` listing:
1. Every generated `chart_##_<description>.png`, its report section, and its primary data source.
2. Any **Omitted Charts** section explicitly explaining any catalog chart skipped due to non-disclosure in primary filings (so Task 5 knows not to expect or fabricate it).

Verify programmatically that every PNG opens cleanly via `PIL.Image.open(...)`, has 300 DPI print resolution (typically `> 40–50 KB`), zero truncated labels, and matches the numbers in `<TICKER>_Financial_Model_<Date>.xlsx`.
