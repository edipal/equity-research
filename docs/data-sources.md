# Data Sources & Connectors

This harness works **out of the box with zero paid subscriptions** using public filings, Python data libraries, web research, and local/Google Drive files — and automatically upgrades if you attach institutional financial MCP servers.

## 1. Free & Public Data Workflows (Default)

### SEC EDGAR (Primary Source for US Equities)
- Company filings (`10-K`, `10-Q`, `8-K`, `DEF 14A`, `20-F`) and XBRL structured financials (`data.sec.gov/api/xbrl/companyfacts/CIK##########.json`) are freely accessible via HTTP.
- SEC requires a descriptive `User-Agent` header (e.g., `YourName your.email@example.com`). Set `SEC_EDGAR_USER_AGENT` in your local `.env` (gitignored) when scripting automated EDGAR pulls.

### Market & Pricing Data
- Use Python (`yfinance`, `pandas`, `requests`) or verified financial web sources for current share prices, historical price charts, beta inputs, and dividend history.
- Always cross-check diluted share counts, debt, and cash directly against the latest `10-Q` or `10-K` balance sheet and footnotes rather than relying solely on third-party aggregators.

### Google Drive & Google Sheets (Optional Antigravity Plugins)
- If you connect the `gdrive` and `gsheets` plugins in Antigravity, the agent can read uploaded filings/transcripts from Drive and sync models or comps tables to Google Sheets.

## 2. Optional Institutional Financial MCP Servers

If you subscribe to institutional data vendors (FactSet, S&P Global / Kensho, Daloopa, LSEG, Morningstar, Moody's, Aiera, MT Newswires, PitchBook):

1. Copy [`config/mcp_config.example.json`](../config/mcp_config.example.json) to `.agents/mcp_config.json` (or configure via Antigravity Settings → MCP Servers).
2. Remove any servers you do not have credentials/subscriptions for so Antigravity does not attempt unauthenticated connections.
3. Authenticate via OAuth or your vendor token when prompted by Antigravity.

> **Security note:** `.agents/mcp_config.json` is listed in `.gitignore` so personal endpoints, headers, or local credentials are never accidentally committed.

## 3. Local Python Environment

Install the Python dependencies used by the modeling, charting, and report-assembly skills:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

For full Excel formula recalculation (`scripts/recalc_xlsx.py`) and `.pptx` / `.docx` PDF rendering, install LibreOffice (`sudo apt-get install libreoffice` on Ubuntu/Debian or `brew install --cask libreoffice` on macOS).
