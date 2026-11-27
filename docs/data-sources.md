# Data Sources & Connectors

This harness works **out of the box with zero paid subscriptions** using public filings, Python data libraries, web research, and local/Google Drive files — and automatically upgrades if you attach institutional financial MCP servers.

## 1. Free & Public Data Workflows (Default)

### SEC EDGAR (Primary Source for US Equities)
- Company filings (`10-K`, `10-Q`, `8-K`, `DEF 14A`, `20-F`) and XBRL structured financials (`https://data.sec.gov/api/xbrl/companyfacts/CIK##########.json`) are freely accessible via HTTP.
- SEC requires a descriptive `User-Agent` header (e.g., `YourName your.email@example.com`). Set `SEC_EDGAR_USER_AGENT` in your local `.env` (gitignored) when scripting automated EDGAR pulls.

### Market & Pricing Data
- Use Python (`yfinance`, `pandas`, `requests`) or verified financial web sources for current share prices, historical price charts, beta inputs, and dividend history.
- Always cross-check diluted share counts, debt, and cash directly against the latest `10-Q` or `10-K` balance sheet and footnotes rather than relying solely on third-party aggregators.
- For stock screens (`idea-generation`), combine SEC EDGAR XBRL (`companyfacts`/`frames`), `yfinance` (LTM multiples default), SEC Form 4 (insider transactions), and FINRA bi-monthly short interest.

### Google Drive & Google Sheets (Optional Antigravity Plugins)
- If you connect the `gdrive` and `gsheets` plugins in Antigravity, the agent can read uploaded filings/transcripts from Drive and sync models or comps tables to Google Sheets.

## 2. Optional Institutional Financial MCP Servers

If you subscribe to institutional data vendors (FactSet, S&P Global / Kensho, Daloopa, LSEG, Morningstar, Moody's, Aiera, MT Newswires, PitchBook):

> **Illustrative templates only:** All vendor URLs in [`config/mcp_config.example.json`](../config/mcp_config.example.json) are **illustrative templates requiring active institutional subscriptions** and vendor-confirmed endpoints/credentials.

1. Copy [`config/mcp_config.example.json`](../config/mcp_config.example.json) to your workspace `.agents/mcp_config.json` (project-scoped, **gitignored**) or global `~/.gemini/config/mcp_config.json` (user-scoped across workspaces), or configure via Antigravity Settings → MCP Servers.
2. Remove any servers you do not have active subscriptions/credentials for so Antigravity does not attempt unauthenticated connections.
3. Each remote HTTP MCP entry uses `"serverUrl"` and supports optional `"headers"` (for API-key / bearer auth) and `"timeoutSeconds"` fields. Replace `<YOUR_VENDOR_TOKEN>` with your literal issued token directly in your local gitignored `.agents/mcp_config.json` (do **not** rely on undocumented `${ENV_VAR}` shell interpolation inside JSON):
   ```json
   {
     "mcpServers": {
       "factset": {
         "serverUrl": "https://mcp.factset.com/mcp",
         "headers": {
           "Authorization": "Bearer <YOUR_VENDOR_TOKEN>"
         },
         "timeoutSeconds": 60
       }
     }
   }
   ```
4. Authenticate via OAuth or your vendor token when prompted by Antigravity.

> **Security note:** `.agents/mcp_config.json` is listed in `.gitignore` so personal endpoints, headers, and literal vendor tokens are never accidentally committed. Never put real credentials inside `config/mcp_config.example.json` or any tracked file.

## 3. Local Python Environment

Install the Python dependencies used by the modeling, charting, and report-assembly skills:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

For full Excel formula recalculation (`scripts/recalc_xlsx.py`) and `.pptx` / `.docx` PDF rendering, install LibreOffice (`sudo apt-get install libreoffice` on Ubuntu/Debian or `brew install --cask libreoffice` on macOS).
