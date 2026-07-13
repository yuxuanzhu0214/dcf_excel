# Unlevered DCF Excel Model Generator

An industrial-grade, fully dynamic, and linkage-complete Unlevered Discounted Cash Flow (DCF) valuation model generator for US equities. It automates financial data extraction, consensus forecasting, model tapering, WACC estimation, sensitivity grids, and checks into a publication-ready Excel spreadsheet.

## 🚀 Quick Start

Run the generator for any stock ticker with a single command:

```bash
./run_dcf.sh <TICKER>
```

*Example:*
```bash
./run_dcf.sh GOOG
```

If you run the script without any arguments, it will interactively prompt you for the ticker:
```bash
./run_dcf.sh
# Enter stock ticker (e.g. AAPL, GOOG, NVDA): RDDT
```

The resulting file will be saved in your current working directory as `<TICKER>_DCF_Model.xlsx`.

### 🌍 Global Terminal Command Setup (Zsh/Bash)

To run this tool from **any folder** in your terminal using a quick command (like `dcf AAPL`), add the following alias to your shell configuration file:

1. **Open your `.zshrc` (Mac default):**
   ```bash
   nano ~/.zshrc
   ```

2. **Add this line at the bottom:**
   ```bash
   alias dcf="/Users/yuxuanzhu/dev/dcf_excel/run_dcf.sh"
   ```

3. **Save and reload your shell:**
   ```bash
   source ~/.zshrc
   ```

4. **Use it anywhere:**
   ```bash
   dcf TSLA
   ```

---

## 🛠️ Installation & Setup

If you need to install the dependencies manually:

1. **Set up virtual environment:**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

2. **Install required packages:**
   ```bash
   pip install openbb openpyxl yfinance pandas google-api-python-client google-auth-httplib2 google-auth-oauthlib
   ```

### ☁️ Optional: Auto-Upload to Google Drive / Google Sheets

If configured, the generator will automatically upload a copy of the Excel sheet and convert it to a native **Google Sheet** (preserving all formulas), printing the share link in the console.

#### 1. Setup Config Folder & Credentials
Create the config directory outside this repository to keep your keys private:
```bash
mkdir -p ~/.config/dcf_excel
```
Save your Google Service Account key file as `google_credentials.json` inside that folder:
`/Users/yuxuanzhu/.config/dcf_excel/google_credentials.json`

#### 2. Configure target folder (Highly Recommended)
Service Accounts operate in their own sandbox. To see the uploaded sheets in your personal Google Drive, you **must** share a Google Drive folder with the Service Account email (found under `"client_email"` in your JSON key) and configure the folder ID in a `config.json` file located at `/Users/yuxuanzhu/.config/dcf_excel/config.json`:

```json
{
  "drive_folder_id": "your_google_drive_folder_id_here"
}
```

---

## 📊 Model Highlights & Methodology

1. **Dual-Class Share Dilution (Gold Standard):** Total shares outstanding are calculated using `Market Cap / Price` to ensure dual-class structures (such as `GOOG`/`GOOGL`, `BRK.A`/`BRK.B`) are fully captured and not undercounted.
2. **Consensus-Driven Convex Decaying Tapering:** Years 1-3 anchor directly to consensus estimates (obtained dynamically from yfinance & implied EPS growth). Years 4-10 apply a gentle convex (power 1.5) decay down to a 3.0% long-term perpetuity rate, avoiding abrupt drops.
3. **Mid-Year Discounting Convention:** Discounting periods set to 0.5, 1.5, ..., 9.5 for explicit cash flows, and 10.0 for terminal value.
4. **Interactive WACC & Sensitivities:** Every value is formula-linked. Change assumptions on the `Assumptions` sheet (in blue) and see the entire valuation bridge and WACC recalculate in real-time.
5. **No Grid Lines & Bloomberg Aesthetic:** Built using clean borders, dark navy headers, blue input formatting, and hidden gridlines on every sheet for a premium presentation.

---

## 📂 Sheet Architecture

*   **Summary:** Executive valuation card showing implied vs. current share price, upside/downside, and enterprise value bridge.
*   **Assumptions:** All active drivers (WACC, terminal growth, capex/revenue, net debt, and share counts).
*   **Historical:** Last 4 years of actual reported figures.
*   **DCF:** Full 10-year explicit projection sheet with NOPAT, UFCF, discount factors, PV of cash flows, and Gordon Growth terminal values.
*   **Sensitivity:** 2D data table showing implied share price at various WACC and terminal growth rate intervals.
*   **Checks:** Verification checks to guarantee model integrity.
*   **Sources & Audit:** Explanations of data sources, definitions, and formulas.
