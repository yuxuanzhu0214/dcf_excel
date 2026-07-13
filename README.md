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

You can configure this using **User OAuth 2.0** (recommended for personal accounts to bypass Service Account quota limits) or a **Service Account**.

---

### Method A: User OAuth 2.0 (Recommended)

This configures the generator to log in using your personal Google account. Uploaded spreadsheets will be directly owned by you in your Google Drive.

#### 1. Generate Client Credentials in Google Cloud Console
1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Select or create a project.
3. Enable both the **Google Drive API** and **Google Sheets API**.
4. Go to the **Google Auth Platform** (or *OAuth consent screen*).
5. Set up the OAuth consent screen with **External** user type.
6. **Crucial:** In the left menu, select **Audience** (or *Test Users*), click **+ Add Users**, and add your own Google email address (the one you will log in with).
7. Go to **Clients** (or *Credentials*), click **+ Create Credentials > OAuth client ID**, and choose **Desktop App**.
8. Download the generated client secret JSON, rename it to **`client_secret.json`**, and save it to the local config folder:
   `/Users/yuxuanzhu/.config/dcf_excel/client_secret.json`

#### 2. Run and Authorize
The first time you run `dcf <TICKER>`, the script will prompt you with a link or open your browser to log in. Click **Allow** (under *Advanced > Go to App (unsafe)*) to authorize the application. A login token will be securely saved locally to:
`/Users/yuxuanzhu/.config/dcf_excel/token.pickle`

---

### Method B: Service Account (Alternative)

#### 1. Setup Service Account Credentials
1. Under **Credentials**, click **+ Create Credentials > Service Account**.
2. Go to the **Keys** tab of the service account, click **Add Key > Create New Key (JSON)**, and download it.
3. Rename the file to **`google_credentials.json`** and save it to:
   `/Users/yuxuanzhu/.config/dcf_excel/google_credentials.json`

#### 2. Share Google Drive Folder
Because Service Accounts have a separate 0-byte sandbox, you **must** share a Google Drive folder with the Service Account’s email address as an **Editor** to allow uploads.

---

### Global Config File (`config.json`)

To specify a target Google Drive folder destination for uploads (under either method), create a configuration file at `/Users/yuxuanzhu/.config/dcf_excel/config.json`:

```json
{
  "drive_folder_id": "your_google_drive_folder_id_here"
}
```

*Note: You can find the Folder ID in the Google Drive URL: `drive.google.com/drive/folders/YOUR_FOLDER_ID`.*

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
