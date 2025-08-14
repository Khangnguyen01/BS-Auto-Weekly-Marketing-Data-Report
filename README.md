# BlueStars – Weekly Marketing Data Report 📈📦

Automate your weekly **Amazon Ads** campaign report collection for **BlueStars** (and **Canamax**), clean the data, and email a zipped Excel bundle to your team — all in one run.

> This README documents the uploaded script: **`BlueStars - Weekly Marketing Data Report.py`**. It explains setup, environment, how it works, configuration knobs, outputs, and troubleshooting.

---

## ✨ What this script does

1. **Finds report links in Gmail**  
   Uses Gmail API to search for Amazon Ad reports with specific subjects, within the last few days.
2. **Logs into Amazon Advertising (headless)**  
   Opens a headless Chrome session, signs in with **email + password + TOTP** (2FA), then reuses cookies.
3. **Downloads reports & parses files**  
   For each email link, downloads the file via `requests` and reads CSV/XLSX into **pandas DataFrames**.
4. **Cleans & normalizes**  
   Standardizes columns for **SP/SB/SD**; fixes currency strings; derives fields (e.g., Campaign Form); converts **CA → USD** (0.76 multiplier).
5. **Enriches with SKU mapping**  
   Looks up SKUs from published Google Sheets (per brand) and extracts SKU from campaign names.
6. **Builds outputs**  
   - If **any campaigns missing SKU** (after ignoring known exceptions): emails a **missing-SKU summary** (no attachments).  
   - Otherwise: exports **six** Excel files (US/CA × SP/SB/SD), zips them, and emails the attachment.
7. **Cleans up temp files** and closes the browser.

---

## 🧩 Tech stack & dependencies

- Python 3.9+
- Gmail API (OAuth) – **read-only scope**
- Selenium (headless Chrome) + `requests` session for authenticated downloads
- Pandas, NumPy, BeautifulSoup (HTML parsing)
- `pyotp` for TOTP 2FA
- Email via Gmail SMTP (app password)

### Install
```bash
python -m venv .venv
# macOS/Linux
source .venv/bin/activate
# Windows (PowerShell)
# .venv\Scripts\Activate.ps1

pip install -U   google-api-python-client google-auth-oauthlib google-auth-httplib2   selenium webdriver-manager requests beautifulsoup4   python-dotenv pyotp pandas numpy openpyxl xlsxwriter fake-useragent
```

> **Chrome/Driver**: The script creates `webdriver.Chrome(options=...)`. Ensure a compatible **ChromeDriver** is available on PATH or adjust the code to use `webdriver_manager`'s auto-install service.

---

## 🔐 Credentials & environment

### 1) Gmail API OAuth
- Scope: `https://www.googleapis.com/auth/gmail.readonly`
- Create OAuth client credentials and download JSON (desktop app recommended).
- Set and confirm paths in the script:
  ```python
  JSON_FILE_PATH = "/path/to/gmail_api.json"
  TOKEN_FILE_PATH = "/path/to/token.pickle"
  ```
- First run opens a local browser to consent and writes `token.pickle` for reuse.

### 2) Amazon login & TOTP
The script loads a `.env` file for login secrets:
```bash
# /path/to/credentials.env
EMAIL="your_amazon_email@example.com"
PASSWORD="your_amazon_password_or_app_password"
TOTP_SECRET="BASE32_TOTP_SECRET"
```
Change the code’s `.env` path:
```python
load_dotenv("/path/to/credentials.env")
```

> **Security tip**: Avoid hardcoding secrets. Use `.env` or a secret manager. Do **not** commit `.env`/tokens.

### 3) Outgoing email
- Sender and recipients (edit in script):
  ```python
  sender_email = "sender@example.com"
  receiver_emails = ["a@example.com", "b@example.com"]
  app_password = "your_gmail_app_password"
  ```
- Uses Gmail SMTP over SSL on port 465.

---

## 🗓️ Data window & subjects

- **Report window**: previous **Sunday → Saturday**. The script computes:
  - `start_date_str = DD.MM`
  - `end_date_str   = DD.MM`
- **Gmail search window**: `after=<now-5days>` and `before=<today>` (epoch seconds).

- **Subjects searched** (maps to internal keys):
  - `Weekly BlueStars US Sponsored Products Campaign report` → `BS_US_SP_link`
  - `Weekly BlueStars US Sponsored Brands Campaign report` → `BS_US_SB_link`
  - `Weekly BlueStars US Sponsored Display Campaign report` → `BS_US_SD_link`
  - `Weekly BlueStars CA Sponsored Products Campaign report` → `BS_CA_SP_link`
  - `Weekly BlueStars CA Sponsored Brands Campaign report` → `BS_CA_SB_link`
  - `Weekly BlueStars CA Sponsored Display Campaign report` → `BS_CA_SD_link`

> The script grabs the **first link** found in the HTML part of each email.

---

## 🧭 How it works (step-by-step)

1. **Authenticate Gmail** → `authenticate_gmail()` builds a Gmail service using OAuth tokens.
2. **Fetch links** → `get_filtered_emails(service)` searches for the above subjects and extracts the **first `<a href>`** URL from each email body.
3. **Login headless** → `web_driver()` launches Chrome headless; navigates to the **first** report link; fills **email / password / TOTP**; submits.
4. **Reuse cookies** → Copies Selenium cookies into a `requests.Session()`.
5. **Download files** → For each report link:
   - `download_file_to_dataframe()` tries **CSV** first (UTF-8), then **XLSX** (`openpyxl`), with retries and 503 handling.
6. **Process DataFrames** → `process_dataframe(df, df_name)`:
   - Detects **Brand** (BlueStars/Canamax) from the link key
   - Detects **Ad Type**: SP / SB / SD
   - Detects **Market**: US / CA
   - Renames metrics per ad type (CTR, CPC, Orders, ROAS, Sales)
   - Cleans money/number columns (strip `$`, `US`, `CA`, `,`) → `float`
   - Adds derived fields: `Cost Type` (CPM if present else CPC), **Campaign Form** (Auto/Exact/Broad/Video… rules), `Campaign Type`
   - **Canada** normalization: multiply `Spend` and `Sales` by **0.76**
   - **SKU mapping**: loads a published Google Sheet (per brand), pulls `SKU`, and matches any SKU contained in the **Campaign Name**
   - Returns ordered columns:
     ```text
     Date, Campaign Type, Campaign Name, Bidding strategy,
     Impressions, Clicks, Spend, Orders, Sales,
     Product Number, Campaign Form, Market, Brand,
     Cost Type, SKU
     ```
7. **SKU checks & outputs**
   - Aggregates campaigns where `SKU` is missing (`'None'`) by **brand**, minus `ignore_cases` set.
   - If any remain → send an email: **"MISSING SKU FOR MULTIPLE BRANDS"** with details (no files).
   - Else → export **six** Excel files (US/CA × SP/SB/SD) named:
     ```text
     BlueStars US SP {DD.MM} - {DD.MM}.xlsx
     BlueStars US SB {DD.MM} - {DD.MM}.xlsx
     BlueStars US SD {DD.MM} - {DD.MM}.xlsx
     BlueStars CA SP {DD.MM} - {DD.MM}.xlsx
     BlueStars CA SB {DD.MM} - {DD.MM}.xlsx
     BlueStars CA SD {DD.MM} - {DD.MM}.xlsx
     ```
     Zip them to:  
     **`Weekly Marketing Data {DD.MM} - {DD.MM}.zip`** and email as attachment.

---

## ▶️ Run it

> The filename contains spaces — wrap it in quotes.

```bash
# macOS/Linux
python "BlueStars - Weekly Marketing Data Report.py"

# Windows (PowerShell)
python ".\BlueStars - Weekly Marketing Data Report.py"
```

The script is fully interactive (no CLI flags). Adjust constants/paths in the code before running.

---

## ⚙️ Configuration knobs (edit in code)

- **OAuth file paths** → `JSON_FILE_PATH`, `TOKEN_FILE_PATH`
- **.env path** → `load_dotenv("/path/to/credentials.env")`
- **Recipients** → `receiver_emails = [...]`
- **Sender & app password** → `sender_email`, `app_password`
- **Subjects & mapping keys** → within `get_filtered_emails()`
- **Ignore SKU cases** → `ignore_cases = {...}`
- **Canada FX factor** → `0.76` multiplier in `process_dataframe()`
- **Selenium timeouts** → waits and `headless` options in `web_driver()`
- **Date window** → computed start/end; Gmail search `report_date = now-5d`

---

## 📦 Outputs

- If SKUs missing → Plaintext summary email (no files).
- If complete → Zip attachment containing 6 Excel files (see names above).

Additionally, the script internally builds and trims a combined dataframe for final checks:
```python
df_combined = pd.concat([...])
# fill/trim/drop columns prior to export
```

---

## 🧯 Troubleshooting

- **Gmail auth fails** → Delete `token.pickle` and re-auth; confirm `JSON_FILE_PATH`
- **No messages found** → Check subject lines, Gmail date window, and sender filters (`noreply@amazon.com` or `no-reply@amazon.com`)
- **Login fails** → Verify EMAIL/PASSWORD/TOTP, and that Amazon’s element IDs (`ap_email`, `ap_password`, `auth-mfa-otpcode`) haven’t changed
- **Driver errors** → Ensure Chrome + matching ChromeDriver; consider `webdriver_manager` to auto-install
- **CSV/XLSX parse errors** → Script auto-falls back; confirm reports aren’t empty or behind additional redirects
- **Email send fails** → Use a **Gmail app password**; ensure SMTP 465 is accessible
- **Wrong currency** → Adjust the **0.76** factor for CA→USD to your current rate
- **SKU not found** → Confirm published Google Sheets are reachable and SKUs appear inside Campaign Name

---

## 🛡️ Security notes

- Prefer environment variables (.env) or a secret manager over hardcoding credentials.
- Treat OAuth JSON, tokens, and zip outputs as **sensitive**; avoid committing to version control.
- Limit OAuth scope to **readonly** for Gmail.
- Consider rotating Gmail app passwords regularly.

---

## 🗺️ Roadmap ideas

- Parameterize via CLI flags (dates, markets, brands)
- Async downloads with retries/backoff and checksum validation
- Centralized logging + run summary
- Push outputs to GDrive/Sheets/BigQuery as an alternative to email
- Add brand/product dictionaries to reduce SKU misses

---

## 🙋 FAQ

**Does it support Canamax?**  
Yes — brand inference is based on the report key (contains **BS** or **CNM**).

**What if an email has multiple links?**  
The script takes the **first** link from the HTML part.

**Can I change recipients dynamically?**  
Edit `receiver_emails` or refactor to read from `.env` or a YAML config.

**Why 0.76 for Canada?**  
A fixed CA→USD normalization in the script. Change it to your FX or remove it.
