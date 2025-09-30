# Invoice Sender

A small local tool to email monthly invoice PDFs to customers using a single Excel workbook as the source of truth.

## What it will do (scope)

- Read an Excel file that contains:
  - Column A: Company name
  - Column B: 5-digit account number(s) - supports multiple accounts per row
  - Column G: Recipient email address(es)
- Find PDF files in an invoices folder whose filename starts with the 5‑digit account number followed by an underscore (e.g., `12345_*.pdf`).
- Email each PDF to the addresses from the matching row.
- Support a dry-run mode (log only, no emails).
- Handle multiple account numbers per company (separated by commas, semicolons, slashes, or spaces).

## Assumptions (simple defaults)

- Account numbers are exactly 5 digits.
- PDF files live in a single folder for the month.
- Multiple emails per row are separated by comma or semicolon.
- Multiple account numbers per row are separated by comma, semicolon, slash, or space.
- Each account number gets its own email with the corresponding PDF attachment.

## Quickstart (Idiot-Proof Instructions)

### Step 1: Open Terminal/Command Prompt
- **Mac**: Press `Cmd + Space`, type "Terminal", press Enter
- **Windows**: Press `Win + R`, type "cmd", press Enter

### Step 2: Navigate to Project
```bash
cd /path/to/invoice_sender
```
*(Replace with your actual folder path)*

### Step 3: Create Virtual Environment
```bash
python -m venv .venv
```

### Step 4: Activate Virtual Environment
- **Mac/Linux**: `source .venv/bin/activate`
- **Windows**: `.venv\Scripts\activate`

### Step 5: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 6: Run the App

**Mac/Linux:**
```bash
streamlit run app.py --server.headless true
```

**Windows:**
```cmd
streamlit run app.py --server.headless true
```

### Step 7: Open Browser
- Go to: `http://localhost:8501`
- Fill in your settings and start sending!

**That's it! 5 commands total.**

## Multiple Account Numbers Per Company

**The system supports companies with multiple account numbers:**

### Excel Format Examples:
- `12345, 67890` - Two accounts separated by comma
- `12345; 67890` - Two accounts separated by semicolon  
- `12345 / 67890` - Two accounts separated by slash
- `12345 67890` - Two accounts separated by space
- `12345, 67890, 11111` - Three accounts

### What Happens:
1. **System finds** `12345_invoice.pdf` → sends to company email
2. **System finds** `67890_invoice.pdf` → sends to company email  
3. **Two separate emails** with different attachments
4. **Same recipients** get all invoices for that company

### PDF File Naming:
- `12345_invoice.pdf` ✅
- `67890_invoice.pdf` ✅
- `12345_january_invoice.pdf` ✅
- `67890_january_invoice.pdf` ✅

## Excel Sheet Support

**The system supports both single-sheet and multi-sheet Excel files:**

### Single Sheet Excel (Testing):
- **Sheet name field**: Leave blank
- **System reads**: First/default sheet
- **Example**: `data/accounts.xlsx` (no sheet name needed)

### Multi-Sheet Excel (Production):
- **Sheet name field**: Enter specific sheet name
- **System reads**: Only the specified sheet
- **Example**: Enter "Combined" to read the "Combined" tab

### Benefits:
- ✅ **Works with existing** single-sheet files
- ✅ **Supports production** multi-tab workbooks
- ✅ **No data copying** required
- ✅ **Settings saved** for next time

## Platform-Specific Instructions

### Mac (Terminal):
```bash
cd /path/to/invoice_sender
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py --server.headless true
```

### Windows (Command Prompt):
```cmd
cd C:\path\to\invoice_sender
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py --server.headless true
```

### Windows (PowerShell):
```powershell
cd C:\path\to\invoice_sender
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py --server.headless true
```

### Linux/WSL:
```bash
cd /path/to/invoice_sender
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py --server.headless true
```

## How to Use the App

1) **Start the app**: `source .venv/bin/activate && streamlit run app.py --server.headless true`
2) **Open browser**: Go to `http://localhost:8501`
3) **Configure paths**:
   - Excel file: `data/accounts.xlsx` (or your file)
   - Excel sheet name: Leave blank for first sheet, or specify like "Combined"
   - Invoices folder: `invoices` (or your folder)
4) **Set up email**:
   - From address: Your email
   - Subject: "Your Invoice" (or customize)
   - Body: "Here is the invoice for account %ACCOUNT%.\n\nThank you."
5) **Configure SMTP**:
   - Host: `smtp.office365.com` (Office 365) or `smtp.gmail.com` (Gmail)
   - Port: `587`
   - User: Your email address
   - Password: Your password or app password
   - Use TLS: ✅ Checked
6) **Rate limiting** (for large batches):
   - Delay between emails: `2.1` seconds (recommended for Office 365)
   - Max retries: `3` attempts
7) **Test and send**:
   - Click "Test SMTP" first
   - Click "Dry Run" to preview
   - Click "Send" to actually send emails

## Restarting the App
- **Stop**: Press `Ctrl+C` in the terminal
- **If Ctrl+C doesn't work**: `pkill -f "streamlit run app.py"`
- **Start**: `source .venv/bin/activate && streamlit run app.py --server.headless true`
- **Or use one-click scripts**: `bash run.sh` (macOS/Linux) or `run.ps1` (Windows)

## Optional CLI usage
- Dry run: `python send_invoices.py --excel data/accounts.xlsx --invoices invoices --dry-run --verbose`
- Send: `python send_invoices.py --excel data/accounts.xlsx --invoices invoices`

## Deploy on a new machine (minimal steps)
1) Clone this repo (GitHub) or copy the folder.
2) Ensure Python 3.10+ is installed.
3) One-click run:
   - Windows: double-click `run.ps1` (or run in PowerShell)
   - Linux/WSL/macOS: `bash run.sh`
4) On first run, it will create a venv, install deps, copy `env.example` to `.env` if missing, then launch the UI.
5) Edit `.env` with SMTP details, set paths in the UI, Test SMTP, Dry Run, then Send.

## Troubleshooting
- If `source .venv/bin/activate` says "No such file or directory", create the venv first: `python3 -m venv .venv` (Linux/WSL/macOS) or `py -3 -m venv .venv` (Windows).
- If `streamlit: command not found`, activate the venv or run `python -m streamlit run app.py`.
- On Windows, if PowerShell blocks `run.ps1`, run: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` then try again.

## How to run the one-click scripts
- Windows PowerShell (from the project folder):
  - `.\run.ps1` (note the `./` prefix)
- WSL/Linux/macOS (from the project folder):
  - `bash run.sh`

## Saving settings from the app
- In the UI, open "Settings persistence" → set `.env path` (defaults to `.env`) → click "Save settings (.env)".
- Next launch, the app will load these values automatically.

## SMTP providers and requirements
- Standard SMTP username/password is supported.
- Office 365 (Exchange Online):
  - Recommended: Use an app password or modern auth SMTP (tenant policy must allow SMTP AUTH for the mailbox or use a dedicated connector). Many orgs disable basic auth; if disabled, ask IT to enable SMTP AUTH for your sender mailbox or provide a relay.
  - Host: `smtp.office365.com`, Port: `587`, TLS: `true`.
- Gmail:
  - Use an App Password (with 2FA enabled). Host: `smtp.gmail.com`, Port: `587`, TLS: `true`.
- Custom domain provider:
  - Use their SMTP host/port with a mailbox that's allowed to send. Some providers require an app password or SMTP relay.

If your Office 365 account has SMTP AUTH enabled, it will work. Otherwise, use a mailbox that permits SMTP (e.g., a domain email you control) or ask IT to enable SMTP AUTH or provide an SMTP relay.