# ⚔️ Epic Seven Guild War Parser

Automates tracking guild war performance using screenshots or live screen capture.  
Supports **Google Sheets** and **Local Excel** modes.

[Download Latest Release](../../releases) · [Report Bug](../../issues)

---

> [!NOTE]
> **Prerequisites:** Python 3.11 or 3.12 must be installed from [python.org](https://www.python.org/downloads/). During installation, ensure **"Add Python to PATH"** is checked.

---

## 🚀 Setup Instructions

### Step 1: Download
Download and extract `GW_Parser_Windows.zip` from the [Releases](../../releases) page.

### Step 2: Google Cloud Setup (Google Sheets Mode Only)

> [!TIP]
> Excel-only users can skip this step. You only need to do this once. It takes about 3 minutes.

To connect to Google Sheets, you need a credentials file. This uses **your own** free Google Cloud quota (it costs $0).

1. Go to the [Google Cloud Console](https://console.cloud.google.com/) and log in.
2. Click the **Select a project** dropdown (top left) → Click **New Project** → Name it `GW Parser` → Click **Create**.
3. In the top search bar, search for **Google Sheets API** → Click it → Click **Enable**.
4. Search again for **Google Drive API** → Click it → Click **Enable**.
5. Go to **APIs & Services** → **Credentials** (left sidebar).
6. Click **+ CREATE CREDENTIALS** (top) → Select **OAuth client ID**.
7. If asked to configure the consent screen, select **External** → Click **Create** → Fill in an App name (e.g., "My GW Parser") → Click **Save and Continue** through the rest.
8. Back on the Credentials page, set the Application type dropdown to **Desktop app** → Name it `GW Parser` → Click **Create**.
9. A popup will appear. Click the **Download JSON** button.
10. Rename the downloaded file to exactly **`client_secret.json`** and place it in your GW Parser folder next to `GW Parser.exe`.

~~~text
📁 GW_Parser_Windows/
├── 📄 GW Parser.exe
├── 📄 client_secret.json  ← Place here
├── 📄 GW Tracker.xlsx
└── 📄 ... (other files)
~~~

### Step 3: Run
Double-click **`GW Parser.exe`**

> [!WARNING]
> **First Run Only:** A console window will appear to set up the environment (~2-3 min). **Do not close this window.** The GUI will open automatically when complete.

---

## 🎮 Usage

| Mode | Requirements | Notes |
|------|--------------|-------|
| **Excel** | `GW Tracker.xlsx` (included) | Close the file in Excel before parsing |
| **Google Sheets** | `client_secret.json` | Browser auth prompt appears on first connect |

**Data Sources:**
- **Screenshot Folder** — Point to a folder of saved war log screenshots
- **Live Capture** — Record your screen while scrolling through in-game logs

> [!NOTE]
> Live Capture is locked to your **primary monitor**.

---

## ❓ Troubleshooting

> [!NOTE]
> **"Workbook is open in Excel"**
> Save and close `GW Tracker.xlsx`, then try again.
> 
> **"Python not found"**
> Ensure Python 3.11+ is installed and added to your system PATH.
> 
> **Export Fails (Excel Mode)**
> PNG export requires Microsoft Excel installed on Windows. Use Google Sheets mode for exports on Mac/Linux.
> 
> **"Token expired or missing"**
> Re-run authentication: Configure → Re-run full setup wizard.
> 
> **OCR models downloading slowly**
> EasyOCR downloads ~100MB of models on first use. Ensure a stable internet connection during the first run.

---

## 📁 File Structure

~~~text
📁 Extracted Folder/
├── GW Parser.exe          # Main launcher
├── GW Tracker.xlsx        # Excel mode template
├── gw_parser_gui.py       # GUI application
├── launcher.py            # EXE entry point
├── setup_venv.py          # Environment setup
├── parser_core.py         # OCR & matching logic
├── ... (other scripts)
│
├── .venv/                 # Created on first run (auto-generated)
├── War Screenshots/       # Screenshot storage (auto-generated)
├── War Records/           # Exported PNG summaries (auto-generated)
└── client_secret.json     # User must add this (Google Sheets only)
~~~
