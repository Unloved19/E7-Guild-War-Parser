```markdown
# ⚔️ Epic Seven Guild War Parser

Automates tracking guild war performance using screenshots or live screen capture.  
Supports **Google Sheets** and **Local Excel** modes.

---

> [!NOTE]
> **Prerequisites:** Python 3.11 or 3.12 must be installed from [python.org](https://www.python.org/downloads/). During installation, ensure **"Add Python to PATH"** is checked.

---

## 🚀 Setup Instructions

### Step 1: Download
Download and extract `GW_Parser_Windows.zip` from the [Releases](../../releases) page.

### Step 2: OAuth Credentials (Google Sheets Mode Only)

> [!TIP]
> Excel-only users can skip this step.

1. Download [`client_secret.json`](https://drive.google.com/file/d/14BRDjGtVm8RTsdlKj0CT-X0Mr4tQewwA/view?usp=sharing)
2. Place it in the extracted folder, next to `GW Parser.exe`

```text
📁 GW_Parser_Windows/
├── 📄 GW Parser.exe
├── 📄 client_secret.json  ← Place here
├── 📄 GW Tracker.xlsx
└── 📄 ... (other files)
```

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

```text
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
```
