```markdown
# Epic Seven Guild War Parser

Automates tracking guild war performance by using screenshots or live recording of the game. Supports both Google Sheets and local Excel modes.

## Prerequisites

1. **Python 3.11 or 3.12** must be installed on your computer. 
   - Download it from [python.org](https://www.python.org/downloads/)
   - **IMPORTANT:** During installation, check the box that says *"Add Python to PATH"*.

## Setup Instructions

### 1. Download the Program
Download and extract the latest `GW_Parser_Windows.zip` from the [Releases](../../releases) page.

### 2. Add OAuth Credentials (Google Sheets Mode Only)
If you plan to use Google Sheets mode, you need the credentials file:
- Download [`client_secret.json`]([https://drive.google.com/file/d/14BRDjGtVm8RTsdlKj0CT-X0Mr4tQewwA/view?usp=sharing])
- Place it directly inside the folder where you extracted the ZIP. (It should be right next to `GW Parser.exe`)

*(Note: If you only plan to use Excel mode, you can skip this step).*

### 3. Run the Program
Double-click **`GW Parser.exe`**.

**First Run Only:** 
The program will automatically set up a Python virtual environment and install all required dependencies. This takes about 2-3 minutes. 
- A console window will appear showing the installation progress.
- **Do not close this window while it is installing.**
- When it finishes, the main GUI window will open automatically.

## Usage

- **Excel Mode:** Uses `GW Tracker.xlsx` (included). Ensure the file is closed in Excel before clicking "Record War".
- **Google Sheets Mode:** Requires the `client_secret.json` from Step 2. You will be prompted to log into your Google account the first time you connect.
- **Data Source:** You can either point the parser to a folder of saved screenshots, or use "Live Screen Capture" to record your screen while you scroll through the in-game war logs.

## Troubleshooting

- **"Workbook is open in Excel"**: Save and close `GW Tracker.xlsx` in Excel, then try again.
- **"Python not found"**: Ensure Python is installed and added to your system PATH.
- **Export Fails (Excel Mode)**: Exporting the PNG summary in Excel mode requires Microsoft Excel to be installed on Windows. If you are on a Mac/Linux, use Google Sheets mode for exports.
- **Live Capture Monitor**: Live capture is locked to your primary monitor. Open the game on your main screen before starting the capture.
```

---