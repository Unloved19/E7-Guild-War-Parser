"""First-run setup wizard for the GW Parser.

Can be run standalone:
    python first_run_setup.py

Or imported by the main parser to auto-trigger on missing config.
"""

import re
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
TOKEN_PATH = BASE_DIR / "token.json"
SECRET_PATH = BASE_DIR / "client_secret.json"

# Replace this with your actual "Make a Copy" link
TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/1gCnRQ0-K0Lvn5KjXmjpgWQZvp59P5oabsR0n1UEm8gk/edit?usp=sharing"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

RE_SHEET_ID = re.compile(r"/d/([a-zA-Z0-9_-]+)")


def ask_yes_no(prompt: str, default: bool = False) -> bool:
    suffix = " [Y/n]: " if default else " [y/N]: "
    while True:
        answer = input(prompt + suffix).strip().lower()
        if answer == "":
            return default
        if answer in {"y", "yes"}:
            return True
        if answer in {"n", "no"}:
            return False
        print("Please answer y or n.")


def extract_spreadsheet_id(url: str) -> str:
    match = RE_SHEET_ID.search(url)
    if not match:
        raise ValueError("Could not find spreadsheet ID in that URL.")
    return match.group(1)


def run_auth() -> bool:
    """Runs the OAuth flow. Returns True on success."""
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
    except ImportError:
        print("Missing auth dependencies. Install with:")
        print("  python -m pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client")
        return False

    if not SECRET_PATH.exists():
        print(f"\nMissing {SECRET_PATH.name}.")
        print("This file should have been included with the program.")
        print("If it's missing, contact the person who gave you this tool.")
        return False

    creds = None
    if TOKEN_PATH.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        except ValueError:
            TOKEN_PATH.unlink()
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired token...", end=" ", flush=True)
            creds.refresh(Request())
            print("Done.")
        else:
            print("\nA browser window will open. Log into the Google account")
            print("that owns your Guild War spreadsheet and click 'Allow'.")
            input("Press ENTER to open the browser...")
            flow = InstalledAppFlow.from_client_secrets_file(str(SECRET_PATH), SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, "w") as f:
            json.dump(json.loads(creds.to_json()), f, indent=2)

    print("Authentication successful.")
    return True


def collect_config(existing: dict = None) -> dict:
    """Walks the user through setup. Returns a config dict."""
    print("=" * 50)
    print("GW Parser - Setup Wizard")
    print("=" * 50)

    # --- Spreadsheet URL ---
    if existing and existing.get("spreadsheet_id"):
        print(f"\nCurrent spreadsheet ID: {existing['spreadsheet_id']}")
        if not ask_yes_no("Change to a different spreadsheet?", default=False):
            spreadsheet_id = existing["spreadsheet_id"]
        else:
            spreadsheet_id = None
    else:
        spreadsheet_id = None

    if not spreadsheet_id:
        print("\nStep 1: Get your Google Sheet")
        print(f"  1. Open this link in your browser: {TEMPLATE_URL}")
        print("  2. Click 'Make a Copy' to create your own copy")
        print("  3. Copy the URL of your new spreadsheet")
        print()
        while True:
            url = input("  Paste your Google Sheet URL here: ").strip()
            try:
                spreadsheet_id = extract_spreadsheet_id(url)
                print(f"  Extracted ID: {spreadsheet_id}")
                break
            except ValueError:
                print("  Invalid URL. Make sure it's a Google Sheets URL.")

    # --- Sheet name ---
    default_sheet = existing.get("source_sheet", "GW Tracker") if existing else "GW Tracker"
    print(f"\nStep 2: Sheet name (tab name in your spreadsheet)")
    sheet_name = input(f"  Sheet name [default: {default_sheet}]: ").strip()
    if not sheet_name:
        sheet_name = default_sheet

    # --- Verify connection ---
    print(f"\nStep 3: Verifying access to '{sheet_name}'...")
    try:
        import gspread
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        
        
        
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        except ValueError:
            creds = None

        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                raise SystemExit("Auth token invalid. Something went wrong.")

        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(spreadsheet_id)

        sheet_names = [ws.title for ws in spreadsheet.worksheets()]
        if sheet_name not in sheet_names:
            print(f"  Available sheets: {', '.join(sheet_names)}")
            print(f"  '{sheet_name}' not found!")
            if not ask_yes_no("Continue anyway?", default=False):
                raise SystemExit("Setup cancelled.")

        print("  Connected successfully.")
    except Exception as exc:
        print(f"  Warning: Could not verify: {exc}")
        if not ask_yes_no("Continue anyway?", default=False):
            raise SystemExit("Setup cancelled.")

    config = {
        "spreadsheet_id": spreadsheet_id,
        "source_sheet": sheet_name,
        "target_sheet": sheet_name,
        "override_col": "C",
        "date_col": 15,
    }

    return config


def print_post_setup_instructions(config: dict) -> None:
    print("\n" + "=" * 50)
    print("Setup complete! Before running the parser:")
    print("=" * 50)
    print(f"  1. Open your Google Sheet")
    print(f"  2. Go to the '{config['target_sheet']}' tab")
    print(f"  3. Add guild member names in Column B (starting at row 5)")
    print(f"  4. Enter the first war date of the season in Column O (row 5)")
    print(f"     Format: DD/MM/YYYY (e.g. 23/03/2026)")
    print(f"  5. Click that cell and drag down to fill all war dates")
    print(f"  6. Run the parser again to record your first war")
    print("=" * 50)


def main():
    if CONFIG_PATH.exists():
        existing = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        print("Found existing configuration.")
        if not ask_yes_no("Reconfigure?", default=False):
            return
    else:
        existing = None

    if not TOKEN_PATH.exists():
        print("No authentication found. Let's set that up first.\n")
        if not run_auth():
            raise SystemExit(1)
        print()

    config = collect_config(existing)
    CONFIG_PATH.write_text(json.dumps(config, indent=2), encoding="utf-8")
    print(f"\nConfig saved to: {CONFIG_PATH.name}")

    print_post_setup_instructions(config)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.")
    except SystemExit:
        raise
    except Exception as exc:
        print(f"\nSetup error: {exc}")
        raise SystemExit(1)