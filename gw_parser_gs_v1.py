from __future__ import annotations
import sys
import subprocess
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import time
import threading
import tempfile
import shutil
try:
    import mss
    import numpy as np
except ImportError as exc:
    raise SystemExit("Missing dependency: mss or numpy. Install with: python -m pip install mss numpy") from exc


import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from gw_parser_core import *

# ============================================================================
# PATHS & FILE LOCATIONS
# ============================================================================

TOKEN_PATH = BASE_DIR / "token.json"
SECRET_PATH = BASE_DIR / "client_secret.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ============================================================================
# ALIAS & FILE UTILITIES
# ============================================================================


def get_gspread_client() -> gspread.Client:
    if not SECRET_PATH.exists():
        raise SystemExit(f"Missing {SECRET_PATH.name}. Run setup_auth.py first.")
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise SystemExit("Token expired or missing. Run setup_auth.py again.")
    return gspread.authorize(creds)


def _col_num_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def fetch_sheet_data(ws) -> List[List]:
    """Fetches the entire sheet in a single API call. Returns a 2D list of strings."""
    return ws.get_all_values()


# ============================================================================
# WORKSHEET WRITING OPERATIONS
# ============================================================================


def clear_target_row_values(
    ws, target_row: int, section_maps: Dict[str, Dict[str, int]]
) -> None:
    updates = []
    for section_name in WRITE_SECTIONS:
        for col in section_maps[section_name].values():
            updates.append({
                "range": f"{_col_num_to_letter(col)}{target_row}",
                "values": [[""]]
            })
    if updates:
        ws.batch_update(updates)


def write_to_rough_sheet(
    ws, target_row: int, normalized: List[Dict[str, object]],
    section_maps: Dict[str, Dict[str, int]],
) -> None:
    updates = []
    for row in normalized:
        name = row["Name"]
        mappings = {
            "Roster Active Status": row["Active flag"],
            "Tokens Tracking": row["Tokens Used"],
            "Offense Wins": row["Offense W"],
            "Offense Draws": row["Offense D"],
            "Offense Losses": row["Offense L"],
            "Defense Wins": row["Defense W"],
            "Defense Draws": row["Defense D"],
            "Defense Losses": row["Defense L"],
        }
        for section_name, value in mappings.items():
            col_map = section_maps[section_name]
            if name not in col_map:
                continue
            updates.append({
                "range": f"{_col_num_to_letter(col_map[name])}{target_row}",
                "values": [[value]]
            })
    if updates:
        ws.batch_update(updates)


# ============================================================================
# MAIN EXECUTION FLOW
# ============================================================================

def main():
    # --- DEBUG CLEANUP PROMPT ---
    if DEBUG_ROOT.exists():
        if ask_yes_no("Debug images from previous runs found. Delete them before starting?", default=True):
            cleanup_debug_images()

    # --- FIRST-RUN / CONFIGURE CHECK ---
    if not CONFIG_PATH.exists():
        print("No configuration found. Running setup wizard...\n")
        subprocess.run([sys.executable, str(BASE_DIR / "first_run_setup.py")], check=True)
        return

    print("Guild War Screenshot Parser -> Workbook Writer")
    run_mode = ask_run_mode()
    
    screenshots_folder, folder_name, video_temp_dir = prompt_and_get_screenshots()

    if not screenshots_folder.exists():
        raise FileNotFoundError(f"Could not find source: {screenshots_folder}")

    aliases_existing = load_aliases(ALIASES_PATH)
    aliases_working = dict(aliases_existing)
    aliases_created_this_run: Dict[str, str] = {}
    config = load_config()
    client = get_gspread_client()
    spreadsheet = client.open_by_key(config["google_sheet"]["spreadsheet_id"])

    source_ws = spreadsheet.worksheet(config["google_sheet"]["source_sheet"])
    target_ws = spreadsheet.worksheet(config["google_sheet"]["target_sheet"])

    print("Fetching sheet data...", end=" ", flush=True)
    source_data = fetch_sheet_data(source_ws)
    target_data = fetch_sheet_data(target_ws)
    print("Done.")

    roster = read_roster(source_data)
    target_row = find_date_row(target_data, folder_name, config["date_col"])
    target_section_maps = build_target_section_maps(target_data)
    overrides = read_column_a_overrides(target_data, roster, config["override_col"])

    try:
        extracted, _contexts = extract_rows_from_folder(
            screenshots_folder, folder_name, roster,
            aliases_existing, aliases_working, aliases_created_this_run, overrides
        )
        normalized = build_normalized_table(roster, extracted)

        prev_active_status = get_previous_war_active_status(target_data, target_row, target_section_maps, config["date_col"])
        resolve_active_inactive_status(normalized, extracted, prev_active_status)

        run_mode = handle_post_processing(run_mode, extracted, normalized, aliases_created_this_run)

        if run_mode == "write":
            if row_has_existing_values(target_data, target_row, target_section_maps):
                if not ask_yes_no("Target row already has values in one or more write sections. Overwrite whole row?", default=False):
                    raise RuntimeError("Run cancelled by user at overwrite prompt.")
            
            clear_target_row_values(target_ws, target_row, target_section_maps)
            write_to_rough_sheet(target_ws, target_row, normalized, target_section_maps)
            
            if aliases_created_this_run:
                save_aliases(ALIASES_PATH, aliases_working)

            print("\nDone.")
            print(f"Folder processed: {folder_name}")
            print(f"Data written to: {config['google_sheet']['target_sheet']} (date row matched from folder name)")
            print(f"New aliases saved: {len(aliases_created_this_run)}")

            try:
                subprocess.run([sys.executable, str(BASE_DIR / "war_record_exporter.py"), folder_name], check=True)
            except subprocess.CalledProcessError:
                print("Warning: War record export failed.")

    finally:
        if video_temp_dir and video_temp_dir.exists():
            shutil.rmtree(video_temp_dir)
            print("Cleaned up temporary capture files.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(1)
    except Exception as exc:
        print(f"\nError: {exc}")
        raise SystemExit(1)