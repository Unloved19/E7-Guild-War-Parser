from __future__ import annotations
import sys
import subprocess
from datetime import date, datetime
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
from parser_core import *

# ============================================================================
# PATHS & FILE LOCATIONS
# ============================================================================

CONFIG_PATH = BASE_DIR / "config.json"
TOKEN_PATH = BASE_DIR / "token.json"
SECRET_PATH = BASE_DIR / "client_secret.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ============================================================================
# ALIAS & FILE UTILITIES
# ============================================================================

def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Missing {CONFIG_PATH.name}. Create it with your spreadsheet ID, sheet names, and override column."
        )
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


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


def col_letter_to_num(letter: str) -> int:
    num = 0
    for ch in letter.upper():
        num = num * 26 + (ord(ch) - ord("A") + 1)
    return num

def _col_num_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def fetch_sheet_data(ws) -> List[List]:
    """Fetches the entire sheet in a single API call. Returns a 2D list of strings."""
    return ws.get_all_values()


def cell_val(data: List[List], row: int, col: int) -> Optional[str]:
    """Get cell value from fetched data. row and col are 1-indexed."""
    if row < 1 or row > len(data):
        return None
    row_data = data[row - 1]
    if col < 1 or col > len(row_data):
        return None
    val = row_data[col - 1]
    return val if val != "" else None

# ============================================================================
# TEXT PROCESSING & NORMALIZATION
# ============================================================================

def sheet_date_to_folder_date_str(value) -> Optional[str]:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date().strftime("%Y%m%d")
    if isinstance(value, date):
        return value.strftime("%Y%m%d")
    if isinstance(value, str):
        value = value.strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y", "%Y/%m/%d", "%Y%m%d"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y%m%d")
            except ValueError:
                continue
    return None


# ============================================================================
# USER INTERACTION & CLI PROMPTS
# ============================================================================

def ask_folder_name() -> str:
    folder_name = input("Enter War Screenshots folder name (e.g. 20260410): ").strip()
    if not RE_FOLDER_NAME.fullmatch(folder_name):
        raise ValueError("Folder name must be 8 digits in YYYYMMDD format.")
    return folder_name


def ask_run_mode() -> str:
    print("\nRun Mode")
    print("1. Dry run")
    print("2. Write to workbook")
    print("3. Cancel")
    while True:
        choice = input("Choose 1/2/3: ").strip()
        if choice == "1":
            return "dry"
        if choice == "2":
            return "write"
        if choice == "3":
            raise KeyboardInterrupt
        print("Invalid choice.")



# ============================================================================
# WORKSHEET READING OPERATIONS
# ============================================================================

def read_roster(data) -> List[str]:
    roster: List[str] = []
    blank_run = 0
    row = 5
    while row <= len(data) + 10:
        value = cell_val(data, row, 2)
        if value is None:
            blank_run += 1
            if blank_run >= 10:
                break
        else:
            blank_run = 0
            name = str(value).strip()
            low = name.lower()
            if low.startswith("notes:") or low.startswith("note:"):
                row += 1
                continue
            roster.append(name)
        row += 1

    if not roster:
        raise RuntimeError("No roster found starting from B5.")
    return roster

    
def read_column_a_overrides(data, roster: List[str], override_col: str) -> Dict[str, str]:
    """Reads the override column to build match redirects for returning players.
    
    When the override column contains a bare player name (e.g. "Kyūbi焔"), it means:
    "OCR text matching this name should map to the roster name in column B"
    (e.g. "Kyūbi焔 (2) (Rejoined 02/5/26)"), overriding the default matcher.
    """
    overrides: Dict[str, str] = {}
    row = 5
    roster_idx = 0
    override_col_num = col_letter_to_num(override_col)

    while row <= len(data) + 10 and roster_idx < len(roster):
        name = cell_val(data, row, 2)
        override_val = cell_val(data, row, override_col_num)

        if name and str(name).strip() in roster:
            override = str(override_val).strip() if override_val else ""
            if override:
                overrides[canonical_key(override)] = str(name).strip()
            roster_idx += 1

        row += 1
    return overrides

def find_date_row(data, folder_name: str, date_col: int = 15) -> int:
    for row in range(5, len(data) + 1):
        val = cell_val(data, row, date_col)
        if sheet_date_to_folder_date_str(val) == folder_name:
            return row
    raise RuntimeError(f"Could not find date {folder_name} in column {date_col}.")


def find_section_starts(data) -> Dict[str, int]:
    starts: Dict[str, int] = {}
    max_col = max((len(r) for r in data), default=0)
    for col in range(1, max_col + 1):
        value = cell_val(data, 2, col)
        if not value:
            continue
        text = str(value).strip()
        if text in SECTION_HEADERS and text not in starts:
            starts[text] = col
    missing = [h for h in WRITE_SECTIONS if h not in starts]
    if missing:
        raise RuntimeError(f"Missing section headers: {missing}")
    return starts


def build_target_section_maps(data) -> Dict[str, Dict[str, int]]:
    starts = find_section_starts(data)
    maps: Dict[str, Dict[str, int]] = {}
    sorted_headers = sorted(starts.items(), key=lambda item: item[1])
    max_col = max((len(r) for r in data), default=0)

    for idx, (header, start_col) in enumerate(sorted_headers):
        next_start = sorted_headers[idx + 1][1] if idx + 1 < len(sorted_headers) else max_col + 1
        name_map: Dict[str, int] = {}
        for col in range(start_col + 1, next_start):
            value = cell_val(data, 2, col)
            if not value:
                continue
            name = str(value).strip()
            if name in SECTION_HEADERS:
                break
            if name != "":
                name_map[name] = col
        maps[header] = name_map
    return maps


# ============================================================================
# WORKSHEET WRITING OPERATIONS
# ============================================================================

def row_has_existing_values(
    data, target_row: int, section_maps: Dict[str, Dict[str, int]]
) -> bool:
    for section_name in WRITE_SECTIONS:
        for col in section_maps[section_name].values():
            if cell_val(data, target_row, col) not in (None, ""):
                return True
    return False


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


def get_previous_war_active_status(
    data, target_row: int, section_maps: Dict[str, Dict[str, int]], date_col: int = 15
) -> Optional[Dict[str, bool]]:
    active_cols = section_maps.get("Roster Active Status", {})
    for row in range(target_row - 1, 4, -1):
        if not cell_val(data, row, date_col):
            continue
        prev_status = {}
        has_any_data = False
        for name, col in active_cols.items():
            val = cell_val(data, row, col)
            if val is not None and val != "":
                has_any_data = True
                prev_status[name] = True
            else:
                prev_status[name] = False
        if has_any_data:
            return prev_status
    return None


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
    
    # --- DATA SOURCE SELECTION ---
    print("\nData Source:")
    print("1. Screenshot Folder")
    print("2. Live Screen Capture (Switch to game and scroll)")
    while True:
        source_choice = input("Choose 1/2: ").strip()
        if source_choice in ("1", "2"):
            break
        print("Invalid choice. Please enter 1 or 2.")
    
    video_temp_dir = None
    screenshots_folder = None
    folder_name = None

    if source_choice == "2":
        # --- LIVE CAPTURE MODE ---
        folder_name = ask_folder_name()
        
        # Countdown to give the user time to switch to the game
        print("\nSwitch to the game and open the Guild War logs.")
        for i in range(10, 0, -1):
            print(f"Recording starts in {i} seconds...", end="\r", flush=True)
            time.sleep(1)
        print("Recording started!                    ")  # Clears the line
        
        video_temp_dir = Path(tempfile.mkdtemp(prefix="gw_live_capture_"))
        
        stop_event = threading.Event()
        capture_thread = threading.Thread(
            target=live_capture_loop, 
            args=(video_temp_dir, stop_event)
        )
        
        capture_thread.start()
        input("\n(Recording in background... Press ENTER here when you are done scrolling)\n")
        
        stop_event.set()
        capture_thread.join()
        
        screenshots_folder = video_temp_dir

    else:
        # --- STANDARD FOLDER MODE ---
        folder_name = ask_folder_name()
        screenshots_folder = SCREENSHOTS_ROOT / folder_name

    if not screenshots_folder.exists():
        raise FileNotFoundError(f"Could not find source: {screenshots_folder}")


    aliases_existing = load_aliases(ALIASES_PATH)
    aliases_working = dict(aliases_existing)
    aliases_created_this_run: Dict[str, str] = {}
    config = load_config()
    client = get_gspread_client()
    spreadsheet = client.open_by_key(config["google_sheet"]["spreadsheet_id"])

    source_ws = spreadsheet.worksheet(config["source_sheet"])
    target_ws = spreadsheet.worksheet(config["target_sheet"])

    print("Fetching sheet data...", end=" ", flush=True)
    source_data = fetch_sheet_data(source_ws)
    target_data = fetch_sheet_data(target_ws)
    print("Done.")


    # THESE MUST BE DEFINED OUTSIDE THE TRY BLOCK
    roster = read_roster(source_data)
    target_row = find_date_row(target_data, folder_name, config["date_col"])
    target_section_maps = build_target_section_maps(target_data)
    overrides = read_column_a_overrides(target_data, roster, config["override_col"])

    # ==========================================
    # TRY BLOCK STARTS HERE
    # ==========================================
    try:
        extracted, _contexts = extract_rows_from_folder(
            screenshots_folder, folder_name, roster,
            aliases_existing, aliases_working, aliases_created_this_run, overrides
        )
        normalized = build_normalized_table(roster, extracted)

        # --- CONTEXT-AWARE ACTIVE/INACTIVE RESOLUTION ---
        prev_active_status = get_previous_war_active_status(target_data, target_row, target_section_maps, config["date_col"])

        for row in normalized:
            display_name = base_roster_name(row["Name"])

            # SCENARIO 1: The "New Recruit Paradox" 
            if row["Active flag"] == 1 and prev_active_status is not None and not prev_active_status.get(row["Name"], False):
                
                # If they used 0 tokens, they didn't attack. Always ask to confirm participation,
                # even if the defense column picked up a base defense stat or OCR noise.
                if row["Tokens Used"] == 0:
                    if ask_yes_no(
                        f"{display_name} is a new recruit with 0 tokens used. Did they actually participate in this war?",
                        default=False,
                    ):
                        row["Notes"] = "Active (New recruit, confirmed 0 stats)"
                    else:
                        print(f"  -> Marking {display_name} as inactive.")
                        row.update({
                            "Active flag": "", "Tokens Used": "", "Offense W": "", "Offense D": "",
                            "Offense L": "", "Defense W": "", "Defense D": "", "Defense L": "",
                            "Notes": "Inactive (New recruit, did not participate in this war)",
                        })
                        extracted = [r for r in extracted if r.matched_name != row["Name"]]
                else:
                    # They used tokens, so they definitely participated
                    row["Notes"] = "Active (New recruit)"
                continue

            # SCENARIO 2: Missing from screenshots entirely
            if row["Notes"] == "Inactive (not found in screenshots)":
                should_prompt = False
                if prev_active_status is None:
                    should_prompt = True
                elif prev_active_status.get(row["Name"], False):
                    should_prompt = True
                    
                if should_prompt:
                    action = prompt_missing_active_player(display_name)
                    if action == "manual":
                        print(f"\nEntering stats for {display_name}:")
                        off_w, off_d, off_l = prompt_stat_line("Offense")
                        def_w, def_d, def_l = prompt_stat_line("Defense")
                        row.update({
                            "Active flag": 1, "Tokens Used": off_w + off_d + off_l,
                            "Offense W": off_w, "Offense D": off_d, "Offense L": off_l,
                            "Defense W": def_w, "Defense D": def_d, "Defense L": def_l,
                            "Notes": "Active (Manually entered stats)",
                        })
                        extracted.append(OCRRow(
                            screenshot="MANUAL", order=len(extracted) + 1,
                            raw_name=display_name, matched_name=row["Name"],
                            off_w=off_w, off_d=off_d, off_l=off_l,
                            def_w=def_w, def_d=def_d, def_l=def_l,
                            note="Manually entered (missing from screenshots)",
                        ))
                    elif action == "inactive":
                        pass
                continue

        # --- TERMINAL OUTPUT ---
        print_results_output_terminal(extracted)
        print_normalized_table_terminal(normalized)
        print_run_summary(extracted, normalized, aliases_created_this_run)

        # --- PLAYER COUNT THRESHOLD CHECK ---
        if len(extracted) < MIN_EXPECTED_MATCHED_PLAYERS:
            print(f"\nWarning: matched unique players {len(extracted)} is below threshold {MIN_EXPECTED_MATCHED_PLAYERS}.")
            if not ask_yes_no("Continue anyway?", default=False):
                raise RuntimeError("Run cancelled because extracted player count was below threshold.")


        # --- DRY RUN CONVERSION LOGIC ---
        if run_mode == "dry":
            print("\nDry run complete. No workbook or alias changes were saved.")
            if ask_yes_no("Do you want to write these results to the workbook now?", default=False):
                print("\nConverting dry run to write mode...")
                run_mode = "write"
            else:
                return

        # --- WRITE TO WORKBOOK ---
        if run_mode == "write":
            if row_has_existing_values(target_data, target_row, target_section_maps):
                if not ask_yes_no(
                    "Target row already has values in one or more write sections. Overwrite whole row?",
                    default=False,
                ):
                    raise RuntimeError("Run cancelled by user at overwrite prompt.")
            
            clear_target_row_values(target_ws, target_row, target_section_maps)
            write_to_rough_sheet(target_ws, target_row, normalized, target_section_maps)
            
            if aliases_created_this_run:
                save_aliases(ALIASES_PATH, aliases_working)
                

            print("\nDone.")
            print(f"Folder processed: {folder_name}")
            print(f"Data written to: {config['target_sheet']} (date row matched from folder name)")
            print(f"New aliases saved: {len(aliases_created_this_run)}")

            # --- Export war record PNG ---
            try:
                subprocess.run(
                    [sys.executable, str(BASE_DIR / "war_record_exporter.py"), folder_name],
                    check=True,
                )
            except subprocess.CalledProcessError:
                print("Warning: War record export failed.")

    # ==========================================
    # FINALLY BLOCK ENSURES CLEANUP
    # ==========================================
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