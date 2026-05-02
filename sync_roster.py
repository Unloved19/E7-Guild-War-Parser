"""Syncs Column B roster names with helper section columns, or creates a new season tab.

Run standalone:
    python sync_roster.py              # Sync add/remove
    python sync_roster.py --new-season # Duplicate tab, clear data, keep roster+formulas
"""

import sys
import json
from pathlib import Path
from typing import Dict, List, Set

import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
TOKEN_PATH = BASE_DIR / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# All 10 sections that have 60 player columns.
# MVP and MVP Count are NOT included — they have no player helper columns.
WRITE_SECTIONS = [
    "Roster Active Status", "Tokens Tracking",
    "Offense Wins", "Offense Draws", "Offense Losses", "Offense Win Rate",
    "Defense Wins", "Defense Draws", "Defense Losses", "Defense Win Rate",
]

SECTION_HEADERS = [
    "Roster Active Status", "Tokens Tracking",
    "Offense Wins", "Offense Draws", "Offense Losses", "Offense Win Rate",
    "MVP", "MVP Count",
    "Defense Wins", "Defense Draws", "Defense Losses", "Defense Win Rate",
]


def load_config() -> dict:
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def save_config(config: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(config, indent=2), encoding="utf-8")


def get_gspread_client() -> gspread.Client:
    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds.valid and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return gspread.authorize(creds)


def cell_val(data: list, row: int, col: int):
    if row < 1 or row > len(data):
        return None
    r = data[row - 1]
    if col < 1 or col > len(r):
        return None
    v = r[col - 1]
    return v if v != "" else None


def read_roster(data: list) -> List[str]:
    roster = []
    blank = 0
    row = 5
    while row <= len(data) + 10:
        val = cell_val(data, row, 2)
        if val is None:
            blank += 1
            if blank >= 10:
                break
        else:
            blank = 0
            name = str(val).strip()
            if name.lower().startswith("notes:"):
                row += 1
                continue
            roster.append(name)
        row += 1
    return roster


def find_section_starts(data: list) -> Dict[str, int]:
    starts = {}
    max_col = max((len(r) for r in data), default=0)
    for col in range(1, max_col + 1):
        val = cell_val(data, 2, col)
        if not val:
            continue
        text = str(val).strip()
        if text in SECTION_HEADERS and text not in starts:
            starts[text] = col
    return starts


def build_section_maps(data: list) -> Dict[str, Dict[str, int]]:
    starts = find_section_starts(data)
    maps = {}
    sorted_h = sorted(starts.items(), key=lambda x: x[1])
    max_col = max((len(r) for r in data), default=0)

    for idx, (header, start_col) in enumerate(sorted_h):
        next_start = sorted_h[idx + 1][1] if idx + 1 < len(sorted_h) else max_col + 1
        name_map = {}
        for col in range(start_col + 1, next_start):
            val = cell_val(data, 2, col)
            if not val:
                continue
            name = str(val).strip()
            if name in SECTION_HEADERS:
                break
            if name != "":
                name_map[name] = col
        maps[header] = name_map
    return maps


def _col_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def sync_roster(ws, data: list, section_maps: Dict[str, Dict[str, int]], roster: List[str]) -> None:
    """Adds new players to helper sections, clears removed players.

    Each section is independent — finds empty slots per section.
    No cross-section alignment needed because formulas reference
    their own section ranges.
    """
    roster_set = set(roster)
    existing: Set[str] = set()
    for section in WRITE_SECTIONS:
        existing.update(section_maps[section].keys())

    to_add = sorted(roster_set - existing)
    to_remove = sorted(existing - roster_set)

    if not to_add and not to_remove:
        print("Roster is already in sync with helper sections.")
        return

    updates = []

    # --- Remove players (clear header + data in their column) ---
    if to_remove:
        print(f"\nClearing {len(to_remove)} removed player(s):")
        for name in to_remove:
            print(f"  - {name}")
        for section in WRITE_SECTIONS:
            for name in to_remove:
                col = section_maps[section].get(name)
                if col:
                    letter = _col_letter(col)
                    updates.append({"range": f"{letter}2", "values": [[""]]})
                    for row in range(5, len(data) + 1):
                        if cell_val(data, row, col):
                            updates.append({"range": f"{letter}{row}", "values": [[""]]})

    # --- Add players (find empty slots independently per section) ---
    if to_add:
        print(f"\nAdding {len(to_add)} new player(s):")
        for name in to_add:
            print(f"  + {name}")
        for section in WRITE_SECTIONS:
            existing_names = set(section_maps[section].keys())
            empty_slots = []
            for col in range(1, len(data[0]) + 1 if data else 1):
                header = cell_val(data, 2, col)
                if header and header not in existing_names and header not in SECTION_HEADERS:
                    empty_slots.append(col)
            if len(empty_slots) < len(to_add):
                print(f"  WARNING: '{section}' only has {len(empty_slots)} empty slot(s), "
                      f"need {len(to_add)}!")
            for i, name in enumerate(to_add):
                if i < len(empty_slots):
                    col = empty_slots[i]
                    updates.append({"range": f"{_col_letter(col)}2", "values": [[name]]})

    if updates:
        ws.batch_update(updates)
        total = len(to_add) * len(WRITE_SECTIONS) + len(updates) - (len(to_add) * len(WRITE_SECTIONS))
        print(f"\n  Updated {len(updates)} cells.")
    print("Done.")


def new_season(client, spreadsheet, config: dict) -> None:
    """Duplicates the current tab, clears war data, keeps roster + formulas."""
    current_name = config["target_sheet"]
    # Try to auto-increment season number (e.g. "GW Tracker 2026-1" -> "GW Tracker 2026-2")
    parts = current_name.rsplit("-", 1)
    if len(parts) == 2 and parts[1].strip().isdigit():
        default_new = f"{parts[0]}-{int(parts[1].strip()) + 1}"
    else:
        default_new = f"{current_name} (2)"

    print(f"\nCurrent season tab: {current_name}")
    new_name = input(f"New season tab name [default: {default_new}]: ").strip()
    if not new_name:
        new_name = default_new

    existing_names = [ws.title for ws in spreadsheet.worksheets()]
    if new_name in existing_names:
        print(f"Error: Tab '{new_name}' already exists.")
        return

    source_ws = spreadsheet.worksheet(current_name)

    print(f"\nDuplicating '{current_name}' as '{new_name}'...", end=" ", flush=True)
    body = {
        "requests": [{
            "duplicateSheet": {
                "sourceSheetId": source_ws.id,
                "insertSheetIndex": 0,
                "newSheetName": new_name,
            }
        }]
    }
    spreadsheet.batch_update(body)
    print("Done.")

    new_ws = spreadsheet.worksheet(new_name)
    data = new_ws.get_all_values()
    section_maps = build_section_maps(data)

    # Clear helper section data (rows 5+) and player headers (row 2)
    print("Clearing helper section data and player names...", end=" ", flush=True)
    updates = []
    for section in WRITE_SECTIONS:
        for name, col in section_maps[section].items():
            letter = _col_letter(col)
            # Clear player name in row 2
            updates.append({"range": f"{letter}2", "values": [[""]]})
            # Clear war data in rows 5+
            for row in range(5, len(data) + 1):
                if cell_val(data, row, col):
                    updates.append({"range": f"{letter}{row}", "values": [[""]]})

    # Clear summary block data (B5:S last row), keep row 2-4 headers and formulas
    for row in range(5, len(data) + 1):
        for col in range(2, 20):  # B to S
            if cell_val(data, row, col):
                updates.append({"range": f"{_col_letter(col)}{row}", "values": [[""]]})

    # Clear date column data (rows 5+)
    date_col = config.get("date_col", 15)
    for row in range(5, len(data) + 1):
        if cell_val(data, row, date_col):
            updates.append({"range": f"{_col_letter(date_col)}{row}", "values": [[""]]})

    if updates:
        new_ws.batch_update(updates)
    print(f"Cleared {len(updates)} cells.")

    # Update config
    config["source_sheet"] = new_name
    config["target_sheet"] = new_name
    save_config(config)

    print(f"\nNew season tab '{new_name}' is ready.")
    print("Roster in Column B and all formulas are preserved.")
    print(f"Config updated to use '{new_name}'.")
    print(f"\nNext steps:")
    print(f"  1. Enter the first war date in the date column (row 5)")
    print(f"  2. Drag down to fill all war dates")
    print(f"  3. Add/remove players in Column B if needed")
    print(f"  4. Run 'Sync roster' to update helper sections")
    print(f"  5. Run the parser to record your first war")


def main():
    mode = "new-season" if "--new-season" in sys.argv else "sync"

    config = load_config()
    client = get_gspread_client()
    spreadsheet = client.open_by_key(config["spreadsheet_id"])

    if mode == "new-season":
        new_season(client, spreadsheet, config)
        return

    ws = spreadsheet.worksheet(config["target_sheet"])

    print(f"Reading '{config['target_sheet']}'...", end=" ", flush=True)
    data = ws.get_all_values()
    print("Done.")

    roster = read_roster(data)
    if not roster:
        print("No roster found in Column B. Add player names first.")
        return

    print(f"Found {len(roster)} player(s) in Column B.")
    section_maps = build_section_maps(data)
    sync_roster(ws, data, section_maps, roster)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nError: {exc}")
        raise SystemExit(1)