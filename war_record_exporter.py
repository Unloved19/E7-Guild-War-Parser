"""Exports a war record summary from Google Sheets to a trimmed PNG.

Can be run standalone:
    python war_record_exporter.py 20260501

Or called from the main parser after a successful write.
"""

import sys
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, List

import requests
import fitz  # PyMuPDF
from PIL import Image, ImageChops, ImageOps
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request


# ============================================================================
# PATHS & CONFIG
# ============================================================================

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
TOKEN_PATH = BASE_DIR / "token.json"
SECRET_PATH = BASE_DIR / "client_secret.json"
WAR_RECORDS_DIR = BASE_DIR / "War Records"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

SEASON_TOTAL_TOKENS = 108
WAR_TOKEN_STEP = 3
EXPORT_RANGE = "B2:S45"
DATE_COL = 15  # Column O


# ============================================================================
# SHARED HELPERS
# ============================================================================

def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"Missing {CONFIG_PATH.name}")
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

def get_auth_token() -> str:
    """Reads the OAuth token directly from file, bypassing gspread internals."""
    if not TOKEN_PATH.exists():
        raise SystemExit("Token expired or missing. Run setup_auth.py again.")
    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise SystemExit("Token expired or missing. Run setup_auth.py again.")
    return creds.token


def cell_val(data: List[List], row: int, col: int) -> Optional[str]:
    if row < 1 or row > len(data):
        return None
    row_data = data[row - 1]
    if col < 1 or col > len(row_data):
        return None
    val = row_data[col - 1]
    return val if val != "" else None


def sheet_date_to_folder_date_str(value) -> Optional[str]:
    if value is None or value == "":
        return None
    if isinstance(value, str):
        value = value.strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y", "%Y/%m/%d", "%Y%m%d"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y%m%d")
            except ValueError:
                continue
    return None


# ============================================================================
# IMAGE PROCESSING
# ============================================================================

def trim_white_margins(img: Image.Image, padding: int = 8) -> Image.Image:
    rgb = img.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg)
    bbox = diff.getbbox()
    if bbox is None:
        return rgb
    left, top, right, bottom = bbox
    cropped = rgb.crop((left, top, right, bottom))
    return ImageOps.expand(cropped, border=padding, fill="white")


# ============================================================================
# TOKEN CALCULATION
# ============================================================================

def read_war_dates(data: List[List]) -> List[str]:
    """Reads all war dates from column O, returns sorted unique YYYYMMDD list."""
    dates = []
    for row in range(5, len(data) + 1):
        val = cell_val(data, row, DATE_COL)
        if val:
            parsed = sheet_date_to_folder_date_str(val)
            if parsed:
                dates.append(parsed)
    return sorted(set(dates))


def calculate_used_tokens(war_dates: List[str], current_date: str) -> int:
    """Calculates tokens used based on war date position in the season."""
    if current_date not in war_dates:
        first = war_dates[0] if war_dates else "none"
        raise RuntimeError(
            f"Date {current_date} not found in sheet dates "
            f"(starts with {first})."
        )
    war_number = war_dates.index(current_date) + 1  # 1-indexed
    return war_number * WAR_TOKEN_STEP


# ============================================================================
# EXPORT LOGIC
# ============================================================================

def export_to_png(
    client: gspread.Client,
    spreadsheet_id: str,
    sheet_name: str,
    folder_name: str,
    dpi: int = 300,
) -> Path:
    ws = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
    gid = ws.id

    data = ws.get_all_values()
    war_dates = read_war_dates(data)
    used = calculate_used_tokens(war_dates, folder_name)

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export"
    params = {
        "format": "pdf",
        "gid": gid,
        "range": EXPORT_RANGE,
    }
    headers = {"Authorization": f"Bearer {get_auth_token()}"}

    print(f"Exporting {EXPORT_RANGE} from {sheet_name}...", end=" ", flush=True)
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    print("Downloaded.", end=" ", flush=True)

    doc = fitz.open(stream=response.content, filetype="pdf")
    page = doc.load_page(0)
    scale = dpi / 72.0
    matrix = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=matrix, alpha=False)

    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    img = trim_white_margins(img, padding=8)

    WAR_RECORDS_DIR.mkdir(parents=True, exist_ok=True)
    filename = f"Zodiac {used}-{SEASON_TOTAL_TOKENS} {folder_name}.png"
    out_path = WAR_RECORDS_DIR / filename

    img.save(str(out_path), "PNG")
    print(f"Saved: {filename}")
    return out_path


# ============================================================================
# ENTRY POINT
# ============================================================================

def main():
    if len(sys.argv) < 2:
        folder_name = input("Enter war date (YYYYMMDD): ").strip()
    else:
        folder_name = sys.argv[1]

    if not re.fullmatch(r"\d{8}", folder_name):
        raise SystemExit("Date must be 8 digits in YYYYMMDD format.")

    config = load_config()
    client = get_gspread_client()

    export_to_png(
        client, config["spreadsheet_id"], config["target_sheet"], folder_name
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nError: {exc}")
        raise SystemExit(1)