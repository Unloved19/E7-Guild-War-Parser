from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Optional, TYPE_CHECKING
import shutil
import fitz  # PyMuPDF
from PIL import Image, ImageChops, ImageOps

if TYPE_CHECKING:
    from watchdog.events import FileSystemEventHandler as RealFileSystemEventHandler
    from watchdog.observers import Observer as RealObserver

WATCHDOG_AVAILABLE = True
try:
    from watchdog.events import FileSystemEventHandler  # type: ignore
    from watchdog.observers import Observer  # type: ignore
except Exception:
    WATCHDOG_AVAILABLE = False
    class FileSystemEventHandler:  # Dummy fallback for runtime compatibility
        pass
    Observer = None

# ============================================================================
# PATHS & FILE LOCATIONS
# ============================================================================

BASE_DIR = Path(r"F:\Game Clients\E7-Guild-War-Parser")
WAR_SCREENSHOTS_DIR = BASE_DIR / "War Screenshots"
WAR_RECORDS_DIR = BASE_DIR / "War Records"
DOWNLOADS_DIR = Path.home() / "Downloads"

# ============================================================================
# PRE-COMPILED REGEX PATTERNS
# ============================================================================

FINAL_PDF_RE = re.compile(r"^Zodiac (\d+)-(\d+) (\d{8})\.(?:pdf|png)$", re.IGNORECASE)
TMP_PNG_RE = re.compile(r"^__tmp__\d+__\.png$", re.IGNORECASE)
GSHEET_EXPORT_RE = re.compile(r"^Zodiac GW Spreadsheet.*\.pdf$", re.IGNORECASE)  # <--- ADD THIS

SEASON_TOTAL_TOKENS = 108
FIRST_TRACKED_DATE = "20260404"
FIRST_TRACKED_USED = 18
WAR_TOKEN_STEP = 3

WAR_DATES = [
    "20260404", "20260406", "20260408", "20260410",
    "20260413", "20260415", "20260417", "20260420", "20260422", "20260424",
    "20260427", "20260429", "20260501", "20260504", "20260506", "20260508",
    "20260511", "20260513", "20260515", "20260518", "20260520", "20260522",
    "20260525", "20260527", "20260529", "20260601", "20260603", "20260605",
    "20260608", "20260610", "20260612",
]

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


def pdf_to_png(pdf_path: Path, output_path: Optional[Path] = None, dpi: int = 300) -> Path:
    pdf_path = pdf_path.resolve()

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    if output_path is None:
        output_path = pdf_path.with_suffix(".png")
    else:
        output_path = output_path.resolve()

    doc = fitz.open(pdf_path)
    try:
        page = doc.load_page(0)
        scale = dpi / 72.0
        matrix = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=matrix, alpha=False)

        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)  # Fixed: list to tuple
        img = trim_white_margins(img, padding=8)
        img.save(output_path, "PNG")
    finally:
        doc.close()

    return output_path


def parse_final_pdf_name(path: Path) -> Optional[tuple[int, int, str]]:
    match = FINAL_PDF_RE.match(path.name)
    if not match:
        return None

    used = int(match.group(1))
    total = int(match.group(2))
    war_date = match.group(3)
    return used, total, war_date


def next_slot_from_existing(folder: Path) -> tuple[int, str]:
    existing = []

    for png in folder.glob("*.png"):
        parsed = parse_final_pdf_name(png)
        if parsed is None:
            continue

        used, total, war_date = parsed
        if total != SEASON_TOTAL_TOKENS:
            continue
        if war_date not in WAR_DATES:
            continue

        existing.append((used, war_date))

    if not existing:
        return FIRST_TRACKED_USED, FIRST_TRACKED_DATE

    latest_used, latest_date = max(existing, key=lambda item: WAR_DATES.index(item[1]))
    next_index = WAR_DATES.index(latest_date) + 1

    if next_index >= len(WAR_DATES):
        raise RuntimeError("No future war dates remain in WAR_DATES.")

    return latest_used + WAR_TOKEN_STEP, WAR_DATES[next_index]


def final_record_stem(used: int, war_date: str) -> str:
    return f"Zodiac {used}-{SEASON_TOTAL_TOKENS} {war_date}"


def rename_pngs_in_folder(folder: Path) -> None:
    png_files = sorted(
        [
            f for f in folder.iterdir()
            if f.is_file()
            and f.suffix.lower() == ".png"
            and not TMP_PNG_RE.match(f.name)
        ],
        key=lambda f: f.name.lower()
    )

    if not png_files:
        return

    changed = False
    temp_files: list[Path] = []

    for i, file in enumerate(png_files, start=1):
        desired_name = f"{folder.name} {i}.png"
        if file.name != desired_name:
            changed = True

        temp_path = file.with_name(f"__tmp__{i}__.png")
        if temp_path.exists():
            temp_path.unlink()

        file.rename(temp_path)
        temp_files.append(temp_path)

    for i, temp_file in enumerate(temp_files, start=1):
        final_path = temp_file.with_name(f"{folder.name} {i}.png")
        if final_path.exists():
            final_path.unlink()
        temp_file.rename(final_path)

    if changed:
        print(f"Renamed screenshots in: {folder.name}")


def rename_war_screenshots(root: Path = WAR_SCREENSHOTS_DIR) -> None:
    if not root.exists():
        raise FileNotFoundError(f"Could not find: {root}")

    found_any = False

    for folder in sorted(root.iterdir(), key=lambda p: p.name):
        if not folder.is_dir() or folder.name.startswith("."):
            continue

        found_any = True
        rename_pngs_in_folder(folder)

    if not found_any:
        print(f"No date folders found in: {root}")

def move_gsheet_export(
    downloads: Path = DOWNLOADS_DIR,
    records: Path = WAR_RECORDS_DIR,
) -> Optional[Path]:
    """Moves the Google Sheets PDF export from Downloads to the War Records folder."""
    if not downloads.exists():
        return None

    matches = [
        p for p in downloads.glob("*.pdf")
        if p.is_file() and GSHEET_EXPORT_RE.match(p.name)
    ]

    if not matches:
        return None

    # Sort by modification time, take the most recent
    matches.sort(key=lambda p: p.stat().st_mtime)
    pdf = matches[-1]

    target = records / pdf.name
    if target.exists():
        return None

    try:
        shutil.move(str(pdf), str(target))
        print(f"Moved: {pdf.name} -> {records.name}/")
        return target
    except Exception as exc:
        print(f"Failed to move {pdf.name}: {exc}")
        return None


def process_war_records(folder: Path = WAR_RECORDS_DIR, dpi: int = 300, overwrite_png: bool = False) -> None:
    if not folder.exists():
        raise FileNotFoundError(f"Could not find: {folder}")

    # Move Google Sheets export from Downloads if present
    move_gsheet_export(records=folder)

    pdfs = sorted(
        [p for p in folder.glob("*.pdf") if p.is_file()],
        key=lambda p: (p.stat().st_mtime, p.name.lower())
    )

    if not pdfs:
        return

    for pdf in pdfs:
        parsed = parse_final_pdf_name(pdf)

        # Already in final format
        if parsed is not None:
            used, total, war_date = parsed

            if total != SEASON_TOTAL_TOKENS:
                continue

            if war_date < FIRST_TRACKED_DATE:
                continue

            png_path = pdf.with_suffix(".png")
            if png_path.exists() and not overwrite_png:
                continue

            out = pdf_to_png(pdf, png_path, dpi=dpi)
            print(f"Converted: {pdf.name} -> {out.name}")
            pdf.unlink()  # <--- ADDED: Delete the PDF
            continue

        # Generic export, rename to next slot then convert
        used, war_date = next_slot_from_existing(folder)
        new_stem = final_record_stem(used, war_date)
        new_pdf = pdf.with_name(f"{new_stem}.pdf")
        new_png = pdf.with_name(f"{new_stem}.png")

        if new_pdf.exists():
            raise FileExistsError(f"Target PDF already exists: {new_pdf.name}")

        generic_png = pdf.with_suffix(".png")

        pdf.rename(new_pdf)
        print(f"Renamed PDF: {pdf.name} -> {new_pdf.name}")

        if generic_png.exists():
            if new_png.exists():
                if overwrite_png:
                    new_png.unlink()
                else:
                    new_pdf.unlink()
                    continue

            generic_png.rename(new_png)
            print(f"Renamed PNG: {generic_png.name} -> {new_png.name}")

            if overwrite_png:
                out = pdf_to_png(new_pdf, new_png, dpi=dpi)
                print(f"Re-converted: {new_pdf.name} -> {out.name}")
            
            new_pdf.unlink()
            continue

        out = pdf_to_png(new_pdf, new_png, dpi=dpi)
        print(f"Converted: {new_pdf.name} -> {out.name}")
        new_pdf.unlink()


class WarRecordsWatcher(FileSystemEventHandler):  # type: ignore
    def __init__(self, folder: Path, dpi: int = 300, overwrite_png: bool = False) -> None:
        self.folder = folder
        self.dpi = dpi
        self.overwrite_png = overwrite_png
        self.last_processed = 0.0

    def _handle(self, path_str: str) -> None:
        path = Path(path_str)

        if path.is_dir() or path.suffix.lower() != ".pdf":
            return

        now = time.time()
        if now - self.last_processed < 1.0:
            return

        self.last_processed = now
        time.sleep(1.0)

        try:
            process_war_records(self.folder, dpi=self.dpi, overwrite_png=self.overwrite_png)
        except Exception as exc:
            print(f"War Records watcher error: {exc}")

    def on_created(self, event) -> None:
        self._handle(event.src_path)

    def on_modified(self, event) -> None:
        self._handle(event.src_path)

    def on_moved(self, event) -> None:
        self._handle(event.dest_path)


class WarScreenshotsWatcher(FileSystemEventHandler):  # type: ignore
    def __init__(self, root: Path) -> None:
        self.root = root
        self.last_processed: dict[Path, float] = {}

    def _handle(self, path_str: str) -> None:
        path = Path(path_str)

        if path.is_dir() or path.suffix.lower() != ".png":
            return

        parent = path.parent
        if parent == self.root:
            return

        now = time.time()
        last = self.last_processed.get(parent, 0.0)

        if now - last < 0.75:
            return

        self.last_processed[parent] = now
        time.sleep(0.5)

        try:
            rename_pngs_in_folder(parent)
        except Exception as exc:
            print(f"War Screenshots watcher error in {parent.name}: {exc}")

    def on_created(self, event) -> None:
        self._handle(event.src_path)

    def on_modified(self, event) -> None:
        self._handle(event.src_path)

    def on_moved(self, event) -> None:
        self._handle(event.dest_path)

class DownloadsWatcher(FileSystemEventHandler):  # type: ignore
    def __init__(self, records: Path, dpi: int = 300, overwrite_png: bool = False) -> None:
        self.records = records
        self.dpi = dpi
        self.overwrite_png = overwrite_png
        self.last_processed = 0.0

    def _handle(self, path_str: str) -> None:
        path = Path(path_str)

        if path.is_dir() or path.suffix.lower() != ".pdf":
            return
            
        if not GSHEET_EXPORT_RE.match(path.name):
            return

        now = time.time()
        if now - self.last_processed < 2.0:
            return

        self.last_processed = now
        time.sleep(2.0)  # Wait for browser to finish writing the download

        try:
            moved = move_gsheet_export(path.parent, self.records)
            if moved:
                process_war_records(self.records, dpi=self.dpi, overwrite_png=self.overwrite_png)
        except Exception as exc:
            print(f"Downloads watcher error: {exc}")

    def on_created(self, event) -> None:
        self._handle(event.src_path)

    def on_modified(self, event) -> None:
        self._handle(event.src_path)

    def on_moved(self, event) -> None:
        self._handle(event.dest_path)

def watch_war_records_and_screenshots(
    war_records: Path = WAR_RECORDS_DIR,
    war_screenshots: Path = WAR_SCREENSHOTS_DIR,
    downloads: Path = DOWNLOADS_DIR,  # <--- ADD THIS
    dpi: int = 300,
    overwrite_png: bool = False,
) -> None:
    if not WATCHDOG_AVAILABLE:
        raise RuntimeError("watchdog is not installed. Run: python -m pip install watchdog")

    assert Observer is not None  # Type checker hint: Observer is available here

    if not war_records.exists():
        raise FileNotFoundError(f"Could not find: {war_records}")
    if not war_screenshots.exists():
        raise FileNotFoundError(f"Could not find: {war_screenshots}")

    observer = Observer()
    observer.schedule(
        WarRecordsWatcher(war_records, dpi=dpi, overwrite_png=overwrite_png),  # type: ignore[arg-type]
        str(war_records),
        recursive=False,
    )
    observer.schedule(
        WarScreenshotsWatcher(war_screenshots),  # type: ignore[arg-type]
        str(war_screenshots),
        recursive=True,
    )
    
    # Watch for Google Sheets exports in Downloads
    if downloads.exists():
        observer.schedule(
            DownloadsWatcher(war_records, dpi=dpi, overwrite_png=overwrite_png),  # type: ignore[arg-type]
            str(downloads),
            recursive=False,
        )
        print(f"Watching Downloads (Google Sheets exports): {downloads}")
    else:
        print(f"Downloads folder not found, skipping watcher: {downloads}")

    observer.start()

    print(f"Watching War Records: {war_records}")
    print(f"Watching War Screenshots: {war_screenshots}")
    print("Press Ctrl+C to stop.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping watcher...")
    finally:
        observer.stop()
        observer.join()

    print(f"Watching War Records: {war_records}")
    print(f"Watching War Screenshots: {war_screenshots}")
    print("Press Ctrl+C to stop.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping watcher...")
    finally:
        observer.stop()
        observer.join()

def cleanup_existing_pdfs(folder: Path = WAR_RECORDS_DIR) -> None:
    """Deletes PDFs that already have a corresponding PNG."""
    if not folder.exists():
        raise FileNotFoundError(f"Could not find: {folder}")

    deleted_count = 0
    for pdf in folder.glob("*.pdf"):
        png_path = pdf.with_suffix(".png")
        if png_path.exists():
            try:
                pdf.unlink()
                print(f"Deleted: {pdf.name}")
                deleted_count += 1
            except Exception as exc:
                print(f"Failed to delete {pdf.name}: {exc}")

    print(f"Cleanup complete. Removed {deleted_count} PDF(s).")
    
def show_menu() -> str:
    print("\nZodiac War Tools")
    print("1. Rename War Screenshots")
    print("2. Rename and Convert War Record PDFs")
    print("3. Watch War Records and War Screenshots")
    print("4. Cleanup Existing PDFs (Keep PNGs only)")
    return input("Choose an option (1/2/3/4): ").strip()


def main() -> int:
    try:
        choice = show_menu()

        if choice == "1":
            rename_war_screenshots()
            print("Done.")
            return 0

        if choice == "2":
            process_war_records()
            print("Done.")
            return 0

        if choice == "3":
            watch_war_records_and_screenshots()
            return 0
        
        if choice == "4":
            cleanup_existing_pdfs()
            print("Done.")
            return 0

        print("Invalid choice.")
        return 1

    except Exception as exc:
        print(f"Error: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())