from __future__ import annotations
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import json
import shutil
import threading
import tempfile
import time

try:
    import cv2
except ImportError as exc:
    raise SystemExit("Missing dependency: opencv-python. Install with: python -m pip install opencv-python") from exc

try:
    import easyocr
except ImportError as exc:
    raise SystemExit("Missing dependency: easyocr. Install with: python -m pip install easyocr") from exc

# ============================================================================
# PATHS & FILE LOCATIONS
# ============================================================================

BASE_DIR = Path(__file__).resolve().parent
SCREENSHOTS_ROOT = BASE_DIR / "War Screenshots"
ALIASES_PATH = BASE_DIR / "war_name_aliases.json"
DEBUG_ROOT = BASE_DIR / "debug_gw_parser"

# ============================================================================
# WORKBOOK CONFIGURATION
# ============================================================================

SECTION_HEADERS = [
    "Roster Active Status", "Tokens Tracking", "Offense Wins", "Offense Draws",
    "Offense Losses", "Offense Win Rate", "MVP", "MVP Count", "Defense Wins",
    "Defense Draws", "Defense Losses", "Defense Win Rate",
]

WRITE_SECTIONS = [
    "Roster Active Status", "Tokens Tracking", "Offense Wins", "Offense Draws",
    "Offense Losses", "Defense Wins", "Defense Draws", "Defense Losses",
]

PROCESSING_NORMALIZED_HEADER = [
    "Name", "Active flag", "Tokens Used", "Offense W", "Offense D", "Offense L",
    "Defense W", "Defense D", "Defense L", "Notes",
]

# ============================================================================
# OCR CONFIGURATION
# ============================================================================

PRIMARY_NAME_LANGS = ["ch_sim", "en"]
JA_FALLBACK_LANGS = ["ja", "en"]
KO_FALLBACK_LANGS = ["ko", "en"]
STATS_LANGS = ["en"]
MIN_EXPECTED_MATCHED_PLAYERS = 28

READER_CACHE: Dict[str, object] = {}


# ============================================================================
# CROP COORDINATES (Tuned for Fullscreen 1080p Monitor Capture)
# ============================================================================

NAME_CROP_X1 = 490
NAME_CROP_X2 = 882
OFF_CROP_X1 = 892
OFF_CROP_X2 = 1102  # Exactly 210px wide
DEF_CROP_X1 = 1252
DEF_CROP_X2 = 1462  # Exactly 210px wide

# ============================================================================
# PRE-COMPILED REGEX PATTERNS (Performance optimization)
# ============================================================================

NAME_TOKEN_RE = re.compile(r"[A-Za-z0-9一-龯ぁ-んァ-ン가-힣·._'\-]+")
RE_WHITESPACE = re.compile(r"\s+")
RE_REMOVED_TAG = re.compile(r"\s+\(Removed [^)]+\)$", re.IGNORECASE)
RE_PARENTHETICAL = re.compile(r"\s+\([^)]*\)$")
RE_RANK_ART1 = re.compile(r"(?i)\b[s$]?\.*rank[\.:,\-\s]*\d+\b")
RE_RANK_ART2 = re.compile(r"(?i)\brank[\.:,\-\s]*\d+\b")
RE_RANK_ART3 = re.compile(r"(?i)\bs\.?rank\b")
RE_NON_ALNUM_CJK = re.compile(r"[^a-z0-9一-龯ぁ-んァ-ン가-힣]+")
RE_FOLDER_NAME = re.compile(r"\d{8}")
RE_STAT_SEPARATOR = re.compile(r"[\s/]+")
RE_PATH_INVALID = re.compile(r'[\\/:*?"<>|]+')
RE_DIGITS = re.compile(r"(\d+)")
RE_WINS_PREFIX = re.compile(r"(\d+)\s*wi", re.IGNORECASE)
RE_DRAW_PREFIX = re.compile(r"(\d+)\s*dra", re.IGNORECASE)
RE_LOSS_PREFIX = re.compile(r"(\d+)\s*lo", re.IGNORECASE)

NOISE_TERMS = frozenset({
    "guild", "war", "logs", "participants", "offense", "defense",
    "previous", "current", "season", "rank", "s.rank",
})

STAT_TERMS = frozenset({"wins", "draw", "draws", "loss", "losses"})


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class OCRRow:
    screenshot: str
    order: int
    raw_name: str
    matched_name: str
    off_w: int
    off_d: int
    off_l: int
    def_w: int
    def_d: int
    def_l: int
    note: str = ""

    @property
    def result_string(self) -> str:
        return (
            f"{self.matched_name} = "
            f"{self.off_w} / {self.off_d} / {self.off_l} and "
            f"{self.def_w} / {self.def_d} / {self.def_l}"
        )


@dataclass
class OCRContext:
    folder_name: str
    screenshot: str
    order: int
    center_y: int
    row_slot: int
    total_slots: int
    name_candidates: List[str]
    off_w: int
    off_d: int
    off_l: int
    def_w: int
    def_d: int
    def_l: int
    name_crop: object
    off_crop: object
    def_crop: object
    debug_export_asked: bool = False
    debug_export_enabled: bool = False

# ============================================================================
# ALIAS & FILE UTILITIES
# ============================================================================

def load_aliases(path: Path) -> Dict[str, str]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_aliases(path: Path, aliases: Dict[str, str]) -> None:
    path.write_text(json.dumps(aliases, ensure_ascii=False, indent=2), encoding="utf-8")


# ============================================================================
# TEXT PROCESSING & NORMALIZATION
# ============================================================================

def cleaned_text(text: str) -> str:
    text = text.replace("\n", " ").replace("\r", " ")
    return RE_WHITESPACE.sub(" ", text).strip()


def base_roster_name(name: str) -> str:
    s = cleaned_text(name)
    s = RE_REMOVED_TAG.sub("", s)
    s = RE_PARENTHETICAL.sub("", s)
    return RE_WHITESPACE.sub(" ", s).strip()


def strip_rank_artifacts(text: str) -> str:
    text = RE_RANK_ART1.sub(" ", text)
    text = RE_RANK_ART2.sub(" ", text)
    text = RE_RANK_ART3.sub(" ", text)
    return cleaned_text(text)


def canonical_key(name: str) -> str:
    base = base_roster_name(name)
    return "".join(ch.lower() for ch in base if ch.isalnum())


def squash_repeats(text: str, max_run: int = 2) -> str:
    if not text:
        return text
    out = []
    last = None
    run = 0
    for ch in text:
        if ch == last:
            run += 1
        else:
            last = ch
            run = 1
        if run <= max_run:
            out.append(ch)
    return "".join(out)


def normalize_name_for_matching(name: str) -> str:
    s = base_roster_name(name)
    s = s.lower()
    s = s.replace("1", "i").replace("|", "i").replace("!", "i")
    s = s.replace("vv", "w")
    s = squash_repeats(s, max_run=2)
    s = RE_NON_ALNUM_CJK.sub("", s)
    return s.strip()

# ============================================================================
# USER INTERACTION & CLI PROMPTS
# ============================================================================

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

def prompt_missing_active_player(display_name: str) -> str:
    print(f"\n{display_name} was active last war but missing from this war's screenshots.")
    print("  1. Manually enter stats (screenshots were missing/broken)")
    print("  2. Mark as Inactive (removed, benched, or didn't participate)")
    print("  3. Stop run to check data")
    while True:
        answer = input("Choose 1/2/3: ").strip()
        if answer == "1":
            return "manual"
        if answer == "2":
            return "inactive"
        if answer == "3":
            raise RuntimeError(f"Run stopped by user while resolving: {display_name}")
        print("Invalid choice.")


def prompt_stat_line(label: str) -> Tuple[int, int, int]:
    while True:
        raw = input(f"  {label} W D L (e.g. 2 0 1 or 2/0/1): ").strip()
        if not raw:
            continue
        parts = RE_STAT_SEPARATOR.split(raw)
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            return int(parts[0]), int(parts[1]), int(parts[2])
        print("    Invalid format. Please enter 3 numbers.")

# ============================================================================
# IMAGE PROCESSING & CROP UTILITIES
# ============================================================================

def preprocess_for_ocr(image):
    if image is None or getattr(image, "size", 0) == 0:
        return image
    if len(image.shape) == 3:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    else:
        gray = image.copy()
    return cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)


def deterministic_row_centers(height: int) -> List[Tuple[int, int]]:
    # Returns tuples of (Name Center Y, Stat Center Y) for each of the 4 rows
    return [
        (287, 266),
        (470, 447),
        (652, 620),
        (825, 797),
    ]


def detect_rank_centers(img) -> List[Tuple[int, int]]:
    height, _width = img.shape[:2]
    return deterministic_row_centers(height)

def crop_bounds(
    img,
    name_center_y: int,
    stat_center_y: int,
    row_slot: Optional[int] = None,
    total_slots: Optional[int] = None,
) -> Tuple[Tuple[int, int], Tuple[int, int], Tuple[int, int]]:
    height, _ = img.shape[:2]

    # Reverted top padding to original, extended bottom by 30px to catch UI scroll overshoot
    name_y1 = max(0, name_center_y - 26)
    name_y2 = min(height, name_center_y + 93)  # Extended bottom by 30px (was 63)

    # Reverted top padding to original, extended bottom by 30px to catch UI scroll overshoot
    stats_y1 = max(0, stat_center_y - 33)
    stats_y2 = min(height, stat_center_y + 185)  # Extended bottom by 30px (was 155)

    if total_slots == 4 and row_slot == 4:
        stats_y2 = min(height, stat_center_y + 198)  # Extended bottom by 30px (was 168)

    if total_slots == 2:
        stats_y2 = min(height, stats_y2 + 10)

    return (name_y1, name_y2), (stats_y1, stats_y2), (stats_y1, stats_y2)

def safe_crop(img, x1: int, y1: int, x2: int, y2: int):
    h, w = img.shape[:2]
    x1 = max(0, min(w, x1))
    x2 = max(0, min(w, x2))
    y1 = max(0, min(h, y1))
    y2 = max(0, min(h, y2))
    if x2 <= x1 or y2 <= y1:
        return None
    return img[y1:y2, x1:x2].copy()

# ============================================================================
# OCR EXECUTION & NAME EXTRACTION
# ============================================================================

def ocr_text(reader, image, allowlist: Optional[str] = None) -> List[Tuple[float, str, float]]:
    if image is None or getattr(image, "size", 0) == 0:
        return []
    processed = preprocess_for_ocr(image)
    kwargs = {"detail": 1, "paragraph": False}
    if allowlist:
        kwargs["allowlist"] = allowlist
    results = reader.readtext(processed, **kwargs)
    lines: List[Tuple[float, str, float]] = []
    for box, text, conf in results:
        text = cleaned_text(str(text))
        if not text:
            continue
        y = sum(pt[1] for pt in box) / len(box)
        lines.append((y, text, float(conf)))
    lines.sort(key=lambda item: item[0])
    return lines


def get_reader(reader_type: str):
    if reader_type not in READER_CACHE:
        lang_map = {
            "name_primary": PRIMARY_NAME_LANGS,
            "name_ja": JA_FALLBACK_LANGS,
            "name_ko": KO_FALLBACK_LANGS,
            "stats": STATS_LANGS,
        }
        if reader_type not in lang_map:
            raise ValueError(f"Unknown reader_type: {reader_type}")
        READER_CACHE[reader_type] = easyocr.Reader(lang_map[reader_type], gpu=False)
    return READER_CACHE[reader_type]


def looks_like_ui_noise(text: str) -> bool:
    low = text.lower()
    return any(term in low for term in NOISE_TERMS)


def name_token_score(token: str) -> int:
    letters = sum(ch.isalpha() for ch in token)
    digits = sum(ch.isdigit() for ch in token)
    cjk = sum(
        "\u4e00" <= ch <= "\u9fff" or "\u3040" <= ch <= "\u30ff" or "\uac00" <= ch <= "\ud7af"
        for ch in token
    )
    punctuation = sum(not ch.isalnum() and ch not in "._-'·" for ch in token)
    score = len(token) + letters * 2 + cjk * 3 - digits - punctuation * 3
    if token.lower() in STAT_TERMS:
        score -= 50
    return score


def best_name_token_from_line(text: str) -> str:
    text = strip_rank_artifacts(text)
    if not text:
        return ""
    tokens = NAME_TOKEN_RE.findall(text)
    filtered = [tok for tok in tokens if not looks_like_ui_noise(tok)]
    if not filtered:
        return ""
    
    # COMBINATION LOGIC: Fix names split by spaces (e.g., "Kyūbi 焰" -> "Kyūbi焰")
    combined = []
    for tok in filtered:
        if combined:
            prev = combined[-1]
            # Check if previous token is primarily Latin/English
            prev_is_latin = prev.isascii() and any(c.isalpha() for c in prev)
            # Check if current token contains CJK (Japanese/Chinese) characters
            curr_is_cjk = any(
                "\u4e00" <= ch <= "\u9fff" or "\u3040" <= ch <= "\u30ff" or "\uac00" <= ch <= "\ud7af"
                for ch in tok
            )
            # If we have Latin followed by CJK, merge them together
            if prev_is_latin and curr_is_cjk:
                combined[-1] = prev + tok
                continue
                
        combined.append(tok)

    return max(combined, key=name_token_score)


def extract_name_text(reader, image) -> str:
    lines = ocr_text(reader, image)
    candidates: List[str] = []
    for _, text, _ in lines:
        cleaned = cleaned_text(text)
        if not cleaned:
            continue
        token = best_name_token_from_line(cleaned)
        if token:
            candidates.append(token)
    if not candidates:
        return ""
    return max(candidates, key=name_token_score)


def parse_name_candidates_from_crop(image, roster: List[str], matcher: Dict[str, str], aliases_working: Dict[str, str]) -> List[str]:
    candidates: List[str] = []
    for reader_type in ("name_primary", "name_ja", "name_ko"):
        reader = get_reader(reader_type)
        text = extract_name_text(reader, image)
        text = cleaned_text(text)
        if not text or text in candidates:
            continue
        candidates.append(text)
        
        # EARLY EXIT: If we found a direct match, skip Japanese/Korean OCR entirely
        if try_direct_name_match(text, roster, matcher, aliases_working) is not None:
            break
            
    return candidates


def try_direct_name_match(
    raw_name: str, roster: List[str], matcher: Dict[str, str], aliases: Dict[str, str]
) -> Optional[Tuple[str, str]]:
    raw_name = cleaned_text(raw_name)
    key = canonical_key(raw_name)

    # Matcher first (includes column C overrides)
    if key in matcher:
        matched = matcher[key]
        note = "Matched to removed-tag name" if base_roster_name(matched) != matched else ""
        if raw_name in aliases and aliases[raw_name] != matched:
            note = "Column C override (superseded alias)"
        return matched, note

    # Alias fallback (only if matcher has no entry for this key)
    if raw_name in aliases and aliases[raw_name] in roster:
        matched = aliases[raw_name]
        note = "Matched to removed-tag name" if base_roster_name(matched) != matched else "Manual alias fix"
        return matched, note

    return None


# ============================================================================
# STATS PARSING (OFFENSE & DEFENSE)
# ============================================================================

def _closest_valid_offense(values: Tuple[int, int, int]) -> Tuple[int, int, int]:
    valid: List[Tuple[int, int, int]] = [
        (w, d, l)
        for w in range(4) for d in range(4) for l in range(4) if w + d + l in (0, 1, 2, 3)
    ]
    return min(
        valid,
        key=lambda cand: abs(cand[0] - values[0]) + abs(cand[1] - values[1]) + abs(cand[2] - values[2])
    )
    
def _preprocess_stats_crop(image, reader):
    """Shared preprocessing for offense/defense stat crops."""
    if image is None or getattr(image, "size", 0) == 0:
        return None, []
    if len(image.shape) == 3:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    else:
        gray = image.copy()
    gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    return gray, reader.readtext(gray, detail=1, paragraph=False)

def _build_detections(results, max_digit: int = 3):
    """Build detection list from OCR results. max_digit differs for offense (3) vs defense (7+)."""
    detections = []
    for box, text, conf in results:
        text = str(text).strip()
        low = text.lower()
        y_center = sum(pt[1] for pt in box) / len(box)
        m = RE_DIGITS.search(text)
        digit = int(m.group(1)) if m else None
        if digit is not None and digit > 3 and len(str(digit)) >= 2:
            digit = int(str(digit)[0])
        if digit is not None and digit > max_digit:
            continue
        label = None
        if "win" in low:
            label = "w"
        elif "dra" in low:
            label = "d"
        elif "los" in low:
            label = "l"
        detections.append({"text": text, "y": y_center, "digit": digit, "label": label})
    return detections

def _apply_strategies_1_and_2(detections, gray_height, proximity_threshold_ratio=0.20):
    """Apply combined text (Strategy 1) and label-to-digit proximity (Strategy 2)."""
    w, d, l = None, None, None
    
    # Strategy 1: Combined "N Wins" / "N Draw(s)" / "N Losses" text
    for det in detections:
        low = det["text"].lower()
        m = RE_WINS_PREFIX.match(low)
        if m and w is None:
            w = int(m.group(1))
            continue
        m = RE_DRAW_PREFIX.match(low)
        if m and d is None:
            d = int(m.group(1))
            continue
        m = RE_LOSS_PREFIX.match(low)
        if m and l is None:
            l = int(m.group(1))
            continue

    # Strategy 2: Match separate label detections to nearest digit
    if w is None or d is None or l is None:
        labels_by_type = {}
        digits_with_y = []
        for det in detections:
            if det["label"]:
                labels_by_type.setdefault(det["label"], []).append(det["y"])
            if det["digit"] is not None:
                digits_with_y.append({"y": det["y"], "digit": det["digit"], "used": False})

        for lbl, target_key in [("w", "w"), ("d", "d"), ("l", "l")]:
            current = {"w": w, "d": d, "l": l}[target_key]
            if current is not None or lbl not in labels_by_type:
                continue
            label_y = labels_by_type[lbl][0]
            best_digit, best_dist, best_idx = None, float("inf"), None
            for idx, item in enumerate(digits_with_y):
                if item["used"]:
                    continue
                dist = abs(item["y"] - label_y)
                if dist < best_dist:
                    best_dist, best_digit, best_idx = dist, item["digit"], idx
            if best_digit is not None and best_dist < gray_height * proximity_threshold_ratio:
                digits_with_y[best_idx]["used"] = True
                if target_key == "w":
                    w = best_digit
                elif target_key == "d":
                    d = best_digit
                else:
                    l = best_digit
    return w, d, l

def parse_offense_from_crop(reader, image) -> Tuple[int, int, int]:
    gray, results = _preprocess_stats_crop(image, reader)
    if gray is None:
        return 0, 0, 0

    detections = _build_detections(results, max_digit=3)
    w, d, l = None, None, None

    w, d, l = _apply_strategies_1_and_2(detections, gray.shape[0])

    # Strategy 3: Position-based fallback (only fills gaps)
    if w is None or d is None or l is None:
        height = gray.shape[0]
        for det in detections:
            if det["digit"] is None:
                continue
            num = det["digit"]
            if num < 0 or num > 3:
                continue
            rel_y = det["y"] / height
            if rel_y < 0.38 and w is None:
                w = num
            elif rel_y < 0.68 and d is None:
                d = num
            elif l is None:
                l = num

    w = w if w is not None else 0
    d = d if d is not None else 0
    l = l if l is not None else 0

    cand = (w, d, l)
    if sum(cand) in (0, 1, 2, 3) and all(0 <= v <= 3 for v in cand):
        return cand
    return _closest_valid_offense(cand)


def parse_defense_from_crop(reader, image) -> Tuple[int, int, int]:
    gray, results = _preprocess_stats_crop(image, reader)
    if gray is None:
        return 0, 0, 0

    detections = _build_detections(results, max_digit=12)
    w, d, l = _apply_strategies_1_and_2(detections, gray.shape[0])

    return (w or 0, d or 0, l or 0)


# ============================================================================
# NAME MATCHING & RESOLUTION
# ============================================================================

def build_roster_matcher(
    roster: List[str], overrides: Optional[Dict[str, str]] = None
) -> Dict[str, str]:
    matcher: Dict[str, str] = {}
    for name in roster:
        matcher[canonical_key(name)] = name
        matcher[canonical_key(base_roster_name(name))] = name

    # Overrides applied last so they take priority over normal entries
    if overrides:
        for key, name in overrides.items():
            matcher[key] = name

    return matcher


def resolve_name_candidates(
    raw_candidates: List[str],
    roster: List[str],
    matcher: Dict[str, str],
    aliases_existing: Dict[str, str],
    aliases_working: Dict[str, str],
    aliases_created_this_run: Dict[str, str],
    ctx: OCRContext,
    roster_normalized_cache: Optional[Dict[str, str]] = None,
) -> Tuple[str, str, str]:
    cleaned_candidates: List[str] = []
    for candidate in raw_candidates:
        candidate = cleaned_text(candidate)
        if candidate and candidate not in cleaned_candidates:
            cleaned_candidates.append(candidate)
    if not cleaned_candidates:
        raise RuntimeError("No OCR name candidates found for row.")

    for candidate in cleaned_candidates:
        direct = try_direct_name_match(candidate, roster, matcher, aliases_working)
        if direct is not None:
            matched_name, note = direct
            return candidate, matched_name, note

    primary_raw = cleaned_candidates[0]
    roster_normalized = roster_normalized_cache if roster_normalized_cache else \
        {name: normalize_name_for_matching(name) for name in roster}
    primary_normalized = normalize_name_for_matching(primary_raw)
    scored: List[Tuple[float, str]] = []
    for name in roster:
        if not primary_normalized or not roster_normalized[name]:
            scored.append((0.0, name))
        else:
            scored.append((SequenceMatcher(None, primary_normalized, roster_normalized[name]).ratio(), name))
    scored.sort(reverse=True, key=lambda item: item[0])
    suggestions = [name for _, name in scored[:8]]

    if scored and scored[0][0] >= 0.86:
        matched = scored[0][1]
        note = "Matched to removed-tag name" if base_roster_name(matched) != matched else "Auto fuzzy match"
        aliases_working[primary_raw] = matched
        if aliases_existing.get(primary_raw) != matched:
            aliases_created_this_run[primary_raw] = matched
        return primary_raw, matched, note

    matched = prompt_manual_name(primary_raw, roster, suggestions, ctx)
    aliases_working[primary_raw] = matched
    if aliases_existing.get(primary_raw) != matched:
        aliases_created_this_run[primary_raw] = matched
    note = "Matched to removed-tag name" if base_roster_name(matched) != matched else "Manual alias fix"
    return primary_raw, matched, note


def prompt_manual_name(raw_name: str, roster: List[str], suggestions: List[str], ctx: OCRContext) -> str:
    print(f"\nUnmatched OCR name on screenshot {ctx.screenshot}, row {ctx.order}")
    print(f"OCR candidates: {ctx.name_candidates}")
    print(f"Selected raw OCR name: {raw_name}")
    print(f"Offense parsed: {ctx.off_w} / {ctx.off_d} / {ctx.off_l}")
    print(f"Defense parsed: {ctx.def_w} / {ctx.def_d} / {ctx.def_l}")
    maybe_offer_debug_export(ctx)
    if suggestions:
        print("Suggestions:")
        for idx, suggestion in enumerate(suggestions, start=1):
            print(f"  {idx}. {suggestion}")
    print("Type the exact workbook name, a suggestion number, '0' to skip, or leave blank to stop.")
    while True:
        answer = input("> ").strip()
        if answer == "":
            raise RuntimeError(f"Name clarification required for: {raw_name}")
        if answer == "0" or answer.lower() == "skip":
            return "__SKIP__"  # <--- ADDED THIS
        if answer.isdigit():
            choice = int(answer)
            if 1 <= choice <= len(suggestions):
                return suggestions[choice - 1]
            print("Invalid suggestion number.")
            continue
        if answer in roster:
            return answer
        print("That exact name is not in the workbook roster. Try again.")


def merge_duplicate_rows(
    rows: List[OCRRow], contexts: Dict[Tuple[str, int], OCRContext]
) -> List[OCRRow]:
    merged: Dict[str, OCRRow] = {}
    ordered: List[OCRRow] = []
    for row in rows:
        existing = merged.get(row.matched_name)
        if existing is None:
            merged[row.matched_name] = row
            ordered.append(row)
            continue

        same = (
            existing.off_w == row.off_w and existing.off_d == row.off_d and existing.off_l == row.off_l
            and existing.def_w == row.def_w and existing.def_d == row.def_d and existing.def_l == row.def_l
        )
        if same:
            continue

        print(f"\nConflicting duplicate rows for {row.matched_name}:")
        print(f"1. {existing.result_string}  [{existing.screenshot}]")
        print(f"2. {row.result_string}  [{row.screenshot}]")
        ctx_existing = contexts.get((existing.screenshot, existing.order))
        ctx_new = contexts.get((row.screenshot, row.order))
        if ctx_existing:
            maybe_offer_debug_export(ctx_existing)
        if ctx_new:
            maybe_offer_debug_export(ctx_new)

        while True:
            answer = input("Choose row to keep (1/2) or 3 to stop run: ").strip()
            if answer == "1":
                break
            if answer == "2":
                merged[row.matched_name] = row
                idx = ordered.index(existing)
                ordered[idx] = row
                break
            if answer == "3":
                raise RuntimeError("Run stopped by user during duplicate conflict resolution.")
            print("Invalid choice.")
    return ordered


def extract_rows_from_folder(
    folder: Path,
    folder_name: str,
    roster: List[str],
    aliases_existing: Dict[str, str],
    aliases_working: Dict[str, str],
    aliases_created_this_run: Dict[str, str],
    overrides: Dict[str, str], 
) -> Tuple[List[OCRRow], Dict[Tuple[str, int], OCRContext]]:
    if not folder.exists():
        raise FileNotFoundError(f"Folder not found: {folder}")
    pngs = sorted(
        [p for p in folder.glob("*.png") if p.is_file()], key=lambda p: p.name.lower()
    )
    if not pngs:
        raise RuntimeError(f"No PNG files found in {folder}")

    matcher = build_roster_matcher(roster, overrides)  # <--- PASS OVERRIDES TO MATCHER BUILDER
    stats_reader = get_reader("stats")
    
    # Pre-load all OCR models into memory to avoid mid-run stuttering
    print("Loading OCR models...", end=" ", flush=True)
    get_reader("name_primary")
    get_reader("name_ja")
    get_reader("name_ko")
    print("Done.")
    roster_normalized_cache = {name: normalize_name_for_matching(name) for name in roster}
    extracted: List[OCRRow] = []
    contexts: Dict[Tuple[str, int], OCRContext] = {}
    order = 1

    for png in pngs:
        img = cv2.imread(str(png))
        if img is None:
            print(f"Warning: could not read image {png.name}, skipping.")
            continue

        centers = detect_rank_centers(img)
        screenshot_start_count = len(extracted)

        for row_slot, (name_center, stat_center) in enumerate(centers, start=1):
            (name_y1, name_y2), (stats_y1, stats_y2), _ = crop_bounds(img, name_center, stat_center, row_slot, len(centers))

            # Perspective correction: Rows 3 and 4 curve inward on mobile screens.
            narrow_offset = 15 if row_slot in (3, 4) else 0

            name_crop = safe_crop(img, NAME_CROP_X1, name_y1, NAME_CROP_X2, name_y2)
            off_crop = safe_crop(img, OFF_CROP_X1, stats_y1, OFF_CROP_X2 - narrow_offset, stats_y2)
            def_crop = safe_crop(img, DEF_CROP_X1, stats_y1, DEF_CROP_X2 - narrow_offset, stats_y2)
            

            name_candidates = parse_name_candidates_from_crop(name_crop, roster, matcher, aliases_working)
            if not name_candidates:
                continue

            off_w, off_d, off_l = parse_offense_from_crop(stats_reader, off_crop)
            def_w, def_d, def_l = parse_defense_from_crop(stats_reader, def_crop)

            ctx = OCRContext(
                folder_name=folder_name,
                screenshot=png.name,
                order=order,
                center_y=name_center,
                row_slot=row_slot,
                total_slots=len(centers),
                name_candidates=name_candidates,
                off_w=off_w, off_d=off_d, off_l=off_l,
                def_w=def_w, def_d=def_d, def_l=def_l,
                name_crop=name_crop, off_crop=off_crop, def_crop=def_crop,
            )
            raw_name, matched_name, note = resolve_name_candidates(
                name_candidates, roster, matcher,
                aliases_existing, aliases_working, aliases_created_this_run, ctx,
                roster_normalized_cache,  # Fixed: no trailing equals sign
            )
            # Skip row if user typed '0' to ignore garbage OCR (like "S.Rank")
            if matched_name == "__SKIP__":
                continue
            row = OCRRow(
                screenshot=png.name, order=order, raw_name=raw_name, matched_name=matched_name,
                off_w=off_w, off_d=off_d, off_l=off_l,
                def_w=def_w, def_d=def_d, def_l=def_l, note=note,
            )
            extracted.append(row)
            contexts[(png.name, order)] = ctx
            order += 1

        if len(extracted) == screenshot_start_count:
            print(f"Warning: no rows extracted from {png.name}, saving screenshot-level debug.")
            save_screenshot_debug(folder_name, png.name, img, "No rows extracted")

    if not extracted:
        raise RuntimeError("No rows were extracted from the screenshots.")
    return merge_duplicate_rows(extracted, contexts), contexts

# ============================================================================
# DEBUG VISUALIZATION (Overlays & Crops)
# ============================================================================

def sanitize_path_component(text: str) -> str:
    text = RE_PATH_INVALID.sub("_", text)
    text = RE_WHITESPACE.sub("_", text.strip())
    return text or "unknown"

def _draw_base_overlay(img, centers):
    """Draw center lines and labels on overlay."""
    overlay = img.copy()
    for idx, (name_y, stat_y) in enumerate(centers, start=1):
        cv2.line(overlay, (0, stat_y), (overlay.shape[1] - 1, stat_y), (0, 255, 255), 1)
        cv2.putText(overlay, f"grid_{idx}={stat_y}", (20, max(20, stat_y - 6)),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2, cv2.LINE_AA)
    return overlay

def save_row_overlay(ctx: OCRContext) -> None:
    folder = DEBUG_ROOT / ctx.folder_name / sanitize_path_component(Path(ctx.screenshot).stem)
    folder.mkdir(parents=True, exist_ok=True)

    screenshot_path = SCREENSHOTS_ROOT / ctx.folder_name / ctx.screenshot
    img = cv2.imread(str(screenshot_path))
    if img is None:
        return

    centers = deterministic_row_centers(img.shape[0])
    overlay = _draw_base_overlay(img, centers)

    # Extract the specific Y coordinates for this row
    name_y, stat_y = centers[ctx.row_slot - 1]
    
    (name_y1, name_y2), (stats_y1, stats_y2), _ = crop_bounds(img, name_y, stat_y, ctx.row_slot, ctx.total_slots)
    
    cv2.rectangle(overlay, (NAME_CROP_X1, name_y1), (NAME_CROP_X2, name_y2), (0, 255, 0), 2)
    cv2.rectangle(overlay, (OFF_CROP_X1, stats_y1), (OFF_CROP_X2, stats_y2), (255, 0, 0), 2)
    cv2.rectangle(overlay, (DEF_CROP_X1, stats_y1), (DEF_CROP_X2, stats_y2), (0, 0, 255), 2)
    cv2.line(overlay, (0, stat_y), (overlay.shape[1] - 1, stat_y), (0, 255, 255), 2)
    cv2.putText(
        overlay, f"row={ctx.order} center_y={stat_y}",
        (20, max(25, stat_y - 10)),
        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2, cv2.LINE_AA,
    )
    cv2.imwrite(str(folder / f"row_{ctx.order:02d}_overlay.png"), overlay)


def save_screenshot_debug(folder_name: str, screenshot: str, img, reason: str) -> None:
    folder = DEBUG_ROOT / folder_name / sanitize_path_component(Path(screenshot).stem)
    folder.mkdir(parents=True, exist_ok=True)

    overlay = img.copy()
    centers = deterministic_row_centers(img.shape[0])
    for idx, (name_y, stat_y) in enumerate(centers, start=1):
        (name_y1, name_y2), (stats_y1, stats_y2), _ = crop_bounds(img, name_y, stat_y, idx, len(centers))
        cv2.line(overlay, (0, stat_y), (overlay.shape[1] - 1, stat_y), (0, 255, 255), 1)
        cv2.rectangle(overlay, (NAME_CROP_X1, name_y1), (NAME_CROP_X2, name_y2), (0, 255, 0), 2)
        cv2.rectangle(overlay, (OFF_CROP_X1, stats_y1), (OFF_CROP_X2, stats_y2), (255, 0, 0), 2)
        cv2.rectangle(overlay, (DEF_CROP_X1, stats_y1), (DEF_CROP_X2, stats_y2), (0, 0, 255), 2)
        cv2.putText(
            overlay, f"grid_{idx}={stat_y}", (20, max(20, stat_y - 6)),
            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2, cv2.LINE_AA,
        )

    cv2.putText(overlay, reason, (20, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2, cv2.LINE_AA)
    cv2.imwrite(str(folder / "screenshot_debug_overlay.png"), overlay)
    print(f"Saved screenshot debug to: {folder}")


def save_debug_crops(ctx: OCRContext) -> None:
    folder = DEBUG_ROOT / ctx.folder_name / sanitize_path_component(Path(ctx.screenshot).stem)
    folder.mkdir(parents=True, exist_ok=True)
    prefix = f"row_{ctx.order:02d}"
    if ctx.name_crop is not None:
        cv2.imwrite(str(folder / f"{prefix}_name.png"), ctx.name_crop)
    if ctx.off_crop is not None:
        cv2.imwrite(str(folder / f"{prefix}_offense.png"), ctx.off_crop)
    if ctx.def_crop is not None:
        cv2.imwrite(str(folder / f"{prefix}_defense.png"), ctx.def_crop)
    save_row_overlay(ctx)
    print(f"Saved debug crops to: {folder}")


def maybe_offer_debug_export(ctx: OCRContext) -> None:
    if ctx.debug_export_enabled:
        save_debug_crops(ctx)
        return
    if ctx.debug_export_asked:
        return
    ctx.debug_export_asked = True
    if ask_yes_no("Save debug crops for troubleshooting?", default=False):
        ctx.debug_export_enabled = True
        save_debug_crops(ctx)
        
def cleanup_debug_images(root: Path = DEBUG_ROOT) -> None:
    """Deletes the entire debug output folder and all saved crops/overlays."""
    if not root.exists():
        return
    try:
        shutil.rmtree(root)
        print(f"Cleaned up previous debug images from: {root}")
    except Exception as exc:
        print(f"Failed to delete debug images: {exc}")

# ============================================================================
# DATA AGGREGATION & OUTPUT
# ============================================================================

def build_normalized_table(
    roster: List[str], rows: List[OCRRow]
) -> List[Dict[str, object]]:
    row_map = {row.matched_name: row for row in rows}
    normalized: List[Dict[str, object]] = []
    for name in roster:
        row = row_map.get(name)
        if row is None:
            normalized.append({
                "Name": name, "Active flag": "", "Tokens Used": "",
                "Offense W": "", "Offense D": "", "Offense L": "",
                "Defense W": "", "Defense D": "", "Defense L": "",
                "Notes": "Inactive (not found in screenshots)",
            })
            continue

        normalized.append({
            "Name": name, "Active flag": 1,
            "Tokens Used": row.off_w + row.off_d + row.off_l,
            "Offense W": row.off_w, "Offense D": row.off_d, "Offense L": row.off_l,
            "Defense W": row.def_w, "Defense D": row.def_d, "Defense L": row.def_l,
            "Notes": row.note,
        })
    return normalized


def print_results_output_terminal(extracted: List[OCRRow]) -> None:
    print("\nResults Output")
    for row in extracted:
        print(row.result_string)


def print_normalized_table_terminal(normalized: List[Dict[str, object]]) -> None:
    print("\nNormalized Table")
    headers = PROCESSING_NORMALIZED_HEADER
    print(" | ".join(headers))
    for row in normalized:
        print(" | ".join(str(row[h]) for h in headers))


def count_inactive(normalized: List[Dict[str, object]]) -> int:
    return sum(1 for row in normalized if row["Active flag"] == "")


def print_run_summary(
    extracted: List[OCRRow], normalized: List[Dict[str, object]],
    aliases_created_this_run: Dict[str, str],
) -> None:
    print("\nSummary")
    print(f"Extracted rows: {len(extracted)}")
    print(f"Unique matched players: {len(extracted)}")
    print(f"Inactive roster members: {count_inactive(normalized)}")
    print(f"New aliases this run: {len(aliases_created_this_run)}")

def live_capture_loop(temp_dir: Path, stop_event: threading.Event) -> int:
    """
    Smart background thread that waits for the user to stop scrolling,
    holds still for 2 seconds, captures 1 frame, and then ignores the screen 
    for 5 seconds to let the user scroll to the next 4 names.
    """
    sct = mss.MSS()
    monitor = sct.monitors[1]
    
    last_frame = None
    saved_count = 0
    
    # Timing controls
    MIN_STILL_TIME = 2.0  # Wait 2 seconds of perfect stillness to capture
    COOLDOWN_TIME = 5.0   # Ignore everything for 5 seconds after a capture
    
    last_change_time = time.time()
    last_save_time = 0.0   # Start at 0 so the first frame can trigger immediately

    print(f"[*] Live capture started. Smart-mode enabled.")
    print("[*] Stop on 4 names for 2 seconds to capture. Cooldown: 5 seconds.")
    print("[*] Come back here and press ENTER when done.\n")

    while not stop_event.is_set():
        current_time = time.time()
        
        img = np.array(sct.grab(monitor))
        frame_bgr = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        is_still = False
        if last_frame is not None:
            diff = cv2.absdiff(frame_bgr, last_frame)
            # Tightened threshold to 0.5% to prevent catching slow UI animations
            if np.count_nonzero(diff > 20) < (frame_bgr.size * 0.005):
                is_still = True

        if not is_still:
            last_change_time = current_time  # Reset timer because screen is moving
            last_frame = frame_bgr
            time.sleep(0.1) # Fast loop while scrolling
            continue

        # Screen IS still
        # Check if we are in the cooldown phase (just captured a frame recently)
        if current_time - last_save_time < COOLDOWN_TIME:
            last_frame = frame_bgr
            time.sleep(0.2)
            continue

        # Check if the screen has been still long enough to trigger a capture
        if current_time - last_change_time >= MIN_STILL_TIME:
            out_path = temp_dir / f"live_{saved_count:04d}.png"
            cv2.imwrite(str(out_path), frame_bgr)
            saved_count += 1
            last_save_time = current_time  # Start the 5-second cooldown
            print("*", end="", flush=True) # Print a star for a saved frame
            
        last_frame = frame_bgr
        time.sleep(0.2) # Don't hammer the CPU while waiting

    sct.close()
    print(f"\n[*] Capture stopped. Saved {saved_count} frames.")
    return saved_count


