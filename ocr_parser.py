"""
Local OCR parser — replaces Claude Vision with Tesseract.
Reads a blotter screenshot and returns the same list-of-dicts format
that parse_image_with_claude() returned.

Product identity comes from CC, not hub. Hub is not parsed.
"""

import re
from pathlib import Path

# ── constants ──────────────────────────────────────────────────────────────────

TESSERACT_CMD = r"C:\Users\AlexLewis\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# Valid CC codes: 2-5 uppercase letters
CC_PATTERN = re.compile(r"^[A-Z]{2,5}$")

# Timestamp: HH:MM:SS + any timezone abbreviation (BST, GMT, UTC, CET …)
TS_RE = re.compile(r"^(\d{2}:\d{2}:\d{2})\s+([A-Z]{2,5})\s+(.*)", re.DOTALL)

# Price at end of line: last number before optional bullet + TT code.
# TT code may be a short code (BLK, AGR) or the word "cancelled".
PRICE_RE = re.compile(
    r"(-?\d+(?:\.\d+)?)"       # price
    r"\s+\S*\s*"               # optional bullet / junk char
    r"([A-Za-z]{2,})"         # TT code: BLK, cancelled, etc.
    r"[^A-Za-z\d]*$",          # allow any trailing punctuation/symbols (`, ', etc.)
    re.IGNORECASE
)

# TT values that indicate a cancelled trade (case-insensitive match)
_CANCEL_TT = {"cancelled", "cancel", "cxl"}

# Qty stuck to first strip word, e.g. "4Bal" "5Aug26" "1Q3"
QTY_STUCK_RE = re.compile(r"^(\d+)([A-Za-z].*)$")

# Strip token patterns (tokens that belong to the strip field)
_STRIP_TOKEN_RE = re.compile(
    r"^("
    r"\d{2}$"                    # bare 2-digit year: 26, 27
    r"|[A-Za-z]+\d{2}"          # MonthYY: Jul26, Aug26, Dec27
    r"|[A-Za-z]+\d{2}/[A-Za-z]*\d{2}"  # spread: Jul26/Aug26
    r"|Bal$"                     # "Bal" (always followed by "Month")
    r"|Month"                    # "Month" or "Month/MonthYY"
    r"|Q[1-4]$"                  # quarter: Q1 Q2 Q3 Q4
    r")$"
)


# ── Hub → CC fallback map ──────────────────────────────────────────────────────
# Used when the CC column is blank in the blotter (common on Bal Month and
# spread-diff rows).  Key = canonical hub name (lowercase for matching).
# Add new entries here whenever a new hub/CC pair is encountered.
HUB_CC_MAP: dict[str, str] = {
    # Naphtha CIF NWE
    "naphtha cif nwe cg":                                   "NEC",
    "naphtha cif nwe cg mini":                              "NAM",
    # Naphtha C&F Japan
    "naphtha c&f japan cg":                                 "NJC",
    # Sing Mogas
    "sing mogas 92 unl (platts)/brent 1st line":            "STB",
    # Far East (naphtha)
    "far east":                                             "AFE",
    # Saudi CP
    "saudi cp":                                             "SCP",
    # Argus Eurobob
    "argus eurobob oxy fob rdam bg":                        "AEO",
    "argus eurobob oxy fob rdam bg mini":                   "AOM",
    # MT B-ETR / MT B-ENT (propane)
    "mt b-etr":                                             "PRL",
    "mt b-ent":                                             "PRN",
    # Conway
    "conway":                                               "PRC",
    # Far East / CIF ARA (EGD)
    "far east/cif ara":                                     "EGD",
}


def _hub_to_cc(hub: str) -> str:
    """
    Derive a CC from a hub string when the CC column was blank.
    Strips OCR junk and takes the first segment before '/' so that
    spread-diff hubs like 'Far East/Far East' resolve correctly.
    """
    # Clean OCR junk from the start
    clean = hub.lstrip("=©®@•~-_ ").strip()
    # Take only the first hub segment (spread diffs have "Hub1/Hub2")
    first = clean.split("/")[0].strip().lower()
    return HUB_CC_MAP.get(first, "")


def _is_strip_token(tok: str) -> bool:
    """Return True if this token is part of a strip/spread designation."""
    # Also accept tokens that contain a slash and end with 2-digit year
    if "/" in tok and re.search(r"\d{2}$", tok):
        return True
    return bool(_STRIP_TOKEN_RE.match(tok))


# ── image → raw text ───────────────────────────────────────────────────────────

def _ocr_image(image_path: str) -> str:
    """Run Tesseract on the image and return the raw text."""
    tess_path = Path(TESSERACT_CMD)
    if not tess_path.exists():
        raise FileNotFoundError(
            f"Tesseract not found at:\n  {TESSERACT_CMD}\n"
            "Install it from https://github.com/UB-Mannheim/tesseract/wiki "
            "or update TESSERACT_CMD in ocr_parser.py."
        )

    try:
        import pytesseract
        from PIL import Image
    except ImportError as e:
        raise ImportError(
            f"Missing dependency: {e}. Run: pip install pytesseract pillow"
        ) from e

    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

    try:
        img = Image.open(image_path)
    except FileNotFoundError:
        raise FileNotFoundError(f"Image file not found: {image_path}")
    except Exception as e:
        raise ValueError(f"Cannot open image '{image_path}': {e}") from e

    img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)

    try:
        return pytesseract.image_to_string(img, config="--psm 6")
    except Exception as e:
        raise RuntimeError(f"Tesseract OCR failed: {e}") from e


# ── line parser ────────────────────────────────────────────────────────────────

def _parse_line(line: str) -> dict | None:
    """
    Parse one OCR text line into a trade row dict.
    Returns None if the line doesn't look like a blotter data row.
    """
    line = line.strip()
    if not line:
        return None

    # 1. Extract timestamp + timezone
    m = TS_RE.match(line)
    if not m:
        return None
    timestamp = m.group(1) + " " + m.group(2)
    rest = m.group(3).strip()

    # 2. Extract price + trade-type code from end
    pm = PRICE_RE.search(rest)
    if not pm:
        return None
    try:
        price = float(pm.group(1))
    except ValueError:
        return None
    tt_code   = pm.group(2).lower()
    cancelled = tt_code in _CANCEL_TT
    rest      = rest[: pm.start()].strip()

    tokens = rest.split()
    if not tokens:
        return None

    # 3 & 4. Extract CC and Qty — supports both column orders:
    #   Old format:  CC  Qty  Strip  Hub   (e.g. "STB 50 Jul26 Sing…")
    #   New format:  Qty CC   Strip  Hub   (e.g. "50 STB Jul26 Sing…")
    cc = ""

    if CC_PATTERN.match(tokens[0]):
        # Old format: CC leads
        cc     = tokens[0]
        tokens = tokens[1:]
        if not tokens:
            return None
        first = tokens[0]
    else:
        # New format: Qty leads
        first = tokens[0]

    sm = QTY_STUCK_RE.match(first)
    if sm:
        qty_str   = sm.group(1)
        remainder = sm.group(2)
        # If the stuck remainder is a CC code (all uppercase, no digits)
        # treat it as CC rather than the first strip word.
        # e.g. "5NJC" → qty=5, cc="NJC"
        # e.g. "4Bal" → qty=4, strip_first="Bal"
        if CC_PATTERN.match(remainder):
            cc           = remainder
            strip_tokens = []
        else:
            strip_tokens = [remainder]
        tokens = tokens[1:]
    else:
        if not first.isdigit():
            return None
        qty_str      = first
        strip_tokens = []
        tokens       = tokens[1:]

    try:
        qty = int(qty_str)
    except ValueError:
        return None

    # After qty, pick up CC if not already found (new format: Qty CC Strip)
    if not cc and tokens and CC_PATTERN.match(tokens[0]):
        cc     = tokens[0]
        tokens = tokens[1:]

    # 5. Strip: consume tokens that match strip patterns.
    #    Everything after the strip is the hub.
    hub_start = len(tokens)
    for i, tok in enumerate(tokens):
        if _is_strip_token(tok):
            strip_tokens.append(tok)
        else:
            hub_start = i
            break

    strip = " ".join(strip_tokens)
    if not strip:
        return None

    hub = " ".join(tokens[hub_start:]).strip()

    # 6. Fill blank CC from hub when the blotter omits it (e.g. Bal Month rows)
    if not cc and hub:
        cc = _hub_to_cc(hub)

    # 7. Diff row: strip contains "/" and price is small (spread differential)
    is_diff = "/" in strip and abs(price) < 100

    return {
        "timestamp":   timestamp,
        "cc":          cc,
        "qty":         qty,
        "strip":       strip,
        "hub":         hub,
        "price":       price,
        "is_diff_row": is_diff,
        "cancelled":   cancelled,
    }


# ── public entry point ─────────────────────────────────────────────────────────

def _join_wrapped_lines(text: str) -> list[str]:
    """
    Tesseract sometimes splits a single blotter row across two lines when the
    hub name is long.  Any line that does NOT begin with a timestamp is a
    continuation of the previous line — join it back on.
    """
    joined: list[str] = []
    current = ""
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            if current:
                joined.append(current)
                current = ""
            continue
        if TS_RE.match(line):          # new trade row starts
            if current:
                joined.append(current)
            current = line
        else:                          # continuation — append to current row
            current = (current + " " + line).strip() if current else line
    if current:
        joined.append(current)
    return joined


def parse_image_local(image_path: str) -> list[dict]:
    """
    Drop-in replacement for parse_image_with_claude().
    Returns a list of raw row dicts from the blotter screenshot.
    Raises FileNotFoundError, ValueError, or RuntimeError on hard failures.
    """
    raw_text = _ocr_image(image_path)
    rows = [r for line in _join_wrapped_lines(raw_text) if (r := _parse_line(line))]
    if not rows:
        preview = "\n".join(raw_text.splitlines()[:8]) or "(empty)"
        raise ValueError(
            f"No trade rows could be parsed from the image.\n\n"
            f"Raw OCR output (first 8 lines):\n{preview}\n\n"
            f"Common causes:\n"
            f"  • The blotter uses a strip format not yet recognised\n"
            f"  • The image is not a blotter screenshot"
        )
    return rows


def ocr_raw_text(image_path: str) -> str:
    """Return the raw Tesseract text for diagnostic purposes."""
    return _ocr_image(image_path)


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys, json
    path = sys.argv[1] if len(sys.argv) > 1 else "Screenshot.png"
    try:
        rows = parse_image_local(path)
        print(f"{len(rows)} rows parsed:\n")
        print(json.dumps(rows, indent=2))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
