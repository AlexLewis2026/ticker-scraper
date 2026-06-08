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

# Price at end of line: last number before optional junk + trade-type code
PRICE_RE = re.compile(r"(-?\d+(?:\.\d+)?)\s+\S*\s*[A-Z]{2,5}\s*$", re.IGNORECASE)

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
    rest = rest[: pm.start()].strip()

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
        qty_str      = sm.group(1)
        strip_tokens = [sm.group(2)]
        tokens       = tokens[1:]
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

    # 6. Diff row: strip contains "/" and price is small (spread differential)
    is_diff = "/" in strip and abs(price) < 100

    return {
        "timestamp":   timestamp,
        "cc":          cc,
        "qty":         qty,
        "strip":       strip,
        "hub":         hub,
        "price":       price,
        "is_diff_row": is_diff,
    }


# ── public entry point ─────────────────────────────────────────────────────────

def parse_image_local(image_path: str) -> list[dict]:
    """
    Drop-in replacement for parse_image_with_claude().
    Returns a list of raw row dicts from the blotter screenshot.
    Raises FileNotFoundError, ValueError, or RuntimeError on hard failures.
    """
    raw_text = _ocr_image(image_path)
    rows = [r for line in raw_text.splitlines() if (r := _parse_line(line))]
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
