"""
Local OCR parser — replaces Claude Vision with Tesseract.
Reads a blotter screenshot and returns the same list-of-dicts format
that parse_image_with_claude() returned.
"""

import re
from pathlib import Path

# ── known constants ────────────────────────────────────────────────────────────

TESSERACT_CMD = r"C:\Users\AlexLewis\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# Hub names always start with one of these words (after stripping OCR junk)
HUB_START_WORDS = {"Naphtha", "Sing", "Argus"}

# Valid CC codes: 2-5 uppercase letters
CC_PATTERN = re.compile(r"^[A-Z]{2,5}$")

# Timestamp: HH:MM:SS + any timezone abbreviation (BST, GMT, UTC, CET, EST …)
TS_RE = re.compile(r"^(\d{2}:\d{2}:\d{2})\s+([A-Z]{2,5})\s+(.*)", re.DOTALL)

# Price at end of line, followed by an optional bullet/junk char then any
# word (trade-type code: BLK, AGR, VOL, …) — case-insensitive
PRICE_RE = re.compile(r"(-?\d+(?:\.\d+)?)\s+\S*\s*[A-Z]{2,5}\s*$", re.IGNORECASE)

# Qty stuck to first strip word, e.g. "4Bal" "5Aug26"
QTY_STUCK_RE = re.compile(r"^(\d+)([A-Za-z].*)$")


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
        from PIL import Image, UnidentifiedImageError
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

    m = TS_RE.match(line)
    if not m:
        return None
    timestamp = m.group(1) + " " + m.group(2)   # e.g. "13:04:42 BST" or "13:04:42 GMT"
    rest = m.group(3).strip()

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

    # Optional leading CC token
    cc = ""
    if CC_PATTERN.match(tokens[0]):
        cc = tokens[0]
        tokens = tokens[1:]

    if not tokens:
        return None

    # Find hub start — first token whose cleaned form is a known hub word
    hub_start_idx = None
    for i, tok in enumerate(tokens):
        if tok.lstrip("=©®@•~-_") in HUB_START_WORDS:
            hub_start_idx = i
            break

    if hub_start_idx is None:
        return None

    qty_strip_tokens = tokens[:hub_start_idx]
    hub_tokens       = tokens[hub_start_idx:]
    hub_tokens[0]    = hub_tokens[0].lstrip("=©®@•~-_")
    hub              = " ".join(hub_tokens)

    if not qty_strip_tokens:
        return None

    first = qty_strip_tokens[0]
    sm    = QTY_STUCK_RE.match(first)
    if sm:
        qty_str      = sm.group(1)
        strip_tokens = [sm.group(2)] + qty_strip_tokens[1:]
    else:
        if not first.isdigit():
            return None
        qty_str      = first
        strip_tokens = qty_strip_tokens[1:]

    try:
        qty = int(qty_str)
    except ValueError:
        return None

    strip = " ".join(strip_tokens)
    if not strip:
        return None

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
            f"  • Timezone shown is not a recognised abbreviation\n"
            f"  • The blotter hub names are not yet in HUB_START_WORDS\n"
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
