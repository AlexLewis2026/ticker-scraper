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

# Valid CC codes: 2-4 uppercase letters (expand as needed)
CC_PATTERN = re.compile(r"^[A-Z]{2,5}$")

# Timestamp pattern: HH:MM:SS BST (digits may be misread, so just grab what's there)
TS_RE = re.compile(r"^(\d{2}:\d{2}:\d{2})\s+BST\s+(.*)", re.DOTALL)

# Price at end of line: optional sign, digits, optional decimal
PRICE_RE = re.compile(r"(-?\d+(?:\.\d+)?)\s+\S*\s*BLK\s*$", re.IGNORECASE)

# Qty stuck to start of strip word, e.g. "4Bal" "5Aug26" "1Q3"
QTY_STUCK_RE = re.compile(r"^(\d+)([A-Za-z].*)$")


# ── image → raw text ───────────────────────────────────────────────────────────

def _ocr_image(image_path: str) -> str:
    import pytesseract
    from PIL import Image

    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

    img = Image.open(image_path)
    # 2× upscale dramatically improves accuracy on UI screenshots
    img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    return pytesseract.image_to_string(img, config="--psm 6")


# ── line parser ────────────────────────────────────────────────────────────────

def _parse_line(line: str) -> dict | None:
    """
    Parse one OCR line into a trade row dict.
    Returns None if the line doesn't look like a data row.
    """
    line = line.strip()
    if not line:
        return None

    # 1. Extract timestamp
    m = TS_RE.match(line)
    if not m:
        return None
    timestamp = m.group(1) + " BST"
    rest = m.group(2).strip()

    # 2. Extract price + BLK from the end
    pm = PRICE_RE.search(rest)
    if not pm:
        return None
    price_str = pm.group(1)
    try:
        price = float(price_str)
    except ValueError:
        return None
    rest = rest[: pm.start()].strip()

    # 3. Split into tokens
    tokens = rest.split()
    if not tokens:
        return None

    # 4. Extract CC (optional leading 2-5 uppercase token)
    cc = ""
    if tokens and CC_PATTERN.match(tokens[0]):
        cc = tokens[0]
        tokens = tokens[1:]

    if not tokens:
        return None

    # 5. Find where the Hub starts (first token in HUB_START_WORDS, after stripping
    #    OCR junk like leading = or © characters)
    hub_start_idx = None
    for i, tok in enumerate(tokens):
        clean = tok.lstrip("=©®@•~-_")
        if clean in HUB_START_WORDS:
            hub_start_idx = i
            break

    if hub_start_idx is None:
        # Can't identify hub — skip
        return None

    qty_strip_tokens = tokens[:hub_start_idx]
    hub_tokens       = tokens[hub_start_idx:]

    # Clean OCR junk from hub start token
    hub_tokens[0] = hub_tokens[0].lstrip("=©®@•~-_")
    hub = " ".join(hub_tokens)

    # 6. Parse qty + strip from qty_strip_tokens
    if not qty_strip_tokens:
        return None

    first = qty_strip_tokens[0]

    # qty might be stuck to first strip word: "4Bal", "5Aug26"
    sm = QTY_STUCK_RE.match(first)
    if sm:
        qty_str   = sm.group(1)
        strip_first = sm.group(2)
        strip_tokens = [strip_first] + qty_strip_tokens[1:]
    else:
        # First token should be pure digits
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

    # 7. Detect diff row: strip contains "/" (spread legs) AND price is small
    #    The "/" in strip means it's a spread identifier row (e.g. "Mar27/Apr27")
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
    """
    raw_text = _ocr_image(image_path)
    rows = []
    for line in raw_text.splitlines():
        row = _parse_line(line)
        if row is not None:
            rows.append(row)
    return rows


# ── CLI test ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys, json
    path = sys.argv[1] if len(sys.argv) > 1 else "Screenshot.png"
    rows = parse_image_local(path)
    print(f"{len(rows)} rows parsed:\n")
    print(json.dumps(rows, indent=2))
