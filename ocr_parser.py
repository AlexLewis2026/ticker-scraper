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

# Submission timestamp — new blotter has Ex.Time then Sub.Time back-to-back.
# Groups capture (HH:MM:SS, TZ) so we can store the value.
_SUB_TS_RE = re.compile(r"^(\d{2}:\d{2}:\d{2})\s+([A-Z]{2,5})\s+")

# Region tokens that appear between CC and Strip in the new blotter format.
# "none" appears when the CC column is blank (e.g. Bal Month diff rows).
_REGION_TOKENS = frozenset({"europe", "singapore", "asia", "americas", "none"})

# Hyphenated strip range, e.g. Mar27-Jun27, Nov26-Oct27 (always a single trade).
_STRIP_RANGE_RE = re.compile(r"^[A-Za-z]{3}\d{2}-[A-Za-z]{3}\d{2}$")

# Qty stuck to first strip word, e.g. "4Bal" "5Aug26" "1Q3"
QTY_STUCK_RE = re.compile(r"^(\d+)([A-Za-z].*)$")

# Strip token patterns (tokens that belong to the strip field)
_STRIP_TOKEN_RE = re.compile(
    r"^("
    r"\d{2}$"                    # bare 2-digit year: 26, 27
    r"|[A-Za-z]+\d{2}"          # MonthYY: Jul26, Aug26, Dec27
    r"|[A-Za-z]+\d{2}/[A-Za-z]*\d{2}"  # spread: Jul26/Aug26
    r"|Bal$"                     # "Bal" (always followed by "Month")
    r"|Month(?:-[A-Za-z]+)?"     # "Month", "Month-ND", etc.
    r"|Cal$"                     # "Cal" (always followed by bare year: Cal 27)
    r"|Q[1-4]$"                  # quarter: Q1 Q2 Q3 Q4
    r")$"
)


# ── Hub → CC fallback map ──────────────────────────────────────────────────────
# Used when the CC column is blank in the blotter (common on Bal Month and
# spread-diff rows).  Key = canonical hub name (lowercase for matching).
# Add new entries here whenever a new hub/CC pair is encountered.
HUB_CC_MAP: dict[str, str] = {
    # ── Naphtha CIF NWE ───────────────────────────────────────────────────────
    "naphtha cif nwe cg":                                              "NEC",
    "naphtha cif nwe cg mini":                                         "NAM",
    # Crack / Brent spreads
    "naphtha cif nwe cg (platts)/brent 1st line":                      "NBB",
    "naphtha cif nwe cg/brent 1st line":                               "NOB",

    # ── Naphtha C&F Japan  (Balmo = NJD) ─────────────────────────────────────
    "naphtha c&f japan cg":                                            "NJC",
    "naphtha c&f japan cg balmo":                                      "NJD",

    # ── Sing Mogas ────────────────────────────────────────────────────────────
    "sing mogas 92 unl (platts)":                                      "SMT",
    "sing mogas 92 unl (platts) mini":                                 "SMV",
    "sing mogas 95 unl (platts)":                                      "SMF",
    # Diff / crack
    "sing mogas 92 unl (platts)/brent 1st line":                       "STB",
    "sing mogas unl 95/92 (platts)":                                   "SMD",

    # ── Argus Eurobob ─────────────────────────────────────────────────────────
    "argus eurobob oxy fob rdam bg":                                   "AEO",
    "argus eurobob oxy fob rdam bg mini":                              "AOM",
    # Crack / Brent spread  (AEB = crack futures, NOT balmo)
    "argus eurobob oxy fob rdam bg/brent 1st line (bbl)":              "AEB",

    # ── RBOB / Gasoline diff ──────────────────────────────────────────────────
    "rbob 1st line/argus eurobob oxy fob rdam bg mini":                "GDQ",

    # ── Far East / other naphtha ──────────────────────────────────────────────
    "far east":                                                        "AFE",
    "far east/cif ara":                                                "EGD",

    # ── Saudi CP ──────────────────────────────────────────────────────────────
    "saudi cp":                                                        "SCP",

    # ── MT propane ────────────────────────────────────────────────────────────
    "mt b-etr":                                                        "PRL",
    "mt b-ent":                                                        "PRN",

    # ── Conway ────────────────────────────────────────────────────────────────
    "conway":                                                          "PRC",
}


def _hub_to_cc(hub: str) -> str:
    """
    Derive a CC from a hub string when the CC column was blank.
    Strips OCR junk and takes the first segment before '/' so that
    spread-diff hubs like 'Far East/Far East' resolve correctly.
    """
    # Clean OCR junk from the start
    clean = hub.lstrip("=©®@•~-_ ").strip()
    # Try full hub name first (e.g. "Sing Mogas 92 Unl (Platts)/Brent 1st Line" → STB)
    full = clean.lower()
    if full in HUB_CC_MAP:
        return HUB_CC_MAP[full]
    # Fall back to first segment before "/" (spread-diff hubs like "Hub1/Hub2")
    first = clean.split("/")[0].strip().lower()
    return HUB_CC_MAP.get(first, "")


def _is_strip_token(tok: str) -> bool:
    """Return True if this token is part of a strip/spread designation."""
    if "/" in tok and re.search(r"\d{2}$", tok):   # spread: Jul26/Aug26
        return True
    if _STRIP_RANGE_RE.match(tok):                  # range:  Mar27-Jun27
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

    # 1b. Capture and skip submission timestamp when present (new blotter format).
    sub_m = _SUB_TS_RE.match(rest)
    if sub_m:
        sub_time = sub_m.group(1) + " " + sub_m.group(2)
        rest = rest[sub_m.end():]
    else:
        sub_time = ""

    # 2. Split on the bullet character (♦ • ◆ ●) that separates Price from Hub+TT.
    #
    #   New blotter:  Qty CC Region Strip [Strategies] Price ♦ Hub TT
    #   Old blotter:  CC Qty Strip Hub Price ♦ TT
    #
    # The bullet is always a single non-word, non-space character surrounded by
    # whitespace. We find it, then derive price from the LEFT and hub+TT from
    # the RIGHT.  Old format: right = just "TT" (hub lives on the left).
    _BULLET_PAT = re.compile(r"\s+([^\w\s])\s+")
    bm = _BULLET_PAT.search(rest)
    if bm:
        left  = rest[: bm.start()].rstrip()
        right = rest[bm.end() :].strip()

        # Right side: last token is TT code; everything before it is hub (new format).
        right_toks = right.split()
        if right_toks and re.match(r"^[A-Za-z]{2,}$", right_toks[-1]):
            tt_code    = right_toks[-1].lower()
            hub_right  = " ".join(right_toks[:-1])
        else:
            return None

        # Price = last number on the left side.
        pm = re.search(r"(-?\d+(?:\.\d+)?)\s*$", left)
        if not pm:
            return None
        try:
            price = float(pm.group(1))
        except ValueError:
            return None
        cancelled = tt_code in _CANCEL_TT
        rest      = left[: pm.start()].strip()
    else:
        # Fallback: old PRICE_RE (no bullet found — unusual but safe).
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
        hub_right = ""

    tokens = rest.split()
    if not tokens:
        return None

    # 3 & 4. Extract CC and Qty.
    #
    # Three column layouts are supported:
    #
    #   Layout A (newest — Product prefix):
    #     Product...  CC  Qty  [Strategies]  Strip
    #     e.g. "Gasoline Futures Spr  AEO  10  spread  Sep26/Oct26"
    #     Product tokens have mixed case; CC is the first pure-uppercase token
    #     immediately followed by a pure-digit token.
    #
    #   Layout B (mid-generation — Qty leads, no Product):
    #     Qty  CC  [Region]  Strip
    #     e.g. "50 AEO Europe Sep26"
    #
    #   Layout C (oldest — CC leads):
    #     CC  Qty  Strip  Hub
    #     e.g. "STB 50 Jul26 Sing Mogas…"
    #
    # Detection order: A → C → B (A and C both start with a (CC, digit) pair
    # but C has CC as tokens[0]; A has product words before CC).

    cc = ""
    qty_str = ""
    product = ""           # Product field text (Layout A only)
    strip_tokens: list[str] = []

    # Primary scan: find first i where tokens[i] is CC followed by a qty token.
    # Qty token is either a plain digit ("50") or digit-stuck-to-strip ("4Bal",
    # "1Q3", "5NJC").  Covers Layout A (Product prefix) and Layout C (CC-first).
    found_cc_qty = False
    for i in range(len(tokens) - 1):
        if not CC_PATTERN.match(tokens[i]):
            continue
        nxt = tokens[i + 1]
        if nxt.isdigit():
            product = " ".join(tokens[:i])   # everything before CC = Product field
            cc, qty_str = tokens[i], nxt
            tokens = tokens[i + 2:]
            found_cc_qty = True
            break
        sm2 = QTY_STUCK_RE.match(nxt)
        if sm2:
            product = " ".join(tokens[:i])
            cc, qty_str = tokens[i], sm2.group(1)
            remainder   = sm2.group(2)
            tokens      = tokens[i + 2:]
            if CC_PATTERN.match(remainder):
                pass   # e.g. "5NJC" → CC already found
            else:
                strip_tokens = [remainder]
            found_cc_qty = True
            break

    if not found_cc_qty:
        # Layout B fallback: qty leads (e.g. "50 AEO Europe Sep26")
        first = tokens[0]
        sm = QTY_STUCK_RE.match(first)
        if sm:
            qty_str   = sm.group(1)
            remainder = sm.group(2)
            if CC_PATTERN.match(remainder):
                cc     = remainder
                tokens = tokens[1:]
            else:
                strip_tokens = [remainder]
                tokens = tokens[1:]
                if tokens and CC_PATTERN.match(tokens[0]):
                    cc     = tokens[0]
                    tokens = tokens[1:]
        elif first.isdigit():
            qty_str = first
            tokens  = tokens[1:]
            if tokens and CC_PATTERN.match(tokens[0]):
                cc     = tokens[0]
                tokens = tokens[1:]
        else:
            # Layout A with blank CC: Product tokens precede a bare digit qty.
            # e.g. "Naphtha Futures Spr 20 spread Bal Month/Jul26"
            # Scan for first pure-digit token; CC stays blank.
            found_digit = False
            for i, tok in enumerate(tokens):
                if tok.isdigit():
                    qty_str = tok
                    tokens  = tokens[i + 1:]
                    found_digit = True
                    break
            if not found_digit:
                return None

    try:
        qty = int(qty_str)
    except ValueError:
        return None

    # Skip legacy Region token if present (old mid-gen format only).
    if tokens and tokens[0].lower() in _REGION_TOKENS:
        tokens = tokens[1:]

    # 5. Strategies column: "spread" or "fly" — appears before the strip.
    strategy = ""
    if tokens and tokens[0].lower() in ("spread", "fly"):
        strategy = tokens[0].lower()
        tokens   = tokens[1:]

    # 6. Strip: consume tokens that match strip patterns.
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

    # Hub: new format provides it after the bullet; old format it's left tokens.
    hub = hub_right if hub_right else " ".join(tokens[hub_start:]).strip()

    # 7. Fill blank CC from hub when the blotter omits it (e.g. Bal Month rows)
    if not cc and hub:
        cc = _hub_to_cc(hub)

    # 8. Diff / fly row detection:
    #    "spread" strategy → spread diff row.
    #    "fly"    strategy → butterfly diff row.
    #    Heuristic fallback for old format (strip has "/" and price is small).
    is_diff = bool(strategy) or ("/" in strip and abs(price) < 100)

    return {
        "timestamp":   timestamp,
        "sub_time":    sub_time,
        "cc":          cc,
        "qty":         qty,
        "strip":       strip,
        "hub":         hub,
        "price":       price,
        "is_diff_row": is_diff,
        "strategy":    strategy,     # "", "spread", or "fly"
        "product":     product,      # Product field text, e.g. "Gasoline Futures Spr"
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
