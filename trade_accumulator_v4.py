#!/usr/bin/env python3
"""
Trade Volume Accumulator  v4
============================
Parses commodity trade blotter screengrabs via Claude Vision and accumulates
volume + pricing data throughout a trading day into a structured Excel workbook.

Scope (v4)
----------
  - Outright trades
  - Same-product spreads (two or more legs, identical CC)
  - Same-timestamp / same-CC / different-qty rows → flagged as unrelated outrights
  - Cross-product situations (different CCs at same timestamp) → each CC treated
    as its own independent trade; no cross-spread logic applied

Usage
-----
    python trade_accumulator.py screenshot.png --output trade_tally.xlsx
    python trade_accumulator.py screenshot.png --dry-run   # preview only

Environment
-----------
    ANTHROPIC_API_KEY   or pass via --api-key
"""

import argparse
import base64
import json
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════════════
#  CLAUDE VISION — raw row extraction
# ═══════════════════════════════════════════════════════════════════════════

VISION_PROMPT = """
You are parsing a commodity trade blotter screenshot.

Return every visible data row as a JSON array — one object per row.
Do NOT group rows. Do NOT interpret spreads. Just return what you see.

Each object:
  "timestamp"   : string   e.g. "16:13:11 BST"
  "cc"          : string   e.g. "APC"
  "qty"         : integer
  "strip"       : string   e.g. "Dec26"
  "hub"         : string   e.g. "CIF ARA"
  "price"       : float
  "is_diff_row" : boolean  — true if this row looks like a spread differential
                             (small absolute value, same timestamp as adjacent
                             leg rows, sometimes negative). False otherwise.

Return ONLY a JSON array. No markdown, no explanation. Ignore header rows.
"""


def parse_image_with_claude(image_path: str, api_key: str) -> list[dict]:
    import urllib.request

    with open(image_path, "rb") as f:
        img_b64 = base64.standard_b64encode(f.read()).decode()

    ext  = Path(image_path).suffix.lower()
    mime = {".png": "image/png", ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg"}.get(ext, "image/png")

    payload = json.dumps({
        "model": "claude-opus-4-5-20251101",
        "max_tokens": 8192,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image",
                 "source": {"type": "base64", "media_type": mime, "data": img_b64}},
                {"type": "text", "text": VISION_PROMPT}
            ]
        }]
    }).encode()

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "x-api-key":         api_key,
            "anthropic-version": "2023-06-01",
            "content-type":      "application/json",
        },
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())

    raw = "".join(b["text"] for b in data["content"] if b["type"] == "text")
    raw = raw.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
    return json.loads(raw)


# ═══════════════════════════════════════════════════════════════════════════
#  TAPS / MOC DETECTION
# ═══════════════════════════════════════════════════════════════════════════

# Outrights for eligible CCs before this time at qualifying price → TAPS
TAPS_CUTOFF = "09:45:00"

# Per-group price ranges and tick increments
# SMT/SMU/SMV/SMS : -0.010 to +0.020, tick 0.01
# NJC/NJD/NJM/NJB : -0.100 to +0.100, tick 0.05
TAPS_GROUPS = {
    "SMT": (-0.010, +0.020),
    "SMU": (-0.010, +0.020),
    "SMV": (-0.010, +0.020),
    "SMS": (-0.010, +0.020),
    "NJC": (-0.100, +0.100),
    "NJD": (-0.100, +0.100),
    "NJM": (-0.100, +0.100),
    "NJB": (-0.100, +0.100),
}
TAPS_CC = set(TAPS_GROUPS.keys())


def _classify_trade_type(trade: dict) -> str:
    """Return 'TAPS' if the trade meets MOC/TAPS criteria, else the original type."""
    if trade.get("trade_type") != "OUTRIGHT":
        return trade["trade_type"]
    cc = trade.get("cc", "")
    if cc not in TAPS_GROUPS:
        return "OUTRIGHT"
    time_part = trade.get("timestamp", "").split()[0]   # "HH:MM:SS"
    if time_part >= TAPS_CUTOFF:
        return "OUTRIGHT"
    legs = trade.get("legs", [])
    if not legs:
        return "OUTRIGHT"
    price = float(legs[0].get("price", 999))
    lo, hi = TAPS_GROUPS[cc]
    if lo <= price <= hi:
        return "TAPS"
    return "OUTRIGHT"


# ═══════════════════════════════════════════════════════════════════════════
#  TRADE GROUPING — same-product only
# ═══════════════════════════════════════════════════════════════════════════

def group_rows_into_trades(raw_rows: list[dict]) -> list[dict]:
    """
    Group raw blotter rows into trades.

    Rules (applied per timestamp, then per CC within that timestamp):

      1. Single leg row            → OUTRIGHT
      2. Multiple legs, same qty   → SPREAD (same-product)
      3. Multiple legs, mixed qty  → each row becomes a flagged OUTRIGHT
         (unrelated trades that happen to share a timestamp)

    Different CCs at the same timestamp are handled independently —
    each CC group goes through rules 1-3 on its own.
    """
    from itertools import groupby

    rows  = sorted(raw_rows, key=lambda r: r["timestamp"])
    trades = []

    for ts, ts_group in groupby(rows, key=lambda r: r["timestamp"]):
        ts_rows = list(ts_group)

        # Split by CC — different CCs always independent at this stage
        cc_map: dict[str, list] = {}
        for row in ts_rows:
            cc_map.setdefault(row["cc"], []).append(row)

        for cc, cc_rows in cc_map.items():
            diff_rows = [r for r in cc_rows if r.get("is_diff_row")]
            leg_rows  = [r for r in cc_rows if not r.get("is_diff_row")]

            if len(leg_rows) == 0:
                continue

            # ── Single leg → outright (or TAPS if criteria met) ───────────
            if len(leg_rows) == 1:
                lr = leg_rows[0]
                t = {
                    "timestamp":    ts,
                    "trade_type":   "OUTRIGHT",
                    "notes":        "",
                    "cc":           cc,
                    "qty":          lr["qty"],
                    "hub":          lr["hub"],
                    "spread_price": None,
                    "legs": [{"strip": lr["strip"], "price": lr["price"]}],
                }
                t["trade_type"] = _classify_trade_type(t)
                if t["trade_type"] == "TAPS":
                    t["notes"] = "TAPS/MOC"
                trades.append(t)

            else:
                qtys = [r["qty"] for r in leg_rows]

                # ── Same qty → same-product spread ────────────────────────
                if len(set(qtys)) == 1:
                    qty = qtys[0]
                    # Use explicit diff row price if present; otherwise imply
                    if diff_rows:
                        sp = diff_rows[0]["price"]
                    else:
                        sp = round(leg_rows[0]["price"] - leg_rows[1]["price"], 6)

                    legs = [{"strip": r["strip"], "price": r["price"]}
                            for r in leg_rows]

                    trades.append({
                        "timestamp":    ts,
                        "trade_type":   "SPREAD",
                        "notes":        "" if diff_rows else "implied diff",
                        "cc":           cc,
                        "qty":          qty,
                        "hub":          leg_rows[0]["hub"],
                        "spread_price": sp,
                        "legs":         legs,
                    })

                # ── Mixed qty → flag each as unrelated outright ───────────
                else:
                    for lr in leg_rows:
                        trades.append({
                            "timestamp":    ts,
                            "trade_type":   "OUTRIGHT",
                            "notes":        "⚠ FLAG: same-timestamp same-CC different qty — verify",
                            "cc":           cc,
                            "qty":          lr["qty"],
                            "hub":          lr["hub"],
                            "spread_price": None,
                            "legs": [{"strip": lr["strip"], "price": lr["price"]}],
                        })

    return trades


# ═══════════════════════════════════════════════════════════════════════════
#  STYLING
# ═══════════════════════════════════════════════════════════════════════════

F_HDR       = PatternFill("solid", fgColor="1F3864")  # navy        — headers
F_OUT       = PatternFill("solid", fgColor="FFFFFF")  # white       — outrights
F_SPREAD    = PatternFill("solid", fgColor="FFF2CC")  # amber       — spreads
F_FLAG      = PatternFill("solid", fgColor="FCE4D6")  # salmon      — flagged
F_CC_BAND   = PatternFill("solid", fgColor="1F3864")  # navy        — CC banner
F_BLK_HDR   = PatternFill("solid", fgColor="D6E4F0")  # steel blue  — block header
F_TRADE_ALT = PatternFill("solid", fgColor="EBF3FB")  # pale blue   — alt trade rows
F_SUMMARY   = PatternFill("solid", fgColor="E2EFDA")  # pale green  — VWAP summary

FN_HDR   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
FN_BODY  = Font(name="Arial", size=10)
FN_BOLD  = Font(name="Arial", bold=True, size=10)
FN_FLAG  = Font(name="Arial", size=10, color="C00000")
FN_TITLE = Font(name="Arial", bold=True, size=13, color="1F3864")

_T = Side(style="thin",   color="BDD7EE")
_M = Side(style="medium", color="2E75B6")
BORD  = Border(left=_T, right=_T, top=_T, bottom=_T)


def _hdr(ws, row_num, labels, fill=None):
    for ci, label in enumerate(labels, 1):
        c = ws.cell(row=row_num, column=ci, value=label)
        c.font      = FN_HDR
        c.fill      = fill or F_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = BORD
    ws.row_dimensions[row_num].height = 30


def _c(ws, r, c_idx, val, fill=None, halign="left",
       bold=False, flag=False, font=None):
    cell = ws.cell(row=r, column=c_idx, value=val)
    cell.fill      = fill or F_OUT
    cell.alignment = Alignment(horizontal=halign, vertical="center")
    cell.border    = BORD
    if flag:
        cell.font = FN_FLAG
    elif bold:
        cell.font = FN_BOLD
    else:
        cell.font = font or FN_BODY
    return cell


# ═══════════════════════════════════════════════════════════════════════════
#  SHEET DEFINITIONS
# ═══════════════════════════════════════════════════════════════════════════

SH_LOG    = "Trade Log"
SH_TALLY  = "Volume Tally"
SH_IMPORT = "Import Log"

# Trade Log columns
# A  Timestamp | B  Trade Type | C  Notes/Flags | D  CC | E  Qty | F  Hub
# G  Spread Price
# H  Leg 1 Strip | I  Leg 1 Price
# J  Leg 2 Strip | K  Leg 2 Price
# L  Leg 3 Strip | M  Leg 3 Price
# N  Source File

LOG_COLS = [
    "Timestamp", "Trade Type", "Notes / Flags", "CC", "Qty", "Hub",
    "Spread Price",
    "Leg 1 Strip", "Leg 1 Price",
    "Leg 2 Strip", "Leg 2 Price",
    "Leg 3 Strip", "Leg 3 Price",
    "Source File",
]
LOG_WIDTHS = {
    "A": 16, "B": 12, "C": 40, "D": 8,  "E": 7,  "F": 16,
    "G": 14,
    "H": 14, "I": 12,
    "J": 14, "K": 12,
    "L": 14, "M": 12,
    "N": 28,
}

# Volume Tally columns
# A  Block label (CC × Strip or CC × Spread pair)
# B  Timestamp (or summary label)
# C  Qty
# D  Price  (per-trade price — never averaged)
# E  Cumul. Volume  (running)
# F  VWAP           (running)

TALLY_COLS   = ["CC × Strip / Spread", "Timestamp", "Qty",
                 "Price", "Cumul. Volume", "VWAP (running)"]
TALLY_WIDTHS = {"A": 30, "B": 18, "C": 12, "D": 14, "E": 14, "F": 16}


# ═══════════════════════════════════════════════════════════════════════════
#  BUILD FRESH WORKBOOK
# ═══════════════════════════════════════════════════════════════════════════

def build_fresh_workbook(path: str):
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(SH_LOG)
    ws.freeze_panes = "A2"
    _hdr(ws, 1, LOG_COLS)
    for col, w in LOG_WIDTHS.items():
        ws.column_dimensions[col].width = w

    wt = wb.create_sheet(SH_TALLY)
    wt.freeze_panes = "A3"
    wt["A1"] = "Volume Tally  —  per-trade prices · cumulative volume · VWAP"
    wt["A1"].font      = FN_TITLE
    wt["A1"].alignment = Alignment(horizontal="left", vertical="center")
    wt.row_dimensions[1].height = 28
    _hdr(wt, 2, TALLY_COLS)
    for col, w in TALLY_WIDTHS.items():
        wt.column_dimensions[col].width = w

    wi = wb.create_sheet(SH_IMPORT)
    _hdr(wi, 1, ["Import Time", "Source File", "Raw Rows",
                  "Trades", "New Added", "Skipped"])
    for col, w in {"A": 22, "B": 32, "C": 12,
                    "D": 10, "E": 12, "F": 10}.items():
        wi.column_dimensions[col].width = w

    wb.save(path)


# ═══════════════════════════════════════════════════════════════════════════
#  DEDUPLICATION KEY
# ═══════════════════════════════════════════════════════════════════════════

def load_existing_keys(path: str) -> set:
    """
    (timestamp, trade_type, cc, leg1_strip) — four fields to safely
    distinguish two legitimate trades of the same type/CC at the same second.
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb[SH_LOG]
        keys = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1] and row[3] and row[7]:
                keys.add((str(row[0]).strip(), str(row[1]).strip(),
                          str(row[3]).strip(), str(row[7]).strip()))
        wb.close()
        return keys
    except Exception:
        return set()


# ═══════════════════════════════════════════════════════════════════════════
#  WRITE ONE TRADE LOG ROW
# ═══════════════════════════════════════════════════════════════════════════

def _write_log_row(ws, t: dict, source_file: str):
    legs = t.get("legs", [])
    tt   = t["trade_type"]
    note = t.get("notes", "")

    fill = (F_FLAG   if "⚠" in note else
            F_SPREAD if tt == "SPREAD" else
            F_OUT)

    r = ws.max_row + 1
    ws.row_dimensions[r].height = 17

    def leg(i, field):
        return legs[i][field] if i < len(legs) else None

    vals = [
        t["timestamp"], tt, note, t["cc"], t["qty"], t.get("hub", ""),
        t.get("spread_price"),
        leg(0, "strip"), leg(0, "price"),
        leg(1, "strip"), leg(1, "price"),
        leg(2, "strip"), leg(2, "price"),
        source_file,
    ]

    RIGHT = {5, 7, 9, 11, 13}  # Qty + price columns (1-indexed)
    for ci, val in enumerate(vals, 1):
        _c(ws, r, ci, val,
           fill=fill,
           halign="right" if ci in RIGHT else "left",
           flag=(ci == 3 and "⚠" in str(val or "")))


# ═══════════════════════════════════════════════════════════════════════════
#  REBUILD VOLUME TALLY
# ═══════════════════════════════════════════════════════════════════════════

def _rebuild_tally(wb):
    ws = wb[SH_LOG]
    wt = wb[SH_TALLY]

    # ── Read all trades from Trade Log ─────────────────────────────────────
    all_trades = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        ts    = str(row[0] or "")
        tt    = str(row[1] or "")
        notes = str(row[2] or "")
        cc    = str(row[3] or "")
        qty   = row[4] or 0
        hub, sp = row[5], row[6]
        legs = []
        for i in range(3):
            base = 7 + i * 2   # cols H,I then J,K then L,M (0-indexed: 7,8,9,10,11,12)
            s, p = row[base], row[base + 1]
            if s:
                legs.append({"strip": s, "price": p})
        all_trades.append({
            "ts": ts, "tt": tt, "cc": cc,
            "qty": qty or 0, "sp": sp, "legs": legs,
        })

    if not all_trades:
        return

    # ── Build records for two bucket types ────────────────────────────────
    #
    #  (a) OUTRIGHT  key = (cc, strip,        "OUTRIGHT")
    #  (b) SPREAD    key = (cc, "L1 / L2 ...", "SPREAD")   — differential only

    from collections import defaultdict
    buckets: dict[tuple, list] = defaultdict(list)

    for t in all_trades:
        cc, qty, legs, tt = t["cc"], t["qty"], t["legs"], t["tt"]

        if tt in ("OUTRIGHT", "TAPS") and legs:
            l = legs[0]
            buckets[(cc, l["strip"], tt)].append(
                {"ts": t["ts"], "qty": qty, "price": l["price"]})

        elif tt == "SPREAD" and t["sp"] is not None:
            # One row per spread trade using the differential price only.
            # Individual legs are not shown separately to avoid duplication.
            diff_label = " / ".join(l["strip"] for l in legs)
            buckets[(cc, diff_label, "SPREAD")].append(
                {"ts": t["ts"], "qty": qty, "price": t["sp"]})

    # Sort each bucket by timestamp
    for k in buckets:
        buckets[k].sort(key=lambda x: x["ts"])

    # ── Sort keys: CC → kind order → strip label ──────────────────────────
    KIND_ORDER = {"OUTRIGHT": 0, "TAPS": 1, "SPREAD": 2}
    sorted_keys = sorted(
        buckets.keys(),
        key=lambda k: (k[0], KIND_ORDER.get(k[2], 9), k[1])
    )

    # ── Clear old tally rows (row 3+) ─────────────────────────────────────
    for rw in wt.iter_rows(min_row=3):
        for c in rw:
            c.value  = None
            c.fill   = F_OUT
            c.border = Border()
            c.font   = FN_BODY

    # ── Write blocks ──────────────────────────────────────────────────────
    r        = 3
    prev_cc  = None

    for key in sorted_keys:
        cc_key, strip_label, kind = key
        recs      = buckets[key]
        trade_fill = F_SPREAD if kind == "SPREAD" else F_FLAG if kind == "TAPS" else F_OUT

        # CC banner (new CC)
        if cc_key != prev_cc:
            if prev_cc is not None:
                r += 1  # blank spacer row between CCs
            for ci in range(1, 7):
                c = wt.cell(row=r, column=ci,
                             value=(cc_key if ci == 1 else None))
                c.fill      = F_CC_BAND
                c.font      = FN_HDR
                c.border    = BORD
                c.alignment = Alignment(horizontal="left", vertical="center")
            wt.row_dimensions[r].height = 22
            r += 1
            prev_cc = cc_key

        # Block sub-header
        kind_label = {"OUTRIGHT": "Outright", "TAPS": "TAPS / MOC", "SPREAD": "Spread"}[kind]
        block_title = f"{strip_label}  [{kind_label}]"

        for ci in range(1, 7):
            c = wt.cell(row=r, column=ci,
                         value=(block_title if ci == 1 else None))
            c.fill      = F_BLK_HDR
            c.font      = FN_BOLD
            c.border    = BORD
            c.alignment = Alignment(horizontal="left", vertical="center")
        wt.row_dimensions[r].height = 18
        r += 1

        # Individual trade rows — running cumvol + VWAP
        cumvol   = 0
        vwap_num = 0.0

        for i, rec in enumerate(recs):
            qty_val   = rec["qty"]   or 0
            price_val = rec["price"] or 0.0
            cumvol   += qty_val
            vwap_num += qty_val * price_val
            running_vwap = round(vwap_num / cumvol, 6) if cumvol else None

            row_fill = F_TRADE_ALT if i % 2 == 0 else trade_fill
            vals     = [None, rec["ts"], qty_val, rec["price"],
                        cumvol, running_vwap]
            aligns   = ["left", "left", "right", "right", "right", "right"]
            for ci, (val, ha) in enumerate(zip(vals, aligns), 1):
                _c(wt, r, ci, val, fill=row_fill, halign=ha)
            wt.row_dimensions[r].height = 16
            r += 1

        # Summary row
        final_vwap = round(vwap_num / cumvol, 6) if cumvol else None
        sum_vals   = ["► Cumul. Vol / VWAP", "", cumvol, "",
                      cumvol, final_vwap]
        sum_aligns = ["left", "left", "right", "right", "right", "right"]
        for ci, (val, ha) in enumerate(zip(sum_vals, sum_aligns), 1):
            _c(wt, r, ci, val, fill=F_SUMMARY, halign=ha, bold=True)
        wt.row_dimensions[r].height = 18
        r += 1


# ═══════════════════════════════════════════════════════════════════════════
#  APPEND TRADES + SAVE
# ═══════════════════════════════════════════════════════════════════════════

def append_trades(path: str, trades: list[dict],
                  raw_row_count: int, source_file: str):
    existing = load_existing_keys(path)
    wb = load_workbook(path)
    ws = wb[SH_LOG]
    wi = wb[SH_IMPORT]

    new_count  = 0
    skip_count = 0

    for t in trades:
        legs  = t.get("legs", [])
        l1    = legs[0]["strip"] if legs else ""
        key   = (t["timestamp"].strip(), t["trade_type"],
                 t["cc"].strip(), l1)
        if key in existing:
            skip_count += 1
            continue
        _write_log_row(ws, t, source_file)
        existing.add(key)
        new_count += 1

    _rebuild_tally(wb)

    log_r = wi.max_row + 1
    for ci, val in enumerate([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        source_file, raw_row_count, len(trades), new_count, skip_count
    ], 1):
        _c(wi, log_r, ci, val,
           halign="left" if ci <= 2 else "center")
    wi.row_dimensions[log_r].height = 17

    wb.save(path)
    return new_count, skip_count


# ═══════════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description="Trade blotter screengrab → Excel accumulator")
    ap.add_argument("image",
                    help="Path to screengrab PNG/JPG")
    ap.add_argument("--output", "-o", default="trade_tally.xlsx",
                    help="Excel file (created on first run, updated thereafter)")
    ap.add_argument("--api-key",
                    default=os.environ.get("ANTHROPIC_API_KEY", ""),
                    help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Parse and print trades without writing to Excel")
    args = ap.parse_args()

    if not args.api_key:
        sys.exit("ERROR: No API key found. "
                 "Set ANTHROPIC_API_KEY or pass --api-key.")
    if not Path(args.image).exists():
        sys.exit(f"ERROR: Image not found: {args.image}")

    print(f"📷  Parsing {args.image} …")
    raw_rows = parse_image_with_claude(args.image, args.api_key)
    print(f"✅  {len(raw_rows)} raw row(s) extracted.")

    trades = group_rows_into_trades(raw_rows)
    print(f"🔗  Grouped into {len(trades)} trade(s).")

    if args.dry_run:
        print(json.dumps(trades, indent=2))
        return

    if not Path(args.output).exists():
        print(f"📄  Creating new workbook: {args.output}")
        build_fresh_workbook(args.output)

    added, skipped = append_trades(
        args.output, trades, len(raw_rows), Path(args.image).name)

    print(f"📊  {added} new trade(s) added, {skipped} duplicate(s) skipped.")
    print(f"💾  Saved → {args.output}")


if __name__ == "__main__":
    main()
