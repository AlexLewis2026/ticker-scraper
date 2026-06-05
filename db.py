"""
Persistent storage for screenshots and scraped trade data.
Uses SQLite — no extra dependencies required.
"""

import json
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH          = Path(__file__).parent / "scraper.db"
SCREENSHOTS_DIR  = Path(__file__).parent / "screenshots"


def _connect() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    return con


def init_db():
    """Create tables and screenshots folder if they don't exist."""
    SCREENSHOTS_DIR.mkdir(exist_ok=True)
    with _connect() as con:
        con.executescript("""
            CREATE TABLE IF NOT EXISTS imports (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                imported_at   TEXT    NOT NULL,
                original_name TEXT    NOT NULL,
                screenshot_path TEXT  NOT NULL,
                raw_row_count INTEGER NOT NULL DEFAULT 0,
                trade_count   INTEGER NOT NULL DEFAULT 0,
                added         INTEGER,
                skipped       INTEGER
            );

            CREATE TABLE IF NOT EXISTS raw_rows (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                import_id INTEGER NOT NULL REFERENCES imports(id),
                row_json  TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS trades (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                import_id INTEGER NOT NULL REFERENCES imports(id),
                trade_json TEXT   NOT NULL
            );
        """)


def _row_key(row: dict) -> tuple:
    """Deduplication key for a single raw OCR row."""
    return (
        row.get("timestamp", ""),
        row.get("cc", ""),
        row.get("qty"),
        row.get("strip", ""),
        row.get("price"),
    )


def known_row_keys() -> set[tuple]:
    """Return the dedup keys of every raw row ever stored."""
    with _connect() as con:
        rows = con.execute("SELECT row_json FROM raw_rows").fetchall()
    return {_row_key(json.loads(r["row_json"])) for r in rows}


def filter_new_rows(raw_rows: list[dict]) -> tuple[list[dict], int]:
    """
    Remove rows already seen in any previous import.
    Returns (new_rows, n_skipped).
    """
    known = known_row_keys()
    new, skipped = [], 0
    seen_this_batch: set[tuple] = set()
    for row in raw_rows:
        k = _row_key(row)
        if k in known or k in seen_this_batch:
            skipped += 1
        else:
            new.append(row)
            seen_this_batch.add(k)
    return new, skipped


def save_import(original_name: str, tmp_image_path: str,
                raw_rows: list[dict], trades: list[dict]) -> int:
    """
    Copy the screenshot into the screenshots/ folder, persist all rows
    and trades, return the new import id.
    """
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix   = Path(original_name).suffix or ".png"
    dest     = SCREENSHOTS_DIR / f"{ts}_{Path(original_name).stem}{suffix}"
    shutil.copy2(tmp_image_path, dest)

    with _connect() as con:
        cur = con.execute(
            """INSERT INTO imports (imported_at, original_name, screenshot_path,
                                   raw_row_count, trade_count)
               VALUES (?, ?, ?, ?, ?)""",
            (datetime.now().isoformat(timespec="seconds"),
             original_name, str(dest), len(raw_rows), len(trades))
        )
        import_id = cur.lastrowid

        con.executemany(
            "INSERT INTO raw_rows (import_id, row_json) VALUES (?, ?)",
            [(import_id, json.dumps(r)) for r in raw_rows]
        )
        con.executemany(
            "INSERT INTO trades (import_id, trade_json) VALUES (?, ?)",
            [(import_id, json.dumps(t)) for t in trades]
        )
    return import_id


def update_save_counts(import_id: int, added: int, skipped: int):
    """Record how many trades were actually written to Excel."""
    with _connect() as con:
        con.execute(
            "UPDATE imports SET added=?, skipped=? WHERE id=?",
            (added, skipped, import_id)
        )


def list_imports() -> list[dict]:
    """Return all imports, newest first."""
    with _connect() as con:
        rows = con.execute(
            """SELECT id, imported_at, original_name, screenshot_path,
                      raw_row_count, trade_count, added, skipped
               FROM imports ORDER BY id DESC"""
        ).fetchall()
    return [dict(r) for r in rows]


def get_import_trades(import_id: int) -> list[dict]:
    with _connect() as con:
        rows = con.execute(
            "SELECT trade_json FROM trades WHERE import_id=? ORDER BY id",
            (import_id,)
        ).fetchall()
    return [json.loads(r["trade_json"]) for r in rows]


def get_screenshot_path(import_id: int) -> Path | None:
    with _connect() as con:
        row = con.execute(
            "SELECT screenshot_path FROM imports WHERE id=?", (import_id,)
        ).fetchone()
    return Path(row["screenshot_path"]) if row else None
