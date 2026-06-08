"""
Persistent storage for screenshots and scraped trade data.
Uses SQLite — no extra dependencies required.
"""

import json
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH         = Path(__file__).parent / "scraper.db"
SCREENSHOTS_DIR = Path(__file__).parent / "screenshots"
ARCHIVE_DIR     = Path(__file__).parent / "archive"


def _connect(db_path: Path | None = None) -> sqlite3.Connection:
    path = db_path or DB_PATH
    try:
        con = sqlite3.connect(path)
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("PRAGMA foreign_keys=ON")
        return con
    except sqlite3.Error as e:
        raise RuntimeError(f"Cannot open database at {path}: {e}") from e


def init_db(db_path: Path | None = None, screenshots_dir: Path | None = None):
    """Create tables and screenshots folder if they don't exist."""
    shots = screenshots_dir or SCREENSHOTS_DIR
    try:
        shots.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        raise RuntimeError(f"Cannot create screenshots directory {shots}: {e}") from e

    with _connect(db_path) as con:
        con.executescript("""
            CREATE TABLE IF NOT EXISTS imports (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                imported_at     TEXT    NOT NULL,
                original_name   TEXT    NOT NULL,
                screenshot_path TEXT    NOT NULL,
                raw_row_count   INTEGER NOT NULL DEFAULT 0,
                trade_count     INTEGER NOT NULL DEFAULT 0,
                added           INTEGER,
                skipped         INTEGER
            );

            CREATE TABLE IF NOT EXISTS raw_rows (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                import_id INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
                row_json  TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS trades (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                import_id  INTEGER NOT NULL REFERENCES imports(id) ON DELETE CASCADE,
                trade_json TEXT    NOT NULL
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


def known_row_keys(db_path: Path | None = None) -> set[tuple]:
    """Return the dedup keys of every raw row ever stored."""
    with _connect(db_path) as con:
        rows = con.execute("SELECT row_json FROM raw_rows").fetchall()
    return {_row_key(json.loads(r["row_json"])) for r in rows}


def filter_new_rows(raw_rows: list[dict],
                    db_path: Path | None = None) -> tuple[list[dict], int]:
    """
    Remove rows already seen in any previous import.
    Also deduplicates within the current batch.
    Returns (new_rows, n_skipped).
    """
    if not raw_rows:
        return [], 0

    known = known_row_keys(db_path)
    new: list[dict] = []
    skipped = 0
    seen_this_batch: set[tuple] = set()

    for row in raw_rows:
        k = _row_key(row)
        if k in known or k in seen_this_batch:
            skipped += 1
        else:
            new.append(row)
            seen_this_batch.add(k)

    return new, skipped


def save_import(original_name: str,
                tmp_image_path: str,
                raw_rows: list[dict],
                trades: list[dict],
                db_path: Path | None = None,
                screenshots_dir: Path | None = None) -> int:
    """
    Copy the screenshot into screenshots/, persist all rows and trades,
    return the new import id.
    Raises RuntimeError if the screenshot cannot be copied or the DB write fails.
    """
    if not original_name:
        raise ValueError("original_name must not be empty")

    shots  = screenshots_dir or SCREENSHOTS_DIR
    ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = Path(original_name).suffix or ".png"
    dest   = shots / f"{ts}_{Path(original_name).stem}{suffix}"

    try:
        shutil.copy2(tmp_image_path, dest)
    except OSError as e:
        raise RuntimeError(f"Could not save screenshot to {dest}: {e}") from e

    try:
        with _connect(db_path) as con:
            cur = con.execute(
                """INSERT INTO imports
                       (imported_at, original_name, screenshot_path,
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
    except sqlite3.Error as e:
        dest.unlink(missing_ok=True)
        raise RuntimeError(f"Database write failed: {e}") from e

    return import_id


def update_save_counts(import_id: int, added: int, skipped: int,
                       db_path: Path | None = None):
    """Record how many trades were actually written to Excel."""
    if import_id <= 0:
        raise ValueError(f"Invalid import_id: {import_id}")
    with _connect(db_path) as con:
        con.execute(
            "UPDATE imports SET added=?, skipped=? WHERE id=?",
            (added, skipped, import_id)
        )


def list_imports(db_path: Path | None = None) -> list[dict]:
    """Return all imports, newest first."""
    with _connect(db_path) as con:
        rows = con.execute(
            """SELECT id, imported_at, original_name, screenshot_path,
                      raw_row_count, trade_count, added, skipped
               FROM imports ORDER BY id DESC"""
        ).fetchall()
    return [dict(r) for r in rows]


def get_import_trades(import_id: int, db_path: Path | None = None) -> list[dict]:
    """Return the grouped trades for a given import."""
    with _connect(db_path) as con:
        rows = con.execute(
            "SELECT trade_json FROM trades WHERE import_id=? ORDER BY id",
            (import_id,)
        ).fetchall()
    return [json.loads(r["trade_json"]) for r in rows]


def get_screenshot_path(import_id: int, db_path: Path | None = None) -> Path | None:
    """Return the stored screenshot path for an import, or None if not found."""
    with _connect(db_path) as con:
        row = con.execute(
            "SELECT screenshot_path FROM imports WHERE id=?", (import_id,)
        ).fetchone()
    return Path(row["screenshot_path"]) if row else None


# ── Day boundary helpers ───────────────────────────────────────────────────────

def get_last_import_date(db_path: Path | None = None) -> str | None:
    """Return the date (YYYY-MM-DD) of the most recent import, or None if empty."""
    path = db_path or DB_PATH
    if not path.exists():
        return None
    try:
        with _connect(db_path) as con:
            row = con.execute(
                "SELECT imported_at FROM imports ORDER BY id DESC LIMIT 1"
            ).fetchone()
        if row:
            return row["imported_at"][:10]   # "YYYY-MM-DD"
    except Exception:
        pass
    return None


def is_new_business_day(db_path: Path | None = None) -> bool:
    """
    Return True if the DB contains data from a previous business day.
    Weekends are skipped — Monday detects Friday's data as stale.
    """
    from datetime import date, timedelta

    last = get_last_import_date(db_path)
    if not last:
        return False

    today = date.today()
    last_date = date.fromisoformat(last)

    if last_date >= today:
        return False

    # Walk backwards from today to find the last business day
    prev = today - timedelta(days=1)
    while prev.weekday() >= 5:          # 5=Sat, 6=Sun
        prev -= timedelta(days=1)

    return last_date < prev or (last_date == prev and today != prev)


def reset_day(excel_path: Path,
              db_path: Path | None = None,
              screenshots_dir: Path | None = None,
              archive_dir: Path | None = None):
    """
    Archive today's Excel file and wipe the DB + screenshots.
    Clears DB tables in-place (avoids Windows file-lock issues with WAL mode).
    Safe to call even if files don't exist.
    """
    from datetime import date

    archive = archive_dir or ARCHIVE_DIR
    shots   = screenshots_dir or SCREENSHOTS_DIR

    archive.mkdir(parents=True, exist_ok=True)
    shots.mkdir(parents=True, exist_ok=True)

    # Archive Excel with the date it covers (last import date or today)
    trade_date = get_last_import_date(db_path) or str(date.today())
    if excel_path.exists():
        dest = archive / f"{trade_date}_trade_tally.xlsx"
        shutil.copy2(excel_path, dest)
        excel_path.unlink()

    # Wipe screenshots
    for f in shots.iterdir():
        try:
            f.unlink()
        except OSError:
            pass

    # Clear DB tables in-place — avoids deleting the file while it may be
    # locked by SQLite's WAL mode on Windows
    with _connect(db_path) as con:
        con.execute("DELETE FROM trades")
        con.execute("DELETE FROM raw_rows")
        con.execute("DELETE FROM imports")
        con.execute("DELETE FROM sqlite_sequence WHERE name IN ('imports','raw_rows','trades')")

    # VACUUM must run outside a transaction
    con2 = _connect(db_path)
    con2.isolation_level = None
    con2.execute("VACUUM")
    con2.close()


def get_all_trades(db_path: Path | None = None) -> list[dict]:
    """
    Return every trade across all imports, oldest first, with the source
    filename attached. Used by the Trade Log and Volume Tally routes so they
    reflect parsed data immediately without needing an Excel save.
    """
    with _connect(db_path) as con:
        rows = con.execute(
            """SELECT t.trade_json, i.original_name
               FROM trades t
               JOIN imports i ON i.id = t.import_id
               ORDER BY t.id ASC"""
        ).fetchall()
    result = []
    for r in rows:
        trade = json.loads(r["trade_json"])
        trade["source_file"] = r["original_name"]
        result.append(trade)
    return result
