"""Tests for db.py — storage, deduplication, retrieval."""

import json
import shutil
import sqlite3
from pathlib import Path

import pytest
import db as database


# ── fixtures ───────────────────────────────────────────────────────────────────

@pytest.fixture()
def store(tmp_path):
    """Isolated DB + screenshots dir for each test."""
    db_path   = tmp_path / "test.db"
    shots_dir = tmp_path / "screenshots"
    database.init_db(db_path=db_path, screenshots_dir=shots_dir)
    return db_path, shots_dir


def _make_image(tmp_path: Path, name="img.png") -> Path:
    """Create a minimal valid PNG for testing."""
    from PIL import Image
    p = tmp_path / name
    Image.new("RGB", (10, 10), color="white").save(p)
    return p


def _sample_rows():
    return [
        {"timestamp": "12:00:00 BST", "cc": "NEC", "qty": 10,
         "strip": "Jul26", "hub": "Naphtha CIF NWE Cg", "price": 700.0,
         "is_diff_row": False},
        {"timestamp": "12:00:00 BST", "cc": "NEC", "qty": 10,
         "strip": "Aug26", "hub": "Naphtha CIF NWE Cg", "price": 690.0,
         "is_diff_row": False},
    ]

def _sample_trades():
    return [{"trade_type": "OUTRIGHT", "cc": "NEC", "qty": 10,
             "timestamp": "12:00:00 BST", "hub": "Naphtha CIF NWE Cg",
             "spread_price": None, "notes": "",
             "legs": [{"strip": "Jul26", "price": 700.0}]}]


# ── init_db ────────────────────────────────────────────────────────────────────

class TestInitDb:

    def test_creates_tables(self, store):
        db_path, _ = store
        con = sqlite3.connect(db_path)
        tables = {r[0] for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
        assert {"imports", "raw_rows", "trades"} <= tables

    def test_creates_screenshots_dir(self, tmp_path):
        shots = tmp_path / "new_dir" / "screenshots"
        database.init_db(db_path=tmp_path / "t.db", screenshots_dir=shots)
        assert shots.exists()

    def test_idempotent(self, store):
        """Calling init_db twice does not raise."""
        db_path, shots = store
        database.init_db(db_path=db_path, screenshots_dir=shots)


# ── save_import ────────────────────────────────────────────────────────────────

class TestSaveImport:

    def test_returns_positive_id(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        iid = database.save_import("test.png", str(img),
                                   _sample_rows(), _sample_trades(),
                                   db_path=db_path, screenshots_dir=shots)
        assert isinstance(iid, int) and iid > 0

    def test_screenshot_copied(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        database.save_import("test.png", str(img), _sample_rows(), _sample_trades(),
                             db_path=db_path, screenshots_dir=shots)
        assert any(shots.iterdir())

    def test_rows_and_trades_persisted(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        iid = database.save_import("test.png", str(img), _sample_rows(), _sample_trades(),
                                   db_path=db_path, screenshots_dir=shots)
        con = sqlite3.connect(db_path)
        n_rows   = con.execute("SELECT COUNT(*) FROM raw_rows  WHERE import_id=?", (iid,)).fetchone()[0]
        n_trades = con.execute("SELECT COUNT(*) FROM trades     WHERE import_id=?", (iid,)).fetchone()[0]
        assert n_rows   == len(_sample_rows())
        assert n_trades == len(_sample_trades())

    def test_empty_name_raises(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        with pytest.raises(ValueError):
            database.save_import("", str(img), [], [], db_path=db_path, screenshots_dir=shots)

    def test_missing_source_image_raises(self, store):
        db_path, shots = store
        with pytest.raises(RuntimeError):
            database.save_import("x.png", "/no/such/file.png", [], [],
                                 db_path=db_path, screenshots_dir=shots)


# ── filter_new_rows ────────────────────────────────────────────────────────────

class TestFilterNewRows:

    def test_empty_input(self, store):
        db_path, shots = store
        new, skipped = database.filter_new_rows([], db_path=db_path)
        assert new == [] and skipped == 0

    def test_all_new_rows_kept(self, store):
        db_path, _ = store
        rows = _sample_rows()
        new, skipped = database.filter_new_rows(rows, db_path=db_path)
        assert len(new) == len(rows) and skipped == 0

    def test_duplicates_within_batch_skipped(self, store):
        db_path, _ = store
        row = _sample_rows()[0]
        new, skipped = database.filter_new_rows([row, row], db_path=db_path)
        assert len(new) == 1 and skipped == 1

    def test_previously_stored_rows_skipped(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        rows = _sample_rows()
        database.save_import("first.png", str(img), rows, [],
                             db_path=db_path, screenshots_dir=shots)
        new, skipped = database.filter_new_rows(rows, db_path=db_path)
        assert len(new) == 0 and skipped == len(rows)

    def test_new_rows_after_previous_import(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        old_rows = _sample_rows()
        database.save_import("first.png", str(img), old_rows, [],
                             db_path=db_path, screenshots_dir=shots)

        new_row = {"timestamp": "13:00:00 BST", "cc": "NEC", "qty": 5,
                   "strip": "Sep26", "hub": "Naphtha CIF NWE Cg", "price": 680.0,
                   "is_diff_row": False}
        new, skipped = database.filter_new_rows(old_rows + [new_row], db_path=db_path)
        assert len(new) == 1 and skipped == len(old_rows)
        assert new[0]["strip"] == "Sep26"


# ── list_imports ───────────────────────────────────────────────────────────────

class TestListImports:

    def test_empty_db(self, store):
        db_path, _ = store
        assert database.list_imports(db_path=db_path) == []

    def test_newest_first(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        database.save_import("a.png", str(img), [], [], db_path=db_path, screenshots_dir=shots)
        database.save_import("b.png", str(img), [], [], db_path=db_path, screenshots_dir=shots)
        imports = database.list_imports(db_path=db_path)
        assert imports[0]["original_name"] == "b.png"
        assert imports[1]["original_name"] == "a.png"

    def test_counts_correct(self, store, tmp_path):
        db_path, shots = store
        img   = _make_image(tmp_path)
        rows  = _sample_rows()
        trades = _sample_trades()
        database.save_import("t.png", str(img), rows, trades,
                             db_path=db_path, screenshots_dir=shots)
        imports = database.list_imports(db_path=db_path)
        assert imports[0]["raw_row_count"] == len(rows)
        assert imports[0]["trade_count"]   == len(trades)


# ── update_save_counts ─────────────────────────────────────────────────────────

class TestUpdateSaveCounts:

    def test_updates_correctly(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        iid = database.save_import("t.png", str(img), [], [],
                                   db_path=db_path, screenshots_dir=shots)
        database.update_save_counts(iid, added=5, skipped=2, db_path=db_path)
        imports = database.list_imports(db_path=db_path)
        assert imports[0]["added"]   == 5
        assert imports[0]["skipped"] == 2

    def test_invalid_id_raises(self, store):
        db_path, _ = store
        with pytest.raises(ValueError):
            database.update_save_counts(0, 1, 0, db_path=db_path)


# ── get_import_trades ──────────────────────────────────────────────────────────

class TestGetImportTrades:

    def test_returns_correct_trades(self, store, tmp_path):
        db_path, shots = store
        img    = _make_image(tmp_path)
        trades = _sample_trades()
        iid    = database.save_import("t.png", str(img), [], trades,
                                      db_path=db_path, screenshots_dir=shots)
        result = database.get_import_trades(iid, db_path=db_path)
        assert len(result) == len(trades)
        assert result[0]["cc"] == trades[0]["cc"]

    def test_unknown_id_returns_empty(self, store):
        db_path, _ = store
        assert database.get_import_trades(9999, db_path=db_path) == []


# ── get_screenshot_path ────────────────────────────────────────────────────────

class TestGetScreenshotPath:

    def test_returns_path(self, store, tmp_path):
        db_path, shots = store
        img = _make_image(tmp_path)
        iid = database.save_import("t.png", str(img), [], [],
                                   db_path=db_path, screenshots_dir=shots)
        path = database.get_screenshot_path(iid, db_path=db_path)
        assert path is not None and path.exists()

    def test_unknown_id_returns_none(self, store):
        db_path, _ = store
        assert database.get_screenshot_path(9999, db_path=db_path) is None
