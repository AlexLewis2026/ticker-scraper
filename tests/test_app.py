"""Tests for Flask routes in app.py."""

import io
import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from PIL import Image

import app as flask_app


# ── fixtures ───────────────────────────────────────────────────────────────────

@pytest.fixture()
def client(tmp_path, monkeypatch):
    """Flask test client with isolated DB and no real OCR/Excel calls."""
    db_path   = tmp_path / "test.db"
    shots_dir = tmp_path / "screenshots"

    import db as database
    database.init_db(db_path=db_path, screenshots_dir=shots_dir)

    # Patch module-level defaults so all no-arg calls hit the tmp DB
    monkeypatch.setattr(database, "DB_PATH",         db_path)
    monkeypatch.setattr(database, "SCREENSHOTS_DIR", shots_dir)

    # Wrap each public function to inject tmp db_path when not provided
    import functools
    for fn_name in ("list_imports", "get_import_trades", "get_screenshot_path",
                    "known_row_keys", "filter_new_rows", "update_save_counts"):
        original = getattr(database, fn_name)
        @functools.wraps(original)
        def _wrapped(*args, _orig=original, _db=db_path, **kwargs):
            kwargs.setdefault("db_path", _db)
            return _orig(*args, **kwargs)
        monkeypatch.setattr(database, fn_name, _wrapped)

    save_orig = database.save_import
    @functools.wraps(save_orig)
    def _save_wrapped(*args, **kwargs):
        kwargs.setdefault("db_path",         db_path)
        kwargs.setdefault("screenshots_dir", shots_dir)
        return save_orig(*args, **kwargs)
    monkeypatch.setattr(database, "save_import", _save_wrapped)

    flask_app.app.config["TESTING"] = True
    flask_app.EXCEL_PATH = tmp_path / "trade_tally_v4.xlsx"

    with flask_app.app.test_client() as c:
        yield c, tmp_path, db_path, shots_dir


def _png_bytes():
    buf = io.BytesIO()
    Image.new("RGB", (10, 10), "white").save(buf, format="PNG")
    buf.seek(0)
    return buf


SAMPLE_TRADES = [
    {"trade_type": "OUTRIGHT", "cc": "NEC", "qty": 10, "timestamp": "12:00:00 BST",
     "hub": "Naphtha CIF NWE Cg", "spread_price": None, "notes": "",
     "legs": [{"strip": "Jul26", "price": 700.0}]}
]

SAMPLE_ROWS = [
    {"timestamp": "12:00:00 BST", "cc": "NEC", "qty": 10, "strip": "Jul26",
     "hub": "Naphtha CIF NWE Cg", "price": 700.0, "is_diff_row": False}
]


# ── GET / ──────────────────────────────────────────────────────────────────────

class TestIndex:
    def test_returns_200(self, client):
        c, *_ = client
        assert c.get("/").status_code == 200

    def test_returns_html(self, client):
        c, *_ = client
        assert b"Trade Accumulator" in c.get("/").data


# ── POST /parse ────────────────────────────────────────────────────────────────

class TestParse:

    def test_no_image_returns_400(self, client):
        c, *_ = client
        r = c.post("/parse", data={})
        assert r.status_code == 400
        assert "error" in r.get_json()

    def test_wrong_extension_returns_400(self, client):
        c, *_ = client
        r = c.post("/parse", data={
            "image": (io.BytesIO(b"data"), "file.pdf")
        }, content_type="multipart/form-data")
        assert r.status_code == 400
        assert "Unsupported" in r.get_json()["error"]

    def test_valid_image_returns_trades(self, client):
        c, tmp_path, db_path, shots_dir = client
        with patch("app.parse_image_local", return_value=SAMPLE_ROWS), \
             patch("app.database.filter_new_rows", return_value=(SAMPLE_ROWS, 0)), \
             patch("app.database.save_import", return_value=1):
            r = c.post("/parse", data={
                "image": (_png_bytes(), "blotter.png")
            }, content_type="multipart/form-data")

        assert r.status_code == 200
        data = r.get_json()
        assert "trades"     in data
        assert "raw_count"  in data
        assert "dup_count"  in data
        assert "import_id"  in data

    def test_ocr_failure_returns_400(self, client):
        c, *_ = client
        with patch("app.parse_image_local", side_effect=ValueError("No trade rows could be parsed")):
            r = c.post("/parse", data={
                "image": (_png_bytes(), "blotter.png")
            }, content_type="multipart/form-data")
        assert r.status_code == 400
        assert "No trade rows" in r.get_json()["error"]

    def test_duplicate_rows_reported(self, client):
        c, *_ = client
        with patch("app.parse_image_local", return_value=SAMPLE_ROWS), \
             patch("app.database.filter_new_rows", return_value=([], 1)), \
             patch("app.database.save_import", return_value=2):
            r = c.post("/parse", data={
                "image": (_png_bytes(), "blotter.png")
            }, content_type="multipart/form-data")
        assert r.status_code == 200
        assert r.get_json()["dup_count"] == 1


# ── POST /save ─────────────────────────────────────────────────────────────────

class TestSave:

    def test_no_body_returns_400(self, client):
        c, *_ = client
        r = c.post("/save", content_type="application/json", data="not json")
        assert r.status_code == 400

    def test_invalid_trades_type_returns_400(self, client):
        c, *_ = client
        r = c.post("/save", json={"trades": "not-a-list"})
        assert r.status_code == 400

    def test_valid_save_returns_counts(self, client):
        c, tmp_path, *_ = client
        with patch("app.append_trades", return_value=(3, 1)), \
             patch("app.build_fresh_workbook"):
            r = c.post("/save", json={
                "trades": SAMPLE_TRADES,
                "raw_count": 1,
                "filename": "test.png",
                "import_id": None,
            })
        assert r.status_code == 200
        data = r.get_json()
        assert data["added"]   == 3
        assert data["skipped"] == 1

    def test_excel_locked_returns_409(self, client):
        c, *_ = client
        with patch("app.append_trades", side_effect=PermissionError("locked")), \
             patch("app.build_fresh_workbook"):
            r = c.post("/save", json={"trades": SAMPLE_TRADES, "filename": "x.png"})
        assert r.status_code == 409
        assert "open in another program" in r.get_json()["error"]


# ── GET /history ───────────────────────────────────────────────────────────────

class TestHistory:

    def test_empty_db_returns_empty_list(self, client):
        c, *_ = client
        r = c.get("/history")
        assert r.status_code == 200
        assert r.get_json()["imports"] == []

    def test_returns_imports_after_save(self, client):
        c, tmp_path, db_path, shots_dir = client
        import db as database
        img_path = tmp_path / "img.png"
        Image.new("RGB", (10, 10)).save(img_path)
        database.save_import("blotter.png", str(img_path), SAMPLE_ROWS, SAMPLE_TRADES,
                             db_path=db_path, screenshots_dir=shots_dir)
        r = c.get("/history")
        assert r.status_code == 200
        assert len(r.get_json()["imports"]) == 1


# ── GET /import/<id>/trades ────────────────────────────────────────────────────

class TestImportTrades:

    def test_unknown_id_returns_empty(self, client):
        c, *_ = client
        r = c.get("/import/9999/trades")
        assert r.status_code == 200
        assert r.get_json()["trades"] == []

    def test_known_id_returns_trades(self, client):
        c, tmp_path, db_path, shots_dir = client
        import db as database
        img_path = tmp_path / "img.png"
        Image.new("RGB", (10, 10)).save(img_path)
        iid = database.save_import("b.png", str(img_path), SAMPLE_ROWS, SAMPLE_TRADES,
                                   db_path=db_path, screenshots_dir=shots_dir)
        r = c.get(f"/import/{iid}/trades")
        assert r.status_code == 200
        assert len(r.get_json()["trades"]) == len(SAMPLE_TRADES)


# ── GET /screenshot/<id> ───────────────────────────────────────────────────────

class TestScreenshot:

    def test_unknown_id_returns_404(self, client):
        c, *_ = client
        assert c.get("/screenshot/9999").status_code == 404

    def test_known_id_returns_image(self, client):
        c, tmp_path, db_path, shots_dir = client
        import db as database
        img_path = tmp_path / "img.png"
        Image.new("RGB", (10, 10)).save(img_path)
        iid = database.save_import("b.png", str(img_path), [], [],
                                   db_path=db_path, screenshots_dir=shots_dir)
        r = c.get(f"/screenshot/{iid}")
        assert r.status_code == 200


# ── GET /log ───────────────────────────────────────────────────────────────────

class TestLog:

    def test_no_excel_returns_empty(self, client):
        c, *_ = client
        r = c.get("/log")
        assert r.status_code == 200
        assert r.get_json()["rows"] == []


# ── GET /tally ─────────────────────────────────────────────────────────────────

class TestTally:

    def test_no_excel_returns_empty(self, client):
        c, *_ = client
        r = c.get("/tally")
        assert r.status_code == 200
        assert r.get_json()["groups"] == []


# ── GET /download ──────────────────────────────────────────────────────────────

class TestDownload:

    def test_no_excel_returns_404(self, client):
        c, *_ = client
        assert c.get("/download").status_code == 404

    def test_excel_present_returns_file(self, client):
        c, tmp_path, *_ = client
        xl = tmp_path / "trade_tally_v4.xlsx"
        xl.write_bytes(b"PK")   # placeholder file
        flask_app.EXCEL_PATH = xl
        r = c.get("/download")
        assert r.status_code == 200
