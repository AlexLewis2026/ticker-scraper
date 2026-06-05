"""Tests for ocr_parser._parse_line and parse_image_local."""

import pytest
from ocr_parser import _parse_line, parse_image_local


# ── _parse_line ────────────────────────────────────────────────────────────────

class TestParseLine:

    def test_outright_with_cc(self):
        line = "13:02:53 BST STB 50 Jul26 Sing Mogas 92 Unl (Platts)/Brent 1st Line 17.25 @ BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["timestamp"] == "13:02:53 BST"
        assert row["cc"]        == "STB"
        assert row["qty"]       == 50
        assert row["strip"]     == "Jul26"
        assert row["price"]     == 17.25
        assert row["is_diff_row"] is False

    def test_outright_without_cc(self):
        line = "13:03:21 BST 25 Jul26 Sing Mogas 92 Unl (Platts)/Brent 1st Line 17.25 @ BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]  == ""
        assert row["qty"] == 25

    def test_qty_stuck_to_strip(self):
        """Qty merged with strip word by OCR, e.g. '4Bal' → qty=4, strip='Bal Month'."""
        line = "13:00:40 BST NEH 4Bal Month Naphtha CIF NWE Cg 712.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "NEH"
        assert row["qty"]   == 4
        assert row["strip"] == "Bal Month"

    def test_diff_row_detected(self):
        """Strip containing '/' with small price → is_diff_row=True."""
        line = "12:58:16 BST NEC 20 Mar27/Apr27 Naphtha CIF NWE Cg 8.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["is_diff_row"] is True
        assert row["strip"] == "Mar27/Apr27"

    def test_diff_row_large_price_not_flagged(self):
        """Strip containing '/' but price >= 100 → not a diff row."""
        line = "12:54:02 BST NJC 5 Aug26/Sep26 Naphtha C&F Japan Cg 500.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["is_diff_row"] is False

    def test_negative_price(self):
        line = "13:04:42 BST NBP 20 Aug26 Naphtha CIF NWE Cg -12.25 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["price"] == -12.25

    def test_naphtha_hub(self):
        line = "12:57:43 BST NEC 10 Mar27 Naphtha CIF NWE Cg 623.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert "Naphtha" in row["hub"]

    def test_argus_hub(self):
        line = "12:53:46 BST AOM 10 Jul26 Argus Eurobob Oxy FOB Rdam Bg Mini 957.95 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert "Argus" in row["hub"]

    def test_ocr_junk_on_hub_stripped(self):
        """Leading '=' or '©' on hub word should be stripped."""
        line = "13:00:40 BST 4Bal Month/Jul26 =Naphtha CIF NWE Cg 11.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["hub"].startswith("Naphtha")

    def test_header_row_ignored(self):
        assert _parse_line("Ex. Time CC Qty Strip Hub Price TT") is None

    def test_empty_line_ignored(self):
        assert _parse_line("") is None
        assert _parse_line("   ") is None

    def test_no_blk_marker_ignored(self):
        assert _parse_line("13:00:00 BST NEC 10 Jul26 Naphtha CIF NWE Cg 700.00") is None

    def test_no_timestamp_ignored(self):
        assert _parse_line("NEC 10 Jul26 Naphtha CIF NWE Cg 700.00 BLK") is None

    def test_unknown_hub_ignored(self):
        """Row whose hub doesn't start with a known word is skipped."""
        assert _parse_line("13:00:00 BST NEC 10 Jul26 UnknownHub Cg 700.00 BLK") is None

    def test_multi_word_strip(self):
        line = "13:00:28 BST NEH 5 Bal Month Naphtha CIF NWE Cg 712.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["strip"] == "Bal Month"
        assert row["qty"]   == 5


# ── parse_image_local ─────────────────────────────────────────────────────────

class TestParseImageLocal:

    def test_missing_file_raises(self):
        with pytest.raises(FileNotFoundError):
            parse_image_local("nonexistent_file.png")

    def test_invalid_image_raises(self, tmp_path):
        bad = tmp_path / "bad.png"
        bad.write_bytes(b"not an image")
        with pytest.raises((ValueError, Exception)):
            parse_image_local(str(bad))
