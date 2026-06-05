"""Tests for ocr_parser._parse_line and parse_image_local."""

import pytest
from ocr_parser import _parse_line, _is_strip_token, parse_image_local


# ── _is_strip_token ───────────────────────────────────────────────────────────

class TestIsStripToken:

    def test_month_year(self):
        assert _is_strip_token("Jul26")
        assert _is_strip_token("Aug26")
        assert _is_strip_token("Dec27")

    def test_spread_slash(self):
        assert _is_strip_token("Jul26/Aug26")
        assert _is_strip_token("Mar27/Apr27")
        assert _is_strip_token("Aug26/Sep26")

    def test_bal_and_month(self):
        assert _is_strip_token("Bal")
        assert _is_strip_token("Month")

    def test_quarter(self):
        assert _is_strip_token("Q1")
        assert _is_strip_token("Q3")
        assert _is_strip_token("Q4")

    def test_bare_year(self):
        assert _is_strip_token("26")
        assert _is_strip_token("27")

    def test_hub_words_not_strip(self):
        assert not _is_strip_token("Far")
        assert not _is_strip_token("Saudi")
        assert not _is_strip_token("Naphtha")
        assert not _is_strip_token("MT")
        assert not _is_strip_token("Sing")
        assert not _is_strip_token("CIF")


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
        assert row["hub"]       == ""

    def test_outright_without_cc(self):
        line = "13:03:21 BST 25 Jul26 Sing Mogas 92 Unl (Platts)/Brent 1st Line 17.25 @ BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]  == ""
        assert row["qty"] == 25

    def test_qty_stuck_to_strip(self):
        line = "13:00:40 BST NEH 4Bal Month Naphtha CIF NWE Cg 712.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["qty"]   == 4
        assert row["strip"] == "Bal Month"

    def test_diff_row_detected(self):
        line = "12:58:16 BST NEC 20 Mar27/Apr27 Naphtha CIF NWE Cg 8.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["is_diff_row"] is True
        assert row["strip"] == "Mar27/Apr27"

    def test_diff_row_large_price_not_flagged(self):
        line = "12:54:02 BST NJC 5 Aug26/Sep26 Naphtha C&F Japan Cg 500.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["is_diff_row"] is False

    def test_negative_price(self):
        line = "13:04:42 BST NBP 20 Aug26 Naphtha CIF NWE Cg -12.25 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["price"] == -12.25

    def test_far_east_hub_ignored(self):
        line = "10:13:01 BST AFE 2 Aug26 Far East 667.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "AFE"
        assert row["qty"]   == 2
        assert row["strip"] == "Aug26"
        assert row["hub"]   == ""

    def test_saudi_cp_hub_ignored(self):
        line = "10:07:33 BST SCP 1 Oct26 Saudi CP 581.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "SCP"
        assert row["strip"] == "Oct26"
        assert row["hub"]   == ""

    def test_mt_betr_hub_ignored(self):
        line = "09:56:15 BST PRL 12 Jul26 MT B-ETR 0.800000 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "PRL"
        assert row["strip"] == "Jul26"
        assert row["hub"]   == ""

    def test_far_east_spread_diff(self):
        line = "10:13:01 BST AFE 2 Jul26/Aug26 Far East 45.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["strip"]     == "Jul26/Aug26"
        assert row["is_diff_row"] is True

    def test_bal_month_slash_spread(self):
        line = "13:00:40 BST 4 Bal Month/Jul26 Naphtha CIF NWE Cg 11.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["strip"]     == "Bal Month/Jul26"
        assert row["is_diff_row"] is True

    def test_quarter_strip(self):
        line = "09:03:20 BST ARR 3 Q4 26 Far East/Naphtha C&F Japan -56.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["strip"] == "Q4 26"
        assert row["qty"]   == 3

    def test_qty_stuck_quarter(self):
        line = "12:51:16 BST NEC 1Q3 26 Naphtha CIF NWE Cg 687.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["qty"]   == 1
        assert row["strip"] == "Q3 26"

    def test_gmt_timezone(self):
        line = "13:02:53 GMT STB 50 Jul26 Sing Mogas 92 Unl 17.25 @ BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["timestamp"] == "13:02:53 GMT"

    def test_utc_timezone(self):
        line = "13:02:53 UTC NEC 10 Jul26 Naphtha CIF NWE Cg 700.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["timestamp"] == "13:02:53 UTC"

    def test_header_row_ignored(self):
        assert _parse_line("Ex. Time CC Qty Strip Hub Price TT") is None

    def test_empty_line_ignored(self):
        assert _parse_line("") is None
        assert _parse_line("   ") is None

    def test_no_price_marker_ignored(self):
        assert _parse_line("13:00:00 BST NEC 10 Jul26 Naphtha CIF NWE Cg 700.00") is None

    def test_no_timestamp_ignored(self):
        assert _parse_line("NEC 10 Jul26 Naphtha CIF NWE Cg 700.00 BLK") is None

    def test_multi_word_strip(self):
        line = "13:00:28 BST NEH 5 Bal Month Naphtha CIF NWE Cg 712.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["strip"] == "Bal Month"
        assert row["qty"]   == 5

    def test_hub_is_always_empty(self):
        """Hub field should always be empty string regardless of blotter hub."""
        for line in [
            "10:13:01 BST AFE 2 Aug26 Far East 667.00 © BLK",
            "10:07:33 BST SCP 1 Oct26 Saudi CP 581.00 © BLK",
            "13:02:53 BST STB 50 Jul26 Sing Mogas 92 Unl 17.25 @ BLK",
            "12:53:46 BST AOM 10 Jul26 Argus Eurobob Oxy 957.95 © BLK",
        ]:
            row = _parse_line(line)
            if row:
                assert row["hub"] == "", f"Expected empty hub, got '{row['hub']}' for: {line}"


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
