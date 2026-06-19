"""Tests for ocr_parser._parse_line and parse_image_local."""

import pytest
from ocr_parser import _parse_line, _is_strip_token, _join_wrapped_lines, parse_image_local


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
        assert _is_strip_token("Month-ND")

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

class TestJoinWrappedLines:

    def test_single_line_unchanged(self):
        text = "14:25:58 BST 10 SMT Jul26 Sing Mogas 111.20 © BLK"
        assert _join_wrapped_lines(text) == [text]

    def test_wrapped_hub_rejoined(self):
        text = "14:25:58 BST 10 SMT Jul26 Sing Mogas 92\nUnl (Platts) 111.20 © BLK"
        lines = _join_wrapped_lines(text)
        assert len(lines) == 1
        assert "Sing Mogas 92 Unl (Platts)" in lines[0]

    def test_multiple_rows_stay_separate(self):
        text = (
            "14:25:58 BST 10 SMT Jul26 Sing Mogas 111.20 © BLK\n"
            "14:26:00 BST 5 NEC Aug26 Naphtha CIF 700.00 © BLK"
        )
        lines = _join_wrapped_lines(text)
        assert len(lines) == 2

    def test_multi_line_wrap_rejoined(self):
        """Three OCR lines for one trade row."""
        text = "14:25:58 BST 10 SMT Jul26 Sing Mogas 92\nUnl (Platts)/Brent\n1st Line 111.20 © BLK"
        lines = _join_wrapped_lines(text)
        assert len(lines) == 1
        assert lines[0].startswith("14:25:58")

    def test_blank_lines_ignored(self):
        text = "14:25:58 BST 10 SMT Jul26 Sing 111.20 © BLK\n\n14:26:00 BST 5 NEC Aug26 Naphtha 700.00 © BLK"
        lines = _join_wrapped_lines(text)
        assert len(lines) == 2

    def test_wrapped_line_fully_parses(self):
        """End-to-end: a wrapped row should parse correctly after joining."""
        text = "14:25:58 BST 10 SMT Jul26 Sing Mogas 92\nUnl (Platts)/Brent 1st Line 111.20 © BLK"
        lines = _join_wrapped_lines(text)
        assert len(lines) == 1
        row = _parse_line(lines[0])
        assert row is not None
        assert row["cc"]    == "SMT"
        assert row["qty"]   == 10
        assert row["strip"] == "Jul26"
        assert row["price"] == 111.20


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
        assert "Sing" in row["hub"]

    def test_outright_without_cc_resolved_from_hub(self):
        """Blank CC is resolved via hub lookup when the hub is in HUB_CC_MAP."""
        line = "13:03:21 BST 25 Jul26 Sing Mogas 92 Unl (Platts)/Brent 1st Line 17.25 @ BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]  == "STB"   # resolved from hub map
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

    def test_diff_row_large_price_flagged(self):
        # A 2-part slash strip is always a spread diff, regardless of price magnitude.
        line = "12:54:02 BST NJC 5 Aug26/Sep26 Naphtha C&F Japan Cg 500.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["is_diff_row"] is True

    def test_negative_price(self):
        line = "13:04:42 BST NBP 20 Aug26 Naphtha CIF NWE Cg -12.25 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["price"] == -12.25

    def test_far_east_hub_parsed(self):
        line = "10:13:01 BST AFE 2 Aug26 Far East 667.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "AFE"
        assert row["qty"]   == 2
        assert row["strip"] == "Aug26"
        assert row["hub"]   == "Far East"

    def test_saudi_cp_hub_parsed(self):
        line = "10:07:33 BST SCP 1 Oct26 Saudi CP 581.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "SCP"
        assert row["strip"] == "Oct26"
        assert row["hub"]   == "Saudi CP"

    def test_mt_betr_hub_parsed(self):
        line = "09:56:15 BST PRL 12 Jul26 MT B-ETR 0.800000 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "PRL"
        assert row["strip"] == "Jul26"
        assert row["hub"]   == "MT B-ETR"

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

    def test_new_format_qty_before_cc(self):
        """New screenshot format: Qty comes before CC."""
        line = "13:02:53 BST 50 STB Jul26 Sing Mogas 92 Unl 17.25 @ BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["qty"]   == 50
        assert row["cc"]    == "STB"
        assert row["strip"] == "Jul26"

    def test_stuck_qty_and_cc(self):
        """OCR merges qty+CC into one token, e.g. '5NJC' — CC must be extracted."""
        line = "09:15:00 BST 5NJC Jul26 Far East 0.000 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["qty"]   == 5
        assert row["cc"]    == "NJC"
        assert row["strip"] == "Jul26"

    def test_stuck_qty_and_cc_with_quarter(self):
        line = "09:28:00 BST 5NJD Q3 26 Far East -0.050 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["qty"]   == 5
        assert row["cc"]    == "NJD"
        assert row["strip"] == "Q3 26"

    def test_stuck_qty_and_strip_word_still_works(self):
        """'4Bal' — lowercase means it's a strip word, not a CC.
        CC is blank in the column but filled from hub lookup (Naphtha → NEC)."""
        line = "13:00:40 BST 4Bal Month Naphtha CIF NWE Cg 712.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["qty"]   == 4
        assert row["cc"]    == "NEC"   # filled via hub lookup
        assert row["strip"] == "Bal Month"

    def test_new_format_quarter_strip(self):
        """New format: Q not misread into qty."""
        line = "09:00:00 BST 5 NJC Q4 26 Far East 0.000 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["qty"]   == 5
        assert row["cc"]    == "NJC"
        assert row["strip"] == "Q4 26"

    def test_new_format_no_cc(self):
        line = "13:03:21 BST 25 Jul26 Sing Mogas 92 Unl 17.25 @ BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["qty"]   == 25
        assert row["cc"]    == ""

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

    def test_hub_captured(self):
        """Hub field should be captured from the first non-strip token onwards."""
        cases = [
            ("10:13:01 BST AFE 2 Aug26 Far East 667.00 © BLK",        "Far East"),
            ("10:07:33 BST SCP 1 Oct26 Saudi CP 581.00 © BLK",        "Saudi CP"),
            ("12:53:46 BST AOM 10 Jul26 Argus Eurobob Oxy 957.95 © BLK", "Argus Eurobob Oxy"),
        ]
        for line, expected_hub in cases:
            row = _parse_line(line)
            assert row is not None, f"Failed to parse: {line}"
            assert row["hub"] == expected_hub, f"Expected hub '{expected_hub}', got '{row['hub']}'"


# ── parse_image_local ─────────────────────────────────────────────────────────

class TestParseImageLocal:

    def test_hub_fills_blank_cc_far_east(self):
        """Bal Month outright where CC column is blank — hub resolves it."""
        line = "09:37:09 BST 4 Bal Month Far East 775.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "AFE"
        assert row["strip"] == "Bal Month"

    def test_hub_fills_blank_cc_spread_diff(self):
        """Spread diff row with blank CC and joined hub — first hub segment used."""
        line = "13:00:40 BST 4Bal Month/Jul26 =Naphtha CIF NWE Cg/Naphtha CIF NWE Cg 11.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "NEC"
        assert row["strip"] == "Bal Month/Jul26"

    def test_hub_does_not_override_existing_cc(self):
        """If CC is already present, hub lookup is not applied."""
        line = "13:00:40 BST NEH 4Bal Month Naphtha CIF NWE Cg 712.50 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"] == "NEH"

    def test_hub_unknown_leaves_cc_blank(self):
        """Hub not in the map → CC stays empty."""
        line = "13:00:00 BST 5 Jul26 Unknown Hub XYZ 700.00 © BLK"
        row = _parse_line(line)
        if row:
            assert row["cc"] == ""

    def test_cancelled_trade_flagged(self):
        line = "13:02:53 BST 50 STB Jul26 Sing Mogas 92 Unl 17.25 © cancelled"
        row = _parse_line(line)
        assert row is not None
        assert row["cancelled"] is True
        assert row["price"] == 17.25

    def test_normal_trade_not_cancelled(self):
        line = "13:02:53 BST 50 STB Jul26 Sing Mogas 92 Unl 17.25 @ BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cancelled"] is False

    def test_missing_file_raises(self):
        with pytest.raises(FileNotFoundError):
            parse_image_local("nonexistent_file.png")

    def test_invalid_image_raises(self, tmp_path):
        bad = tmp_path / "bad.png"
        bad.write_bytes(b"not an image")
        with pytest.raises((ValueError, Exception)):
            parse_image_local(str(bad))


# ── New blotter format (Ex.Time + Sub.Time + Region + Strategies) ─────────────

class TestNewBlotterFormat:

    def test_sub_timestamp_skipped(self):
        """Second timestamp (Sub.Time) is skipped; Ex.Time is kept."""
        line = "12:20:57 BST 12:26:32 BST 6 AEO Europe Aug26 858.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["timestamp"] == "12:20:57 BST"
        assert row["qty"]   == 6
        assert row["cc"]    == "AEO"
        assert row["strip"] == "Aug26"
        assert row["price"] == 858.00

    def test_region_europe_skipped(self):
        """'Europe' between CC and Strip is discarded."""
        line = "12:20:57 BST 12:26:32 BST 6 AEO Europe Jul26/Aug26 spread 23.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]          == "AEO"
        assert row["strip"]       == "Jul26/Aug26"
        assert row["is_diff_row"] is True
        assert row["price"]       == 23.00

    def test_region_singapore_skipped(self):
        """'Singapore' between CC and Strip is discarded."""
        line = "11:37:18 BST 11:39:40 BST 5 NJC Singapore Jul26 673.00 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["cc"]    == "NJC"
        assert row["strip"] == "Jul26"
        assert row["price"] == 673.00

    def test_explicit_spread_keyword_sets_diff_row(self):
        """'spread' in Strategies column sets is_diff_row without needing heuristic."""
        line = "11:29:46 BST 11:34:07 BST 50 SMT Singapore Jul26/Aug26 spread 4.05 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["is_diff_row"] is True
        assert row["cc"]    == "SMT"
        assert row["strip"] == "Jul26/Aug26"

    def test_cal_strip_token(self):
        """'Cal 27' is recognised as a strip."""
        line = "12:09:41 BST 12:20:04 BST 5 NBB Europe Cal 27 -9.40 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["strip"] == "Cal 27"
        assert row["qty"]   == 5
        assert row["cc"]    == "NBB"
        assert row["price"] == -9.40

    def test_quarter_with_region(self):
        """Q3 26 strip parsed correctly when Region column is present."""
        line = "12:15:01 BST 12:25:52 BST 25 AEB Europe Q3 26 23.250 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["strip"] == "Q3 26"
        assert row["qty"]   == 25
        assert row["cc"]    == "AEB"

    def test_new_format_explicit_spread_large_price_flagged(self):
        """Explicit 'spread' keyword flags diff row even if price > 100."""
        line = "11:41:01 BST 11:48:14 BST 10 AEO Europe Oct26/Nov26 spread 33.25 © BLK"
        row = _parse_line(line)
        assert row is not None
        assert row["is_diff_row"] is True
        assert row["strip"] == "Oct26/Nov26"
