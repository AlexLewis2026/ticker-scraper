"""Tests for trade_accumulator_v4.group_rows_into_trades."""

import pytest
from trade_accumulator_v4 import group_rows_into_trades, TAPS_CC, _strip_sort_key


def _row(ts, cc, qty, strip, price, is_diff=False, hub="Naphtha CIF NWE Cg"):
    return {"timestamp": ts, "cc": cc, "qty": qty, "strip": strip,
            "hub": hub, "price": price, "is_diff_row": is_diff}


class TestGroupRows:

    def test_single_leg_is_outright(self):
        rows = [_row("12:00:00 BST", "NEC", 10, "Jul26", 700.0)]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 1
        assert trades[0]["trade_type"] == "OUTRIGHT"
        assert trades[0]["qty"]        == 10
        assert trades[0]["cc"]         == "NEC"
        assert len(trades[0]["legs"])  == 1
        assert trades[0]["legs"][0]["strip"] == "Jul26"

    def test_two_legs_same_qty_is_spread(self):
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Mar27", 620.0),
            _row("12:00:00 BST", "NEC", 10, "Apr27", 612.0),
        ]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 1
        assert trades[0]["trade_type"] == "SPREAD"
        assert len(trades[0]["legs"]) == 2

    def test_spread_implied_diff_price(self):
        """When no explicit diff row, spread price is leg[0].price - leg[1].price."""
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Mar27", 620.0),
            _row("12:00:00 BST", "NEC", 10, "Apr27", 612.0),
        ]
        trades = group_rows_into_trades(rows)
        assert trades[0]["spread_price"] == pytest.approx(8.0)
        assert trades[0]["notes"] == "implied diff"

    def test_spread_explicit_diff_row(self):
        """Explicit diff row price takes precedence over implied."""
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Mar27",      620.0),
            _row("12:00:00 BST", "NEC", 10, "Apr27",      612.0),
            _row("12:00:00 BST", "NEC", 10, "Mar27/Apr27", 8.5, is_diff=True),
        ]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 1
        assert trades[0]["spread_price"] == 8.5
        assert trades[0]["notes"]        == ""

    def test_two_legs_different_qty_is_unequal_spread(self):
        """Rule 4: 2 legs same CC/timestamp but different qty → SPREAD with implied diff."""
        rows = [
            _row("12:00:00 BST", "NEC", 75, "Jul26", 700.0),
            _row("12:00:00 BST", "NEC", 25, "Aug26", 690.0),
        ]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 1
        assert trades[0]["trade_type"] == "SPREAD"
        assert trades[0]["spread_price"] == pytest.approx(10.0)
        assert "unequal" in trades[0]["notes"]

    def test_different_ccs_at_same_timestamp_are_independent(self):
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Jul26", 700.0),
            _row("12:00:00 BST", "NJC", 10, "Jul26", 730.0),
        ]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 2
        assert all(t["trade_type"] == "OUTRIGHT" for t in trades)
        ccs = {t["cc"] for t in trades}
        assert ccs == {"NEC", "NJC"}

    def test_three_legs_same_qty_is_spread(self):
        rows = [
            _row("12:00:00 BST", "NEC", 5, "Jul26", 700.0),
            _row("12:00:00 BST", "NEC", 5, "Aug26", 690.0),
            _row("12:00:00 BST", "NEC", 5, "Sep26", 680.0),
        ]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 1
        assert trades[0]["trade_type"] == "SPREAD"
        assert len(trades[0]["legs"]) == 3

    def test_empty_input(self):
        assert group_rows_into_trades([]) == []

    def test_different_timestamps_produce_separate_trades(self):
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Jul26", 700.0),
            _row("12:01:00 BST", "NEC", 10, "Jul26", 701.0),
        ]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 2

    def test_outright_spread_price_is_none(self):
        rows = [_row("12:00:00 BST", "NEC", 10, "Jul26", 700.0)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["spread_price"] is None

class TestCancelledTrades:

    def test_cancelled_row_becomes_cancelled_trade(self):
        rows = [_row("13:00:00 BST", "NEC", 10, "Jul26", 700.0)]
        rows[0]["cancelled"] = True
        trades = group_rows_into_trades(rows)
        assert len(trades) == 1
        assert trades[0]["trade_type"] == "CANCELLED"

    def test_cancelled_not_mixed_with_active_spread(self):
        """A cancelled row at the same timestamp as an active leg should not form a spread."""
        rows = [
            {**_row("12:00:00 BST", "NEC", 10, "Jul26", 700.0), "cancelled": False},
            {**_row("12:00:00 BST", "NEC", 10, "Aug26", 690.0), "cancelled": True},
        ]
        trades = group_rows_into_trades(rows)
        types = {t["trade_type"] for t in trades}
        assert "CANCELLED" in types
        assert "OUTRIGHT" in types
        assert "SPREAD" not in types

    def test_active_row_not_affected_by_cancelled_flag(self):
        rows = [_row("12:00:00 BST", "NEC", 10, "Jul26", 700.0)]
        rows[0]["cancelled"] = False
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "OUTRIGHT"


class TestHubMisc:

    def test_hub_carried_through(self):
        rows = [_row("12:00:00 BST", "STB", 25, "Jul26", 17.0,
                     hub="Sing Mogas 92 Unl (Platts)/Brent 1st Line")]
        trades = group_rows_into_trades(rows)
        assert trades[0]["hub"] == "Sing Mogas 92 Unl (Platts)/Brent 1st Line"


class TestTapsDetection:

    # ── SM group: exact values -0.020, -0.010, 0.000, +0.010, +0.020 ─────

    def test_smt_zero_is_taps(self):
        rows = [_row("09:29:59 BST", "SMT", 10, "Aug26", 0.000)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"
        assert "TAPS" in trades[0]["notes"]

    def test_smt_minus_020_is_taps(self):
        rows = [_row("09:00:00 BST", "SMU", 10, "Aug26", -0.020)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_smt_minus_010_is_taps(self):
        rows = [_row("09:00:00 BST", "SMV", 10, "Aug26", -0.010)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_smt_plus_010_is_taps(self):
        rows = [_row("09:00:00 BST", "SMS", 10, "Aug26", +0.010)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_smt_plus_020_is_taps(self):
        rows = [_row("09:00:00 BST", "SMT", 10, "Aug26", +0.020)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_smt_non_taps_price_not_taps(self):
        rows = [_row("09:00:00 BST", "SMS", 10, "Aug26", +0.030)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "OUTRIGHT"

    # ── NJ group: exact values -0.100, -0.050, 0.000, +0.050, +0.100 ─────

    def test_njc_zero_is_taps(self):
        rows = [_row("09:00:00 BST", "NJC", 5, "Jul26", 0.000)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_njc_minus_100_is_taps(self):
        rows = [_row("09:00:00 BST", "NJD", 5, "Jul26", -0.100)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_njc_minus_050_is_taps(self):
        rows = [_row("09:00:00 BST", "NJM", 5, "Jul26", -0.050)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_njc_plus_050_is_taps(self):
        rows = [_row("09:29:59 BST", "NJB", 5, "Jul26", +0.050)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_njc_plus_100_is_taps(self):
        rows = [_row("09:29:59 BST", "NJC", 5, "Jul26", +0.100)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_njc_non_taps_price_not_taps(self):
        rows = [_row("09:00:00 BST", "NJC", 5, "Jul26", +0.075)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "OUTRIGHT"

    def test_njc_flat_price_not_taps(self):
        rows = [_row("09:00:00 BST", "NJC", 5, "Jul26", 700.0)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "OUTRIGHT"

    # ── Time and CC boundary conditions ───────────────────────────────────

    def test_taps_at_cutoff_time(self):
        """09:30:00 exactly is included."""
        rows = [_row("09:30:00 BST", "NJC", 5, "Jul26", 0.000)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_not_taps_after_cutoff(self):
        rows = [_row("09:30:01 BST", "NJC", 5, "Jul26", 0.000)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "OUTRIGHT"

    def test_just_before_cutoff_is_taps(self):
        rows = [_row("09:29:59 BST", "NJC", 5, "Jul26", 0.000)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "TAPS"

    def test_not_taps_wrong_cc(self):
        rows = [_row("09:00:00 BST", "NEC", 5, "Jul26", 0.000)]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "OUTRIGHT"

    def test_taps_cc_set_correct(self):
        assert TAPS_CC == {"SMT", "SMU", "SMV", "SMS", "NJC", "NJD", "NJM", "NJB"}

    def test_spread_never_taps(self):

        rows = [
            _row("09:00:00 BST", "NJC", 5, "Jul26", 0.000),
            _row("09:00:00 BST", "NJC", 5, "Aug26", 0.000),
        ]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "SPREAD"


# ── Butterfly trades ──────────────────────────────────────────────────────────

class TestButterflyTrades:

    def test_butterfly_detected(self):
        """Middle leg double the outer legs → BUTTERFLY; fly = (L1-L2) - (L2-L3)."""
        rows = [
            _row("10:51:00 BST", "NJC", 15, "Mar27", 655.00),
            _row("10:51:00 BST", "NJC", 30, "Apr27", 645.50),
            _row("10:51:00 BST", "NJC", 15, "May27", 638.50),
        ]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 1
        assert trades[0]["trade_type"] == "BUTTERFLY"
        assert trades[0]["spread_price"] == pytest.approx(2.50)
        assert trades[0]["qty"] == 15

    def test_butterfly_legs_sorted_sequentially(self):
        """Legs come out in strip date order regardless of input order."""
        rows = [
            _row("10:51:00 BST", "NJC", 15, "May27", 638.50),
            _row("10:51:00 BST", "NJC", 30, "Apr27", 645.50),
            _row("10:51:00 BST", "NJC", 15, "Mar27", 655.00),
        ]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "BUTTERFLY"
        strips = [l["strip"] for l in trades[0]["legs"]]
        assert strips == ["Mar27", "Apr27", "May27"]

    def test_three_equal_legs_not_butterfly(self):
        """Three legs of equal qty → SPREAD, not BUTTERFLY."""
        rows = [
            _row("12:00:00 BST", "NEC", 5, "Jul26", 700.0),
            _row("12:00:00 BST", "NEC", 5, "Aug26", 690.0),
            _row("12:00:00 BST", "NEC", 5, "Sep26", 680.0),
        ]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "SPREAD"

    def test_butterfly_wrong_middle_qty_not_butterfly(self):
        """Middle leg 1.5× outer → not a butterfly; flags as outrights."""
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Jul26", 700.0),
            _row("12:00:00 BST", "NEC", 15, "Aug26", 690.0),
            _row("12:00:00 BST", "NEC", 10, "Sep26", 680.0),
        ]
        trades = group_rows_into_trades(rows)
        # Falls into the 3+ mixed-qty → flagged outrights branch
        assert all(t["trade_type"] == "OUTRIGHT" for t in trades)
        assert all("⚠" in t["notes"] for t in trades)


# ── Orphaned Bal Month ────────────────────────────────────────────────────────

class TestOrphanedBalMonth:

    def _diff_row(self, ts, cc, qty, strip, price):
        return {"timestamp": ts, "cc": cc, "qty": qty, "strip": strip,
                "hub": "Sing Mogas 92 Unl (Platts)", "price": price,
                "is_diff_row": True, "cancelled": False}

    def test_balmo_with_diff_synthesises_spread(self):
        """Single Bal Month leg + diff row → SPREAD with synthesised leg."""
        rows = [
            _row("08:42:46 BST", "SMU", 50, "Bal Month", 117.40),
            self._diff_row("08:42:46 BST", "SMU", 50, "Bal Month/Jul26", 6.50),
        ]
        # Mark first row as not diff
        rows[0]["is_diff_row"] = False
        trades = group_rows_into_trades(rows)
        assert len(trades) == 1
        t = trades[0]
        assert t["trade_type"] == "SPREAD"
        assert t["spread_price"] == pytest.approx(6.50)
        strips = [l["strip"] for l in t["legs"]]
        assert "Bal Month" in strips
        assert "Jul26" in strips
        # Synthesised Jul26 price = 117.40 - 6.50 = 110.90
        jul_leg = next(l for l in t["legs"] if l["strip"] == "Jul26")
        assert jul_leg["price"] == pytest.approx(110.90)

    def test_balmo_negative_diff_adds(self):
        """Negative diff: Leg2 = BalMo - (negative) = BalMo + |diff|."""
        rows = [
            _row("08:42:46 BST", "SMU", 50, "Bal Month", 117.40),
            self._diff_row("08:42:46 BST", "SMU", 50, "Bal Month/Jul26", -2.50),
        ]
        rows[0]["is_diff_row"] = False
        trades = group_rows_into_trades(rows)
        jul_leg = next(l for l in trades[0]["legs"] if l["strip"] == "Jul26")
        assert jul_leg["price"] == pytest.approx(119.90)

    def test_single_balmo_no_diff_stays_outright(self):
        """Bal Month with no diff row → still an OUTRIGHT."""
        rows = [_row("08:42:46 BST", "SMU", 50, "Bal Month", 117.40)]
        rows[0]["is_diff_row"] = False
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] in ("OUTRIGHT", "TAPS")


# ── Sequential spread legs ────────────────────────────────────────────────────

class TestSequentialLegs:

    def test_spread_legs_sorted_by_strip(self):
        """Legs always come out in chronological strip order."""
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Aug26", 690.0),
            _row("12:00:00 BST", "NEC", 10, "Mar27", 620.0),
        ]
        trades = group_rows_into_trades(rows)
        assert trades[0]["trade_type"] == "SPREAD"
        strips = [l["strip"] for l in trades[0]["legs"]]
        assert strips == ["Aug26", "Mar27"]

    def test_bal_month_always_first_leg(self):
        """Bal Month sorts before any named month."""
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Jul26", 700.0),
            _row("12:00:00 BST", "NEC", 10, "Bal Month", 710.0),
        ]
        trades = group_rows_into_trades(rows)
        assert trades[0]["legs"][0]["strip"] == "Bal Month"


# ── Strip sort key ────────────────────────────────────────────────────────────

class TestStripSortKey:

    def test_bal_month_first(self):
        assert _strip_sort_key("Bal Month") < _strip_sort_key("Jul26")

    def test_month_year_order(self):
        assert _strip_sort_key("Jul26") < _strip_sort_key("Aug26")
        assert _strip_sort_key("Aug26") < _strip_sort_key("Mar27")

    def test_quarter_order(self):
        assert _strip_sort_key("Q1 26") < _strip_sort_key("Q3 26")
        assert _strip_sort_key("Q3 26") < _strip_sort_key("Q1 27")
