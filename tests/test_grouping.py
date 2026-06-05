"""Tests for trade_accumulator_v4.group_rows_into_trades."""

import pytest
from trade_accumulator_v4 import group_rows_into_trades


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

    def test_mixed_qty_flags_as_outrights(self):
        rows = [
            _row("12:00:00 BST", "NEC", 10, "Jul26", 700.0),
            _row("12:00:00 BST", "NEC", 20, "Aug26", 690.0),
        ]
        trades = group_rows_into_trades(rows)
        assert len(trades) == 2
        assert all(t["trade_type"] == "OUTRIGHT" for t in trades)
        assert all("⚠" in t["notes"] for t in trades)

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

    def test_hub_carried_through(self):
        rows = [_row("12:00:00 BST", "STB", 25, "Jul26", 17.0,
                     hub="Sing Mogas 92 Unl (Platts)/Brent 1st Line")]
        trades = group_rows_into_trades(rows)
        assert trades[0]["hub"] == "Sing Mogas 92 Unl (Platts)/Brent 1st Line"
