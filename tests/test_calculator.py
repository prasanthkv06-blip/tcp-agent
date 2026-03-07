"""
Tests for calculator.py
========================
Covers:
  - Interpolation (kept from original — 8 tests)
  - Weather exclusion formula (Rule 2)
  - Segment detection (Rule 3)
  - Boundary pro-rating (Rule 3)
  - Per-segment performance calculations
  - Voyage totals
"""

import sys
from pathlib import Path

import numpy as np
import pytest

# Ensure the app dir is importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


# =============================================================================
# Interpolation  (Appendix A — preserved)
# =============================================================================

class TestInterpolation:
    def test_exact_speed(self):
        from calculator import interpolate_consumption
        # 19.5 kts laden base gas = 68.1 (from config)
        assert interpolate_consumption(19.5, "laden", "base_gas") == 68.1

    def test_exact_speed_ballast(self):
        from calculator import interpolate_consumption
        # 19.5 kts ballast base gas = 67.2
        assert interpolate_consumption(19.5, "ballast", "base_gas") == 67.2

    def test_interpolation_midpoint(self):
        from calculator import interpolate_consumption
        # 14.25 is midpoint between 14.0 (29.7) and 14.5 (32.1)
        result = interpolate_consumption(14.25, "laden", "base_gas")
        expected = 29.7 + 0.5 * (32.1 - 29.7)
        assert abs(result - expected) < 0.01

    def test_clamp_below_minimum(self):
        from calculator import interpolate_consumption
        # Below 12.0 should clamp to 12.0 value
        result = interpolate_consumption(10.0, "laden", "base_gas")
        assert result == 21.8

    def test_clamp_above_maximum(self):
        from calculator import interpolate_consumption
        result = interpolate_consumption(25.0, "laden", "base_gas")
        assert result == 68.1

    def test_pilot_component(self):
        from calculator import interpolate_consumption
        result = interpolate_consumption(19.5, "laden", "pilot")
        assert result == 1.2

    def test_guaranteed_daily_consumption(self):
        from calculator import get_guaranteed_daily_consumption
        result = get_guaranteed_daily_consumption(19.5, "laden")
        assert result["base_gas_mt"] == 68.1
        assert result["pilot_mt"] == 1.2
        assert result["boiler_mt"] == 1.3
        assert abs(result["total_fuel_mt"] - (68.1 + 1.2 + 1.3)) < 0.01

    def test_ageing_factor(self):
        from calculator import get_guaranteed_daily_consumption
        result = get_guaranteed_daily_consumption(19.5, "laden", ageing_factor=1.05)
        assert abs(result["base_gas_mt"] - 68.1 * 1.05) < 0.01


# =============================================================================
# Weather Exclusion  (Rule 2)
# =============================================================================

class TestWeatherExclusion:
    """Tests for compute_weather_exclusion()."""

    def test_basic_exclusion(self):
        """BF5=12 hrs > 6 threshold → exclude proportional fuel."""
        from calculator import compute_weather_exclusion
        # daily_cons=2.4 MT, steaming=24 hrs, bf5=12 hrs
        result = compute_weather_exclusion(2.4, 24.0, 12.0)
        # (2.4 / 24) × 12 = 1.2
        assert abs(result - 1.2) < 0.001

    def test_below_threshold_returns_zero(self):
        """BF5=5 hrs ≤ 6 threshold → no exclusion."""
        from calculator import compute_weather_exclusion
        result = compute_weather_exclusion(2.4, 24.0, 5.0)
        assert result == 0.0

    def test_exactly_at_threshold_returns_zero(self):
        """BF5=6 hrs = threshold → no exclusion (must be > threshold)."""
        from calculator import compute_weather_exclusion
        result = compute_weather_exclusion(2.4, 24.0, 6.0)
        assert result == 0.0

    def test_bf5_capped_at_steaming_hours(self):
        """BF5=24 hrs but steaming=23 hrs → cap at 23 (timezone scenario)."""
        from calculator import compute_weather_exclusion
        # daily_cons=0.60, steaming=23, bf5=24
        result = compute_weather_exclusion(0.60, 23.0, 24.0)
        # (0.60 / 23) × MIN(24, 23) = (0.60 / 23) × 23 = 0.60
        assert abs(result - 0.60) < 0.001

    def test_bf5_equals_steaming_excludes_full_day(self):
        """When BF5 = steaming hours, entire day is excluded."""
        from calculator import compute_weather_exclusion
        result = compute_weather_exclusion(3.0, 24.0, 24.0)
        # (3.0 / 24) × 24 = 3.0
        assert abs(result - 3.0) < 0.001

    def test_zero_steaming_returns_zero(self):
        """Zero steaming hours → no exclusion (avoid division by zero)."""
        from calculator import compute_weather_exclusion
        result = compute_weather_exclusion(2.4, 0.0, 12.0)
        assert result == 0.0

    def test_zero_consumption_returns_zero(self):
        """Zero daily consumption → no exclusion."""
        from calculator import compute_weather_exclusion
        result = compute_weather_exclusion(0.0, 24.0, 12.0)
        assert result == 0.0

    def test_negative_consumption_returns_zero(self):
        """Negative daily consumption → no exclusion."""
        from calculator import compute_weather_exclusion
        result = compute_weather_exclusion(-1.0, 24.0, 12.0)
        assert result == 0.0


class TestTagWeatherExclusions:
    """Tests for tag_weather_exclusions() — in-place pre-tagging."""

    def test_tags_excluded_row(self):
        """Row with BF5 > 6 gets weather_excluded=True and correct amounts."""
        from calculator import tag_weather_exclusions
        rows = [{
            "bf5_hours": 12.0,
            "steaming_hrs": 24.0,
            "mgo_daily": 2.4,
            "lng_daily": 100.0,
            "vlsfo_daily": 1.0,
            "distance": 400.0,
        }]
        tag_weather_exclusions(rows)
        assert rows[0]["weather_excluded"] is True
        assert abs(rows[0]["excl_mgo"] - 1.2) < 0.001
        assert abs(rows[0]["excl_lng"] - 50.0) < 0.001
        assert abs(rows[0]["excl_vlsfo"] - 0.5) < 0.001

    def test_tags_non_excluded_row(self):
        """Row with BF5 ≤ 6 gets weather_excluded=False and zero amounts."""
        from calculator import tag_weather_exclusions
        rows = [{
            "bf5_hours": 3.0,
            "steaming_hrs": 24.0,
            "mgo_daily": 2.4,
            "lng_daily": 100.0,
            "vlsfo_daily": 1.0,
            "distance": 400.0,
        }]
        tag_weather_exclusions(rows)
        assert rows[0]["weather_excluded"] is False
        assert rows[0]["excl_mgo"] == 0.0
        assert rows[0]["excl_lng"] == 0.0
        assert rows[0]["excl_vlsfo"] == 0.0

    def test_mixed_rows(self):
        """Mixture of excluded and non-excluded rows."""
        from calculator import tag_weather_exclusions
        rows = [
            {"bf5_hours": 0.0, "steaming_hrs": 24.0, "mgo_daily": 1.0,
             "lng_daily": 50.0, "vlsfo_daily": 0.5, "distance": 400.0},
            {"bf5_hours": 10.0, "steaming_hrs": 24.0, "mgo_daily": 1.2,
             "lng_daily": 60.0, "vlsfo_daily": 0.6, "distance": 380.0},
            {"bf5_hours": 4.0, "steaming_hrs": 24.0, "mgo_daily": 1.0,
             "lng_daily": 50.0, "vlsfo_daily": 0.5, "distance": 400.0},
        ]
        tag_weather_exclusions(rows)
        assert rows[0]["weather_excluded"] is False
        assert rows[1]["weather_excluded"] is True
        assert rows[2]["weather_excluded"] is False
        # Only row 1 has exclusion amounts
        assert rows[1]["excl_mgo"] > 0
        assert rows[0]["excl_mgo"] == 0.0
        assert rows[2]["excl_mgo"] == 0.0


class TestApplyWeatherExclusions:
    """Tests for apply_weather_exclusions() — voyage-level totals."""

    def test_sums_exclusions(self):
        """Sums exclusion amounts across all rows."""
        from calculator import apply_weather_exclusions
        rows = [
            {"bf5_hours": 12.0, "steaming_hrs": 24.0, "mgo_daily": 2.0,
             "lng_daily": 100.0, "vlsfo_daily": 1.0, "distance": 400.0},
            {"bf5_hours": 8.0, "steaming_hrs": 24.0, "mgo_daily": 3.0,
             "lng_daily": 150.0, "vlsfo_daily": 1.5, "distance": 380.0},
            {"bf5_hours": 2.0, "steaming_hrs": 24.0, "mgo_daily": 2.0,
             "lng_daily": 100.0, "vlsfo_daily": 1.0, "distance": 400.0},
        ]
        result = apply_weather_exclusions(rows)
        # Row 0: (2.0/24)*12 = 1.0, Row 1: (3.0/24)*8 = 1.0, Row 2: no exclusion
        assert abs(result["excluded_mgo"] - 2.0) < 0.01
        assert result["total_bf5_hours"] == 22.0

    def test_no_exclusions(self):
        """No rows exceed threshold → all zeros."""
        from calculator import apply_weather_exclusions
        rows = [
            {"bf5_hours": 0.0, "steaming_hrs": 24.0, "mgo_daily": 2.0,
             "lng_daily": 100.0, "vlsfo_daily": 1.0, "distance": 400.0},
            {"bf5_hours": 5.0, "steaming_hrs": 24.0, "mgo_daily": 2.0,
             "lng_daily": 100.0, "vlsfo_daily": 1.0, "distance": 400.0},
        ]
        result = apply_weather_exclusions(rows)
        assert result["excluded_mgo"] == 0.0
        assert result["excluded_lng"] == 0.0
        assert result["excluded_vlsfo"] == 0.0
        assert len(result["excluded_rows"]) == 0


# =============================================================================
# Segment Detection  (Rule 3)
# =============================================================================

class TestDetectSegments:
    """Tests for detect_segments()."""

    def _make_rows(self, n, **overrides):
        """Create n simple daily_rows with optional per-row overrides."""
        rows = []
        for i in range(n):
            row = {
                "df_idx": i,
                "datetime": f"2024-12-{5+i:02d} 12:00:00",
                "ordered_speed": 19.5,
                "next_port": "ALIAGA",
                "fuel_mode": "LNG ONLY",
                "voyage_order_rev": "",
                "rev_start_time": None,
                "rev_gmt_offset": 0,
                "rev_speed": 0.0,
                "rev_sat": None,
                "distance": 400.0,
                "steaming_hrs": 24.0,
                "mgo_daily": 1.0,
                "lng_daily": 100.0,
                "vlsfo_daily": 0.5,
                "bf5_hours": 0.0,
                "avg_speed": 17.5,
            }
            # Apply per-row overrides
            for key, vals in overrides.items():
                if i < len(vals):
                    row[key] = vals[i]
            rows.append(row)
        return rows

    def test_single_segment_no_changes(self):
        """No field changes → 1 segment covering all rows."""
        from calculator import detect_segments
        rows = self._make_rows(5)
        segments = detect_segments(rows, "2024-12-05 08:00", "2024-12-09 14:00")
        assert len(segments) == 1
        assert segments[0]["start_row_idx"] == 0
        assert segments[0]["end_row_idx"] == 4
        assert segments[0]["instructed_speed"] == 19.5

    def test_field_change_next_port(self):
        """Next port changes mid-voyage → 2 segments."""
        from calculator import detect_segments
        rows = self._make_rows(6, next_port=[
            "ALIAGA", "ALIAGA", "ALIAGA",
            "DORTYOL", "DORTYOL", "DORTYOL",
        ])
        segments = detect_segments(rows, "2024-12-05 08:00", "2024-12-10 14:00")
        assert len(segments) == 2
        # Seg 1 ends at row 2 (field change boundary → row goes to new segment)
        assert segments[0]["end_row_idx"] == 2
        assert segments[0]["boundary_type"] == "field_change"
        # Seg 2 starts at row 3
        assert segments[1]["start_row_idx"] == 3
        assert segments[1]["end_row_idx"] == 5

    def test_field_change_fuel_mode(self):
        """Fuel mode changes → new segment."""
        from calculator import detect_segments
        rows = self._make_rows(4, fuel_mode=[
            "LNG ONLY", "LNG ONLY", "FUEL MIX", "FUEL MIX",
        ])
        segments = detect_segments(rows, "2024-12-05 08:00", "2024-12-08 14:00")
        assert len(segments) == 2

    def test_voyage_order_revision(self):
        """BM='yes' on a row → shared boundary, new segment."""
        from calculator import detect_segments
        rows = self._make_rows(6, voyage_order_rev=[
            "", "", "", "yes", "", "",
        ], rev_start_time=[
            None, None, None, "2024-12-08 06:00:00", None, None,
        ], rev_speed=[
            0, 0, 0, 18.2, 0, 0,
        ])
        segments = detect_segments(rows, "2024-12-05 08:00", "2024-12-10 14:00")
        assert len(segments) == 2
        # Seg 1: shared boundary at row 3
        assert segments[0]["end_row_idx"] == 3
        assert segments[0]["is_boundary_shared"] is True
        assert segments[0]["boundary_type"] == "voyage_order"
        # Seg 2: starts at row 3 (same row), new speed
        assert segments[1]["start_row_idx"] == 3
        assert segments[1]["instructed_speed"] == 18.2

    def test_arrival_row_field_change_ignored(self):
        """Field change on the ARRIVAL row (last row) should NOT create new segment."""
        from calculator import detect_segments
        rows = self._make_rows(5, next_port=[
            "ALIAGA", "ALIAGA", "ALIAGA", "ALIAGA", "DORTYOL",
        ])
        segments = detect_segments(rows, "2024-12-05 08:00", "2024-12-09 14:00")
        # Should be 1 segment, not 2 (ARRIVAL row change is ignored)
        assert len(segments) == 1

    def test_empty_rows(self):
        """Empty daily_rows → empty segments."""
        from calculator import detect_segments
        assert detect_segments([], "2024-12-05", "2024-12-09") == []


# =============================================================================
# Boundary Pro-Rating  (Rule 3)
# =============================================================================

class TestBoundaryProRating:
    """Tests for prorate_boundary_row()."""

    def test_basic_prorate(self):
        """Speed-weighted pro-rate with known hours and speeds."""
        from calculator import prorate_boundary_row
        rows = [
            {"datetime": "2024-12-09 12:00", "distance": 100.0,
             "steaming_hrs": 24.0, "mgo_daily": 2.0, "lng_daily": 100.0,
             "vlsfo_daily": 1.0, "vlsfo_g1_daily": 0.5, "vlsfo_g2_daily": 0.5,
             "bf5_hours": 0.0, "mgo_boiler": 0.1, "mgo_pilot": 0.05,
             "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
             "vlsfo_g2_boiler": 0.02, "vlsfo_g2_pilot": 0.01,
             "gcu_lng": 0.0, "reliq_hours": 5.0, "subcooler_hours": 3.0,
             "excl_hours": 0.0, "excl_mgo": 0.0, "excl_lng": 0.0,
             "excl_vlsfo": 0.0, "excl_distance": 0.0, "weather_excluded": False,
             "df_idx": 5, "remarks": ""},
            {"datetime": "2024-12-10 12:00", "distance": 374.2,
             "steaming_hrs": 24.0, "mgo_daily": 2.5, "lng_daily": 150.0,
             "vlsfo_daily": 1.2, "vlsfo_g1_daily": 0.6, "vlsfo_g2_daily": 0.6,
             "bf5_hours": 0.0, "mgo_boiler": 0.1, "mgo_pilot": 0.05,
             "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
             "vlsfo_g2_boiler": 0.02, "vlsfo_g2_pilot": 0.01,
             "gcu_lng": 0.0, "reliq_hours": 5.0, "subcooler_hours": 3.0,
             "excl_hours": 0.0, "excl_mgo": 0.0, "excl_lng": 0.0,
             "excl_vlsfo": 0.0, "excl_distance": 0.0, "weather_excluded": False,
             "df_idx": 6, "remarks": ""},
        ]
        # Boundary at row 1:
        # prev_dt = 2024-12-09 12:00, rev_dt = 2024-12-10 03:00, curr_dt = 2024-12-10 12:00
        # hours_before = 15, hours_after = 9
        # speed_before = 19.5, speed_after = 18.2
        # theo_before = 19.5 * 15 = 292.5
        # theo_after  = 18.2 * 9  = 163.8
        # total = 456.3
        # r_before = 292.5 / 456.3 ≈ 0.6410
        # r_after  = 163.8 / 456.3 ≈ 0.3590
        before, after = prorate_boundary_row(
            rows, 1,
            "2024-12-09 12:00", "2024-12-10 03:00", "2024-12-10 12:00",
            19.5, 18.2,
        )
        r_before = (19.5 * 15) / (19.5 * 15 + 18.2 * 9)
        r_after = 1 - r_before

        assert abs(before["distance"] - 374.2 * r_before) < 0.1
        assert abs(after["distance"] - 374.2 * r_after) < 0.1
        assert abs(before["distance"] + after["distance"] - 374.2) < 0.01

    def test_prorate_preserves_total(self):
        """Before + after portions always sum to original values."""
        from calculator import prorate_boundary_row
        row = {
            "datetime": "2024-12-10 12:00", "distance": 500.0,
            "steaming_hrs": 24.0, "mgo_daily": 3.0, "lng_daily": 200.0,
            "vlsfo_daily": 2.0, "vlsfo_g1_daily": 1.0, "vlsfo_g2_daily": 1.0,
            "bf5_hours": 8.0, "mgo_boiler": 0.15, "mgo_pilot": 0.06,
            "vlsfo_g1_boiler": 0.03, "vlsfo_g1_pilot": 0.02,
            "vlsfo_g2_boiler": 0.03, "vlsfo_g2_pilot": 0.02,
            "gcu_lng": 5.0, "reliq_hours": 10.0, "subcooler_hours": 8.0,
            "excl_hours": 2.67, "excl_mgo": 0.33, "excl_lng": 22.2,
            "excl_vlsfo": 0.22, "excl_distance": 55.6, "weather_excluded": True,
            "df_idx": 3, "remarks": "test",
        }
        before, after = prorate_boundary_row(
            [{"datetime": "2024-12-09 12:00"}, row], 1,
            "2024-12-09 12:00", "2024-12-10 00:00", "2024-12-10 12:00",
            19.0, 17.0,
        )
        for key in ["distance", "steaming_hrs", "mgo_daily", "lng_daily",
                     "vlsfo_daily", "excl_mgo", "excl_lng"]:
            total = before[key] + after[key]
            assert abs(total - row[key]) < 0.001, f"{key}: {total} != {row[key]}"


# =============================================================================
# make_portion  (helper for pro-rating)
# =============================================================================

class TestMakePortion:
    def test_ratio_zero(self):
        """Ratio=0 → all values are 0."""
        from calculator import _make_portion
        row = {
            "distance": 400, "steaming_hrs": 24, "mgo_daily": 2.0,
            "lng_daily": 100, "vlsfo_daily": 1.0, "vlsfo_g1_daily": 0.5,
            "vlsfo_g2_daily": 0.5, "bf5_hours": 0, "mgo_boiler": 0.1,
            "mgo_pilot": 0.05, "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
            "vlsfo_g2_boiler": 0.02, "vlsfo_g2_pilot": 0.01,
            "gcu_lng": 0, "reliq_hours": 5, "subcooler_hours": 3,
            "excl_hours": 0, "excl_mgo": 0, "excl_lng": 0,
            "excl_vlsfo": 0, "excl_distance": 0, "weather_excluded": False,
            "df_idx": 1, "datetime": "2024-12-06", "remarks": "",
        }
        portion = _make_portion(row, 0.0)
        assert portion["distance"] == 0.0
        assert portion["mgo_daily"] == 0.0
        assert portion["is_prorated"] is True

    def test_ratio_one(self):
        """Ratio=1 → values equal original."""
        from calculator import _make_portion
        row = {
            "distance": 400, "steaming_hrs": 24, "mgo_daily": 2.0,
            "lng_daily": 100, "vlsfo_daily": 1.0, "vlsfo_g1_daily": 0.5,
            "vlsfo_g2_daily": 0.5, "bf5_hours": 0, "mgo_boiler": 0.1,
            "mgo_pilot": 0.05, "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
            "vlsfo_g2_boiler": 0.02, "vlsfo_g2_pilot": 0.01,
            "gcu_lng": 0, "reliq_hours": 5, "subcooler_hours": 3,
            "excl_hours": 0, "excl_mgo": 0, "excl_lng": 0,
            "excl_vlsfo": 0, "excl_distance": 0, "weather_excluded": False,
            "df_idx": 1, "datetime": "2024-12-06", "remarks": "",
        }
        portion = _make_portion(row, 1.0)
        assert portion["distance"] == 400.0
        assert portion["mgo_daily"] == 2.0


# =============================================================================
# Per-Segment Performance  (compute_segment_data)
# =============================================================================

class TestComputeSegmentData:
    """Tests for compute_segment_data()."""

    def test_basic_segment(self):
        """Simple segment with known values."""
        from calculator import compute_segment_data
        seg_info = {
            "start_datetime": "2024-12-05 08:00:00",
            "end_datetime": "2024-12-07 14:00:00",
            "instructed_speed": 19.5,
            "fuel_mode": "LNG ONLY",
        }
        seg_rows = [
            {"distance": 0, "steaming_hrs": 0, "mgo_daily": 0, "lng_daily": 0,
             "vlsfo_daily": 0, "vlsfo_g1_daily": 0, "vlsfo_g2_daily": 0,
             "mgo_boiler": 0, "mgo_pilot": 0,
             "vlsfo_g1_boiler": 0, "vlsfo_g1_pilot": 0,
             "vlsfo_g2_boiler": 0, "vlsfo_g2_pilot": 0,
             "gcu_lng": 0, "reliq_hours": 0, "subcooler_hours": 0,
             "reliq_load": 0, "subcooler_load": 0,
             "bf5_hours": 0, "excl_hours": 0, "excl_mgo": 0,
             "excl_lng": 0, "excl_vlsfo": 0, "excl_distance": 0,
             "weather_excluded": False, "remarks": "", "datetime": "2024-12-05"},
            {"distance": 450, "steaming_hrs": 24, "mgo_daily": 2.0, "lng_daily": 150.0,
             "vlsfo_daily": 1.0, "vlsfo_g1_daily": 0.6, "vlsfo_g2_daily": 0.4,
             "mgo_boiler": 0.1, "mgo_pilot": 0.05,
             "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
             "vlsfo_g2_boiler": 0.02, "vlsfo_g2_pilot": 0.01,
             "gcu_lng": 0, "reliq_hours": 10, "subcooler_hours": 5,
             "reliq_load": 50, "subcooler_load": 30,
             "bf5_hours": 0, "excl_hours": 0, "excl_mgo": 0,
             "excl_lng": 0, "excl_vlsfo": 0, "excl_distance": 0,
             "weather_excluded": False, "remarks": "", "datetime": "2024-12-06"},
            {"distance": 420, "steaming_hrs": 24, "mgo_daily": 1.8, "lng_daily": 140.0,
             "vlsfo_daily": 0.8, "vlsfo_g1_daily": 0.5, "vlsfo_g2_daily": 0.3,
             "mgo_boiler": 0.1, "mgo_pilot": 0.05,
             "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
             "vlsfo_g2_boiler": 0.02, "vlsfo_g2_pilot": 0.01,
             "gcu_lng": 0, "reliq_hours": 10, "subcooler_hours": 5,
             "reliq_load": 50, "subcooler_load": 30,
             "bf5_hours": 0, "excl_hours": 0, "excl_mgo": 0,
             "excl_lng": 0, "excl_vlsfo": 0, "excl_distance": 0,
             "weather_excluded": False, "remarks": "test remark", "datetime": "2024-12-07"},
        ]
        result = compute_segment_data(seg_info, seg_rows)

        # Distance
        assert abs(result["distance"] - 870) < 0.1
        # Fuel
        assert abs(result["mgo_consumed"] - 3.8) < 0.01
        assert abs(result["lng_consumed"] - 290.0) < 0.1
        assert abs(result["vlsfo_consumed"] - 1.8) < 0.01
        # Duration: Dec 5 08:00 to Dec 7 14:00 = 2.25 days
        assert abs(result["duration_days"] - 2.25) < 0.01
        # Speed = 870 / (2.25 * 24) = 16.11 kts (no exclusions)
        expected_speed = 870 / (2.25 * 24)
        assert abs(result["actual_avg_speed"] - expected_speed) < 0.1
        # Reference speed = MIN(actual, instructed) = MIN(16.11, 19.5) = 16.11
        assert result["reference_speed"] <= result["instructed_speed"]

    def test_segment_with_weather_exclusions(self):
        """Segment with weather exclusions reduces net fuel and distance."""
        from calculator import compute_segment_data
        seg_info = {
            "start_datetime": "2024-12-05 08:00:00",
            "end_datetime": "2024-12-07 08:00:00",
            "instructed_speed": 19.5,
            "fuel_mode": "LNG ONLY",
        }
        seg_rows = [
            {"distance": 0, "steaming_hrs": 0, "mgo_daily": 0, "lng_daily": 0,
             "vlsfo_daily": 0, "vlsfo_g1_daily": 0, "vlsfo_g2_daily": 0,
             "mgo_boiler": 0, "mgo_pilot": 0,
             "vlsfo_g1_boiler": 0, "vlsfo_g1_pilot": 0,
             "vlsfo_g2_boiler": 0, "vlsfo_g2_pilot": 0,
             "gcu_lng": 0, "reliq_hours": 0, "subcooler_hours": 0,
             "reliq_load": 0, "subcooler_load": 0,
             "bf5_hours": 0, "excl_hours": 0, "excl_mgo": 0,
             "excl_lng": 0, "excl_vlsfo": 0, "excl_distance": 0,
             "weather_excluded": False, "remarks": "", "datetime": "2024-12-05"},
            {"distance": 400, "steaming_hrs": 24, "mgo_daily": 2.4, "lng_daily": 120,
             "vlsfo_daily": 1.0, "vlsfo_g1_daily": 1.0, "vlsfo_g2_daily": 0,
             "mgo_boiler": 0.1, "mgo_pilot": 0.05,
             "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
             "vlsfo_g2_boiler": 0, "vlsfo_g2_pilot": 0,
             "gcu_lng": 0, "reliq_hours": 0, "subcooler_hours": 0,
             "reliq_load": 0, "subcooler_load": 0,
             "bf5_hours": 12, "excl_hours": 12, "excl_mgo": 1.2,
             "excl_lng": 60, "excl_vlsfo": 0.5, "excl_distance": 200,
             "weather_excluded": True, "remarks": "", "datetime": "2024-12-06"},
        ]
        result = compute_segment_data(seg_info, seg_rows)
        # Net fuel = actual - excluded
        assert abs(result["net_mgo"] - (2.4 - 1.2)) < 0.01
        assert abs(result["net_lng"] - (120 - 60)) < 0.1
        # Net distance = 400 - 200 = 200
        assert abs(result["net_distance"] - 200) < 0.1


# =============================================================================
# Voyage Totals  (compute_voyage_totals)
# =============================================================================

class TestComputeVoyageTotals:
    """Tests for compute_voyage_totals()."""

    def test_sums_across_segments(self):
        """Totals sum distance, fuel, exclusions across segments."""
        from calculator import compute_voyage_totals
        segments = [
            {
                "start_datetime": "2024-12-05 08:00",
                "end_datetime": "2024-12-10 06:00",
                "distance": 1762.7, "duration_days": 4.92,
                "weather_excl_hours": 12.0, "other_excl_hours": 0.0,
                "total_excl_hours": 12.0,
                "weather_excl_distance": 200.0,
                "regulatory_excl_hours": 0.0, "regulatory_excl_distance": 0.0,
                "total_speed_excl_hours": 12.0, "total_speed_excl_distance": 200.0,
                "net_duration_days": 4.42, "net_distance": 1562.7,
                "actual_avg_speed": 14.73,
                "instructed_speed": 19.5,
                "reference_speed": 14.73,
                "lng_consumed": 800.0, "mgo_consumed": 15.0, "vlsfo_consumed": 8.0,
                "vlsfo_g1_consumed": 5.0, "vlsfo_g2_consumed": 3.0,
                "mgo_pilot": 1.0, "mgo_boiler": 2.0, "mgo_propulsion": 12.0,
                "vlsfo_pilot": 0.5, "vlsfo_boiler": 1.0, "vlsfo_propulsion": 6.5,
                "excl_lng_weather": 100.0, "excl_mgo_weather": 3.0,
                "excl_vlsfo_weather": 2.0,
                "excl_lng_other": 0, "excl_mgo_other": 0, "excl_vlsfo_other": 0,
                "excl_lng_total": 100.0, "excl_mgo_total": 3.0,
                "excl_vlsfo_total": 2.0,
                "net_lng": 700.0, "net_mgo": 12.0, "net_vlsfo": 6.0,
                "gcu_total": 0.0, "gcu_used": False, "gcu_dates": [],
                "reliq_hours": 48.0, "reliq_avg_load": 45.0,
                "weather_bf5_hours": 24.0,
                "fuel_mode": "LNG ONLY",
                "remarks": [],
            },
            {
                "start_datetime": "2024-12-10 06:00",
                "end_datetime": "2024-12-22 04:30",
                "distance": 5028.2, "duration_days": 11.94,
                "weather_excl_hours": 20.0, "other_excl_hours": 0.0,
                "total_excl_hours": 20.0,
                "weather_excl_distance": 350.0,
                "regulatory_excl_hours": 0.0, "regulatory_excl_distance": 0.0,
                "total_speed_excl_hours": 20.0, "total_speed_excl_distance": 350.0,
                "net_duration_days": 11.11, "net_distance": 4678.2,
                "actual_avg_speed": 17.55,
                "instructed_speed": 18.2,
                "reference_speed": 17.55,
                "lng_consumed": 1803.0, "mgo_consumed": 26.3, "vlsfo_consumed": 11.75,
                "vlsfo_g1_consumed": 7.0, "vlsfo_g2_consumed": 4.75,
                "mgo_pilot": 2.5, "mgo_boiler": 5.0, "mgo_propulsion": 18.8,
                "vlsfo_pilot": 1.0, "vlsfo_boiler": 2.5, "vlsfo_propulsion": 8.25,
                "excl_lng_weather": 315.0, "excl_mgo_weather": 7.26,
                "excl_vlsfo_weather": 3.25,
                "excl_lng_other": 0, "excl_mgo_other": 0, "excl_vlsfo_other": 0,
                "excl_lng_total": 315.0, "excl_mgo_total": 7.26,
                "excl_vlsfo_total": 3.25,
                "net_lng": 1488.0, "net_mgo": 19.04, "net_vlsfo": 8.5,
                "gcu_total": 0.0, "gcu_used": False, "gcu_dates": [],
                "reliq_hours": 96.0, "reliq_avg_load": 50.0,
                "weather_bf5_hours": 40.0,
                "fuel_mode": "LNG ONLY",
                "remarks": [],
            },
        ]
        totals = compute_voyage_totals(segments)

        # Distance
        assert abs(totals["distance"] - 6790.9) < 0.1
        # Fuel
        assert abs(totals["lng_consumed"] - 2603.0) < 0.1
        assert abs(totals["mgo_consumed"] - 41.3) < 0.1
        assert abs(totals["vlsfo_consumed"] - 19.75) < 0.1
        # Exclusions
        assert abs(totals["excl_mgo_weather"] - 10.26) < 0.01
        # Net
        assert abs(totals["net_mgo"] - 31.04) < 0.01
        # Speed = total net_dist / (total net_dur * 24)
        expected_speed = totals["net_distance"] / (totals["net_duration_days"] * 24)
        assert abs(totals["actual_avg_speed"] - expected_speed) < 0.1
        # Instructed speed: varies → None
        assert totals["instructed_speed"] is None

    def test_empty_segments(self):
        """Empty segments list → empty dict."""
        from calculator import compute_voyage_totals
        assert compute_voyage_totals([]) == {}


# =============================================================================
# Speed Formula
# =============================================================================

class TestSpeedFormula:
    """Verify speed = net_distance / (net_duration_days × 24)."""

    def test_speed_formula(self):
        from calculator import compute_segment_data
        seg_info = {
            "start_datetime": "2024-12-05 00:00:00",
            "end_datetime": "2024-12-10 00:00:00",  # 5 days
            "instructed_speed": 18.0,
            "fuel_mode": "LNG ONLY",
        }
        # Total distance = 2000 nm, no exclusions
        seg_rows = [
            {"distance": 0, "steaming_hrs": 0, "mgo_daily": 0, "lng_daily": 0,
             "vlsfo_daily": 0, "vlsfo_g1_daily": 0, "vlsfo_g2_daily": 0,
             "mgo_boiler": 0, "mgo_pilot": 0,
             "vlsfo_g1_boiler": 0, "vlsfo_g1_pilot": 0,
             "vlsfo_g2_boiler": 0, "vlsfo_g2_pilot": 0,
             "gcu_lng": 0, "reliq_hours": 0, "subcooler_hours": 0,
             "reliq_load": 0, "subcooler_load": 0,
             "bf5_hours": 0, "excl_hours": 0, "excl_mgo": 0,
             "excl_lng": 0, "excl_vlsfo": 0, "excl_distance": 0,
             "weather_excluded": False, "remarks": "", "datetime": "2024-12-05"},
        ]
        # Add 5 rows each with 400 nm
        for d in range(1, 6):
            seg_rows.append({
                "distance": 400, "steaming_hrs": 24, "mgo_daily": 2, "lng_daily": 100,
                "vlsfo_daily": 1, "vlsfo_g1_daily": 0.5, "vlsfo_g2_daily": 0.5,
                "mgo_boiler": 0.1, "mgo_pilot": 0.05,
                "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
                "vlsfo_g2_boiler": 0.02, "vlsfo_g2_pilot": 0.01,
                "gcu_lng": 0, "reliq_hours": 0, "subcooler_hours": 0,
                "reliq_load": 0, "subcooler_load": 0,
                "bf5_hours": 0, "excl_hours": 0, "excl_mgo": 0,
                "excl_lng": 0, "excl_vlsfo": 0, "excl_distance": 0,
                "weather_excluded": False, "remarks": "",
                "datetime": f"2024-12-{5+d:02d}",
            })

        result = compute_segment_data(seg_info, seg_rows)
        # Speed = 2000 / (5 * 24) = 16.667 kts
        expected = 2000.0 / (5.0 * 24.0)
        assert abs(result["actual_avg_speed"] - expected) < 0.01

    def test_reference_speed_is_min(self):
        """Reference speed = MIN(actual, instructed)."""
        from calculator import compute_segment_data
        seg_info = {
            "start_datetime": "2024-12-05 00:00:00",
            "end_datetime": "2024-12-06 00:00:00",  # 1 day
            "instructed_speed": 15.0,
            "fuel_mode": "LNG ONLY",
        }
        # Distance 480 nm in 1 day → actual = 480/24 = 20 kts > instructed 15
        seg_rows = [{
            "distance": 480, "steaming_hrs": 24, "mgo_daily": 2, "lng_daily": 100,
            "vlsfo_daily": 1, "vlsfo_g1_daily": 0.5, "vlsfo_g2_daily": 0.5,
            "mgo_boiler": 0.1, "mgo_pilot": 0.05,
            "vlsfo_g1_boiler": 0.02, "vlsfo_g1_pilot": 0.01,
            "vlsfo_g2_boiler": 0.02, "vlsfo_g2_pilot": 0.01,
            "gcu_lng": 0, "reliq_hours": 0, "subcooler_hours": 0,
            "reliq_load": 0, "subcooler_load": 0,
            "bf5_hours": 0, "excl_hours": 0, "excl_mgo": 0,
            "excl_lng": 0, "excl_vlsfo": 0, "excl_distance": 0,
            "weather_excluded": False, "remarks": "", "datetime": "2024-12-05",
        }]
        result = compute_segment_data(seg_info, seg_rows)
        # Actual = 20 kts, Instructed = 15 kts → Reference = 15
        assert result["reference_speed"] == 15.0


# =============================================================================
# Speed Anomaly Detection (Rule A)
# =============================================================================

class TestSpeedAnomalyDetection:
    """Tests for detect_speed_anomalies() — Rule A."""

    def test_no_anomaly_constant_speed(self):
        """Constant speed → no anomalies flagged."""
        from calculator import detect_speed_anomalies
        rows = [
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-01", "df_idx": 0, "report_type": "NOON"},
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-02", "df_idx": 1, "report_type": "NOON"},
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-03", "df_idx": 2, "report_type": "NOON"},
        ]
        assert detect_speed_anomalies(rows) == []

    def test_zero_speed_flagged(self):
        """A row with 0 speed after normal rows → flagged."""
        from calculator import detect_speed_anomalies
        rows = [
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-01", "df_idx": 0, "report_type": "NOON"},
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-02", "df_idx": 1, "report_type": "NOON"},
            {"distance": 0, "steaming_hrs": 24, "avg_speed": 0.0,
             "datetime": "2024-12-03", "df_idx": 2, "report_type": "NOON"},
        ]
        flagged = detect_speed_anomalies(rows)
        assert len(flagged) == 1
        assert flagged[0]["df_idx"] == 2
        assert flagged[0]["avg_speed"] == 0.0

    def test_slow_row_below_threshold(self):
        """A row with speed < 10% of weighted avg → flagged."""
        from calculator import detect_speed_anomalies
        rows = [
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-01", "df_idx": 0, "report_type": "NOON"},
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-02", "df_idx": 1, "report_type": "NOON"},
            {"distance": 10, "steaming_hrs": 24, "avg_speed": 0.4,
             "datetime": "2024-12-03", "df_idx": 2, "report_type": "NOON"},
        ]
        flagged = detect_speed_anomalies(rows)
        assert len(flagged) == 1
        assert flagged[0]["avg_speed"] == 0.4

    def test_slow_row_above_threshold_not_flagged(self):
        """A row with speed just above 10% of weighted avg → NOT flagged."""
        from calculator import detect_speed_anomalies
        # Weighted avg after 2 rows: 800/48 = 16.67 kts
        # 10% threshold = 1.667 kts → speed 2.0 should NOT be flagged
        rows = [
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-01", "df_idx": 0, "report_type": "NOON"},
            {"distance": 400, "steaming_hrs": 24, "avg_speed": 16.7,
             "datetime": "2024-12-02", "df_idx": 1, "report_type": "NOON"},
            {"distance": 48, "steaming_hrs": 24, "avg_speed": 2.0,
             "datetime": "2024-12-03", "df_idx": 2, "report_type": "NOON"},
        ]
        flagged = detect_speed_anomalies(rows)
        assert len(flagged) == 0

    def test_empty_rows(self):
        """Empty input → no anomalies."""
        from calculator import detect_speed_anomalies
        assert detect_speed_anomalies([]) == []
