"""
Integration tests — end-to-end pipeline with synthetic data.
=============================================================
Verifies that the full pipeline (load → detect → extract → compute → totals)
produces correct results for a synthetic Excel file.
"""

import sys
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import COL


# =============================================================================
# Helpers — create a synthetic Excel file
# =============================================================================

def _create_synthetic_excel(tmp_path: Path, n_noon: int = 5) -> Path:
    """
    Create a synthetic 340-column Excel file with a single LADEN voyage:
    DEPARTURE + n_noon NOON + ARRIVAL.

    ROBs decrease linearly so that:
      MGO: 2 MT/day, LNG: 150 m³/day, VLSFO G1: 1 MT/day, G2: 0.5 MT/day
    Distance: 420 nm per NOON, 380 nm at ARRIVAL, 0 at DEPARTURE
    """
    n_rows = 2 + n_noon  # DEP + NOONs + ARR
    n_cols = 340
    data = np.full((n_rows, n_cols), np.nan, dtype=object)

    base_date = pd.Timestamp("2024-12-05 08:00:00")

    for i in range(n_rows):
        dt = base_date + pd.Timedelta(days=i)

        if i == 0:
            rt = "DEPARTURE"
            dist = 0.0
            steaming = 0.0
        elif i == n_rows - 1:
            rt = "ARRIVAL"
            dist = 380.0
            steaming = 20.0
        else:
            rt = "NOON"
            dist = 420.0
            steaming = 24.0

        data[i, COL["report_type"]] = rt
        data[i, COL["voyage_no"]] = "99"
        data[i, COL["portcall_type"]] = "load"
        data[i, COL["fuel_mode"]] = "LNG ONLY"
        data[i, COL["datetime"]] = str(dt)
        data[i, COL["distance"]] = dist
        data[i, COL["steaming_hrs"]] = steaming
        data[i, COL["ordered_speed"]] = 19.5
        data[i, COL["avg_speed"]] = 17.5 if i > 0 else 0.0
        data[i, COL["next_port"]] = "TESTPORT"

        # ROBs — decrease linearly
        data[i, COL["mgo_rob"]] = 200.0 - i * 2.0
        data[i, COL["lng_rob"]] = 8000.0 - i * 150.0
        data[i, COL["vlsfo_g1_rob"]] = 80.0 - i * 1.0
        data[i, COL["vlsfo_g2_rob"]] = 40.0 - i * 0.5

        # BF5 hours — rows 2 and 4 have bad weather (>6 hrs)
        if i in (2, 4):
            data[i, COL["bf5_hours"]] = 12.0
        else:
            data[i, COL["bf5_hours"]] = 0.0

        # Boiler / pilot / GCU
        data[i, COL["mgo_boiler"]] = 0.05 if i > 0 else 0.0
        data[i, COL["mgo_pilot"]] = 0.02 if i > 0 else 0.0
        data[i, COL["vlsfo_g1_boiler"]] = 0.01
        data[i, COL["vlsfo_g1_pilot"]] = 0.005
        data[i, COL["vlsfo_g2_boiler"]] = 0.01
        data[i, COL["vlsfo_g2_pilot"]] = 0.005
        data[i, COL["gcu_lng"]] = 0.0

        # Reliq / Subcooler
        data[i, COL["reliq_hours"]] = 10.0 if i > 0 else 0.0
        data[i, COL["reliq_load"]] = 50.0 if i > 0 else 0.0
        data[i, COL["subcooler_hours"]] = 5.0 if i > 0 else 0.0
        data[i, COL["subcooler_load"]] = 30.0 if i > 0 else 0.0

        # Density / LCV
        data[i, COL["cargo_density"]] = 460.0
        data[i, COL["lcv"]] = 50.0

        # Segmentation columns (no changes → single segment)
        data[i, COL["voyage_order_rev"]] = ""
        data[i, COL["remarks"]] = ""

    df = pd.DataFrame(data)
    xlsx_path = tmp_path / "test_noon_report.xlsx"
    df.to_excel(xlsx_path, index=False, sheet_name="Sheet")
    return xlsx_path


# =============================================================================
# Integration Tests
# =============================================================================

class TestFullPipeline:
    """End-to-end: Excel → detect voyages → extract → compute → totals."""

    @pytest.fixture
    def synthetic_excel(self, tmp_path):
        return _create_synthetic_excel(tmp_path, n_noon=5)

    def test_detect_single_voyage(self, synthetic_excel):
        """Pipeline detects exactly 1 voyage from synthetic data."""
        from data_extractor import load_raw_excel, detect_voyages
        df = load_raw_excel(synthetic_excel)
        voyages = detect_voyages(df)
        assert len(voyages) == 1
        assert voyages[0]["voyage_type"] == "LADEN"
        assert voyages[0]["voyage_no"] == "99"

    def test_rob_diff_fuel(self, synthetic_excel):
        """ROB differences match expected values."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        df = load_raw_excel(synthetic_excel)
        voyages = detect_voyages(df)
        v = voyages[0]
        vd = extract_voyage_data(df, v["dep_row"], v["arr_row"])

        n_total = 7  # DEP + 5 NOON + ARR → index span = 6
        expected_mgo = 6 * 2.0  # 12.0 MT
        expected_lng = 6 * 150.0  # 900.0 m³
        expected_vlsfo_g1 = 6 * 1.0  # 6.0 MT
        expected_vlsfo_g2 = 6 * 0.5  # 3.0 MT

        assert abs(vd["mgo_consumed"] - expected_mgo) < 0.01
        assert abs(vd["lng_consumed"] - expected_lng) < 0.1
        assert abs(vd["vlsfo_g1_consumed"] - expected_vlsfo_g1) < 0.01
        assert abs(vd["vlsfo_g2_consumed"] - expected_vlsfo_g2) < 0.01
        assert abs(vd["vlsfo_consumed"] - (expected_vlsfo_g1 + expected_vlsfo_g2)) < 0.01

    def test_distance(self, synthetic_excel):
        """Total distance = sum of all Col P values."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        df = load_raw_excel(synthetic_excel)
        voyages = detect_voyages(df)
        v = voyages[0]
        vd = extract_voyage_data(df, v["dep_row"], v["arr_row"])
        # DEP=0 + 5*420 + ARR=380 = 2480
        assert abs(vd["total_distance"] - 2480.0) < 0.1

    def test_single_segment_detected(self, synthetic_excel):
        """No field changes → 1 segment."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        from calculator import compute_all_segments
        df = load_raw_excel(synthetic_excel)
        voyages = detect_voyages(df)
        v = voyages[0]
        vd = extract_voyage_data(df, v["dep_row"], v["arr_row"])
        computed = compute_all_segments(vd)
        assert len(computed["segments"]) == 1

    def test_weather_exclusions_applied(self, synthetic_excel):
        """Rows 2 and 4 have BF5=12 > threshold=6 → exclusions on those days."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        from calculator import compute_all_segments
        df = load_raw_excel(synthetic_excel)
        voyages = detect_voyages(df)
        v = voyages[0]
        vd = extract_voyage_data(df, v["dep_row"], v["arr_row"])
        computed = compute_all_segments(vd)
        totals = computed["totals"]

        # There should be weather exclusions
        assert totals["excl_mgo_weather"] > 0
        assert totals["excl_lng_weather"] > 0
        assert totals["excl_vlsfo_weather"] > 0

        # Net fuel = actual - excluded
        assert abs(totals["net_mgo"] - (totals["mgo_consumed"] - totals["excl_mgo_weather"])) < 0.01
        assert abs(totals["net_lng"] - (totals["lng_consumed"] - totals["excl_lng_weather"])) < 0.01

    def test_speed_formula(self, synthetic_excel):
        """Speed = net_distance / (net_duration_days × 24)."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        from calculator import compute_all_segments
        df = load_raw_excel(synthetic_excel)
        voyages = detect_voyages(df)
        v = voyages[0]
        vd = extract_voyage_data(df, v["dep_row"], v["arr_row"])
        computed = compute_all_segments(vd)
        totals = computed["totals"]

        expected_speed = totals["net_distance"] / (totals["net_duration_days"] * 24)
        assert abs(totals["actual_avg_speed"] - expected_speed) < 0.01

    def test_totals_consistent(self, synthetic_excel):
        """Totals match sum of segment values (single segment case)."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        from calculator import compute_all_segments
        df = load_raw_excel(synthetic_excel)
        voyages = detect_voyages(df)
        v = voyages[0]
        vd = extract_voyage_data(df, v["dep_row"], v["arr_row"])
        computed = compute_all_segments(vd)

        seg = computed["segments"][0]
        totals = computed["totals"]

        # With 1 segment, totals should equal segment values
        assert abs(totals["distance"] - seg["distance"]) < 0.01
        assert abs(totals["mgo_consumed"] - seg["mgo_consumed"]) < 0.01
        assert abs(totals["lng_consumed"] - seg["lng_consumed"]) < 0.01


class TestMultiSegmentPipeline:
    """Integration test with a voyage order revision creating 2 segments."""

    @pytest.fixture
    def two_segment_excel(self, tmp_path):
        """Create Excel with a voyage order revision at row 3 (NOON 3)."""
        n_noon = 5
        n_rows = 2 + n_noon
        n_cols = 340
        data = np.full((n_rows, n_cols), np.nan, dtype=object)

        base_date = pd.Timestamp("2024-12-05 08:00:00")

        for i in range(n_rows):
            dt = base_date + pd.Timedelta(days=i)

            if i == 0:
                rt = "DEPARTURE"
                dist = 0.0
                steaming = 0.0
            elif i == n_rows - 1:
                rt = "ARRIVAL"
                dist = 380.0
                steaming = 20.0
            else:
                rt = "NOON"
                dist = 420.0
                steaming = 24.0

            data[i, COL["report_type"]] = rt
            data[i, COL["voyage_no"]] = "11"
            data[i, COL["portcall_type"]] = "load"
            data[i, COL["fuel_mode"]] = "LNG ONLY"
            data[i, COL["datetime"]] = str(dt)
            data[i, COL["distance"]] = dist
            data[i, COL["steaming_hrs"]] = steaming
            data[i, COL["ordered_speed"]] = 19.5
            data[i, COL["avg_speed"]] = 17.5 if i > 0 else 0.0
            data[i, COL["next_port"]] = "ALIAGA"
            data[i, COL["mgo_rob"]] = 200.0 - i * 2.0
            data[i, COL["lng_rob"]] = 8000.0 - i * 150.0
            data[i, COL["vlsfo_g1_rob"]] = 80.0 - i * 1.0
            data[i, COL["vlsfo_g2_rob"]] = 40.0 - i * 0.5
            data[i, COL["bf5_hours"]] = 0.0
            data[i, COL["mgo_boiler"]] = 0.05 if i > 0 else 0.0
            data[i, COL["mgo_pilot"]] = 0.02 if i > 0 else 0.0
            data[i, COL["vlsfo_g1_boiler"]] = 0.01
            data[i, COL["vlsfo_g1_pilot"]] = 0.005
            data[i, COL["vlsfo_g2_boiler"]] = 0.01
            data[i, COL["vlsfo_g2_pilot"]] = 0.005
            data[i, COL["gcu_lng"]] = 0.0
            data[i, COL["reliq_hours"]] = 10.0 if i > 0 else 0.0
            data[i, COL["reliq_load"]] = 50.0 if i > 0 else 0.0
            data[i, COL["subcooler_hours"]] = 5.0 if i > 0 else 0.0
            data[i, COL["subcooler_load"]] = 30.0 if i > 0 else 0.0
            data[i, COL["cargo_density"]] = 460.0
            data[i, COL["lcv"]] = 50.0
            data[i, COL["voyage_order_rev"]] = ""
            data[i, COL["remarks"]] = ""

        # Add voyage order revision at row 3 (NOON 3)
        rev_row = 3
        rev_dt = base_date + pd.Timedelta(days=3, hours=-2)  # 2 hours before report
        data[rev_row, COL["voyage_order_rev"]] = "yes"
        data[rev_row, COL["rev_start_time"]] = str(rev_dt)
        data[rev_row, COL["rev_gmt_offset"]] = 0
        data[rev_row, COL["rev_speed"]] = 18.0
        data[rev_row, COL["rev_sat"]] = "2024-12-15 06:00:00"

        df = pd.DataFrame(data)
        xlsx_path = tmp_path / "test_two_seg.xlsx"
        df.to_excel(xlsx_path, index=False, sheet_name="Sheet")
        return xlsx_path

    def test_two_segments_detected(self, two_segment_excel):
        """Voyage order revision → 2 segments."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        from calculator import compute_all_segments
        df = load_raw_excel(two_segment_excel)
        voyages = detect_voyages(df)
        vd = extract_voyage_data(df, voyages[0]["dep_row"], voyages[0]["arr_row"])
        computed = compute_all_segments(vd)
        assert len(computed["segments"]) == 2

    def test_segment_speeds_differ(self, two_segment_excel):
        """Seg 1 has 19.5 kts instructed, Seg 2 has 18.0 kts."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        from calculator import compute_all_segments
        df = load_raw_excel(two_segment_excel)
        voyages = detect_voyages(df)
        vd = extract_voyage_data(df, voyages[0]["dep_row"], voyages[0]["arr_row"])
        computed = compute_all_segments(vd)
        seg_info = computed["segment_info"]
        assert seg_info[0]["instructed_speed"] == 19.5
        assert seg_info[1]["instructed_speed"] == 18.0

    def test_totals_distance_consistent(self, two_segment_excel):
        """
        Sum of segment distances should approximately equal total distance.
        (Small rounding diffs from pro-rating are acceptable.)
        """
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        from calculator import compute_all_segments
        df = load_raw_excel(two_segment_excel)
        voyages = detect_voyages(df)
        vd = extract_voyage_data(df, voyages[0]["dep_row"], voyages[0]["arr_row"])
        computed = compute_all_segments(vd)

        seg_dist = sum(s["distance"] for s in computed["segments"])
        assert abs(seg_dist - computed["totals"]["distance"]) < 0.1

    def test_totals_fuel_consistent(self, two_segment_excel):
        """Sum of segment fuel should equal totals."""
        from data_extractor import load_raw_excel, detect_voyages, extract_voyage_data
        from calculator import compute_all_segments
        df = load_raw_excel(two_segment_excel)
        voyages = detect_voyages(df)
        vd = extract_voyage_data(df, voyages[0]["dep_row"], voyages[0]["arr_row"])
        computed = compute_all_segments(vd)

        for fuel in ["mgo_consumed", "lng_consumed", "vlsfo_consumed"]:
            seg_total = sum(s[fuel] for s in computed["segments"])
            assert abs(seg_total - computed["totals"][fuel]) < 0.1, \
                f"{fuel}: segments sum {seg_total} != total {computed['totals'][fuel]}"
