"""
Tests for data_extractor.py
============================
Covers:
  - Voyage detection (single, multiple, incomplete)
  - ROB-diff fuel consumption (MGO, LNG, VLSFO both grades)
  - Distance sum
  - Steaming hours sum
  - Daily row construction (ROB-diff per row)
  - Reliq/Subcooler aggregation
  - Auxiliary data extraction (boiler, pilot, GCU, density, LCV)
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

# Ensure the app dir is importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import COL


# =============================================================================
# Helpers — build synthetic DataFrames matching 340-column noon-report format
# =============================================================================

def _make_df(rows_data: list[dict]) -> pd.DataFrame:
    """
    Create a synthetic DataFrame with 340 columns.

    Each row_data dict maps COL keys (e.g. 'report_type', 'distance')
    to values. Columns not set are filled with NaN.
    """
    n_cols = 340
    n_rows = len(rows_data)
    data = np.full((n_rows, n_cols), np.nan, dtype=object)

    for i, rd in enumerate(rows_data):
        for key, val in rd.items():
            if key in COL:
                data[i, COL[key]] = val
            else:
                raise KeyError(f"Unknown COL key: {key}")

    df = pd.DataFrame(data)
    return df


def _basic_voyage_rows(n_noon=3):
    """
    Create rows for a simple single-voyage test:
    DEPARTURE + n_noon NOON reports + ARRIVAL.
    """
    rows = []

    # DEPARTURE row
    dep = {
        "report_type": "DEPARTURE",
        "voyage_no": "11",
        "portcall_type": "load",
        "fuel_mode": "LNG ONLY",
        "datetime": "2024-12-05 08:00:00",
        "distance": 0.0,
        "steaming_hrs": 0.0,
        "ordered_speed": 19.5,
        "mgo_rob": 100.0,
        "lng_rob": 5000.0,
        "vlsfo_g1_rob": 50.0,
        "vlsfo_g2_rob": 30.0,
        "next_port": "ALIAGA",
        "bf5_hours": 0.0,
        "mgo_boiler": 0.0,
        "mgo_pilot": 0.0,
        "vlsfo_g1_boiler": 0.0,
        "vlsfo_g1_pilot": 0.0,
        "vlsfo_g2_boiler": 0.0,
        "vlsfo_g2_pilot": 0.0,
        "gcu_lng": 0.0,
        "reliq_hours": 0.0,
        "reliq_load": 0.0,
        "subcooler_hours": 0.0,
        "subcooler_load": 0.0,
        "cargo_density": 460.0,
        "lcv": 50.0,
        "avg_speed": 0.0,
        "voyage_order_rev": "",
        "remarks": "",
    }
    rows.append(dep)

    # NOON rows
    for d in range(1, n_noon + 1):
        noon = {
            "report_type": "NOON",
            "voyage_no": "11",
            "portcall_type": "load",
            "fuel_mode": "LNG ONLY",
            "datetime": f"2024-12-{5+d:02d} 12:00:00",
            "distance": 400.0,
            "steaming_hrs": 24.0,
            "ordered_speed": 19.5,
            "mgo_rob": 100.0 - d * 2.0,    # 2 MT/day
            "lng_rob": 5000.0 - d * 150.0,  # 150 m3/day
            "vlsfo_g1_rob": 50.0 - d * 1.0, # 1 MT/day
            "vlsfo_g2_rob": 30.0 - d * 0.5, # 0.5 MT/day
            "next_port": "ALIAGA",
            "bf5_hours": 0.0,
            "mgo_boiler": 0.1,
            "mgo_pilot": 0.05,
            "vlsfo_g1_boiler": 0.02,
            "vlsfo_g1_pilot": 0.01,
            "vlsfo_g2_boiler": 0.02,
            "vlsfo_g2_pilot": 0.01,
            "gcu_lng": 0.0,
            "reliq_hours": 8.0,
            "reliq_load": 50.0,
            "subcooler_hours": 4.0,
            "subcooler_load": 30.0,
            "cargo_density": 460.0,
            "lcv": 50.0,
            "avg_speed": 16.7,
            "voyage_order_rev": "",
            "remarks": "",
        }
        rows.append(noon)

    # ARRIVAL row
    d = n_noon + 1
    arr = {
        "report_type": "ARRIVAL",
        "voyage_no": "11",
        "portcall_type": "load",
        "fuel_mode": "LNG ONLY",
        "datetime": f"2024-12-{5+d:02d} 14:00:00",
        "distance": 380.0,
        "steaming_hrs": 20.0,
        "ordered_speed": 19.5,
        "mgo_rob": 100.0 - d * 2.0,
        "lng_rob": 5000.0 - d * 150.0,
        "vlsfo_g1_rob": 50.0 - d * 1.0,
        "vlsfo_g2_rob": 30.0 - d * 0.5,
        "next_port": "ALIAGA",
        "bf5_hours": 0.0,
        "mgo_boiler": 0.1,
        "mgo_pilot": 0.05,
        "vlsfo_g1_boiler": 0.02,
        "vlsfo_g1_pilot": 0.01,
        "vlsfo_g2_boiler": 0.02,
        "vlsfo_g2_pilot": 0.01,
        "gcu_lng": 0.0,
        "reliq_hours": 6.0,
        "reliq_load": 50.0,
        "subcooler_hours": 3.0,
        "subcooler_load": 30.0,
        "cargo_density": 460.0,
        "lcv": 50.0,
        "avg_speed": 16.5,
        "voyage_order_rev": "",
        "remarks": "",
    }
    rows.append(arr)

    return rows


# =============================================================================
# Voyage Detection
# =============================================================================

class TestDetectVoyages:
    """Tests for detect_voyages()."""

    def test_single_voyage(self):
        """Simple DEPARTURE → noons → ARRIVAL → 1 voyage."""
        from data_extractor import detect_voyages
        df = _make_df(_basic_voyage_rows(3))
        voyages = detect_voyages(df)
        assert len(voyages) == 1
        v = voyages[0]
        assert v["voyage_no"] == "11"
        assert v["voyage_type"] == "LADEN"
        assert v["dep_row"] == 0
        assert v["arr_row"] == 4  # DEP + 3 NOON + ARR

    def test_multiple_voyages(self):
        """Two DEPARTURE→ARRIVAL pairs → 2 voyages."""
        from data_extractor import detect_voyages
        rows1 = _basic_voyage_rows(2)
        rows2 = _basic_voyage_rows(2)
        # Change second voyage to BALLAST
        for r in rows2:
            r["portcall_type"] = "discharge"
            r["voyage_no"] = "12"
        # Adjust datetimes so they follow
        for i, r in enumerate(rows2):
            r["datetime"] = f"2024-12-{20+i:02d} 12:00:00"

        df = _make_df(rows1 + rows2)
        voyages = detect_voyages(df)
        assert len(voyages) == 2
        assert voyages[0]["voyage_type"] == "LADEN"
        assert voyages[1]["voyage_type"] == "BALLAST"
        assert voyages[1]["voyage_no"] == "12"

    def test_incomplete_voyage(self):
        """DEPARTURE with no ARRIVAL → 0 voyages + warning."""
        from data_extractor import detect_voyages
        rows = _basic_voyage_rows(3)
        # Remove ARRIVAL row
        rows = [r for r in rows if r["report_type"] != "ARRIVAL"]
        df = _make_df(rows)
        voyages = detect_voyages(df)
        assert len(voyages) == 0

    def test_no_departure(self):
        """No DEPARTURE rows → 0 voyages."""
        from data_extractor import detect_voyages
        rows = _basic_voyage_rows(2)
        rows = [r for r in rows if r["report_type"] != "DEPARTURE"]
        df = _make_df(rows)
        voyages = detect_voyages(df)
        assert len(voyages) == 0

    def test_voyage_type_laden(self):
        """Portcall starting with 'load' → LADEN."""
        from data_extractor import detect_voyages
        df = _make_df(_basic_voyage_rows(1))
        voyages = detect_voyages(df)
        assert voyages[0]["voyage_type"] == "LADEN"

    def test_voyage_type_ballast(self):
        """Portcall starting with 'discharge' → BALLAST."""
        from data_extractor import detect_voyages
        rows = _basic_voyage_rows(1)
        for r in rows:
            r["portcall_type"] = "discharge"
        df = _make_df(rows)
        voyages = detect_voyages(df)
        assert voyages[0]["voyage_type"] == "BALLAST"


# =============================================================================
# Voyage Data Extraction (Rule 1: ROB difference)
# =============================================================================

class TestExtractVoyageData:
    """Tests for extract_voyage_data()."""

    def test_rob_diff_mgo(self):
        """MGO consumed = ROB at DEPARTURE minus ROB at ARRIVAL."""
        from data_extractor import extract_voyage_data
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        vd = extract_voyage_data(df, 0, 4)
        # MGO: 100.0 - (100.0 - 4*2.0) = 8.0 MT
        assert abs(vd["mgo_consumed"] - 8.0) < 0.01

    def test_rob_diff_lng(self):
        """LNG consumed = ROB at DEPARTURE minus ROB at ARRIVAL."""
        from data_extractor import extract_voyage_data
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        vd = extract_voyage_data(df, 0, 4)
        # LNG: 5000.0 - (5000.0 - 4*150) = 600.0 m3
        assert abs(vd["lng_consumed"] - 600.0) < 0.1

    def test_rob_diff_vlsfo_two_grades(self):
        """VLSFO consumed = Grade 1 + Grade 2 ROB diffs."""
        from data_extractor import extract_voyage_data
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        vd = extract_voyage_data(df, 0, 4)
        # VLSFO G1: 50 - (50 - 4*1) = 4.0 MT
        # VLSFO G2: 30 - (30 - 4*0.5) = 2.0 MT
        # Total: 6.0 MT
        assert abs(vd["vlsfo_g1_consumed"] - 4.0) < 0.01
        assert abs(vd["vlsfo_g2_consumed"] - 2.0) < 0.01
        assert abs(vd["vlsfo_consumed"] - 6.0) < 0.01

    def test_distance_sum(self):
        """Total distance = sum Col P from DEPARTURE to ARRIVAL (inclusive)."""
        from data_extractor import extract_voyage_data
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        vd = extract_voyage_data(df, 0, 4)
        # DEP=0 + NOON1=400 + NOON2=400 + NOON3=400 + ARR=380 = 1580
        assert abs(vd["total_distance"] - 1580.0) < 0.1

    def test_steaming_hours_sum(self):
        """Total steaming hours = sum Col AT2."""
        from data_extractor import extract_voyage_data
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        vd = extract_voyage_data(df, 0, 4)
        # DEP=0 + NOON1=24 + NOON2=24 + NOON3=24 + ARR=20 = 92
        assert abs(vd["total_steaming_hrs"] - 92.0) < 0.1

    def test_daily_rows_count(self):
        """daily_rows should have one entry per df row from DEP to ARR."""
        from data_extractor import extract_voyage_data
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        vd = extract_voyage_data(df, 0, 4)
        assert len(vd["daily_rows"]) == 5  # DEP + 3 NOON + ARR

    def test_daily_rows_rob_diff(self):
        """First row (DEPARTURE) has 0 daily consumption; subsequent use ROB diff."""
        from data_extractor import extract_voyage_data
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        vd = extract_voyage_data(df, 0, 4)
        dr = vd["daily_rows"]
        # Row 0 (DEPARTURE): daily = 0
        assert dr[0]["mgo_daily"] == 0.0
        assert dr[0]["lng_daily"] == 0.0
        # Row 1 (first NOON): MGO = 100 - 98 = 2.0
        assert abs(dr[1]["mgo_daily"] - 2.0) < 0.01
        # Row 1: LNG = 5000 - 4850 = 150
        assert abs(dr[1]["lng_daily"] - 150.0) < 0.1
        # Row 1: VLSFO G1 = 50 - 49 = 1.0, G2 = 30 - 29.5 = 0.5
        assert abs(dr[1]["vlsfo_g1_daily"] - 1.0) < 0.01
        assert abs(dr[1]["vlsfo_g2_daily"] - 0.5) < 0.01
        assert abs(dr[1]["vlsfo_daily"] - 1.5) < 0.01


# =============================================================================
# Auxiliary Data (Rule 4)
# =============================================================================

class TestExtractAuxiliary:
    """Tests for extract_auxiliary()."""

    def test_boiler_pilot_sums(self):
        """Boiler and pilot values are summed across all rows."""
        from data_extractor import extract_auxiliary
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        aux = extract_auxiliary(df, 0, 4)
        # 4 rows with boiler (not DEP which is 0)
        # MGO boiler: 3*0.1 + 0.1 = 0.4 (3 NOON + ARR, DEP=0)
        assert abs(aux["mgo_boiler_total"] - 0.4) < 0.01
        assert abs(aux["mgo_pilot_total"] - 0.2) < 0.01

    def test_reliq_subcooler(self):
        """Reliq hours = sum(IL+IQ), load = avg(IN+IR non-zero)."""
        from data_extractor import extract_auxiliary
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        aux = extract_auxiliary(df, 0, 4)
        # Hours: DEP=0, NOON1=8+4=12, NOON2=12, NOON3=12, ARR=6+3=9
        # Total = 0 + 12 + 12 + 12 + 9 = 45
        assert abs(aux["reliq_total_hours"] - 45.0) < 0.1
        # Load: all non-zero values are 50 (reliq) and 30 (subcooler) = avg 40
        # DEP has 0 load → excluded
        assert abs(aux["reliq_avg_load"] - 40.0) < 0.1

    def test_gcu_detection(self):
        """GCU total and dates detected correctly."""
        from data_extractor import extract_auxiliary
        rows = _basic_voyage_rows(3)
        # Add some GCU on row 2
        rows[2]["gcu_lng"] = 5.0
        df = _make_df(rows)
        aux = extract_auxiliary(df, 0, 4)
        assert aux["gcu_used"] is True
        assert abs(aux["gcu_lng_total"] - 5.0) < 0.01
        assert len(aux["gcu_dates"]) == 1

    def test_no_gcu(self):
        """No GCU usage → gcu_used=False."""
        from data_extractor import extract_auxiliary
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        aux = extract_auxiliary(df, 0, 4)
        assert aux["gcu_used"] is False
        assert aux["gcu_lng_total"] == 0.0

    def test_cargo_density_constant(self):
        """Constant density → use as-is."""
        from data_extractor import extract_auxiliary
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        aux = extract_auxiliary(df, 0, 4)
        assert aux["cargo_density"] == 460.0

    def test_lcv_constant(self):
        """Constant LCV → use as-is."""
        from data_extractor import extract_auxiliary
        rows = _basic_voyage_rows(3)
        df = _make_df(rows)
        aux = extract_auxiliary(df, 0, 4)
        assert aux["lcv"] == 50.0


# =============================================================================
# Load Raw Excel
# =============================================================================

class TestLoadRawExcel:
    """Tests for load_raw_excel()."""

    def test_file_not_found(self):
        """Missing file → FileNotFoundError."""
        from data_extractor import load_raw_excel
        with pytest.raises(FileNotFoundError):
            load_raw_excel("/nonexistent/path/data.xlsx")


# =============================================================================
# safe_float helper
# =============================================================================

class TestSafeFloat:
    """Tests for _safe_float()."""

    def test_nan_returns_zero(self):
        from data_extractor import _safe_float
        assert _safe_float(float("nan")) == 0.0

    def test_none_returns_zero(self):
        from data_extractor import _safe_float
        assert _safe_float(None) == 0.0

    def test_string_number(self):
        from data_extractor import _safe_float
        assert _safe_float("42.5") == 42.5

    def test_invalid_string(self):
        from data_extractor import _safe_float
        assert _safe_float("abc") == 0.0

    def test_normal_float(self):
        from data_extractor import _safe_float
        assert _safe_float(3.14) == 3.14


# =============================================================================
# Mid-Voyage Bunkering Detection (Rule B)
# =============================================================================

class TestMergeBunkeringStops:
    """Tests for merge_bunkering_stops() — Rule B."""

    def test_single_voyage_no_merge(self):
        """Single voyage → returned unchanged with empty intermediate_stops."""
        from data_extractor import merge_bunkering_stops
        import pandas as pd
        from config import COL

        # Minimal df with 2 rows
        n_cols = max(COL.values()) + 1
        data = [[None] * n_cols for _ in range(2)]
        # Set next_port at dep row (0) and arr row (1)
        data[0][COL["next_port"]] = "Port A"
        data[1][COL["next_port"]] = "Port B"
        df = pd.DataFrame(data)

        voyages = [{
            "voyage_no": "V001",
            "voyage_type": "LADEN",
            "fuel_mode": "LNG ONLY",
            "dep_row": 0,
            "arr_row": 1,
            "dep_datetime": "2024-12-01 06:00",
            "arr_datetime": "2024-12-05 18:00",
        }]
        result = merge_bunkering_stops(df, voyages)
        assert len(result) == 1
        assert result[0]["intermediate_stops"] == []

    def test_two_voyages_same_port_merged(self):
        """Two voyages with same next_port at boundary → merged into one."""
        from data_extractor import merge_bunkering_stops
        import pandas as pd
        from config import COL

        n_cols = max(COL.values()) + 1
        data = [[None] * n_cols for _ in range(4)]
        # V1: dep=0, arr=1. V2: dep=2, arr=3
        # V1 arr next_port = V2 dep next_port = "Singapore" → merge
        data[0][COL["next_port"]] = "Singapore"
        data[1][COL["next_port"]] = "Singapore"
        data[2][COL["next_port"]] = "Singapore"
        data[3][COL["next_port"]] = "Final Port"

        # Set datetime, ROB values for stop calculation
        data[1][COL["datetime"]] = "2024-12-03 18:00"  # arr of V1
        data[2][COL["datetime"]] = "2024-12-04 06:00"  # dep of V2
        data[1][COL["mgo_rob"]] = 100.0
        data[2][COL["mgo_rob"]] = 98.0
        data[1][COL["lng_rob"]] = 5000.0
        data[2][COL["lng_rob"]] = 4950.0
        data[1][COL["vlsfo_g1_rob"]] = 50.0
        data[2][COL["vlsfo_g1_rob"]] = 49.0
        data[1][COL["vlsfo_g2_rob"]] = 30.0
        data[2][COL["vlsfo_g2_rob"]] = 29.0

        df = pd.DataFrame(data)

        voyages = [
            {"voyage_no": "V001", "voyage_type": "LADEN", "fuel_mode": "LNG ONLY",
             "dep_row": 0, "arr_row": 1,
             "dep_datetime": "2024-12-01 06:00", "arr_datetime": "2024-12-03 18:00"},
            {"voyage_no": "V002", "voyage_type": "LADEN", "fuel_mode": "LNG ONLY",
             "dep_row": 2, "arr_row": 3,
             "dep_datetime": "2024-12-04 06:00", "arr_datetime": "2024-12-08 18:00"},
        ]
        result = merge_bunkering_stops(df, voyages)
        assert len(result) == 1  # merged into one voyage
        assert result[0]["dep_row"] == 0
        assert result[0]["arr_row"] == 3
        assert len(result[0]["intermediate_stops"]) == 1

        stop = result[0]["intermediate_stops"][0]
        assert stop["port_name"] == "Singapore"
        assert abs(stop["mgo_consumed"] - 2.0) < 0.01
        assert abs(stop["lng_consumed"] - 50.0) < 0.1

    def test_two_voyages_different_port_not_merged(self):
        """Two voyages with different next_port → NOT merged."""
        from data_extractor import merge_bunkering_stops
        import pandas as pd
        from config import COL

        n_cols = max(COL.values()) + 1
        data = [[None] * n_cols for _ in range(4)]
        data[0][COL["next_port"]] = "Port A"
        data[1][COL["next_port"]] = "Port B"  # V1 arr
        data[2][COL["next_port"]] = "Port C"  # V2 dep — different!
        data[3][COL["next_port"]] = "Port D"
        df = pd.DataFrame(data)

        voyages = [
            {"voyage_no": "V001", "voyage_type": "LADEN", "fuel_mode": "LNG ONLY",
             "dep_row": 0, "arr_row": 1,
             "dep_datetime": "2024-12-01 06:00", "arr_datetime": "2024-12-03 18:00"},
            {"voyage_no": "V002", "voyage_type": "BALLAST", "fuel_mode": "LNG ONLY",
             "dep_row": 2, "arr_row": 3,
             "dep_datetime": "2024-12-04 06:00", "arr_datetime": "2024-12-08 18:00"},
        ]
        result = merge_bunkering_stops(df, voyages)
        assert len(result) == 2
        assert result[0]["intermediate_stops"] == []
        assert result[1]["intermediate_stops"] == []

    def test_empty_voyages(self):
        """Empty voyages list → empty result."""
        from data_extractor import merge_bunkering_stops
        import pandas as pd
        df = pd.DataFrame()
        assert merge_bunkering_stops(df, []) == []
