"""
data_extractor.py
=================
Loads raw noon-report Excel and extracts voyage data using index-based
column access and the ROB difference method.

Implements:
  - Rule 1: Voyage boundary detection (DEPARTURE → ARRIVAL)
  - Rule 1: Fuel consumption via ROB difference (MGO, LNG, VLSFO two grades)
  - Rule 1: Reliq/Subcooler aggregation (IL+IQ hours, IN+IR load)
  - Rule 4: Boiler, pilot, GCU, density, LCV extraction
"""

from __future__ import annotations

import logging
from pathlib import Path

import numpy as np
import pandas as pd

from config import COL, resolve_columns

logger = logging.getLogger(__name__)


# =============================================================================
# 1. Raw Excel Loading
# =============================================================================

def load_raw_excel(
    excel_path: str | Path,
    sheet_name: str = "Sheet",
) -> pd.DataFrame:
    """
    Load raw noon-report Excel file.

    Returns the DataFrame with positional (integer) column access.
    No filtering, no date parsing — raw load only.
    """
    path = Path(excel_path)
    if not path.exists():
        raise FileNotFoundError(f"Raw Excel file not found: {path}")

    df = pd.read_excel(path, sheet_name=sheet_name)
    logger.info(
        "Loaded %d rows × %d columns from '%s' (sheet='%s')",
        len(df), df.shape[1], path.name, sheet_name,
    )

    # Auto-detect column indices from header names
    try:
        resolve_columns(df)
        logger.info("Column auto-detection succeeded.")
    except ValueError as e:
        logger.warning(
            "Column auto-detection failed — using fallback indices. %s", e
        )

    return df


# =============================================================================
# 2. Voyage Detection  (Rule 1: DEPARTURE → ARRIVAL boundaries)
# =============================================================================

def detect_voyages(df: pd.DataFrame) -> list[dict]:
    """
    Scan Col AC (report_type, index 28) for DEPARTURE/ARRIVAL markers.

    Returns a list of voyage dicts, each containing:
      - dep_row:      int — DataFrame row index of DEPARTURE
      - arr_row:      int — DataFrame row index of ARRIVAL
      - voyage_no:    str — from Col B (index 1)
      - voyage_type:  str — "LADEN" or "BALLAST" from Col G (index 6)
      - fuel_mode:    str — from Col K (index 10)
      - dep_datetime: str — departure date/time from Col C (index 2)
      - arr_datetime: str — arrival date/time from Col C (index 2)
    """
    rt_col = COL["report_type"]
    voyages = []
    dep_row = None

    for i in range(len(df)):
        rt = str(df.iloc[i, rt_col]).strip().upper() if pd.notna(df.iloc[i, rt_col]) else ""

        if rt == "DEPARTURE":
            dep_row = i
            logger.debug("DEPARTURE found at row %d", i)

        elif rt == "ARRIVAL" and dep_row is not None:
            # Determine voyage type from Col G of the DEPARTURE row
            portcall = str(df.iloc[dep_row, COL["portcall_type"]]).strip().lower()
            if portcall.startswith("load"):
                voyage_type = "LADEN"
            elif portcall.startswith("discharge"):
                voyage_type = "BALLAST"
            else:
                voyage_type = portcall.upper()
                logger.warning(
                    "Unknown portcall type '%s' at row %d — using as-is",
                    portcall, dep_row,
                )

            voyage_no = str(df.iloc[dep_row, COL["voyage_no"]]) if pd.notna(
                df.iloc[dep_row, COL["voyage_no"]]
            ) else "Unknown"

            fuel_mode = str(df.iloc[dep_row, COL["fuel_mode"]]) if pd.notna(
                df.iloc[dep_row, COL["fuel_mode"]]
            ) else "Unknown"

            dep_dt = str(df.iloc[dep_row, COL["datetime"]])
            arr_dt = str(df.iloc[i, COL["datetime"]])

            voyages.append({
                "dep_row":      dep_row,
                "arr_row":      i,
                "voyage_no":    voyage_no,
                "voyage_type":  voyage_type,
                "fuel_mode":    fuel_mode,
                "dep_datetime": dep_dt,
                "arr_datetime": arr_dt,
            })

            logger.info(
                "Voyage %s detected: rows %d→%d, type=%s, %s → %s",
                voyage_no, dep_row, i, voyage_type, dep_dt[:10], arr_dt[:10],
            )
            dep_row = None  # reset for next voyage

    if dep_row is not None:
        logger.warning(
            "Incomplete voyage: DEPARTURE at row %d with no matching ARRIVAL",
            dep_row,
        )

    logger.info("Total voyages detected: %d", len(voyages))
    return voyages


# =============================================================================
# 2b. Mid-Voyage Bunkering Detection  (Rule B)
# =============================================================================

def _get_next_port(df: pd.DataFrame, row_idx: int) -> str | None:
    """Get cleaned next_port value from a row."""
    val = df.iloc[row_idx, COL["next_port"]]
    if pd.notna(val) and str(val).strip():
        return str(val).strip()
    return None


def _build_stop_data(df: pd.DataFrame, arr_row: int, dep_row: int) -> dict:
    """
    Build data for an intermediate stop (bunkering/port call).

    Fuel consumed during stop = ROB at intermediate ARRIVAL − ROB at next DEPARTURE.
    """
    arr_dt = pd.to_datetime(df.iloc[arr_row, COL["datetime"]], errors="coerce")
    dep_dt = pd.to_datetime(df.iloc[dep_row, COL["datetime"]], errors="coerce")

    duration_hours = 0.0
    if pd.notna(arr_dt) and pd.notna(dep_dt):
        duration_hours = (dep_dt - arr_dt).total_seconds() / 3600.0

    mgo_consumed = (
        _safe_float(df.iloc[arr_row, COL["mgo_rob"]])
        - _safe_float(df.iloc[dep_row, COL["mgo_rob"]])
    )
    lng_consumed = (
        _safe_float(df.iloc[arr_row, COL["lng_rob"]])
        - _safe_float(df.iloc[dep_row, COL["lng_rob"]])
    )
    vlsfo_g1 = (
        _safe_float(df.iloc[arr_row, COL["vlsfo_g1_rob"]])
        - _safe_float(df.iloc[dep_row, COL["vlsfo_g1_rob"]])
    )
    vlsfo_g2 = (
        _safe_float(df.iloc[arr_row, COL["vlsfo_g2_rob"]])
        - _safe_float(df.iloc[dep_row, COL["vlsfo_g2_rob"]])
    )

    port_name = _get_next_port(df, arr_row) or "Unknown"

    return {
        "arr_datetime":      str(arr_dt)[:16] if pd.notna(arr_dt) else "",
        "dep_datetime":      str(dep_dt)[:16] if pd.notna(dep_dt) else "",
        "port_name":         port_name,
        "duration_hours":    round(duration_hours, 2),
        "arr_row":           arr_row,
        "dep_row":           dep_row,
        "mgo_consumed":      round(mgo_consumed, 3),
        "lng_consumed":      round(lng_consumed, 3),
        "vlsfo_consumed":    round(vlsfo_g1 + vlsfo_g2, 3),
        "vlsfo_g1_consumed": round(vlsfo_g1, 3),
        "vlsfo_g2_consumed": round(vlsfo_g2, 3),
    }


def merge_bunkering_stops(
    df: pd.DataFrame,
    voyages: list[dict],
) -> list[dict]:
    """
    Merge consecutive voyage pairs that represent mid-voyage bunkering stops.

    Rule B: If voyage[i] is immediately followed by voyage[i+1], AND
    next_port at voyage[i].arr_row == next_port at voyage[i+1].dep_row,
    then merge into one voyage with an intermediate_stops entry.

    Returns list[dict] — each voyage now has an 'intermediate_stops' key.
    """
    if len(voyages) <= 1:
        for v in voyages:
            v["intermediate_stops"] = []
        return voyages

    merged = []
    i = 0

    while i < len(voyages):
        current = voyages[i].copy()
        current["intermediate_stops"] = []

        # Look ahead: can we merge the next voyage into this one?
        while i + 1 < len(voyages):
            next_voy = voyages[i + 1]

            arr_next_port = _get_next_port(df, current["arr_row"])
            next_dep_next_port = _get_next_port(df, next_voy["dep_row"])

            if (
                arr_next_port
                and next_dep_next_port
                and arr_next_port == next_dep_next_port
            ):
                # Build stop data for the intermediate period
                stop = _build_stop_data(df, current["arr_row"], next_voy["dep_row"])
                current["intermediate_stops"].append(stop)

                # Extend current voyage to end at the next voyage's ARRIVAL
                current["arr_row"] = next_voy["arr_row"]
                current["arr_datetime"] = next_voy["arr_datetime"]

                logger.info(
                    "Merged bunkering stop: rows %d→%d at port '%s' "
                    "(duration: %.1f hrs)",
                    stop["arr_row"], stop["dep_row"], stop["port_name"],
                    stop["duration_hours"],
                )
                i += 1
            else:
                break

        merged.append(current)
        i += 1

    return merged


def tag_intermediate_stops(daily_rows: list[dict], stops: list[dict]) -> None:
    """Tag daily_rows that fall within intermediate stop periods (in-place)."""
    for row in daily_rows:
        row["is_intermediate_stop"] = False
        for stop in stops:
            if stop["arr_row"] <= row.get("df_idx", -1) <= stop["dep_row"]:
                row["is_intermediate_stop"] = True
                row["stop_port"] = stop["port_name"]
                break


# =============================================================================
# 3. Voyage Data Extraction  (Rule 1: ROB difference, distance sum)
# =============================================================================

def _safe_float(value) -> float:
    """Convert value to float, returning 0.0 for NaN/None."""
    if pd.isna(value):
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def extract_voyage_data(df: pd.DataFrame, dep_row: int, arr_row: int) -> dict:
    """
    Extract all voyage-level data for a single DEPARTURE→ARRIVAL span.

    Rule 1: Uses ROB difference method for fuel consumption.
    All column access is by index from config.COL.

    Returns dict with:
      - total_distance:    float — sum of Col P from dep to arr
      - mgo_consumed:      float — ROB at DEP minus ROB at ARR (Col CB)
      - lng_consumed:      float — ROB at DEP minus ROB at ARR (Col CU)
      - vlsfo_consumed:    float — Grade 1 (Col EG) + Grade 2 (Col GL)
      - vlsfo_g1_consumed: float — Grade 1 only
      - vlsfo_g2_consumed: float — Grade 2 only
      - total_steaming_hrs: float — sum of Col AT2 (steaming hours)
      - daily_rows:        list[dict] — per-row data for segment/exclusion calcs
      - dep_datetime:      parsed datetime
      - arr_datetime:      parsed datetime
    """
    # --- Total distance: sum Col P (index 15) ---
    total_distance = 0.0
    for i in range(dep_row, arr_row + 1):
        total_distance += _safe_float(df.iloc[i, COL["distance"]])

    # --- Fuel consumption: ROB difference (Rule 1) ---
    mgo_dep = _safe_float(df.iloc[dep_row, COL["mgo_rob"]])
    mgo_arr = _safe_float(df.iloc[arr_row, COL["mgo_rob"]])
    mgo_consumed = mgo_dep - mgo_arr

    lng_dep = _safe_float(df.iloc[dep_row, COL["lng_rob"]])
    lng_arr = _safe_float(df.iloc[arr_row, COL["lng_rob"]])
    lng_consumed = lng_dep - lng_arr

    vlsfo_g1_dep = _safe_float(df.iloc[dep_row, COL["vlsfo_g1_rob"]])
    vlsfo_g1_arr = _safe_float(df.iloc[arr_row, COL["vlsfo_g1_rob"]])
    vlsfo_g1_consumed = vlsfo_g1_dep - vlsfo_g1_arr

    vlsfo_g2_dep = _safe_float(df.iloc[dep_row, COL["vlsfo_g2_rob"]])
    vlsfo_g2_arr = _safe_float(df.iloc[arr_row, COL["vlsfo_g2_rob"]])
    vlsfo_g2_consumed = vlsfo_g2_dep - vlsfo_g2_arr

    vlsfo_consumed = vlsfo_g1_consumed + vlsfo_g2_consumed

    # --- Total steaming hours ---
    total_steaming_hrs = 0.0
    for i in range(dep_row, arr_row + 1):
        total_steaming_hrs += _safe_float(df.iloc[i, COL["steaming_hrs"]])

    # --- Daily rows: per-row data (ROB diffs, steaming, BF5, etc.) ---
    daily_rows = _build_daily_rows(df, dep_row, arr_row)

    # --- Date/time ---
    dep_datetime = pd.to_datetime(df.iloc[dep_row, COL["datetime"]], errors="coerce")
    arr_datetime = pd.to_datetime(df.iloc[arr_row, COL["datetime"]], errors="coerce")

    return {
        "total_distance":     total_distance,
        "mgo_consumed":       mgo_consumed,
        "lng_consumed":       lng_consumed,
        "vlsfo_consumed":     vlsfo_consumed,
        "vlsfo_g1_consumed":  vlsfo_g1_consumed,
        "vlsfo_g2_consumed":  vlsfo_g2_consumed,
        "total_steaming_hrs": total_steaming_hrs,
        "daily_rows":         daily_rows,
        "dep_datetime":       dep_datetime,
        "arr_datetime":       arr_datetime,
        "dep_row":            dep_row,
        "arr_row":            arr_row,
    }


def _build_daily_rows(df: pd.DataFrame, dep_row: int, arr_row: int) -> list[dict]:
    """
    Build per-row data for the voyage, including ROB-derived daily consumption.

    Daily consumption = ROB(previous row) − ROB(current row)
    First row (DEPARTURE) has no daily consumption (set to 0).
    """
    rows = []

    for i in range(dep_row, arr_row + 1):
        row = {
            "df_idx":       i,
            "datetime":     df.iloc[i, COL["datetime"]],
            "report_type":  str(df.iloc[i, COL["report_type"]]).strip()
                            if pd.notna(df.iloc[i, COL["report_type"]]) else "",
            "distance":     _safe_float(df.iloc[i, COL["distance"]]),
            "steaming_hrs": _safe_float(df.iloc[i, COL["steaming_hrs"]]),
            "bf5_hours":    _safe_float(df.iloc[i, COL["bf5_hours"]]),
            "avg_speed":    _safe_float(df.iloc[i, COL["avg_speed"]]),

            # Current ROB values
            "mgo_rob":       _safe_float(df.iloc[i, COL["mgo_rob"]]),
            "lng_rob":       _safe_float(df.iloc[i, COL["lng_rob"]]),
            "vlsfo_g1_rob":  _safe_float(df.iloc[i, COL["vlsfo_g1_rob"]]),
            "vlsfo_g2_rob":  _safe_float(df.iloc[i, COL["vlsfo_g2_rob"]]),

            # Segment detection columns (Rule 3)
            "next_port":     df.iloc[i, COL["next_port"]]
                             if pd.notna(df.iloc[i, COL["next_port"]]) else None,
            "fuel_mode":     df.iloc[i, COL["fuel_mode"]]
                             if pd.notna(df.iloc[i, COL["fuel_mode"]]) else None,
            "eta_next":      df.iloc[i, COL["eta_next"]]
                             if pd.notna(df.iloc[i, COL["eta_next"]]) else None,
            "voyage_order_rev": str(df.iloc[i, COL["voyage_order_rev"]]).strip().lower()
                             if pd.notna(df.iloc[i, COL["voyage_order_rev"]]) else "",
            "rev_start_time": df.iloc[i, COL["rev_start_time"]]
                             if pd.notna(df.iloc[i, COL["rev_start_time"]]) else None,
            "rev_gmt_offset": _safe_float(df.iloc[i, COL["rev_gmt_offset"]]),
            "rev_speed":     _safe_float(df.iloc[i, COL["rev_speed"]]),
            "rev_sat":       df.iloc[i, COL["rev_sat"]]
                             if pd.notna(df.iloc[i, COL["rev_sat"]]) else None,

            # Auxiliary columns (Rule 4)
            "mgo_boiler":     _safe_float(df.iloc[i, COL["mgo_boiler"]]),
            "mgo_pilot":      _safe_float(df.iloc[i, COL["mgo_pilot"]]),
            "vlsfo_g1_boiler": _safe_float(df.iloc[i, COL["vlsfo_g1_boiler"]]),
            "vlsfo_g1_pilot":  _safe_float(df.iloc[i, COL["vlsfo_g1_pilot"]]),
            "vlsfo_g2_boiler": _safe_float(df.iloc[i, COL["vlsfo_g2_boiler"]]),
            "vlsfo_g2_pilot":  _safe_float(df.iloc[i, COL["vlsfo_g2_pilot"]]),
            "gcu_lng":        _safe_float(df.iloc[i, COL["gcu_lng"]]),
            "reliq_hours":    _safe_float(df.iloc[i, COL["reliq_hours"]]),
            "reliq_load":     _safe_float(df.iloc[i, COL["reliq_load"]]),
            "subcooler_hours": _safe_float(df.iloc[i, COL["subcooler_hours"]]),
            "subcooler_load": _safe_float(df.iloc[i, COL["subcooler_load"]]),

            # Remarks
            "remarks":        str(df.iloc[i, COL["remarks"]])
                              if pd.notna(df.iloc[i, COL["remarks"]]) else "",

            # Ordered speed (for first segment)
            "ordered_speed":  _safe_float(df.iloc[i, COL["ordered_speed"]]),
        }

        # --- Daily consumption via ROB difference (Rule 1) ---
        if i > dep_row:
            prev = i - 1
            row["mgo_daily"] = (
                _safe_float(df.iloc[prev, COL["mgo_rob"]])
                - _safe_float(df.iloc[i, COL["mgo_rob"]])
            )
            row["lng_daily"] = (
                _safe_float(df.iloc[prev, COL["lng_rob"]])
                - _safe_float(df.iloc[i, COL["lng_rob"]])
            )
            row["vlsfo_g1_daily"] = (
                _safe_float(df.iloc[prev, COL["vlsfo_g1_rob"]])
                - _safe_float(df.iloc[i, COL["vlsfo_g1_rob"]])
            )
            row["vlsfo_g2_daily"] = (
                _safe_float(df.iloc[prev, COL["vlsfo_g2_rob"]])
                - _safe_float(df.iloc[i, COL["vlsfo_g2_rob"]])
            )
            row["vlsfo_daily"] = row["vlsfo_g1_daily"] + row["vlsfo_g2_daily"]
        else:
            # DEPARTURE row — no consumption yet
            row["mgo_daily"] = 0.0
            row["lng_daily"] = 0.0
            row["vlsfo_g1_daily"] = 0.0
            row["vlsfo_g2_daily"] = 0.0
            row["vlsfo_daily"] = 0.0

        rows.append(row)

    return rows


# =============================================================================
# 4. Auxiliary Data Extraction  (Rule 4)
# =============================================================================

def extract_auxiliary(df: pd.DataFrame, dep_row: int, arr_row: int) -> dict:
    """
    Extract auxiliary consumption data for the voyage (Rule 4).

    Returns dict with:
      - mgo_boiler_total, mgo_pilot_total
      - vlsfo_boiler_total (Grade 1 + Grade 2), vlsfo_pilot_total
      - gcu_lng_total, gcu_used (bool), gcu_dates (list of date strings)
      - cargo_density, lcv (constant or averaged)
      - reliq_total_hours, reliq_avg_load
    """
    mgo_boiler = 0.0
    mgo_pilot = 0.0
    vlsfo_boiler = 0.0
    vlsfo_pilot = 0.0
    gcu_total = 0.0
    gcu_dates = []

    reliq_hours_total = 0.0
    load_values = []

    density_values = []
    lcv_values = []

    for i in range(dep_row, arr_row + 1):
        # Boiler / Pilot sums (Rule 4)
        mgo_boiler += _safe_float(df.iloc[i, COL["mgo_boiler"]])
        mgo_pilot  += _safe_float(df.iloc[i, COL["mgo_pilot"]])
        vlsfo_boiler += (
            _safe_float(df.iloc[i, COL["vlsfo_g1_boiler"]])
            + _safe_float(df.iloc[i, COL["vlsfo_g2_boiler"]])
        )
        vlsfo_pilot += (
            _safe_float(df.iloc[i, COL["vlsfo_g1_pilot"]])
            + _safe_float(df.iloc[i, COL["vlsfo_g2_pilot"]])
        )

        # GCU (Rule 4)
        gcu_val = _safe_float(df.iloc[i, COL["gcu_lng"]])
        gcu_total += gcu_val
        if gcu_val > 0:
            dt = df.iloc[i, COL["datetime"]]
            gcu_dates.append(str(dt)[:10] if pd.notna(dt) else f"row_{i}")

        # Reliq + Subcooler hours (Rule 1: IL + IQ)
        reliq_hours_total += (
            _safe_float(df.iloc[i, COL["reliq_hours"]])
            + _safe_float(df.iloc[i, COL["subcooler_hours"]])
        )

        # Reliq + Subcooler load (Rule 1: avg of IN + IR, non-zero only)
        for load_col in [COL["reliq_load"], COL["subcooler_load"]]:
            v = _safe_float(df.iloc[i, load_col])
            if v > 0:
                load_values.append(v)

        # Density & LCV (Rule 4)
        d = df.iloc[i, COL["cargo_density"]]
        if pd.notna(d):
            try:
                dv = float(d)
                if dv > 0:
                    density_values.append(dv)
            except (ValueError, TypeError):
                pass

        l_val = df.iloc[i, COL["lcv"]]
        if pd.notna(l_val):
            try:
                lv = float(l_val)
                if lv > 0:
                    lcv_values.append(lv)
            except (ValueError, TypeError):
                pass

    # Density: constant → as-is, varying → average (Rule 4)
    if density_values:
        unique_d = set(density_values)
        cargo_density = density_values[0] if len(unique_d) == 1 else float(np.mean(density_values))
    else:
        cargo_density = None

    # LCV: same logic (Rule 4)
    if lcv_values:
        unique_l = set(lcv_values)
        lcv = lcv_values[0] if len(unique_l) == 1 else float(np.mean(lcv_values))
    else:
        lcv = None

    # Average load (non-zero values only) (Rule 1)
    reliq_avg_load = float(np.mean(load_values)) if load_values else 0.0

    return {
        "mgo_boiler_total":   mgo_boiler,
        "mgo_pilot_total":    mgo_pilot,
        "vlsfo_boiler_total": vlsfo_boiler,
        "vlsfo_pilot_total":  vlsfo_pilot,
        "gcu_lng_total":      gcu_total,
        "gcu_used":           gcu_total > 0,
        "gcu_dates":          gcu_dates,
        "cargo_density":      cargo_density,
        "lcv":                lcv,
        "reliq_total_hours":  reliq_hours_total,
        "reliq_avg_load":     reliq_avg_load,
    }
