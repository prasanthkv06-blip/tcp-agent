"""
calculator.py
=============
Vessel performance calculation engine.

Implements:
  - Rule 2: Weather exclusion (BF>5 for >6 hours, steaming hours cap)
  - Rule 3: Voyage segmentation (J/K/L changes, voyage order revisions)
  - Rule 3: Boundary row pro-rating (speed-weighted ratio)
  - Speed/consumption curve interpolation (TCP Appendix A)
  - Per-segment and voyage-level performance calculations
"""

from __future__ import annotations

import logging
from typing import Any

import numpy as np
import pandas as pd

from config import WEATHER_BF5_THRESHOLD, load_vessel_config

logger = logging.getLogger(__name__)


# =============================================================================
# 1. Speed/Consumption Curve Interpolation  (Appendix A)
# =============================================================================

def interpolate_consumption(
    speed: float,
    condition: str = "laden",
    component: str = "base_gas",
    vessel_config: dict | None = None,
) -> float:
    """
    Interpolate guaranteed daily fuel consumption from the TCP
    speed/consumption table (Part II, Table 6).

    Per Appendix A, Article 2(a)(iii): if the speed is between two
    table speeds, interpolate arithmetically.

    Parameters
    ----------
    speed         : Achieved or ordered speed in knots.
    condition     : 'laden' or 'ballast'.
    component     : 'base_gas' or 'pilot'.
    vessel_config : Vessel configuration dict (from load_vessel_config).

    Returns
    -------
    float  Guaranteed consumption in MT/day.
    """
    if vessel_config is None:
        vessel_config = load_vessel_config()

    table = vessel_config["speed_consumption_table"]
    speeds_sorted = sorted(table.keys())

    # Clamp to table range
    if speed <= speeds_sorted[0]:
        speed = speeds_sorted[0]
    elif speed >= speeds_sorted[-1]:
        speed = speeds_sorted[-1]

    # Column indices: (laden_gas, laden_pilot, ballast_gas, ballast_pilot)
    if condition.lower().startswith("l"):
        gas_idx, pilot_idx = 0, 1
    else:
        gas_idx, pilot_idx = 2, 3

    idx = pilot_idx if "pilot" in component.lower() else gas_idx

    # Exact match?
    if speed in table:
        return table[speed][idx]

    # Find bracketing speeds and interpolate
    lower_speed = max(s for s in speeds_sorted if s <= speed)
    upper_speed = min(s for s in speeds_sorted if s >= speed)

    if lower_speed == upper_speed:
        return table[lower_speed][idx]

    lower_val = table[lower_speed][idx]
    upper_val = table[upper_speed][idx]

    fraction = (speed - lower_speed) / (upper_speed - lower_speed)
    return lower_val + fraction * (upper_val - lower_val)


def get_guaranteed_daily_consumption(
    speed: float,
    condition: str = "laden",
    ageing_factor: float = 1.0,
    vessel_config: dict | None = None,
) -> dict[str, float]:
    """
    Get the full guaranteed daily fuel consumption for a given speed.

    Returns dict with keys: base_gas_mt, pilot_mt, boiler_mt,
    total_gas_mt, total_fuel_mt (in MT/day).
    """
    if vessel_config is None:
        vessel_config = load_vessel_config()

    base_gas = interpolate_consumption(speed, condition, "base_gas", vessel_config)
    pilot = interpolate_consumption(speed, condition, "pilot", vessel_config)

    if condition.lower().startswith("l"):
        boiler = vessel_config["boiler_cons_laden_mt"]
    else:
        boiler = vessel_config["boiler_cons_ballast_mt"]

    # Apply ageing factor to base gas consumption
    base_gas *= ageing_factor

    return {
        "base_gas_mt":   round(base_gas, 3),
        "pilot_mt":      round(pilot, 3),
        "boiler_mt":     round(boiler, 3),
        "total_gas_mt":  round(base_gas + boiler, 3),
        "total_fuel_mt": round(base_gas + boiler + pilot, 3),
    }


# =============================================================================
# 2. Weather Exclusion  (Rule 2)
# =============================================================================

def compute_weather_exclusion(
    daily_cons: float,
    steaming_hrs: float,
    bf5_hours: float,
    threshold: float = WEATHER_BF5_THRESHOLD,
) -> float:
    """
    Calculate excluded fuel for a single day due to adverse weather.

    Rule 2:
      If BF5 hours > threshold (6 hrs):
        Excluded = (daily_cons / steaming_hrs) × MIN(bf5_hours, steaming_hrs)

    Steaming hours is the ONLY denominator (not 24), due to timezone changes.
    BF5 hours are capped at steaming hours to prevent overestimation.

    Returns 0.0 if BF5 hours ≤ threshold or steaming_hrs ≤ 0.
    """
    if bf5_hours <= threshold or steaming_hrs <= 0 or daily_cons <= 0:
        return 0.0

    excluded_hours = min(bf5_hours, steaming_hrs)
    return (daily_cons / steaming_hrs) * excluded_hours


def tag_weather_exclusions(daily_rows: list[dict]) -> None:
    """
    Pre-compute weather exclusion for each row (modifies rows in-place).

    Must be called on ORIGINAL daily_rows BEFORE any pro-rating, so that
    threshold checks use the full day's BF5 hours and exclusion amounts
    can later be pro-rated correctly.
    """
    for row in daily_rows:
        bf5 = row.get("bf5_hours", 0.0)
        steaming = row.get("steaming_hrs", 0.0)

        if bf5 > WEATHER_BF5_THRESHOLD and steaming > 0:
            excl_hrs = min(bf5, steaming)
            row["weather_excluded"] = True
            row["excl_hours"] = excl_hrs
            row["excl_mgo"] = compute_weather_exclusion(
                row.get("mgo_daily", 0.0), steaming, bf5
            )
            row["excl_lng"] = compute_weather_exclusion(
                row.get("lng_daily", 0.0), steaming, bf5
            )
            row["excl_vlsfo"] = compute_weather_exclusion(
                row.get("vlsfo_daily", 0.0), steaming, bf5
            )
            row["excl_distance"] = (
                (row.get("distance", 0.0) / steaming * excl_hrs)
                if steaming > 0 else 0.0
            )
        else:
            row["weather_excluded"] = False
            row["excl_hours"] = 0.0
            row["excl_mgo"] = 0.0
            row["excl_lng"] = 0.0
            row["excl_vlsfo"] = 0.0
            row["excl_distance"] = 0.0


def apply_weather_exclusions(daily_rows: list[dict]) -> dict:
    """
    Apply Rule 2 weather exclusion across all daily rows.

    Uses ROB-derived daily consumption (from data_extractor daily_rows).
    Steaming hours as denominator, BF5 capped at steaming hours.

    Returns dict with:
      - excluded_mgo, excluded_lng, excluded_vlsfo  (totals)
      - excluded_hours, excluded_distance
      - excluded_rows (list of detail dicts)
      - total_bf5_hours
    """
    # Tag if not already done
    if daily_rows and "weather_excluded" not in daily_rows[0]:
        tag_weather_exclusions(daily_rows)

    excluded_rows = [
        {
            "df_idx":       r.get("df_idx"),
            "datetime":     r.get("datetime"),
            "bf5_hours":    r.get("bf5_hours", 0),
            "steaming_hrs": r.get("steaming_hrs", 0),
            "excl_hours":   r.get("excl_hours", 0),
            "excl_mgo":     r.get("excl_mgo", 0),
            "excl_lng":     r.get("excl_lng", 0),
            "excl_vlsfo":   r.get("excl_vlsfo", 0),
            "excl_distance": r.get("excl_distance", 0),
        }
        for r in daily_rows if r.get("weather_excluded")
    ]

    return {
        "excluded_mgo":      sum(r.get("excl_mgo", 0) for r in daily_rows),
        "excluded_lng":      sum(r.get("excl_lng", 0) for r in daily_rows),
        "excluded_vlsfo":    sum(r.get("excl_vlsfo", 0) for r in daily_rows),
        "excluded_hours":    sum(r.get("excl_hours", 0) for r in daily_rows),
        "excluded_distance": sum(r.get("excl_distance", 0) for r in daily_rows),
        "excluded_rows":     excluded_rows,
        "total_bf5_hours":   sum(r.get("bf5_hours", 0) for r in daily_rows),
    }


# =============================================================================
# 2b. Speed Anomaly Detection  (Rule A)
# =============================================================================

def detect_speed_anomalies(
    daily_rows: list[dict],
    threshold_ratio: float = 0.10,
) -> list[dict]:
    """
    Detect rows where vessel likely stopped or drastically reduced speed.

    Rule A:
      For each row, compute weighted average speed from departure up to that row:
        weighted_avg = cumulative_distance / cumulative_steaming_hours
      If row's avg_speed < threshold_ratio × weighted_avg → flag it.

    Parameters
    ----------
    daily_rows      : list[dict] from data_extractor._build_daily_rows
    threshold_ratio : fraction of weighted avg below which a row is flagged (default 0.10)

    Returns
    -------
    list[dict] — flagged rows, each with:
      - df_idx, datetime, avg_speed, weighted_avg, report_type
    """
    flagged = []
    cum_distance = 0.0
    cum_steaming = 0.0

    for row in daily_rows:
        cum_distance += row.get("distance", 0.0)
        cum_steaming += row.get("steaming_hrs", 0.0)

        if cum_steaming <= 0:
            continue

        weighted_avg = cum_distance / cum_steaming
        if weighted_avg <= 0:
            continue

        avg_speed = row.get("avg_speed", 0.0)
        if avg_speed < threshold_ratio * weighted_avg:
            flagged.append({
                "df_idx":       row.get("df_idx"),
                "datetime":     row.get("datetime"),
                "avg_speed":    avg_speed,
                "weighted_avg": round(weighted_avg, 3),
                "report_type":  row.get("report_type", ""),
            })

    return flagged


# =============================================================================
# 3. Voyage Segmentation  (Rule 3)
# =============================================================================

def _parse_datetime(value, gmt_offset=None):
    """Parse a datetime value, adjusting for GMT offset if provided."""
    if value is None:
        return None
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return None
    if gmt_offset is not None and gmt_offset != 0:
        dt = dt - pd.Timedelta(hours=float(gmt_offset))
    return dt


def detect_segments(
    daily_rows: list[dict],
    dep_datetime,
    arr_datetime,
) -> list[dict]:
    """
    Detect voyage segments based on Rule 3.

    Monitors:
      - Col BM (voyage_order_rev) = "yes" → voyage order revision
        → New segment starts at Col BN datetime
        → Boundary row is shared (pro-rated)
      - Col J (next_port), K (fuel_mode) for field changes
        → Boundary row belongs to new segment

    Segment 1 always starts at DEPARTURE.
    Last segment always ends at ARRIVAL.

    Parameters
    ----------
    daily_rows   : list[dict] from data_extractor._build_daily_rows
    dep_datetime : DEPARTURE datetime
    arr_datetime : ARRIVAL datetime

    Returns
    -------
    list[dict] — each segment with start/end rows, datetimes, speed, etc.
    """
    if not daily_rows:
        return []

    segments = []
    current_start = 0
    current_speed = daily_rows[0].get("ordered_speed", 0.0)
    current_fuel_mode = daily_rows[0].get("fuel_mode")
    current_next_port = daily_rows[0].get("next_port")

    last_idx = len(daily_rows) - 1  # ARRIVAL row — never triggers a new segment

    for i in range(1, len(daily_rows)):
        row = daily_rows[i]
        boundary_type = None
        new_speed = None
        new_sat = None

        # Skip the ARRIVAL row for segment detection —
        # it's always the end of the final segment
        if i == last_idx:
            continue

        # Priority 1: Voyage order revision (BM = "yes")
        if row.get("voyage_order_rev", "").lower() == "yes":
            boundary_type = "voyage_order"
            new_speed = row.get("rev_speed", 0.0)
            new_sat = row.get("rev_sat")
            logger.info(
                "Voyage order revision at daily_rows[%d] (df_idx=%d), "
                "new speed=%.1f kts",
                i, row.get("df_idx", i), new_speed or 0,
            )

        # Priority 2: Field changes (next_port, fuel_mode)
        if not boundary_type:
            if (row.get("next_port") is not None
                    and row.get("next_port") != current_next_port):
                boundary_type = "field_change"
                logger.info(
                    "Next port change at daily_rows[%d]: '%s' → '%s'",
                    i, current_next_port, row.get("next_port"),
                )
            elif (row.get("fuel_mode") is not None
                  and row.get("fuel_mode") != current_fuel_mode):
                boundary_type = "field_change"
                logger.info(
                    "Fuel mode change at daily_rows[%d]: '%s' → '%s'",
                    i, current_fuel_mode, row.get("fuel_mode"),
                )

        if boundary_type:
            # --- Close the current segment ---

            # Determine segment end datetime
            if boundary_type == "voyage_order":
                rev_dt = row.get("rev_start_time")
                gmt_off = row.get("rev_gmt_offset", 0)
                end_dt = _parse_datetime(rev_dt, gmt_off)
                if end_dt is None:
                    end_dt = _parse_datetime(row.get("datetime"))
            else:
                end_dt = _parse_datetime(row.get("datetime"))

            # Determine segment start datetime
            if not segments:
                start_dt = _parse_datetime(dep_datetime)
            else:
                start_dt = segments[-1]["end_datetime"]

            # For voyage_order: boundary row is shared (end_row_idx = i)
            # For field_change: boundary row goes to NEXT segment (end_row_idx = i-1)
            if boundary_type == "voyage_order":
                end_row = i
                is_shared = True
            else:
                end_row = i - 1
                is_shared = False

            segments.append({
                "start_row_idx":     current_start,
                "end_row_idx":       end_row,
                "start_datetime":    start_dt,
                "end_datetime":      end_dt,
                "instructed_speed":  current_speed,
                "fuel_mode":         current_fuel_mode,
                "boundary_type":     boundary_type,
                "is_boundary_shared": is_shared,
                "new_sat":           new_sat,
            })

            # --- Start new segment ---
            current_start = i
            if new_speed and new_speed > 0:
                current_speed = new_speed
            current_fuel_mode = row.get("fuel_mode", current_fuel_mode)
            current_next_port = row.get("next_port", current_next_port)

    # --- Close the final segment (ends at ARRIVAL) ---
    if not segments:
        start_dt = _parse_datetime(dep_datetime)
    else:
        start_dt = segments[-1]["end_datetime"]

    segments.append({
        "start_row_idx":     current_start,
        "end_row_idx":       len(daily_rows) - 1,
        "start_datetime":    start_dt,
        "end_datetime":      _parse_datetime(arr_datetime),
        "instructed_speed":  current_speed,
        "fuel_mode":         current_fuel_mode,
        "boundary_type":     "final",
        "is_boundary_shared": False,
        "new_sat":           None,
    })

    logger.info("Detected %d segment(s)", len(segments))
    for idx, seg in enumerate(segments):
        logger.info(
            "  Segment %d: rows[%d:%d], speed=%.1f kts, %s → %s",
            idx + 1, seg["start_row_idx"], seg["end_row_idx"],
            seg["instructed_speed"],
            str(seg["start_datetime"])[:16],
            str(seg["end_datetime"])[:16],
        )

    return segments


# =============================================================================
# 4. Boundary Row Pro-Rating  (Rule 3)
# =============================================================================

def prorate_boundary_row(
    daily_rows: list[dict],
    boundary_idx: int,
    prev_row_datetime,
    rev_start_datetime,
    current_row_datetime,
    speed_before: float,
    speed_after: float,
) -> tuple[dict, dict]:
    """
    Pro-rate a boundary row's distance and fuel between two segments
    using speed-weighted ratios (Rule 3).

    When a voyage order change (BM="yes") occurs mid-day:
      1. Calculate hours before/after the revision start time
      2. Theoretical distance = speed × hours for each portion
      3. Ratio = theoretical / total_theoretical
      4. Apply ratio to actual Col P distance and ROB-diff fuel

    Parameters
    ----------
    daily_rows          : full daily_rows list
    boundary_idx        : index into daily_rows of the boundary row
    prev_row_datetime   : datetime of the previous row
    rev_start_datetime  : datetime of the revision start (Col BN)
    current_row_datetime: datetime of the boundary row
    speed_before        : speed for the ending segment (Col AT avg speed)
    speed_after         : revised speed for the new segment (Col BS)

    Returns
    -------
    (before_portion, after_portion) — dicts with pro-rated data
    """
    row = daily_rows[boundary_idx]

    prev_dt = _parse_datetime(prev_row_datetime)
    rev_dt = _parse_datetime(rev_start_datetime)
    curr_dt = _parse_datetime(current_row_datetime)

    if prev_dt is None or rev_dt is None or curr_dt is None:
        logger.warning(
            "Cannot pro-rate boundary row %d — missing datetime(s). "
            "Assigning entire row to Segment N.",
            boundary_idx,
        )
        return _make_portion(row, 1.0), _make_portion(row, 0.0)

    # Calculate hours before and after the revision start
    hours_before = max((rev_dt - prev_dt).total_seconds() / 3600.0, 0.0)
    hours_after = max((curr_dt - rev_dt).total_seconds() / 3600.0, 0.0)

    if hours_before + hours_after <= 0:
        return _make_portion(row, 1.0), _make_portion(row, 0.0)

    # Theoretical distances
    theo_before = speed_before * hours_before
    theo_after = speed_after * hours_after
    theo_total = theo_before + theo_after

    if theo_total <= 0:
        r_before, r_after = 0.5, 0.5
    else:
        r_before = theo_before / theo_total
        r_after = theo_after / theo_total

    logger.info(
        "Boundary row %d pro-rate: %.1fh @ %.1fkts + %.1fh @ %.1fkts "
        "→ ratios %.3f / %.3f",
        boundary_idx, hours_before, speed_before,
        hours_after, speed_after, r_before, r_after,
    )

    return _make_portion(row, r_before), _make_portion(row, r_after)


def _make_portion(row: dict, ratio: float) -> dict:
    """Create a pro-rated copy of a daily row's numeric data."""
    return {
        "distance":         row.get("distance", 0.0) * ratio,
        "steaming_hrs":     row.get("steaming_hrs", 0.0) * ratio,
        "mgo_daily":        row.get("mgo_daily", 0.0) * ratio,
        "lng_daily":        row.get("lng_daily", 0.0) * ratio,
        "vlsfo_daily":      row.get("vlsfo_daily", 0.0) * ratio,
        "vlsfo_g1_daily":   row.get("vlsfo_g1_daily", 0.0) * ratio,
        "vlsfo_g2_daily":   row.get("vlsfo_g2_daily", 0.0) * ratio,
        "bf5_hours":        row.get("bf5_hours", 0.0) * ratio,
        "mgo_boiler":       row.get("mgo_boiler", 0.0) * ratio,
        "mgo_pilot":        row.get("mgo_pilot", 0.0) * ratio,
        "vlsfo_g1_boiler":  row.get("vlsfo_g1_boiler", 0.0) * ratio,
        "vlsfo_g1_pilot":   row.get("vlsfo_g1_pilot", 0.0) * ratio,
        "vlsfo_g2_boiler":  row.get("vlsfo_g2_boiler", 0.0) * ratio,
        "vlsfo_g2_pilot":   row.get("vlsfo_g2_pilot", 0.0) * ratio,
        "gcu_lng":          row.get("gcu_lng", 0.0) * ratio,
        "reliq_hours":      row.get("reliq_hours", 0.0) * ratio,
        "subcooler_hours":  row.get("subcooler_hours", 0.0) * ratio,
        # Weather exclusion tags (pre-computed, now pro-rated)
        "excl_hours":       row.get("excl_hours", 0.0) * ratio,
        "excl_mgo":         row.get("excl_mgo", 0.0) * ratio,
        "excl_lng":         row.get("excl_lng", 0.0) * ratio,
        "excl_vlsfo":       row.get("excl_vlsfo", 0.0) * ratio,
        "excl_distance":    row.get("excl_distance", 0.0) * ratio,
        "weather_excluded": row.get("weather_excluded", False),
        # Metadata (not pro-rated)
        "ratio":            ratio,
        "is_prorated":      True,
        "df_idx":           row.get("df_idx"),
        "datetime":         row.get("datetime"),
        "remarks":          row.get("remarks", ""),
    }


# =============================================================================
# 5. Build Segment Rows  (assigns daily rows to segments with pro-rating)
# =============================================================================

def build_segment_rows(
    segments: list[dict],
    daily_rows: list[dict],
) -> list[list[dict]]:
    """
    Assign daily rows to segments, handling boundary row pro-rating.

    For voyage order boundaries (is_boundary_shared=True):
      - Previous segment's last row: "before" portion
      - Next segment's first row: "after" portion

    For field change boundaries:
      - Previous segment ends at row i-1 (no pro-rating)
      - Next segment starts at row i (full row)

    Returns a list of lists, one per segment.
    """
    if not segments:
        return []

    result = []

    for seg_idx, seg in enumerate(segments):
        start = seg["start_row_idx"]
        end = seg["end_row_idx"]
        seg_rows = []

        for i in range(start, end + 1):
            row = daily_rows[i].copy()

            # --- Handle shared boundary at END of this segment ---
            if (i == end
                    and seg.get("is_boundary_shared")
                    and seg_idx < len(segments) - 1):
                brow = daily_rows[i]
                if brow.get("voyage_order_rev", "").lower() == "yes":
                    prev_dt = (
                        daily_rows[i - 1].get("datetime") if i > 0
                        else seg["start_datetime"]
                    )
                    rev_dt = brow.get("rev_start_time")
                    curr_dt = brow.get("datetime")
                    avg_spd = brow.get("avg_speed", 0.0)
                    speed_before = (
                        avg_spd if avg_spd > 0
                        else seg["instructed_speed"]
                    )
                    speed_after = brow.get("rev_speed", 0.0)

                    if rev_dt and speed_after > 0:
                        before, _ = prorate_boundary_row(
                            daily_rows, i, prev_dt, rev_dt, curr_dt,
                            speed_before, speed_after,
                        )
                        row.update(before)
                        row["is_prorated"] = True

            # --- Handle shared boundary at START of this segment ---
            elif i == start and seg_idx > 0:
                prev_seg = segments[seg_idx - 1]
                if prev_seg.get("is_boundary_shared"):
                    brow = daily_rows[i]
                    if brow.get("voyage_order_rev", "").lower() == "yes":
                        prev_dt = (
                            daily_rows[i - 1].get("datetime") if i > 0
                            else prev_seg["start_datetime"]
                        )
                        rev_dt = brow.get("rev_start_time")
                        curr_dt = brow.get("datetime")
                        avg_spd = brow.get("avg_speed", 0.0)
                        speed_before = (
                            avg_spd if avg_spd > 0
                            else prev_seg["instructed_speed"]
                        )
                        speed_after = brow.get("rev_speed", 0.0)

                        if rev_dt and speed_after > 0:
                            _, after = prorate_boundary_row(
                                daily_rows, i, prev_dt, rev_dt, curr_dt,
                                speed_before, speed_after,
                            )
                            row.update(after)
                            row["is_prorated"] = True

            seg_rows.append(row)

        result.append(seg_rows)

    return result


# =============================================================================
# 6. Per-Segment Performance Calculations
# =============================================================================

def compute_segment_data(
    seg_info: dict,
    seg_rows: list[dict],
    vessel_config: dict | None = None,
) -> dict:
    """
    Compute all performance data for a single segment.

    seg_rows should already have pro-rated boundary rows and
    weather exclusion tags.

    Returns dict with all fields needed for the standard template:
      Basic info, Exclusions, Speed, Actual Fuel, Fuel Exclusions,
      Net Fuel, GCU, Reliq, Remarks.
    """
    if vessel_config is None:
        vessel_config = load_vessel_config()

    # --- Sum distance, fuel, steaming hours ---
    distance = sum(r.get("distance", 0) for r in seg_rows)
    steaming = sum(r.get("steaming_hrs", 0) for r in seg_rows)
    mgo_consumed = sum(r.get("mgo_daily", 0) for r in seg_rows)
    lng_consumed = sum(r.get("lng_daily", 0) for r in seg_rows)
    vlsfo_consumed = sum(r.get("vlsfo_daily", 0) for r in seg_rows)
    vlsfo_g1_consumed = sum(r.get("vlsfo_g1_daily", 0) for r in seg_rows)
    vlsfo_g2_consumed = sum(r.get("vlsfo_g2_daily", 0) for r in seg_rows)

    # --- Boiler / Pilot breakdown (Rule 4) ---
    mgo_boiler = sum(r.get("mgo_boiler", 0) for r in seg_rows)
    mgo_pilot = sum(r.get("mgo_pilot", 0) for r in seg_rows)
    vlsfo_boiler = sum(
        r.get("vlsfo_g1_boiler", 0) + r.get("vlsfo_g2_boiler", 0)
        for r in seg_rows
    )
    vlsfo_pilot = sum(
        r.get("vlsfo_g1_pilot", 0) + r.get("vlsfo_g2_pilot", 0)
        for r in seg_rows
    )

    # Propulsion = total - pilot - boiler
    mgo_propulsion = max(mgo_consumed - mgo_pilot - mgo_boiler, 0.0)
    vlsfo_propulsion = max(vlsfo_consumed - vlsfo_pilot - vlsfo_boiler, 0.0)

    # --- GCU ---
    gcu_total = sum(r.get("gcu_lng", 0) for r in seg_rows)
    gcu_dates = [
        str(r.get("datetime"))[:10]
        for r in seg_rows
        if r.get("gcu_lng", 0) > 0
    ]

    # --- Reliq/Subcooler (Rule 1: IL+IQ hours, avg IN+IR load) ---
    reliq_hours = sum(
        r.get("reliq_hours", 0) + r.get("subcooler_hours", 0)
        for r in seg_rows
    )
    load_vals = []
    for r in seg_rows:
        for key in ("reliq_load", "subcooler_load"):
            v = r.get(key, 0)
            if v and v > 0:
                load_vals.append(v)
    reliq_avg_load = float(np.mean(load_vals)) if load_vals else 0.0

    # --- Remarks ---
    remarks = [
        r.get("remarks", "")
        for r in seg_rows
        if r.get("remarks", "")
    ]

    # --- Duration ---
    start_dt = seg_info.get("start_datetime")
    end_dt = seg_info.get("end_datetime")
    if start_dt is not None and end_dt is not None:
        start_ts = pd.Timestamp(start_dt)
        end_ts = pd.Timestamp(end_dt)
        duration_days = (end_ts - start_ts).total_seconds() / 86400.0
    else:
        duration_days = steaming / 24.0 if steaming > 0 else 0.0

    # --- Weather exclusions (pre-computed, summed from tags) ---
    excl_hours = sum(r.get("excl_hours", 0) for r in seg_rows)
    excl_mgo = sum(r.get("excl_mgo", 0) for r in seg_rows)
    excl_lng = sum(r.get("excl_lng", 0) for r in seg_rows)
    excl_vlsfo = sum(r.get("excl_vlsfo", 0) for r in seg_rows)
    excl_distance = sum(r.get("excl_distance", 0) for r in seg_rows)
    total_bf5 = sum(r.get("bf5_hours", 0) for r in seg_rows)

    weather_excl_rows = [
        {
            "df_idx": r.get("df_idx"),
            "datetime": r.get("datetime"),
            "bf5_hours": r.get("bf5_hours", 0),
            "steaming_hrs": r.get("steaming_hrs", 0),
            "excl_hours": r.get("excl_hours", 0),
            "excl_mgo": r.get("excl_mgo", 0),
            "excl_lng": r.get("excl_lng", 0),
            "excl_vlsfo": r.get("excl_vlsfo", 0),
        }
        for r in seg_rows if r.get("weather_excluded")
    ]

    # --- Speed calculations (Standard Template formulas) ---
    net_duration_days = duration_days - (excl_hours / 24.0)
    net_distance = distance - excl_distance

    # Actual Avg Speed = NetDistance / (NetDuration_days × 24)
    if net_duration_days > 0:
        actual_speed = net_distance / (net_duration_days * 24.0)
    else:
        actual_speed = 0.0

    # Reference Speed = MIN(Actual, Instructed)
    instructed = seg_info.get("instructed_speed", 0.0)
    reference_speed = (
        min(actual_speed, instructed) if instructed > 0
        else actual_speed
    )

    # --- Net fuel (actual - excluded) ---
    net_mgo = mgo_consumed - excl_mgo
    net_lng = lng_consumed - excl_lng
    net_vlsfo = vlsfo_consumed - excl_vlsfo

    return {
        # Basic segment info
        "start_datetime":       start_dt,
        "end_datetime":         end_dt,
        "duration_days":        round(duration_days, 4),
        "distance":             round(distance, 2),
        "instructed_speed":     instructed,
        "fuel_mode":            seg_info.get("fuel_mode", ""),

        # Exclusions per segment
        "weather_bf5_hours":    round(total_bf5, 2),
        "weather_excl_hours":   round(excl_hours, 2),
        "other_excl_hours":     0.0,        # placeholder for future use
        "total_excl_hours":     round(excl_hours, 2),

        # Speed exclusions
        "weather_excl_distance": round(excl_distance, 2),
        "regulatory_excl_hours": 0.0,
        "regulatory_excl_distance": 0.0,
        "total_speed_excl_hours": round(excl_hours, 2),
        "total_speed_excl_distance": round(excl_distance, 2),

        # Calculated speed
        "net_duration_days":    round(net_duration_days, 4),
        "net_distance":         round(net_distance, 2),
        "actual_avg_speed":     round(actual_speed, 3),
        "reference_speed":      round(reference_speed, 3),

        # Actual fuel consumed
        "lng_consumed":         round(lng_consumed, 2),
        "mgo_consumed":         round(mgo_consumed, 2),
        "vlsfo_consumed":       round(vlsfo_consumed, 2),
        "vlsfo_g1_consumed":    round(vlsfo_g1_consumed, 2),
        "vlsfo_g2_consumed":    round(vlsfo_g2_consumed, 2),
        "mgo_pilot":            round(mgo_pilot, 2),
        "mgo_boiler":           round(mgo_boiler, 2),
        "mgo_propulsion":       round(mgo_propulsion, 2),
        "vlsfo_pilot":          round(vlsfo_pilot, 2),
        "vlsfo_boiler":         round(vlsfo_boiler, 2),
        "vlsfo_propulsion":     round(vlsfo_propulsion, 2),

        # Fuel consumption exclusions
        "excl_lng_weather":     round(excl_lng, 2),
        "excl_mgo_weather":     round(excl_mgo, 2),
        "excl_vlsfo_weather":   round(excl_vlsfo, 2),
        "excl_lng_other":       0.0,
        "excl_mgo_other":       0.0,
        "excl_vlsfo_other":     0.0,
        "excl_lng_total":       round(excl_lng, 2),
        "excl_mgo_total":       round(excl_mgo, 2),
        "excl_vlsfo_total":     round(excl_vlsfo, 2),

        # Net fuel quantities
        "net_lng":              round(net_lng, 2),
        "net_mgo":              round(net_mgo, 2),
        "net_vlsfo":            round(net_vlsfo, 2),

        # GCU compliance
        "gcu_total":            round(gcu_total, 2),
        "gcu_used":             gcu_total > 0,
        "gcu_dates":            gcu_dates,

        # Reliq/Subcooler
        "reliq_hours":          round(reliq_hours, 2),
        "reliq_avg_load":       round(reliq_avg_load, 2),

        # Remarks
        "remarks":              remarks,

        # Detail for weather exclusion rows
        "weather_excluded_rows": weather_excl_rows,
    }


# =============================================================================
# 7. Voyage Totals  (sum across segments)
# =============================================================================

def compute_voyage_totals(segment_data: list[dict]) -> dict:
    """
    Sum across all segments to produce the Total column for the template.
    """
    if not segment_data:
        return {}

    # Fields that are simply summed across segments
    sum_fields = [
        "distance", "duration_days",
        "weather_excl_hours", "other_excl_hours", "total_excl_hours",
        "weather_excl_distance",
        "regulatory_excl_hours", "regulatory_excl_distance",
        "total_speed_excl_hours", "total_speed_excl_distance",
        "net_duration_days", "net_distance",
        "lng_consumed", "mgo_consumed", "vlsfo_consumed",
        "vlsfo_g1_consumed", "vlsfo_g2_consumed",
        "mgo_pilot", "mgo_boiler", "mgo_propulsion",
        "vlsfo_pilot", "vlsfo_boiler", "vlsfo_propulsion",
        "excl_lng_weather", "excl_mgo_weather", "excl_vlsfo_weather",
        "excl_lng_other", "excl_mgo_other", "excl_vlsfo_other",
        "excl_lng_total", "excl_mgo_total", "excl_vlsfo_total",
        "net_lng", "net_mgo", "net_vlsfo",
        "gcu_total", "reliq_hours",
        "weather_bf5_hours",
    ]

    totals = {}
    for field in sum_fields:
        totals[field] = round(
            sum(s.get(field, 0) for s in segment_data), 2
        )

    # --- Derived fields ---

    # Actual Avg Speed = Total Net Distance / (Total Net Duration × 24)
    net_dur = totals.get("net_duration_days", 0)
    net_dist = totals.get("net_distance", 0)
    if net_dur > 0:
        totals["actual_avg_speed"] = round(net_dist / (net_dur * 24), 3)
    else:
        totals["actual_avg_speed"] = 0.0

    # Instructed speed: N/A if segments differ
    speeds = [s.get("instructed_speed", 0) for s in segment_data]
    if len(set(speeds)) == 1:
        totals["instructed_speed"] = speeds[0]
    else:
        totals["instructed_speed"] = None  # varies across segments

    # Reference speed for total = MIN(actual_total_speed, <lowest instructed>)
    # Per standard template: use the total actual speed as reference
    totals["reference_speed"] = totals["actual_avg_speed"]

    # Date range
    totals["start_datetime"] = segment_data[0].get("start_datetime")
    totals["end_datetime"] = segment_data[-1].get("end_datetime")

    # Fuel mode
    modes = [s.get("fuel_mode", "") for s in segment_data]
    totals["fuel_mode"] = modes[0] if len(set(modes)) == 1 else "Mixed"

    # GCU
    totals["gcu_used"] = any(s.get("gcu_used", False) for s in segment_data)
    totals["gcu_dates"] = []
    for s in segment_data:
        totals["gcu_dates"].extend(s.get("gcu_dates", []))

    # Reliq avg load: mean of non-zero segment values
    loads = [
        s.get("reliq_avg_load", 0)
        for s in segment_data
        if s.get("reliq_avg_load", 0) > 0
    ]
    totals["reliq_avg_load"] = (
        round(float(np.mean(loads)), 2) if loads else 0.0
    )

    # Remarks: concatenate all
    totals["remarks"] = []
    for s in segment_data:
        totals["remarks"].extend(s.get("remarks", []))

    return totals


# =============================================================================
# 8. Main Orchestrator
# =============================================================================

def compute_all_segments(
    voyage_data: dict,
    vessel_config: dict | None = None,
) -> dict:
    """
    Main orchestrator: detect segments, pro-rate boundaries, compute data.

    Parameters
    ----------
    voyage_data   : dict from data_extractor.extract_voyage_data()
    vessel_config : vessel config dict (from load_vessel_config)

    Returns
    -------
    dict with:
      - segments:     list[dict] — per-segment computed data
      - totals:       dict — voyage totals (sum across segments)
      - segment_info: list[dict] — segment metadata from detect_segments
    """
    if vessel_config is None:
        vessel_config = load_vessel_config()

    daily_rows = voyage_data["daily_rows"]
    dep_dt = voyage_data["dep_datetime"]
    arr_dt = voyage_data["arr_datetime"]

    # Step 1: Pre-tag weather exclusions on ORIGINAL rows (Rule 2)
    #         Must be done before pro-rating so threshold checks use full-day BF5
    tag_weather_exclusions(daily_rows)

    # Step 2: Detect segments (Rule 3)
    segment_info = detect_segments(daily_rows, dep_dt, arr_dt)

    # Step 3: Build segment rows with boundary pro-rating (Rule 3)
    all_seg_rows = build_segment_rows(segment_info, daily_rows)

    # Step 4: Compute data per segment
    segment_data = []
    for i, (info, rows) in enumerate(zip(segment_info, all_seg_rows)):
        data = compute_segment_data(info, rows, vessel_config)
        segment_data.append(data)
        logger.info(
            "Segment %d: dist=%.1f nm, dur=%.2f days, speed=%.2f kts, "
            "MGO=%.2f MT, LNG=%.2f m³, VLSFO=%.2f MT",
            i + 1, data["distance"], data["duration_days"],
            data["actual_avg_speed"],
            data["mgo_consumed"], data["lng_consumed"], data["vlsfo_consumed"],
        )

    # Step 5: Compute voyage totals
    totals = compute_voyage_totals(segment_data)

    # Step 6: Detect speed anomalies (Rule A — informational)
    speed_anomalies = detect_speed_anomalies(daily_rows)
    if speed_anomalies:
        logger.info(
            "Speed anomalies detected: %d flagged row(s)", len(speed_anomalies)
        )

    return {
        "segments":        segment_data,
        "totals":          totals,
        "segment_info":    segment_info,
        "speed_anomalies": speed_anomalies,
    }
