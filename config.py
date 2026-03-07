# =============================================================================
# config.py  –  Column mappings, vessel parameters, and app settings
# =============================================================================
# Columns are resolved by HEADER NAME at runtime (auto-detection).
# If columns are added/removed/reordered, the app still works as long as
# the header names remain the same.
# Fallback: hardcoded indices are used only if auto-detection hasn't run.
# =============================================================================

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Header Name Map  —  internal key → exact column header in raw Excel
# ---------------------------------------------------------------------------
# These header names are the AUTHORITATIVE source.  The 0-based indices in
# COL (below) are just fallback defaults; resolve_columns() overwrites them.

COL_HEADERS = {
    # ── Voyage identification (Rule 1) ──────────────────────────────────────
    "voyage_no":        "Voyage No",
    "datetime":         "Date/Time (UTC)",
    "portcall_type":    "Type of Portcall",
    "next_port":        "Next Port",
    "fuel_mode":        "Fuel Mode",
    "eta_next":         "ETA Next Port (LT)",
    "ordered_speed":    "Ordered/Required Speed",
    "distance":         "Distance Made Good (NM)",
    "cargo_density":    "Cargo Density (kg/m3)",
    "lcv":              "Actual LCV (MJ/Kg)",
    "report_type":      "Report Type",

    # ── Steaming & speed ────────────────────────────────────────────────────
    "steaming_hrs":     "Steaming Time (Hrs.)",
    "avg_speed":        "Avg. Speed (Kts)",

    # ── Voyage order revision (Rule 3) ──────────────────────────────────────
    "voyage_order_rev": "Revision In Voyage Orders Since Last Report?",
    "rev_start_time":   "Revised Orders Start Time (LT)",
    "rev_gmt_offset":   "Start GMT Offset",
    "rev_sat":          "Required Arrival Time (LT)",
    "rev_speed":        "Ordered Speed (kts)",

    # ── Fuel ROB columns — PRIMARY source (Rule 1: ROB difference) ─────────
    "mgo_rob":          "LSMGO ROB (MT)",
    "mgo_boiler":       "LSMGO Standard Boiler Cons. (MT)",
    "mgo_pilot":        "LSMGO Pilot Flame (MT)",
    "lng_rob":          "LNG ROB (m3)",
    "gcu_lng":          "LNG Combusted in GCU (m3)",

    # ── VLSFO Grade 1 (Rule 1 & 4) ─────────────────────────────────────────
    "vlsfo_g1_rob":     "VLSFO ROB (MT)",
    "vlsfo_g1_boiler":  "VLSFO Standard Boiler Cons. (MT)",
    "vlsfo_g1_pilot":   "VLSFO Pilot Flame (MT)",

    # ── VLSFO Grade 2 (Rule 1 & 4) ─────────────────────────────────────────
    "vlsfo_g2_rob":     "VLSFO_GTE_80 ROB (MT)",
    "vlsfo_g2_boiler":  "VLSFO_GTE_80 Standard Boiler Cons. (MT)",
    "vlsfo_g2_pilot":   "VLSFO_GTE_80 Pilot Flame (MT)",

    # ── Reliquefaction / Subcooler — two systems (Rule 1) ───────────────────
    "reliq_hours":      "Hours in use",
    "reliq_load":       "Load of Reliq %",
    "subcooler_hours":  "Hours in use.1",
    "subcooler_load":   "Load of Subcooler %",

    # ── Weather (Rule 2) ───────────────────────────────────────────────────
    "wind_force":       "Wind Force(Bft.) (T)",
    "bf5_hours":        "Wind Force Above 5bft. (hrs.)",

    # ── Remarks ─────────────────────────────────────────────────────────────
    "remarks":          "Remarks",
}


# ---------------------------------------------------------------------------
# Fallback Column Index Map  (0-based, for the standard 340-column layout)
# ---------------------------------------------------------------------------
# Used when resolve_columns() hasn't been called (e.g. in unit tests with
# synthetic DataFrames that use integer column indices).

COL = {
    "voyage_no":        1,
    "datetime":         2,
    "portcall_type":    6,
    "next_port":        9,
    "fuel_mode":       10,
    "eta_next":        11,
    "ordered_speed":   13,
    "distance":        15,
    "cargo_density":   26,
    "lcv":             27,
    "report_type":     28,
    "steaming_hrs":    44,
    "avg_speed":       45,
    "voyage_order_rev": 64,
    "rev_start_time":   65,
    "rev_gmt_offset":   66,
    "rev_sat":          67,
    "rev_speed":        70,
    "mgo_rob":         79,
    "mgo_boiler":      86,
    "mgo_pilot":       89,
    "lng_rob":         98,
    "gcu_lng":        114,
    "vlsfo_g1_rob":   136,
    "vlsfo_g1_boiler":143,
    "vlsfo_g1_pilot": 146,
    "vlsfo_g2_rob":   193,
    "vlsfo_g2_boiler":200,
    "vlsfo_g2_pilot": 203,
    "reliq_hours":    245,
    "reliq_load":     247,
    "subcooler_hours":250,
    "subcooler_load": 251,
    "wind_force":     259,
    "bf5_hours":      270,
    "remarks":        339,
}


# ---------------------------------------------------------------------------
# Auto-Detection: resolve column indices from header names
# ---------------------------------------------------------------------------

def resolve_columns(df) -> dict[str, int]:
    """
    Scan the DataFrame's header row and resolve column indices by name.

    Updates the global COL dict in-place so all downstream code (data_extractor,
    calculator, etc.) automatically uses the correct indices.

    Returns the resolved COL dict.

    Raises ValueError if any required column is not found.
    """
    global COL

    headers = list(df.columns)
    resolved = {}
    missing = []

    for key, header_name in COL_HEADERS.items():
        try:
            idx = headers.index(header_name)
            resolved[key] = idx
        except ValueError:
            missing.append(f"  {key}: '{header_name}'")

    if missing:
        msg = "Column auto-detection failed. Missing columns:\n" + "\n".join(missing)
        raise ValueError(msg)

    # Update global COL so all modules use the resolved indices
    COL.update(resolved)

    logger.info(
        "Column auto-detection OK: %d columns resolved from header names.",
        len(resolved),
    )

    return COL

# ---------------------------------------------------------------------------
# Weather exclusion threshold  (Rule 2)
# If BF5 hours > this value on any day → excluded period
# ---------------------------------------------------------------------------
WEATHER_BF5_THRESHOLD = 6    # hours (changed from 12 to 6 per user rule)

# ---------------------------------------------------------------------------
# GCU compliance  (Clause 23.5b)
# ---------------------------------------------------------------------------
MIN_SPEED_NO_GCU = 12.0      # knots – GCU should not be used above this speed

# ---------------------------------------------------------------------------
# LCV assumptions  (template defaults, may be overridden by raw data)
# ---------------------------------------------------------------------------
LCV_GAS_KJ_KG   = 50000
LCV_MGO_KJ_KG   = 42700
LCV_VLSFO_KJ_KG = 40600

# ---------------------------------------------------------------------------
# LLM settings  (retained for tcp_parser.py)
# ---------------------------------------------------------------------------
LLM_MODEL       = "gpt-4o"
LLM_TEMPERATURE = 0
LOCAL_LLM_URL   = "http://localhost:11434/api/generate"

# ---------------------------------------------------------------------------
# AI Analyst settings  (Claude-powered review)
# ---------------------------------------------------------------------------
AI_ANALYST_MODEL      = "claude-sonnet-4-6"
AI_ANALYST_MAX_TOKENS = 4096
AI_ANALYST_TEMPERATURE = 0

# ---------------------------------------------------------------------------
# Warranty validation ranges  (retained for tcp_parser.py)
# ---------------------------------------------------------------------------
WARRANTY_RANGES = {
    "speed":             (8,  25),
    "boil_off_rate_pct": (0.01, 0.3),
    "nbog_m3":           (0,   300),
    "lsmgo_me_cons_mt":  (0,   120),
    "lng_me_cons_m3":    (0,  3000),
}


# =============================================================================
# Vessel Configuration
# =============================================================================

def load_vessel_config(vessel_name: str = "Id'Asah") -> dict:
    """
    Load vessel-specific parameters by name.

    Returns dict with keys:
      name, imo, flag, year_built, gross_capacity_cbm, foe_factor,
      bor_laden_pct, bor_ballast_pct, service_speed_laden, minimum_speed,
      boiler_cons_laden_mt, boiler_cons_ballast_mt,
      speed_consumption_table, ageing_factors, port_cons
    """
    vessels = _VESSEL_REGISTRY()
    if vessel_name not in vessels:
        available = ", ".join(vessels.keys())
        raise ValueError(
            f"Unknown vessel '{vessel_name}'. Available: {available}"
        )
    return vessels[vessel_name]


def _VESSEL_REGISTRY() -> dict:
    """Registry of known vessel configurations."""
    return {
        "Id'Asah": {
            "name":              "Id'Asah",
            "imo":               "9977220",
            "flag":              "Marseille/France",
            "year_built":        2024,
            "gross_capacity_cbm": 174_221.98,
            "foe_factor":        0.484,
            "bor_laden_pct":     0.085,
            "bor_ballast_pct":   0.055,
            "service_speed_laden":  19.5,
            "service_speed_ballast": None,
            "minimum_speed":     12.0,
            "boiler_cons_laden_mt":   1.3,
            "boiler_cons_ballast_mt": 1.3,
            "speed_consumption_table": {
                # speed: (laden_gas, laden_pilot, ballast_gas, ballast_pilot) MT/day
                19.5: (68.1, 1.2, 67.2, 1.1),
                19.0: (63.4, 1.2, 62.6, 1.1),
                18.5: (59.0, 1.2, 58.2, 1.1),
                18.0: (54.8, 1.2, 54.1, 1.1),
                17.5: (50.9, 1.2, 50.2, 1.1),
                17.0: (47.2, 1.2, 46.6, 1.1),
                16.5: (43.7, 1.2, 43.2, 1.1),
                16.0: (40.5, 1.2, 40.0, 1.1),
                15.5: (37.5, 1.2, 37.0, 1.1),
                15.0: (34.7, 1.2, 34.3, 1.1),
                14.5: (32.1, 1.2, 31.7, 1.1),
                14.0: (29.7, 1.2, 29.3, 1.1),
                13.5: (27.5, 1.2, 27.1, 1.1),
                13.0: (25.4, 1.2, 25.1, 1.1),
                12.5: (23.5, 1.2, 23.2, 1.1),
                12.0: (21.8, 1.2, 21.5, 1.1),
            },
            "ageing_factors": {1: 1.0, 2: 1.0, 3: 1.0},
            "port_cons": {
                "loading":     {"lng_mt": 15.0, "lsmgo_mt": 0.7, "vlsfo_mt": 5.3},
                "discharging": {"lng_mt": 26.0, "lsmgo_mt": 0.7, "vlsfo_mt": 5.2},
            },
        },
        # Add more vessels here as needed
    }
