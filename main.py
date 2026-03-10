"""
main.py
=======
Orchestrates the vessel performance report generation pipeline:

  1. Load raw noon-report Excel (index-based, no filtering)
  2. Auto-detect voyages (DEPARTURE → ARRIVAL boundaries)
  3. For each voyage:
     a. Extract voyage data (Rule 1: ROB-diff fuel, distance, steaming hrs)
     b. Detect segments (Rule 3: J/K/L changes, BM="yes" voyage orders)
     c. Pro-rate boundary rows (Rule 3: speed-weighted ratio)
     d. Per segment: compute fuel, weather exclusions (Rule 2), speed,
        boiler/pilot (Rule 4)
     e. Compute voyage totals
  4. Generate output Excel (one sheet per voyage leg, standard template)

Usage
-----
python main.py --input raw_data.xlsx --output report.xlsx
python main.py --input raw_data.xlsx --output report.xlsx --vessel "Id'Asah" --verbose
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env", override=True)
except ImportError:
    pass


# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

def _setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    fmt = "%(asctime)s  %(levelname)-8s  %(name)s  -  %(message)s"
    logging.basicConfig(
        level=level, format=fmt,
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler("vessel_performance.log", mode="a"),
        ],
    )


logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# CLI argument parser
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Vessel Performance Report Generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument(
        "--input", "-i", metavar="FILE", required=True,
        help="Path to raw noon-report Excel file",
    )
    p.add_argument(
        "--output", "-o", metavar="FILE", default="voyage_report.xlsx",
        help="Output report file path (default: voyage_report.xlsx)",
    )
    p.add_argument(
        "--vessel", metavar="NAME", default="Id'Asah",
        help="Vessel name for config lookup (default: Id'Asah)",
    )
    p.add_argument(
        "--sheet", metavar="NAME", default="Sheet",
        help="Excel sheet name to read (default: Sheet)",
    )
    p.add_argument(
        "--verbose", "-v", action="store_true",
        help="Enable DEBUG-level logging",
    )
    p.add_argument(
        "--no-ai", action="store_true",
        help="Skip AI Analyst review (faster)",
    )
    return p


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def run(args: argparse.Namespace) -> None:
    """Execute the full voyage report generation pipeline."""

    from config import load_vessel_config
    from data_extractor import (
        load_raw_excel, detect_voyages, extract_voyage_data, extract_auxiliary,
        merge_bunkering_stops, tag_intermediate_stops,
    )
    from calculator import compute_all_segments
    from template_filler import fill_template
    from ai_analyst import review_voyage

    # ── Load vessel configuration ────────────────────────────────────────
    vessel_config = load_vessel_config(args.vessel)

    # ── 0. Validate input ────────────────────────────────────────────────
    input_path = Path(args.input)
    if not input_path.exists():
        logger.error("Input file not found: %s", input_path)
        sys.exit(1)

    output_path = Path(args.output)

    print("\n" + "=" * 60)
    print("  VESSEL PERFORMANCE REPORT GENERATOR")
    print("=" * 60)
    print(f"  Vessel  : {vessel_config['name']}")
    print(f"  Input   : {input_path.name}")
    print(f"  Output  : {output_path}")
    print("=" * 60 + "\n")

    # ── 1. Load raw Excel ────────────────────────────────────────────────
    print("[1/5]  Loading raw noon-report data ...")
    df = load_raw_excel(input_path, sheet_name=args.sheet)
    print(f"  OK  {len(df)} rows × {df.shape[1]} columns loaded.\n")

    # ── 2. Auto-detect voyages ───────────────────────────────────────────
    print("[2/5]  Detecting voyages (DEPARTURE → ARRIVAL) ...")
    voyages = detect_voyages(df)

    if not voyages:
        logger.error(
            "No complete voyages found (no DEPARTURE/ARRIVAL pairs in Col AC). "
            "Check the raw data."
        )
        sys.exit(1)

    # Rule B: merge intermediate bunkering stops
    voyages = merge_bunkering_stops(df, voyages)

    for v in voyages:
        print(
            f"  Voyage {v['voyage_no']}: {v['voyage_type']}  "
            f"({v['dep_datetime'][:10]} → {v['arr_datetime'][:10]})"
        )
        for stop in v.get("intermediate_stops", []):
            print(
                f"    Bunkering stop: {stop['port_name']} "
                f"({stop['arr_datetime']} → {stop['dep_datetime']}, "
                f"{stop['duration_hours']:.1f} hrs)"
            )
    print()

    # ── 3. Process each voyage ───────────────────────────────────────────
    print("[3/5]  Processing voyages ...")
    voyage_results = []

    for v in voyages:
        print(f"\n  --- Voyage {v['voyage_no']} ({v['voyage_type']}) ---")

        # Extract voyage data (Rule 1)
        vd = extract_voyage_data(df, v["dep_row"], v["arr_row"])
        aux = extract_auxiliary(df, v["dep_row"], v["arr_row"])

        # Tag intermediate stop rows (Rule B)
        stops = v.get("intermediate_stops", [])
        if stops:
            tag_intermediate_stops(vd["daily_rows"], stops)

        print(f"    Distance: {vd['total_distance']:.1f} nm")
        print(f"    MGO: {vd['mgo_consumed']:.2f} MT, "
              f"LNG: {vd['lng_consumed']:.2f} m³, "
              f"VLSFO: {vd['vlsfo_consumed']:.2f} MT")

        # Compute segments + weather exclusions + speed (Rules 2-4)
        computed = compute_all_segments(vd, vessel_config)
        n_seg = len(computed["segments"])
        totals = computed["totals"]

        print(f"    Segments: {n_seg}")
        for i, seg in enumerate(computed["segments"]):
            print(
                f"      Seg {i+1}: {seg['distance']:.0f} nm, "
                f"{seg['instructed_speed']:.1f} kts ordered, "
                f"{seg['actual_avg_speed']:.2f} kts actual"
            )

        print(f"    Weather exclusions: "
              f"MGO {totals['excl_mgo_weather']:.2f} MT, "
              f"VLSFO {totals['excl_vlsfo_weather']:.2f} MT, "
              f"LNG {totals['excl_lng_weather']:.2f} m³")

        # Speed anomalies (Rule A)
        anomalies = computed.get("speed_anomalies", [])
        if anomalies:
            print(f"    Speed anomalies: {len(anomalies)} flagged row(s)")
            for a in anomalies:
                print(f"      {str(a['datetime'])[:16]}: "
                      f"speed={a['avg_speed']:.1f} kts "
                      f"(weighted avg: {a['weighted_avg']:.1f} kts)")

        # Build metadata for template
        from config import COL
        discharge_port = ""
        if v["dep_row"] < len(df):
            dp = df.iloc[v["dep_row"], COL["next_port"]]
            if dp and str(dp) != "nan":
                discharge_port = str(dp)

        metadata = {
            "voyage_no":           v["voyage_no"],
            "voyage_type":         v["voyage_type"],
            "fuel_mode":           v["fuel_mode"],
            "load_port":           "",
            "discharge_port":      discharge_port,
            "distance":            totals["distance"],
            "duration_days":       totals["duration_days"],
            "dep_datetime":        v["dep_datetime"],
            "arr_datetime":        v["arr_datetime"],
            "dep_row":             v["dep_row"],
            "arr_row":             v["arr_row"],
            "cargo_density":       aux.get("cargo_density"),
            "lcv":                 aux.get("lcv"),
            "charter_year":        1,
            "intermediate_stops":  stops,
        }

        voyage_results.append({
            "computed":  computed,
            "metadata":  metadata,
            "auxiliary": aux,
        })

    # ── 4. AI Analyst review ────────────────────────────────────────────
    if getattr(args, "no_ai", False):
        print("\n[4/5]  AI Analyst review skipped (--no-ai)")
    else:
        print("\n[4/5]  AI Analyst reviewing report ...")
        for vr in voyage_results:
            alerts = review_voyage(df, vr, vessel_config)
            vr["ai_alerts"] = alerts
            if alerts:
                errors = sum(1 for a in alerts if a.get("severity") == "error")
                warnings = sum(1 for a in alerts if a.get("severity") == "warning")
                infos = sum(1 for a in alerts if a.get("severity") == "info")
                print(f"    Voyage {vr['metadata']['voyage_no']}: "
                      f"{errors} error(s), {warnings} warning(s), {infos} info(s)")
                for a in alerts:
                    icon = {"error": "!!", "warning": "!", "info": "i"}.get(
                        a.get("severity", ""), "?"
                    )
                    print(f"      [{icon}] {a.get('message', '')}")
            else:
                print(f"    Voyage {vr['metadata']['voyage_no']}: "
                      f"No AI review (API key not set or review skipped)")

    # ── 5. Generate output Excel ─────────────────────────────────────────
    print(f"\n[5/5]  Generating report: {output_path} ...")
    fill_template(output_path, voyage_results, vessel_config)
    print(f"\n  OK  Report saved to: {output_path.resolve()}")

    # Print summary
    print("\n" + "=" * 60)
    print("  SUMMARY")
    print("=" * 60)
    for vr in voyage_results:
        m = vr["metadata"]
        t = vr["computed"]["totals"]
        print(f"\n  Voyage {m['voyage_no']} ({m['voyage_type']})")
        print(f"    Distance:      {t['distance']:.1f} nm")
        print(f"    Duration:      {t['duration_days']:.2f} days")
        print(f"    Avg Speed:     {t['actual_avg_speed']:.2f} kts")
        print(f"    LNG consumed:  {t['lng_consumed']:.2f} m³  "
              f"(net: {t['net_lng']:.2f} m³)")
        print(f"    MGO consumed:  {t['mgo_consumed']:.2f} MT  "
              f"(net: {t['net_mgo']:.2f} MT)")
        print(f"    VLSFO consumed:{t['vlsfo_consumed']:.2f} MT  "
              f"(net: {t['net_vlsfo']:.2f} MT)")
    print("\n" + "=" * 60 + "\n")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()
    _setup_logging(args.verbose)

    try:
        run(args)
    except KeyboardInterrupt:
        print("\nAborted by user.")
        sys.exit(0)
    except Exception as exc:
        logger.exception("Unhandled error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
