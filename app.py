"""
TCP Agent — Vessel Performance Report Generator (Web UI)
========================================================
Streamlit app that wraps the CLI pipeline into a browser-based interface.

Run locally:  streamlit run app.py
Deploy:       Push to GitHub → connect to Streamlit Community Cloud
"""

import logging
import os
import tempfile
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env", override=True)
except ImportError:
    pass

import streamlit as st

# Streamlit Cloud secrets → environment variables
# (Streamlit Cloud stores secrets in st.secrets, but our code reads os.environ)
try:
    for key in ("ANTHROPIC_API_KEY", "OPENAI_API_KEY"):
        if key in st.secrets and key not in os.environ:
            os.environ[key] = st.secrets[key]
except Exception:
    pass

# -- Pipeline imports (same directory) --
from config import load_vessel_config, parse_fuel_table_csv, COL
from data_extractor import (
    load_raw_excel, detect_voyages, extract_voyage_data, extract_auxiliary,
    merge_bunkering_stops, tag_intermediate_stops,
)
from calculator import compute_all_segments
from template_filler import fill_template
from ai_analyst import review_voyage
from highlight_report import generate_highlighted_report

logger = logging.getLogger(__name__)


# =============================================================================
# Page config
# =============================================================================

st.set_page_config(
    page_title="TCP Agent",
    page_icon="\u2693",  # anchor emoji
    layout="centered",
)


# =============================================================================
# UI
# =============================================================================

st.title("\u2693 TCP Agent")
st.markdown("**Vessel Performance Report Generator**")
st.markdown("Upload a raw noon-report Excel file to generate a voyage performance report.")

st.divider()

# -- File upload (before sidebar so we can auto-detect vessel name) --
uploaded_file = st.file_uploader(
    "Upload Raw Noon-Report Excel",
    type=["xlsx", "xls"],
    help="The 340-column noon-report Excel file with DEPARTURE/ARRIVAL markers.",
)

# -- Auto-detect vessel name from uploaded file (Rule 6) --
if uploaded_file is not None and st.session_state.get("_last_file") != uploaded_file.name:
    try:
        import pandas as pd
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as _tmp:
            _tmp.write(uploaded_file.getvalue())
            _scan_path = _tmp.name
        _scan_df = pd.read_excel(_scan_path, sheet_name=0, header=0)
        # Look for "Vessel Name" column
        _vessel_col = None
        for _c in _scan_df.columns:
            if str(_c).strip().lower() == "vessel name":
                _vessel_col = _c
                break
        if _vessel_col is not None:
            _raw_names = _scan_df[_vessel_col].dropna().unique()
            if len(_raw_names) > 0:
                st.session_state["detected_vessel"] = str(_raw_names[0]).strip()
        st.session_state["_last_file"] = uploaded_file.name
        Path(_scan_path).unlink(missing_ok=True)
    except Exception:
        pass

# -- Known vessel registry for fuel table matching --
from config import _VESSEL_REGISTRY
_known_vessels = list(_VESSEL_REGISTRY().keys())

# -- Sidebar settings --
with st.sidebar:
    st.header("Settings")
    _detected = st.session_state.get("detected_vessel", "")
    vessel_name = st.text_input("Vessel Name", value=_detected, placeholder="e.g. Id'Asah")

    # Show fuel table status
    if vessel_name.strip():
        if vessel_name.strip() in _known_vessels:
            st.caption(f"Fuel table: **{vessel_name.strip()}** (from database)")
        else:
            st.caption("Fuel table: not found in database — upload CSV or Id'Asah defaults will be used")

    sheet_name = st.text_input("Excel Sheet Name", value="Sheet")
    enable_ai_review = st.checkbox("Enable AI Review", value=False,
                                   help="Uses Claude AI to review the report. Slower but adds expert analysis.")
    st.divider()
    st.caption("Fuel Table (optional — overrides database)")
    fuel_csv = st.file_uploader(
        "Upload CSV: speed, laden_gas, laden_pilot, ballast_gas, ballast_pilot",
        type=["csv"],
    )

if uploaded_file is not None:
    st.success(f"Uploaded: **{uploaded_file.name}** ({uploaded_file.size / 1024:.0f} KB)")

    if st.button("Generate Report", type="primary", use_container_width=True):

        with st.spinner("Processing..."):
            try:
                # -- Save uploaded file to temp path --
                with tempfile.NamedTemporaryFile(
                    suffix=".xlsx", delete=False
                ) as tmp_in:
                    tmp_in.write(uploaded_file.getvalue())
                    tmp_input_path = tmp_in.name

                # -- 1. Load raw Excel --
                df = load_raw_excel(tmp_input_path, sheet_name=sheet_name)

                # -- 2. Detect voyages + merge bunkering stops --
                voyages = detect_voyages(df)

                if not voyages:
                    st.error(
                        "No complete voyages found. "
                        "Check that the file has DEPARTURE/ARRIVAL markers in Col AC."
                    )
                    st.stop()

                voyages = merge_bunkering_stops(df, voyages)

                # Rule 6: Auto-detect vessel name from data if not entered
                effective_vessel_name = vessel_name.strip()
                if not effective_vessel_name:
                    raw_name = voyages[0].get("vessel_name", "")
                    if raw_name:
                        effective_vessel_name = raw_name
                    else:
                        st.error("Please enter a Vessel Name in the sidebar.")
                        st.stop()

                # -- Load vessel config --
                vessel_config = load_vessel_config(effective_vessel_name)

                # Override fuel table if CSV uploaded
                if fuel_csv is not None:
                    fuel_csv.seek(0)
                    vessel_config["speed_consumption_table"] = parse_fuel_table_csv(fuel_csv)

                # -- 3. Process each voyage --
                voyage_results = []

                for v in voyages:
                    vd = extract_voyage_data(df, v["dep_row"], v["arr_row"])
                    aux = extract_auxiliary(df, v["dep_row"], v["arr_row"])

                    # Tag intermediate stop rows (Rule B)
                    stops = v.get("intermediate_stops", [])
                    if stops:
                        tag_intermediate_stops(vd["daily_rows"], stops)

                    computed = compute_all_segments(vd, vessel_config)
                    totals = computed["totals"]

                    # Rule 6: Use extracted port details
                    metadata = {
                        "voyage_no":          v["voyage_no"],
                        "voyage_type":        v["voyage_type"],
                        "fuel_mode":          v["fuel_mode"],
                        "vessel_name":        v.get("vessel_name", ""),
                        "load_port":          v.get("last_port", ""),
                        "discharge_port":     v.get("next_port", ""),
                        "distance":           totals["distance"],
                        "duration_days":      totals["duration_days"],
                        "dep_datetime":       v["dep_datetime"],
                        "arr_datetime":       v["arr_datetime"],
                        "dep_row":            v["dep_row"],
                        "arr_row":            v["arr_row"],
                        "cargo_density":      aux.get("cargo_density"),
                        "lcv":                aux.get("lcv"),
                        "charter_year":       1,
                        "intermediate_stops": stops,
                    }

                    voyage_results.append({
                        "computed":  computed,
                        "metadata":  metadata,
                        "auxiliary": aux,
                    })

                # -- 4. AI Analyst review (optional) --
                if enable_ai_review:
                    ai_status = st.empty()
                    ai_status.info("AI Analyst is reviewing the report...")

                    for vr in voyage_results:
                        alerts = review_voyage(df, vr, vessel_config)
                        vr["ai_alerts"] = alerts

                    ai_status.empty()

                # -- 5. Generate output Excel --
                with tempfile.NamedTemporaryFile(
                    suffix=".xlsx", delete=False
                ) as tmp_out:
                    tmp_output_path = tmp_out.name

                fill_template(tmp_output_path, voyage_results, vessel_config)

                # -- Read output file for download --
                with open(tmp_output_path, "rb") as f:
                    output_bytes = f.read()

                # -- 6. Generate highlighted raw data report --
                with tempfile.NamedTemporaryFile(
                    suffix=".xlsx", delete=False
                ) as tmp_hl:
                    tmp_highlight_path = tmp_hl.name

                generate_highlighted_report(
                    tmp_input_path, tmp_highlight_path,
                    voyages, sheet_name=sheet_name,
                )

                with open(tmp_highlight_path, "rb") as f:
                    highlight_bytes = f.read()

                # -- Clean up temp files --
                Path(tmp_input_path).unlink(missing_ok=True)
                Path(tmp_output_path).unlink(missing_ok=True)
                Path(tmp_highlight_path).unlink(missing_ok=True)

                # -- Display results --
                st.divider()
                st.subheader("Voyage Summary")

                # Build summary table rows
                import pandas as pd

                summary_rows = []
                for vr in voyage_results:
                    m = vr["metadata"]
                    t = vr["computed"]["totals"]
                    segments = vr["computed"]["segments"]

                    if len(segments) <= 1:
                        # Single segment — one row
                        summary_rows.append({
                            "Voyage": m["voyage_no"],
                            "Type": m["voyage_type"],
                            "From": m.get("load_port", "") or "-",
                            "To": m.get("discharge_port", "") or "-",
                            "Departure": str(m["dep_datetime"])[:16],
                            "Arrival": str(m["arr_datetime"])[:16],
                            "Distance (nm)": f"{t['distance']:.1f}",
                            "Duration (days)": f"{t['duration_days']:.2f}",
                            "Avg Speed (kts)": f"{t['actual_avg_speed']:.2f}",
                            "LNG (m\u00b3)": f"{t['lng_consumed']:.1f}",
                            "MGO (MT)": f"{t['mgo_consumed']:.2f}",
                            "VLSFO (MT)": f"{t['vlsfo_consumed']:.2f}",
                        })
                    else:
                        # Multiple segments — one row per segment + voyage total
                        for i, seg in enumerate(segments, 1):
                            summary_rows.append({
                                "Voyage": f"{m['voyage_no']} Seg {i}",
                                "Type": m["voyage_type"],
                                "From": m.get("load_port", "") or "-" if i == 1 else "-",
                                "To": m.get("discharge_port", "") or "-" if i == len(segments) else "-",
                                "Departure": str(seg["start_datetime"])[:16],
                                "Arrival": str(seg["end_datetime"])[:16],
                                "Distance (nm)": f"{seg['distance']:.1f}",
                                "Duration (days)": f"{seg['duration_days']:.2f}",
                                "Avg Speed (kts)": f"{seg['actual_avg_speed']:.2f}",
                                "LNG (m\u00b3)": f"{seg['lng_consumed']:.1f}",
                                "MGO (MT)": f"{seg['mgo_consumed']:.2f}",
                                "VLSFO (MT)": f"{seg['vlsfo_consumed']:.2f}",
                            })
                        # Voyage total row
                        summary_rows.append({
                            "Voyage": f"{m['voyage_no']} TOTAL",
                            "Type": m["voyage_type"],
                            "From": m.get("load_port", "") or "-",
                            "To": m.get("discharge_port", "") or "-",
                            "Departure": str(m["dep_datetime"])[:16],
                            "Arrival": str(m["arr_datetime"])[:16],
                            "Distance (nm)": f"{t['distance']:.1f}",
                            "Duration (days)": f"{t['duration_days']:.2f}",
                            "Avg Speed (kts)": f"{t['actual_avg_speed']:.2f}",
                            "LNG (m\u00b3)": f"{t['lng_consumed']:.1f}",
                            "MGO (MT)": f"{t['mgo_consumed']:.2f}",
                            "VLSFO (MT)": f"{t['vlsfo_consumed']:.2f}",
                        })

                summary_df = pd.DataFrame(summary_rows)
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

                # Alerts below the table
                for vr in voyage_results:
                    m = vr["metadata"]

                    # Speed anomalies (Rule A)
                    anomalies = vr["computed"].get("speed_anomalies", [])
                    if anomalies:
                        st.warning(
                            f"Voyage {m['voyage_no']}: "
                            f"{len(anomalies)} speed anomaly(ies) detected"
                        )

                    # Intermediate stops (Rule B)
                    stops = m.get("intermediate_stops", [])
                    if stops:
                        st.info(
                            f"Voyage {m['voyage_no']}: "
                            f"{len(stops)} mid-voyage bunkering stop(s)"
                        )

                    # Weather exclusions
                    t = vr["computed"]["totals"]
                    if t.get("excl_mgo_weather", 0) > 0:
                        st.caption(
                            f"Voyage {m['voyage_no']} weather exclusions: "
                            f"MGO {t['excl_mgo_weather']:.2f} MT, "
                            f"LNG {t['excl_lng_weather']:.1f} m\u00b3, "
                            f"VLSFO {t['excl_vlsfo_weather']:.2f} MT"
                        )

                # -- AI Review section --
                all_alerts = []
                for vr in voyage_results:
                    for a in vr.get("ai_alerts", []):
                        a["voyage_no"] = vr["metadata"]["voyage_no"]
                        all_alerts.append(a)

                if all_alerts:
                    st.divider()
                    errors = [a for a in all_alerts if a.get("severity") == "error"]
                    warnings = [a for a in all_alerts if a.get("severity") == "warning"]
                    infos = [a for a in all_alerts if a.get("severity") == "info"]

                    st.markdown(
                        f"**AI Review** — {len(errors)} error(s), "
                        f"{len(warnings)} warning(s), {len(infos)} info"
                    )
                    for a in errors:
                        st.error(
                            f"**[{a.get('category', '')}]** "
                            f"{a.get('message', '')}\n\n"
                            f"_{a.get('details', '')}_"
                        )
                    for a in warnings:
                        st.warning(
                            f"**[{a.get('category', '')}]** "
                            f"{a.get('message', '')}\n\n"
                            f"_{a.get('details', '')}_"
                        )
                    for a in infos:
                        st.info(
                            f"**[{a.get('category', '')}]** "
                            f"{a.get('message', '')}\n\n"
                            f"_{a.get('details', '')}_"
                        )

                # -- Download buttons --
                st.divider()
                base_name = (
                    uploaded_file.name.replace(".xlsx", "")
                    .replace(".xls", "")
                )

                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="Download Voyage Report",
                        data=output_bytes,
                        file_name=f"{base_name}_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                with col2:
                    st.download_button(
                        label="Download Highlighted Raw Data",
                        data=highlight_bytes,
                        file_name=f"{base_name}_highlighted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                st.success("Reports generated successfully!")

            except Exception as e:
                st.error(f"Error processing file: {e}")
                logger.exception("Processing error: %s", e)

else:
    st.info("Drag and drop your raw Excel file above, or click Browse.")
