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
from config import load_vessel_config, COL
from data_extractor import (
    load_raw_excel, detect_voyages, extract_voyage_data, extract_auxiliary,
    merge_bunkering_stops, tag_intermediate_stops,
)
from calculator import compute_all_segments
from template_filler import fill_template
from ai_analyst import review_voyage

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

# -- Sidebar settings --
with st.sidebar:
    st.header("Settings")
    vessel_name = st.text_input("Vessel Name", value="", placeholder="e.g. Id'Asah")
    sheet_name = st.text_input("Excel Sheet Name", value="Sheet")

# -- File upload --
uploaded_file = st.file_uploader(
    "Upload Raw Noon-Report Excel",
    type=["xlsx", "xls"],
    help="The 340-column noon-report Excel file with DEPARTURE/ARRIVAL markers.",
)

if uploaded_file is not None:
    st.success(f"Uploaded: **{uploaded_file.name}** ({uploaded_file.size / 1024:.0f} KB)")

    if st.button("Generate Report", type="primary", use_container_width=True):

        if not vessel_name.strip():
            st.error("Please enter a Vessel Name in the sidebar.")
            st.stop()

        with st.spinner("Processing..."):
            try:
                # -- Load vessel config --
                vessel_config = load_vessel_config(vessel_name)

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

                    # Metadata for template
                    discharge_port = ""
                    dp = df.iloc[v["dep_row"], COL["next_port"]]
                    if dp and str(dp) != "nan":
                        discharge_port = str(dp)

                    metadata = {
                        "voyage_no":          v["voyage_no"],
                        "voyage_type":        v["voyage_type"],
                        "fuel_mode":          v["fuel_mode"],
                        "load_port":          "",
                        "discharge_port":     discharge_port,
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

                # -- 4. AI Analyst review --
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

                # -- Clean up temp files --
                Path(tmp_input_path).unlink(missing_ok=True)
                Path(tmp_output_path).unlink(missing_ok=True)

                # -- Display results --
                st.divider()
                st.subheader("Voyage Summary")

                for vr in voyage_results:
                    m = vr["metadata"]
                    t = vr["computed"]["totals"]
                    n_seg = len(vr["computed"]["segments"])

                    with st.expander(
                        f"Voyage {m['voyage_no']} \u2014 {m['voyage_type']}",
                        expanded=True,
                    ):
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Distance", f"{t['distance']:.1f} nm")
                        col2.metric("Duration", f"{t['duration_days']:.2f} days")
                        col3.metric("Avg Speed", f"{t['actual_avg_speed']:.2f} kts")

                        col4, col5, col6 = st.columns(3)
                        col4.metric("LNG", f"{t['lng_consumed']:.1f} m\u00b3")
                        col5.metric("MGO", f"{t['mgo_consumed']:.2f} MT")
                        col6.metric("VLSFO", f"{t['vlsfo_consumed']:.2f} MT")

                        if t.get("excl_mgo_weather", 0) > 0:
                            st.caption(
                                f"Weather exclusions: "
                                f"MGO {t['excl_mgo_weather']:.2f} MT, "
                                f"LNG {t['excl_lng_weather']:.1f} m\u00b3, "
                                f"VLSFO {t['excl_vlsfo_weather']:.2f} MT"
                            )

                        st.caption(f"{n_seg} segment(s) detected")

                        # Speed anomalies (Rule A)
                        anomalies = vr["computed"].get("speed_anomalies", [])
                        if anomalies:
                            st.warning(
                                f"{len(anomalies)} speed anomaly(ies) detected"
                            )
                            for a in anomalies:
                                st.caption(
                                    f"  {str(a['datetime'])[:16]}: "
                                    f"speed={a['avg_speed']:.1f} kts "
                                    f"(weighted avg: {a['weighted_avg']:.1f} kts)"
                                )

                        # Intermediate stops (Rule B)
                        stops = m.get("intermediate_stops", [])
                        if stops:
                            st.info(
                                f"{len(stops)} mid-voyage bunkering stop(s)"
                            )
                            for s in stops:
                                st.caption(
                                    f"  {s['port_name']}: "
                                    f"{s['arr_datetime']} \u2192 {s['dep_datetime']} "
                                    f"({s['duration_hours']:.1f} hrs, "
                                    f"LNG: {s['lng_consumed']:.1f} m\u00b3, "
                                    f"MGO: {s['mgo_consumed']:.2f} MT)"
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

                    with st.expander(
                        f"AI Review ({len(errors)} errors, "
                        f"{len(warnings)} warnings, {len(infos)} info)",
                        expanded=True,
                    ):
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

                # -- Download button --
                st.divider()
                output_filename = (
                    uploaded_file.name.replace(".xlsx", "")
                    .replace(".xls", "")
                    + "_report.xlsx"
                )
                st.download_button(
                    label="Download Report",
                    data=output_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

                st.success("Report generated successfully!")

            except Exception as e:
                st.error(f"Error processing file: {e}")
                logger.exception("Processing error: %s", e)

else:
    st.info("Drag and drop your raw Excel file above, or click Browse.")
