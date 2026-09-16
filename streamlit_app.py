"""Streamlit front-end for the Accenture 'Closed with Insuff' reporting pipeline.

Runs the same login -> export -> filter -> report pipeline as `app.py`, with
an optional upload for the Check_unique_name -> Component mapping so business
users can update it without touching code.
"""
import traceback

import pandas as pd
import streamlit as st

from app import (
    DEFAULT_COMPONENT_MAP,
    DOWNLOAD_DIR,
    compute_date_range,
    find_default_mapping_file,
    load_component_map,
    run_pipeline,
)

st.set_page_config(page_title="Accenture Closed-with-Insuff Reporting", layout="wide")
st.title("Accenture Closed-with-Insuff Reporting")
st.caption(
    "Logs into the AuthBridge MIS query browser, pulls the daily checks export, "
    "the Advance Tracker and Antecedent Details exports, and builds the "
    "formatted Closed-with-Insufficiency report."
)

if "results" not in st.session_state:
    st.session_state.results = None

with st.sidebar:
    st.header("Component mapping")
    st.caption("Columns required: 'Check_unique_name' and 'Mapped Name'.")
    uploaded_mapping = st.file_uploader("Upload mapping.xlsx", type=["xlsx"])

    default_mapping_path = find_default_mapping_file(DOWNLOAD_DIR)
    if uploaded_mapping is not None:
        component_map = load_component_map(uploaded_mapping)
        mapping_source = f"uploaded file: {uploaded_mapping.name}"
    elif default_mapping_path is not None:
        component_map = load_component_map(default_mapping_path)
        mapping_source = f"found on disk: {default_mapping_path.name}"
    else:
        component_map = DEFAULT_COMPONENT_MAP
        mapping_source = "built-in default (no mapping.xlsx found or uploaded)"

    st.caption(mapping_source)
    st.dataframe(
        pd.DataFrame(component_map.items(), columns=["Check_unique_name", "Mapped Name"]),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()
    from_date, to_date = compute_date_range()
    st.metric("From date", from_date)
    st.metric("To date", to_date)
    run_clicked = st.button("Run pipeline", type="primary", use_container_width=True)

if run_clicked:
    log_lines = []
    status_box = st.status("Running pipeline...", expanded=True)

    def log(message):
        log_lines.append(message)
        status_box.write(message)

    try:
        results = run_pipeline(component_map=component_map, log=log)
        status_box.update(label="Pipeline finished", state="complete", expanded=False)
        st.session_state.results = results
    except Exception:
        status_box.update(label="Pipeline failed", state="error", expanded=True)
        st.error("The pipeline raised an error - see details below.")
        st.code(traceback.format_exc())
        st.session_state.results = None

results = st.session_state.results
if results:
    st.subheader("Summary")
    col1, col2, col3 = st.columns(3)
    col1.metric("Date range", f"{results['from_date']} -> {results['to_date']}")
    col2.metric("Closed with Insufficiency rows", results.get("closed_rows", 0))
    col3.metric("Report file", results["report_xlsx"].name if "report_xlsx" in results else "-")

    if "report_xlsx" in results:
        xlsx_path = results["report_xlsx"]
        st.subheader("Report preview")
        preview_df = pd.read_excel(xlsx_path)
        st.dataframe(preview_df, use_container_width=True)

        with open(xlsx_path, "rb") as f:
            st.download_button(
                "Download formatted report (.xlsx)",
                data=f.read(),
                file_name=xlsx_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    raw_files = {
        "Daily checks export": results.get("daily_csv"),
        "Advance Tracker export": results.get("advance_csv"),
        "Antecedent Details export": results.get("antecedent_csv"),
    }
    raw_files = {label: path for label, path in raw_files.items() if path is not None}
    if raw_files:
        with st.expander("Raw CSV exports"):
            for label, path in raw_files.items():
                with open(path, "rb") as f:
                    st.download_button(f"Download {label} ({path.name})", data=f.read(), file_name=path.name)
else:
    st.info("Upload a mapping file if needed, then click **Run pipeline** in the sidebar to start.")
