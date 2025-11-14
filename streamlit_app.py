# streamlit_app.py
import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import io
from src.generate_universal_report import generate_report  # assume your engine available as function

st.set_page_config(page_title="Universal Excel Automation", layout="centered")
st.title("Universal Excel Automation — Upload Excel and get Report")

uploaded_file = st.file_uploader("Upload an Excel file (.xlsx)", type=["xlsx"], accept_multiple_files=False)

if uploaded_file is not None:
    with st.spinner("Generating report..."):
        # save uploaded to temp path
        tmp_in = Path(tempfile.gettempdir()) / uploaded_file.name
        with open(tmp_in, "wb") as f:
            f.write(uploaded_file.getbuffer())

        out_path = Path(tempfile.gettempdir()) / f"{Path(uploaded_file.name).stem}_report.xlsx"
        try:
            generate_report(tmp_in, out_path)  # uses your existing function
            with open(out_path, "rb") as f:
                st.success("Report generated!")
                st.download_button("Download Report", f.read(), file_name=out_path.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error("Error: " + str(e))
            raise
