import streamlit as st
from pathlib import Path
import pandas as pd


def sheet_selector():
    if "file" in st.query_params:
        file_name = st.query_params["file"]
        file_path = Path("data/input") / file_name
        if file_path.exists():
            wb = pd.read_excel(file_path, sheet_name=None)

            sheet_names = list(wb.keys())
            selected_sheet = st.selectbox("Select a sheet", sheet_names)

            if selected_sheet:                
                st.query_params["sheet"] = selected_sheet
                st.success(f"Selected sheet: {selected_sheet}")
    else:
        st.warning(
            "No file selected. Please provide a file parameter in the URL.")

# Usage example (can be placed in your main Streamlit app file):
# sheet_selector()
