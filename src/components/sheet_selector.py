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

            # Get the default sheet from query params
            default_sheet = st.query_params.get("sheet")

            # Find the index of the default sheet
            default_index = 0
            if default_sheet in sheet_names:
                default_index = sheet_names.index(default_sheet)

            selected_sheet = st.selectbox(
                "Select a sheet", sheet_names, index=default_index)

            if selected_sheet:
                st.query_params["sheet"] = selected_sheet
                st.success(f"Selected sheet: {selected_sheet}")
    else:
        st.warning(
            "No file selected. Please provide a file parameter in the URL.")

# Usage example (can be placed in your main Streamlit app file):
# sheet_selector()
