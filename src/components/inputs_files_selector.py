import streamlit as st
from pathlib import Path
import os
from urllib.parse import urlencode


def input_files_selector(mode="link"):
    # Get the list of files in the data/input directory
    input_dir = Path("data/input")
    input_files = [f for f in os.listdir(
        input_dir) if f.endswith(('.xlsx', '.xls'))]

    # Display the list of files using Streamlit
    st.header("Excel Files in data/input")
    if input_files:
        for file in input_files:
            if mode == "link":
                st.link_button(f"Preview {file}", f"/preview?file={file}")
            elif mode == "select":
                if st.button(f"Select {file}"):
                    st.query_params["file"] = file
                    st.success(f"Selected file: {file}")
    else:
        st.write("No Excel files found in the input directory.")
