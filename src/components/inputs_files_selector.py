import streamlit as st
from pathlib import Path
import os
from urllib.parse import urlencode


def input_files_selector(mode="link"):
    # Get the list of files in the data/input directory
    input_dir = Path("data/input")
    input_files = [f for f in os.listdir(
        input_dir) if f.endswith(('.xlsx', '.xls'))]

    # Get the default file from query params
    default_file = st.query_params.get("file")

    # Display the list of files using Streamlit
    st.header("Excel Files in data/input")
    if input_files:
        for file in input_files:
            if mode == "link":
                # Create a query string with both file and sheet parameters
                query_params = {"file": file}
                if "sheet" in st.query_params:
                    query_params["sheet"] = st.query_params["sheet"]
                query_string = urlencode(query_params)

                st.link_button(f"Preview {file}", f"/preview?{query_string}")
            elif mode == "select":
                # If this file is the default, use a success message instead of a button
                if file == default_file:
                    st.success(f"Selected file: {file}")
                else:
                    if st.button(f"Select {file}"):
                        st.query_params["file"] = file
                        st.rerun()  # Rerun the app to reflect the change
    else:
        st.write("No Excel files found in the input directory.")
