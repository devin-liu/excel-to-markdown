import streamlit as st
from pathlib import Path
import os

def input_files_selector():
    # Get the list of files in the data/input directory
    input_dir = Path("data/input")
    input_files = [f for f in os.listdir(input_dir) if f.endswith(('.xlsx', '.xls'))]

    # Display the list of files using Streamlit
    st.header("Excel Files in data/input")
    if input_files:
        for file in input_files:        
            st.link_button(f"Preview {file}", f"/preview?file={file}")
    else:
        st.write("No Excel files found in the input directory.")

# You can call this component in your main Streamlit app like this:
# input_files_selector()