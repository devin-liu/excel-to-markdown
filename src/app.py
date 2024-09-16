import streamlit as st
import os
from pathlib import Path

# Get the list of files in the data/input directory
input_dir = Path("data/input")
input_files = [f for f in os.listdir(
    input_dir) if f.endswith(('.xlsx', '.xls'))]

# Display the list of files using Streamlit
st.header("Excel Files in data/input")
if input_files:
    for file in input_files:
        st.write(f"- {file}")
else:
    st.write("No Excel files found in the input directory.")


# Get the list of files in the data/output directory
output_dir = Path("data/output")
output_files = [f for f in os.listdir(output_dir) if f.endswith('.md')]

# Display the list of output files using Streamlit
st.header("Markdown Files in data/output")
if output_files:
    for file in output_files:
        with open(output_dir / file, 'r') as f:
            content = f.read()         
        st.write(f"- {file}")    
else:
    st.write("No Markdown files found in the output directory.")

