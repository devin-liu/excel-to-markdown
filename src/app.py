import streamlit as st
import os
from pathlib import Path
from components.inputs_files_selector import input_files_selector


input_files_selector()

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

