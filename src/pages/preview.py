import streamlit as st
from pathlib import Path
import pandas as pd
from excel_to_markdown.markdown_generator import dataframe_to_markdown

st.set_page_config(layout="wide")


def preview_excel():
    file_name = st.query_params.get("file")
    if file_name:
        file_path = Path("data/input") / file_name
        if file_path.exists():
            wb = pd.read_excel(file_path, sheet_name=None)

            for sheet_name, df in wb.items():
                st.subheader(f"{sheet_name}")

                # Add row and column range selectors
                col1, col2 = st.columns(2)
                with col1:
                    start_row = st.number_input(
                        f"Start Row for {sheet_name}",
                        min_value=0,
                        max_value=len(df) - 1,
                        value=0,
                    )
                    end_row = st.number_input(
                        f"End Row for {sheet_name}",
                        min_value=start_row,
                        max_value=len(df) - 1,
                        value=len(df) - 1,
                    )
                with col2:
                    start_col = st.number_input(
                        f"Start Column for {sheet_name}",
                        min_value=0,
                        max_value=len(df.columns) - 1,
                        value=0,
                    )
                    end_col = st.number_input(
                        f"End Column for {sheet_name}",
                        min_value=start_col,
                        max_value=len(df.columns) - 1,
                        value=len(df.columns) - 1,
                    )

                # Add button to edit selected range
                # if st.button(f"Edit selected range of {sheet_name}"):
                #     edit_url = f"/edit?file={file_name}&sheet={sheet_name}&start_row={start_row}&end_row={end_row}&start_col={start_col}&end_col={end_col}"
                #     st.switch_page("pages/edit.py")
                edit_url = f"/edit?file={file_name}&sheet={sheet_name}&start_row={start_row}&end_row={end_row}&start_col={start_col}&end_col={end_col}"
                st.link_button(f"Edit selected range of {sheet_name}", edit_url)

                # Display the entire DataFrame
                st.dataframe(df, use_container_width=True)

        else:
            st.error("File not found.")
    else:
        st.error("No file specified.")


if __name__ == "__main__":
    preview_excel()
