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
                    start_row = st.number_input(f"Start Row for {sheet_name}", min_value=0, max_value=len(df)-1, value=0)
                    end_row = st.number_input(f"End Row for {sheet_name}", min_value=start_row, max_value=len(df)-1, value=len(df)-1)
                with col2:
                    start_col = st.number_input(f"Start Column for {sheet_name}", min_value=0, max_value=len(df.columns)-1, value=0)
                    end_col = st.number_input(f"End Column for {sheet_name}", min_value=start_col, max_value=len(df.columns)-1, value=len(df.columns)-1)

                # Generate markdown and create a link to view it
                if st.button(f"View markdown preview of {sheet_name}"):
                    selected_df = df.iloc[start_row:end_row+1, start_col:end_col+1]
                    markdown = dataframe_to_markdown(selected_df)
                    st.markdown(markdown)
                    st.download_button(
                        label="Download Markdown",
                        data=markdown,
                        file_name=f"{sheet_name}_selected.md",
                        mime="text/markdown"
                    )

                st.dataframe(df, use_container_width=True)

        else:
            st.error("File not found.")
    else:
        st.error("No file specified.")


if __name__ == "__main__":
    preview_excel()
