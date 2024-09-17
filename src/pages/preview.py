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
            # wb = openpyxl.load_workbook(file_path, read_only=True)
            # sheet = wb.active
            wb = pd.read_excel(file_path, sheet_name=None)

            for sheet_name, df in wb.items():
                st.subheader(f"{sheet_name}")

                # Generate markdown and create a link to view it
                if st.button(f"View markdown preview of {sheet_name}"):
                    markdown = dataframe_to_markdown(df)
                    st.markdown(markdown)
                    st.download_button(
                        label="Download Markdown",
                        data=markdown,
                        file_name=f"{sheet_name}.md",
                        mime="text/markdown"
                    )

                st.dataframe(df, use_container_width=True)

        else:
            st.error("File not found.")
    else:
        st.error("No file specified.")


if __name__ == "__main__":
    preview_excel()
