import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
import pandas as pd
from components.inputs_files_selector import input_files_selector
from components.sheet_selector import sheet_selector
from utils import get_file_and_sheet

def create_aggrid(df, selection_mode='multiple'):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(editable=True)
    gb.configure_selection(selection_mode=selection_mode, use_checkbox=True)
    grid_options = gb.build()

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        fit_columns_on_grid_load=True,
        key="aggrid",
        editable=True,
    )
    return grid_response


def generate_markdown(selected_rows_df):
    markdown = ""
    for index, row in selected_rows_df.iterrows():
        markdown += f"## {row[0]}\n\n"
        for col_name, value in row.iteritems():
            if pd.notna(value):
                markdown += f"**{col_name}**: {value}\n\n"
    return markdown


def select_range_to_markdown():
    st.title("Select Range to Markdown")

    df, sheet_name = get_file_and_sheet()
    
    st.subheader(f"Row to Document - {sheet_name}")

    grid_response = create_aggrid(df, sheet_name)

    selected_rows = grid_response['selected_rows']

    if selected_rows:
        selected_rows_df = pd.DataFrame(selected_rows)
        st.write("Selected Rows:")
        st.dataframe(selected_rows_df)

        if st.button("Generate Markdown"):
            markdown = generate_markdown(selected_rows_df)
            st.markdown("### Markdown Preview")
            st.markdown(markdown)

            st.download_button(
                label="Download Markdown",
                data=markdown,
                file_name="selected_rows.md",
                mime="text/markdown"
            )
    else:
        st.info("Please select at least one row to generate Markdown.")


if __name__ == "__main__":
    input_files_selector(mode="select")
    if "file" in st.query_params:
        sheet_selector()
    if "file" in st.query_params and "sheet" in st.query_params:
        select_range_to_markdown()
