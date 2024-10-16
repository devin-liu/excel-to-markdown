import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
import pandas as pd
from components.inputs_files_selector import input_files_selector
from components.sheet_selector import sheet_selector
from utils import get_file_and_sheet, sanitize_filename


def select_range_to_markdown():
    st.title("Select Range to Markdown")

    st.text("Select rows where values exist")

    df, sheet_name = get_file_and_sheet()

    # Use the first row as values
    df = df.reset_index(drop=True)
    # Convert column names to strings
    df.columns = [str(i) for i in range(len(df.columns))]

    st.subheader(f"Range to Document - {sheet_name}")

    grid_response = create_aggrid(df, sheet_name)

    selected_rows = grid_response['selected_rows']

    default_file_name = f"{sheet_name}_document"

    file_name_input = st.text_input(
        "Enter the file name for the markdown document", value=default_file_name)

    if selected_rows is not None and len(selected_rows) > 0:
        selected_rows_df = pd.DataFrame(selected_rows)
        st.write("Selected Rows:")
        st.dataframe(selected_rows_df)

        if st.button("Generate Markdown"):

            markdown_file_name = file_name_input + ".md"
            markdown = generate_markdown(selected_rows_df, file_name_input)
            st.markdown("### Markdown Preview")
            st.markdown(markdown)

            st.download_button(
                label="Download Markdown",
                data=markdown,
                file_name=sanitize_filename(markdown_file_name),
                mime="text/markdown"
            )
    else:
        st.info("Please select at least one row to generate Markdown.")


def generate_markdown(selected_rows_df, title=''):
    markdown = ""
    if title:
        markdown += f"## {title}\n\n"
    for index, row in selected_rows_df.iterrows():
        for col_name, value in row.items():
            if pd.notna(value):
                markdown += f"{value}\n\n"
    return markdown


def create_aggrid(df, sheet_name, selection_mode='multiple'):
    # Ensure column names are strings
    df.columns = df.columns.astype(str)

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(editable=True)
    gb.configure_selection(selection_mode=selection_mode, use_checkbox=True)

    grid_options = gb.build()

    grid_response = AgGrid(
        df,  # Include all rows in the grid
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        fit_columns_on_grid_load=True,
        # Use sheet_name in the key to make it unique
        key=f"aggrid_{sheet_name}",
        editable=True,
    )

    return grid_response


if __name__ == "__main__":
    input_files_selector(mode="select")
    if "file" in st.query_params:
        sheet_selector()
    if "file" in st.query_params and "sheet" in st.query_params:
        select_range_to_markdown()
