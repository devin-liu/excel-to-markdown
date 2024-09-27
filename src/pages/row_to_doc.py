import streamlit as st
from utils import load_excel_file
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
import pandas as pd
from components.inputs_files_selector import input_files_selector
from components.sheet_selector import sheet_selector
import io
import os
import re
from pathlib import Path


def create_aggrid(df, sheet_name, selection_mode='multiple'):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(editable=True)  # Make all columns editable
    gb.configure_selection(selection_mode=selection_mode, use_checkbox=True)

    grid_options = gb.build()

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        fit_columns_on_grid_load=True,
        allow_unsafe_jscode=True,
        reload_data=False,
        key=f"{sheet_name}_aggrid",
        editable=True,  # Enable editing for the entire grid
    )

    return grid_response


def row_to_doc():
    file_name = st.query_params.get("file")
    sheet_name = st.query_params.get("sheet")

    wb = load_excel_file(file_name)
    if wb is None:
        st.error("File not found.")
        return

    if sheet_name not in wb:
        st.error(f"Sheet '{sheet_name}' not found in the file.")
        return

    df = wb[sheet_name]
    st.subheader(f"Row to Document - {sheet_name}")

    grid_response = create_aggrid(df, sheet_name)

    df = grid_response['data']

    selected_rows_df = None

    # if a row is selected, is not None, and is not empty, show the row
    if grid_response['selected_rows'] is not None and not pd.DataFrame(grid_response['selected_rows']).empty:
        # make a df from the selected rows, with the first row as the header
        # the first column as the index
        selected_rows_df = pd.DataFrame(grid_response['selected_rows'])
        selected_rows_df.set_index(selected_rows_df.columns[0], inplace=True)

        pinned_columns = [
            column for column in grid_response['columns_state'] if column.get('pinned') is not None]

        if pinned_columns:  # Check if there are pinned columns
            # Get the ID of the first pinned column
            first_pinned_column = pinned_columns[0]['colId']

            # get the location of the first pinned column
            first_index = selected_rows_df.columns.get_loc(first_pinned_column)

            # Keep columns from the first pinned column onward
            selected_rows_df = selected_rows_df.iloc[:, first_index:]

            st.write(f"First pinned column: {first_pinned_column}")
        else:
            st.warning("Please pin at least one column.")
            return

        st.write(selected_rows_df)

    else:
        st.warning("Please select at least one row and pin at least one column.")
        return

    # Column selection
    columns = selected_rows_df.columns.tolist()

    # Select the starting column for answers
    start_column = st.selectbox(
        "Select the starting column for answers", columns)

    # Determine valid answer columns based on the selected start column
    start_index = columns.index(start_column)
    answer_columns = columns[start_index:]

    # Row selection for answer labels
    answer_rows = selected_rows_df.index.tolist()
    answer_row_selector = st.selectbox(
        "Select the row for the answer labels", answer_rows)

    answer_row_index = selected_rows_df.index.get_loc(answer_row_selector)

    default_file_name = f"{sheet_name}_document"

    file_name_input = st.text_input(
        "Enter the file name for the markdown document", value=default_file_name)
    markdown_file_name = file_name_input + ".md"

    if st.button("Generate Markdown Preview"):
        markdown = ""
        for idx in range(answer_row_index + 1, len(selected_rows_df)):
            row = selected_rows_df.iloc[idx]
            question = row[first_pinned_column]
            answers = row[answer_columns]
            answer_labels = selected_rows_df.loc[answer_row_selector,
                                                 answer_columns]
            markdown += generate_markdown(question, answers, answer_labels)

        st.download_button(
            label="Download Markdown",
            data=markdown,
            file_name=sanitize_filename(markdown_file_name),
            mime="text/markdown"
        )
        st.markdown("### Markdown Preview")
        st.markdown(markdown)


def generate_markdown(question, answers, answer_labels):
    markdown = f"## {question}\n\n"
    for label, answer in zip(answer_labels, answers):
        if pd.notna(answer):
            markdown += f"**{label}**: {answer}\n\n"
    return markdown


def sanitize_filename(filename):
    # Remove or replace special characters
    sanitized = re.sub(r'[^\w\-_\. ]', '', filename)
    # Replace spaces with underscores
    sanitized = sanitized.replace(' ', '_')
    # Ensure the filename is not empty after sanitization
    return sanitized or 'untitled'


if __name__ == "__main__":
    input_files_selector(mode="select")
    if "file" in st.query_params:
        sheet_selector()
    if "file" in st.query_params and "sheet" in st.query_params:
        row_to_doc()
