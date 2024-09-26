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
    if grid_response['selected_rows'] is not None and not grid_response['selected_rows'].empty:
        # make a df from the selected rows, with the first row as the header
        # the first column as the index
        selected_rows_df = pd.DataFrame(grid_response['selected_rows'])
        selected_rows_df.set_index(selected_rows_df.columns[0], inplace=True)

        pinned_columns = [
            column for column in grid_response['columns_state'] if column['pinned'] is not None]

        # st.write(pinned_columns)
        if pinned_columns:  # Check if there are pinned columns
            # Get the ID of the first pinned column
            st.write(pinned_columns)
            first_pinned_column = pinned_columns[0]['colId']

            # get the location of the first pinned column
            first_index = selected_rows_df.columns.get_loc(first_pinned_column)

            st.write(first_index)

            # Keep columns from the first pinned column onward
            selected_rows_df = selected_rows_df.iloc[:, first_index:]

            st.write(first_pinned_column)

        # use the first pinned column as the start of selected_rows_df, remove columns to the left

        st.write(selected_rows_df)

        # st.write(grid_response['selected_rows'])

    # Column selection
    columns = df.columns.tolist()
    question_column = st.selectbox(
        "Select the question label column", columns)

    # Select the starting column for answers
    start_column = st.selectbox(
        "Select the starting column for answers", columns)

    # Determine valid ending columns based on the selected start column
    start_index = columns.index(start_column)
    # All columns from the start column to the end
    answer_columns = columns[start_index:]

    # Row selection for question and answer
    # Only rows with questions

    if selected_rows_df is not None:

        question_rows = selected_rows_df.index.to_list()  # All rows for answers
        question_row_selector = st.selectbox(
            "Select the row for the question", question_rows)

        answer_rows = selected_rows_df.index.to_list()  # All rows for answers
        answer_row_selector = st.selectbox(
            "Select the row for the answers", answer_rows)

        question_row_index = selected_rows_df.index.get_loc(
            question_row_selector)
        answer_row_index = selected_rows_df.index.get_loc(answer_row_selector)

        first_question_column_value = selected_rows_df.iloc[question_row_index][question_column]
        default_file_name = f"{sheet_name}_{first_question_column_value}"

        file_name_input = st.text_input(
            "Enter the file name for the markdown document", value=default_file_name)
        markdown_file_name = file_name_input + ".md"

        if st.button("Generate Markdown Preview"):
            if question_column in answer_columns:
                st.error(
                    "Please select different columns for questions and answers.")
            else:
                markdown = generate_markdown(
                    selected_rows_df, first_question_column_value, answer_columns, question_row_index, answer_row_index)
                st.download_button(
                    label="Download Markdown",
                    data=markdown,
                    file_name=sanitize_filename(markdown_file_name),
                    mime="text/markdown"
                )
                st.markdown("### Markdown Preview")
                st.markdown(markdown)


# Updated parameter name
def generate_markdown(df, question, answer_columns, question_row_selector, answer_row_selector):
    markdown = ""
    # Extract the question from the specified row
    markdown += f"## {question}\n\n"

    # Iterate through the selected answer columns
    for column in answer_columns:  # Use answer_columns instead of answer_column
        # Use answer_row_selector for the answer row
        answer = df.iloc[answer_row_selector][column]
        value = df.iloc[question_row_selector][column]
        if pd.notna(answer):
            markdown += f"{answer}: {value}\n\n"

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
