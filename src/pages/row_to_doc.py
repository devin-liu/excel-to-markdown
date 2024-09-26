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

    # Display the grid
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(editable=True)  # Make all columns editable
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

    # Update the dataframe with edited values
    df = grid_response['data']

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
    valid_answer_columns = columns[start_index:]

    answer_columns = st.multiselect(  # Allow multiple answer columns selection
        # Default to the first valid column
        "Select the answer columns", valid_answer_columns, default=[valid_answer_columns[0]])

    # Row selection for question and answer
    # Only rows with questions
    question_rows = df[df[question_column].notna()].index.to_list()
    question_row_selector = st.selectbox(
        "Select the row for the question", question_rows)

    answer_rows = df.index.to_list()  # All rows for answers
    answer_row_selector = st.selectbox(
        "Select the row for the answers", answer_rows)

    # Handle both integer and label-based indices for question and answer rows
    if isinstance(df.index, pd.RangeIndex):
        # If index is a default RangeIndex, use the selected values directly
        question_row_index = question_row_selector
        answer_row_index = answer_row_selector
    else:
        # If index is not a default RangeIndex, find the integer location
        question_row_index = df.index.get_loc(question_row_selector)
        answer_row_index = df.index.get_loc(answer_row_selector)

    first_question_column_value = df.iloc[question_row_index][question_column]
    default_file_name = f"{first_question_column_value}"

    file_name_input = st.text_input(
        "Enter the file name for the markdown document", value=default_file_name)
    markdown_file_name = file_name_input + ".md"

    # create a button to iterate through the columns and combine into a markdown file

    if st.button(f"Generate {len(answer_columns)} markdown files"):
        if question_column in answer_columns:
            st.error("Please select different columns for questions and answers.")
        else:
            for column in answer_columns:
                # Get the value of the row at the answer_row_index for the current column
                file_name = df.iloc[answer_row_index][column]
                file_name = f"{file_name}_qa.md"
                markdown = generate_markdown(
                    df, question_column, answer_columns, question_row_selector, answer_row_index)

                # Download the markdown file
                output_dir = "./data/output/"

                # Sanitize the file name to remove special characters
                file_name = sanitize_filename(file_name)

                full_file_name = output_dir + file_name

                # Check if directory exists
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)

                # Check if file exists and create the file if it doesn't exist
                if not os.path.exists(full_file_name):
                    with open(full_file_name, "w") as f:
                        f.write(markdown)
                else:
                    st.error(f"File {file_name} already exists.")

    if st.button("Generate Markdown Preview"):
        if question_column in answer_columns:
            st.error("Please select different columns for questions and answers.")
        else:
            markdown = generate_markdown(
                df, question_column, answer_columns, question_row_selector, answer_row_selector)
            st.download_button(
                label="Download Markdown",
                data=markdown,
                file_name=markdown_file_name,
                mime="text/markdown"
            )
            st.markdown("### Markdown Preview")
            st.markdown(markdown)


def generate_markdown(df, question_column, answer_columns, question_row_selector, answer_row_selector):  # Updated parameter name
    markdown = ""
    # Extract the question from the specified row
    question = df.iloc[question_row_selector][question_column]
    markdown += f"## {question}\n\n"

    # Iterate through the selected answer columns
    for column in answer_columns:  # Use answer_columns instead of answer_column
        answer = df.iloc[answer_row_selector][column]  # Use answer_row_selector for the answer row
        if pd.notna(answer):
            markdown += f"{column}: {answer}\n"

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
