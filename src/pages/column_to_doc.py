import streamlit as st
from utils import load_excel_file
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
import pandas as pd
from components.inputs_files_selector import input_files_selector
from components.sheet_selector import sheet_selector
import os
import re


def column_to_doc():
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
    st.subheader(f"Column to Document - {sheet_name}")

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

    # row selection
    rows = df.index.to_list()
    row_selector = st.selectbox(
        "Select the row to start generating the markdown from", rows)

    # Column selection
    columns = df.columns.tolist()
    question_column = st.selectbox(
        "Select the question column", columns)
    answer_column = st.selectbox(
        "Select the answer column", columns)

    # Handle both integer and label-based indices
    if isinstance(df.index, pd.RangeIndex):
        # If index is a default RangeIndex, use the selected value directly
        row_index = row_selector
    else:
        # If index is not a default RangeIndex, find the integer location
        row_index = df.index.get_loc(row_selector)

    first_answer_column_value = df.iloc[row_index][answer_column]
    default_file_name = f"{first_answer_column_value}_qa"

    file_name_input = st.text_input(
        "Enter the file name for the markdown document", value=default_file_name)
    markdown_file_name = file_name_input + ".md"

    # make a smaller dataframe that removes the columns before the answer column
    answer_columns = df.columns[df.columns.get_loc(answer_column):]

    # create a button to iterate through the columns and generate a markdown file for each column

    if st.button(f"Generate {len(answer_columns)} markdown files"):
        if question_column == answer_column:
            st.error("Please select different columns for questions and answers.")
        else:
            for column in answer_columns:
                # get the value of the row at the row_index for the current column
                file_name = df.iloc[row_index][column]
                file_name = f"{file_name}_qa.md"
                markdown = generate_markdown(
                    # Convert row_index to int
                    df, question_column, column, row_index)

                # download the markdown file

                output_dir = "./data/output/"

                # sanitize the file name to remove special characters using a library
                file_name = sanitize_filename(file_name)

                full_file_name = output_dir + file_name

                # check if directory exists
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)

                # check if file exists
                # create the file if it doesn't exist
                if not os.path.exists(full_file_name):
                    with open(full_file_name, "w") as f:
                        f.write(markdown)
                else:
                    st.error(f"File {file_name} already exists.")


    if st.button("Generate Markdown Preview"):
        if question_column == answer_column:
            st.error("Please select different columns for questions and answers.")
        else:

            # fill NA values in the answer column with empty strings
            df[answer_column] = df[answer_column].fillna("")

            markdown = generate_markdown(
                df, question_column, answer_column, row_selector)
            st.download_button(
                label="Download Markdown",
                data=markdown,
                file_name=markdown_file_name,
                mime="text/markdown"
            )
            st.markdown("### Markdown Preview")
            st.markdown(markdown)


def generate_markdown(df, question_column, answer_column, row_selector):
    markdown = ""
    for _, row in df.iterrows():
        if int(_) < int(row_selector):
            continue
        question = row[question_column]
        answer = row[answer_column]
        if pd.notna(question) and pd.notna(answer):
            markdown += f"## {question}\n\n{answer}\n\n"
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
        column_to_doc()
