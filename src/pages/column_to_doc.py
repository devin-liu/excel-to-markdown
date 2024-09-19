import streamlit as st
from utils import load_excel_file
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
import pandas as pd
from components.inputs_files_selector import input_files_selector
from components.sheet_selector import sheet_selector


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

    if st.button("Generate Markdown Preview"):
        if question_column == answer_column:
            st.error("Please select different columns for questions and answers.")
        else:

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
        if _ < row_selector:
            continue
        question = row[question_column]
        answer = row[answer_column]
        if pd.notna(question) and pd.notna(answer):
            markdown += f"## {question}\n\n{answer}\n\n"
    return markdown


if __name__ == "__main__":
    input_files_selector(mode="select")
    if "file" in st.query_params:
        sheet_selector()
    if "file" in st.query_params and "sheet" in st.query_params:
        column_to_doc()
