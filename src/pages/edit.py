import streamlit as st
from utils import load_excel_file, get_selected_range
from excel_to_markdown.markdown_generator import dataframe_to_markdown


def edit_excel():
    file_name = st.query_params.get("file")
    sheet_name = st.query_params.get("sheet")
    start_row = int(st.query_params.get("start_row", 0))
    end_row = int(st.query_params.get("end_row", 0))
    start_col = int(st.query_params.get("start_col", 0))
    end_col = int(st.query_params.get("end_col", 0))

    wb = load_excel_file(file_name)
    if wb is None:
        st.error("File not found.")
        return

    for current_sheet, df in wb.items():
        if sheet_name and current_sheet != sheet_name:
            continue

        st.subheader(f"{current_sheet}")

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            disabled=True,  # Prevent editing to focus on selection
            hide_index=False,
            column_config={col: st.column_config.Column(
                required=True) for col in df.columns},
            key=f"{current_sheet}_data_editor"  # Set a unique key
        )

        # Get the selected rows and columns using the new key
        selected_rows = st.session_state.get(
            f"{current_sheet}_data_editor_selected_rows", [])
        selected_columns = st.session_state.get(
            f"{current_sheet}_data_editor_selected_columns", [])

        # Ensure valid values for start_row, end_row, start_col, end_col
        if not selected_rows:
            start_row = 0
            end_row = len(df) - 1
        if not selected_columns:
            start_col = 0
            end_col = len(df.columns) - 1

        # Display the current selection
        st.write(f"Selected range: Rows {start_row} to {
                 end_row}, Columns {start_col} to {end_col}")

        col1, col2 = st.columns(2)
        with col1:
            start_row = st.number_input(f"Start Row for {
                                        current_sheet}", min_value=0, max_value=len(df)-1, value=start_row)
            end_row = st.number_input(f"End Row for {
                                      current_sheet}", min_value=start_row, max_value=len(df)-1, value=end_row)
        with col2:
            start_col = st.number_input(f"Start Column for {
                                        current_sheet}", min_value=0, max_value=len(df.columns)-1, value=start_col)
            end_col = st.number_input(f"End Column for {
                                      current_sheet}", min_value=start_col, max_value=len(df.columns)-1, value=end_col)

        if st.button(f"View markdown preview of {current_sheet}"):
            selected_df = edited_df.iloc[start_row:end_row +
                                         1, start_col:end_col+1]
            markdown = dataframe_to_markdown(selected_df)
            st.markdown(markdown)
            st.download_button(
                label="Download Markdown",
                data=markdown,
                file_name=f"{current_sheet}_selected.md",
                mime="text/markdown"
            )

        if sheet_name:
            break  # Exit the loop after processing the specified sheet


if __name__ == "__main__":
    edit_excel()
