import streamlit as st
from utils import load_excel_file
from excel_to_markdown.markdown_generator import dataframe_to_markdown
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
import pandas as pd


def edit_excel():
    file_name = st.query_params.get("file")
    sheet_name = st.query_params.get("sheet")

    wb = load_excel_file(file_name)
    if wb is None:
        st.error("File not found.")
        return

    for current_sheet, df in wb.items():
        if sheet_name and current_sheet != sheet_name:
            continue

        st.subheader(f"{current_sheet}")

        # Initialize session state keys
        start_row_key = f"{current_sheet}_start_row"
        end_row_key = f"{current_sheet}_end_row"
        start_col_key = f"{current_sheet}_start_col"
        end_col_key = f"{current_sheet}_end_col"

        # Initialize selection keys if not present
        if start_row_key not in st.session_state:
            st.session_state[start_row_key] = 0
        if end_row_key not in st.session_state:
            st.session_state[end_row_key] = len(df) - 1
        if start_col_key not in st.session_state:
            st.session_state[start_col_key] = 0
        if end_col_key not in st.session_state:
            st.session_state[end_col_key] = len(df.columns) - 1

        # Build grid options with selection
        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_selection(selection_mode='multiple', use_checkbox=True)
        grid_options = gb.build()

        # Display the grid
        grid_response = AgGrid(
            df,
            gridOptions=grid_options,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            update_mode=GridUpdateMode.SELECTION_CHANGED | GridUpdateMode.VALUE_CHANGED,
            fit_columns_on_grid_load=True,
            allow_unsafe_jscode=True,
            reload_data=False,
            key=f"{current_sheet}_aggrid_{st.session_state.get('aggrid_key', 0)}"
        )

        selected_rows = grid_response['selected_rows']

        # Initialize selected_indices
        selected_indices = []

        if isinstance(selected_rows, pd.DataFrame):
            if not selected_rows.empty:
                selected_indices = selected_rows.index.tolist()
        elif isinstance(selected_rows, list):
            if len(selected_rows) > 0:
                selected_indices = [
                    int(row['_selectedRowNodeInfo']['nodeRowIndex']) for row in selected_rows]
        else:
            selected_indices = []

        # Synchronize selections to session state
        selection_changed = False
        if selected_indices:
            new_start_row = min(selected_indices)
            new_end_row = max(selected_indices)
            if (new_start_row != st.session_state[start_row_key] or
                    new_end_row != st.session_state[end_row_key]):
                st.session_state[start_row_key] = new_start_row
                st.session_state[end_row_key] = new_end_row
                selection_changed = True
        else:
            if st.session_state[start_row_key] != 0 or st.session_state[end_row_key] != len(df) - 1:
                st.session_state[start_row_key] = 0
                st.session_state[end_row_key] = len(df) - 1
                selection_changed = True

        # Retrieve current values from session state
        start_row = st.session_state[start_row_key]
        end_row = st.session_state[end_row_key]
        start_col = st.session_state[start_col_key]
        end_col = st.session_state[end_col_key]

        # Display the current selection
        st.write(f"Selected range: Rows {start_row} to {end_row}, Columns {start_col} to {end_col}")

        col1, col2 = st.columns(2)
        with col1:
            start_row_input = st.number_input(
                f"Start Row for {current_sheet}",
                min_value=0,
                max_value=len(df)-1,
                value=int(start_row),
                key=f"{start_row_key}_input"
            )
            end_row_input = st.number_input(
                f"End Row for {current_sheet}",
                min_value=int(start_row_input),
                max_value=len(df)-1,
                value=int(end_row),
                key=f"{end_row_key}_input"
            )
        with col2:
            start_col_input = st.number_input(
                f"Start Column for {current_sheet}",
                min_value=0,
                max_value=len(df.columns)-1,
                value=int(start_col),
                key=f"{start_col_key}_input"
            )
            end_col_input = st.number_input(
                f"End Column for {current_sheet}",
                min_value=int(start_col_input),
                max_value=len(df.columns)-1,
                value=int(end_col),
                key=f"{end_col_key}_input"
            )

        # # Update session state if number inputs changed
        # inputs_changed = False
        # if (start_row_input != st.session_state[start_row_key] or
        #         end_row_input != st.session_state[end_row_key]):
        #     st.session_state[start_row_key] = start_row_input
        #     st.session_state[end_row_key] = end_row_input
        #     inputs_changed = True

        # if inputs_changed or selection_changed:
        #     pre_selected_indices = list(range(
        #         int(st.session_state[start_row_key]), int(st.session_state[end_row_key]) + 1))

        #     gb = GridOptionsBuilder.from_dataframe(df)
        #     gb.configure_selection(
        #         selection_mode='multiple',
        #         use_checkbox=True,
        #         pre_selected_rows=pre_selected_indices
        #     )
        #     grid_options = gb.build()

        #     st.session_state['aggrid_key'] = st.session_state.get('aggrid_key', 0) + 1

        #     # Re-render grid with updated selection
        #     grid_response = AgGrid(
        #         df,
        #         gridOptions=grid_options,
        #         data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        #         update_mode=GridUpdateMode.SELECTION_CHANGED | GridUpdateMode.VALUE_CHANGED,
        #         fit_columns_on_grid_load=True,
        #         allow_unsafe_jscode=True,
        #         reload_data=True,
        #         key=f"{current_sheet}_aggrid_{st.session_state['aggrid_key']}"
        #     )

        if st.button(f"View markdown preview of {current_sheet}"):
            selected_df = df.iloc[
                int(st.session_state[start_row_key]):int(st.session_state[end_row_key])+1,
                int(st.session_state[start_col_key]):int(st.session_state[end_col_key])+1
            ]
            markdown = dataframe_to_markdown(selected_df)
            st.markdown(markdown)
            st.download_button(
                label="Download Markdown",
                data=markdown,
                file_name=f"{current_sheet}_selected.md",
                mime="text/markdown"
            )

        if sheet_name:
            break


if __name__ == "__main__":
    edit_excel()
