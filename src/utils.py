import pandas as pd
from pathlib import Path

def load_excel_file(file_name):
    file_path = Path("data/input") / file_name
    if file_path.exists():
        return pd.read_excel(file_path, sheet_name=None)
    return None

def get_selected_range(df, sheet_name, selected_rows, selected_columns):
    start_row = min(selected_rows) if selected_rows else 0
    end_row = max(selected_rows) if selected_rows else len(df) - 1
    start_col = min(selected_columns) if selected_columns else 0
    end_col = max(selected_columns) if selected_columns else len(df.columns) - 1
    return start_row, end_row, start_col, end_col