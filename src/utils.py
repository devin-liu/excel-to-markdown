import streamlit as st
import pandas as pd
from pathlib import Path
import re


def load_excel_file(file_name):
    file_path = Path("data/input") / file_name
    if file_path.exists():
        return pd.read_excel(file_path, sheet_name=None)
    return None


def get_selected_range(df, sheet_name, selected_rows, selected_columns):
    start_row = min(selected_rows) if selected_rows else 0
    end_row = max(selected_rows) if selected_rows else len(df) - 1
    start_col = min(selected_columns) if selected_columns else 0
    end_col = max(selected_columns) if selected_columns else len(
        df.columns) - 1
    return start_row, end_row, start_col, end_col


def get_file_and_sheet():
    file_name = st.query_params.get("file")
    sheet_name = st.query_params.get("sheet")

    wb = load_excel_file(file_name)
    if wb is None:
        st.error("File not found.")
        return None, None

    if sheet_name not in wb:
        st.error(f"Sheet '{sheet_name}' not found in the file.")
        return None, None

    df = wb[sheet_name]
    return df, sheet_name


def sanitize_filename(filename):
    # Remove or replace special characters
    sanitized = re.sub(r'[^\w\-_\. ]', '', filename)
    # Replace spaces with underscores
    sanitized = sanitized.replace(' ', '_')
    # Ensure the filename is not empty after sanitization
    return sanitized or 'untitled'
