# excel_to_markdown/detector.py

import pandas as pd

def detect_table_start(df):
    """
    Detect the starting row of the table by finding the first row
    that is completely filled within the non-null columns.
    """
    # Identify columns that have any non-null values
    non_null_columns = df.columns[df.notnull().any(axis=0)]
    if len(non_null_columns) == 0:
        return None  # No non-null columns found

    # Get the indices of the leftmost and rightmost non-null columns
    left_col_index = df.columns.get_loc(non_null_columns[0])
    right_col_index = df.columns.get_loc(non_null_columns[-1])

    # Iterate through each row to find the first fully populated row within the non-null column range
    for idx, row in df.iterrows():
        row_slice = row.iloc[left_col_index:right_col_index + 1]
        if row_slice.notnull().all() and not row_slice.astype(str).str.strip().eq('').any():
            return idx

    return None


def get_table_region(df):
    """
    Improved logic to detect all relevant columns based on the presence of header names and non-null values.
    """
    start_row = detect_table_start(df)
    if start_row is not None:
        print(f"Automatically detected table starting at row {start_row + 1}.")
        # Consider all columns that have at least a certain percentage of non-null values as part of the table
        threshold = 0.5  # At least 50% non-null to be considered
        valid_cols = [col for col in df.columns if df[col].notnull().mean() > threshold]
        return start_row, valid_cols
    else:
        print("Automatic table detection failed.")
        # Manual detection as fallback
        while True:
            try:
                headers_row = int(input("Enter the header row number (1-based index): ")) - 1
                cols_input = input("Enter the columns to include (e.g., A:D or 1-4): ")
                from .parser import parse_columns
                usecols = parse_columns(cols_input, df)
                return headers_row, usecols
            except Exception as e:
                print(f"Invalid input: {e}. Please try again.")

