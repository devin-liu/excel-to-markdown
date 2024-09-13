# excel_to_markdown/detector.py

import pandas as pd

def detect_table_start(df):
    """
    Detect the starting row of the table by finding the first row
    that is completely filled to the right (i.e., no nulls up to the last used column).

    Args:
        df (pd.DataFrame): The DataFrame representing the entire sheet.

    Returns:
        int or None: The index of the header row if found, else None.
    """
    # Determine the last column with any non-null value
    non_null_counts = df.notnull().sum(axis=0)
    last_col = non_null_counts.idxmax()
    last_col_index = df.columns.get_loc(last_col)

    # Iterate through each row to find the first fully populated row up to last_col_index
    for idx, row in df.iterrows():
        row_slice = row.iloc[:last_col_index + 1]
        if row_slice.notnull().all():
            return idx

    return None

def get_table_region(df):
    """
    Attempt to detect the table region. If unsuccessful, prompt the user for input.

    Args:
        df (pd.DataFrame): The DataFrame representing the entire sheet.

    Returns:
        tuple: (headers_row (int), usecols (list))
    """
    start_row = detect_table_start(df)
    if start_row is not None:
        print(f"Automatically detected table starting at row {start_row + 1}.")
        # Detect columns with data up to the last used column
        non_null_counts = df.notnull().sum(axis=0)
        last_col = non_null_counts.idxmax()
        last_col_index = df.columns.get_loc(last_col)
        usecols = df.columns[:last_col_index + 1].tolist()
        return start_row, usecols
    else:
        print("Automatic table detection failed.")
        # Prompt user for input
        while True:
            try:
                headers_row = int(input("Enter the header row number (1-based index): ")) - 1
                cols_input = input("Enter the columns to include (e.g., A:D or 1-4): ")
                from .parser import parse_columns
                usecols = parse_columns(cols_input, df)
                return headers_row, usecols
            except Exception as e:
                print(f"Invalid input: {e}. Please try again.")
