from .utils import column_letter_to_index


def parse_columns(cols_input, df):
    """
    Parse the user input for columns.

    Args:
        cols_input (str): User input specifying columns (e.g., "A:D" or "1-4").

    Returns:
        list: List of column names based on the input.
    """
    cols_input = cols_input.replace(" ", "").upper()
    if ':' in cols_input:
        start, end = cols_input.split(':')
    elif '-' in cols_input:
        start, end = cols_input.split('-')
    else:
        start, end = cols_input, cols_input

    if start.isalpha() and end.isalpha():
        start_idx = column_letter_to_index(start)
        end_idx = column_letter_to_index(end)
    elif start.isdigit() and end.isdigit():
        start_idx = int(start) - 1
        end_idx = int(end) - 1
    else:
        raise ValueError("Mixed or invalid column specification.")

    if start_idx > end_idx:
        raise ValueError("Start column comes after end column.")

    selected_columns = df.columns[start_idx:end_idx + 1].tolist()
    return selected_columns
