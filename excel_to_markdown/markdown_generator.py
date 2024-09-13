# excel_to_markdown/markdown_generator.py
import pandas as pd


def dataframe_to_markdown(df):
    """
    Convert a pandas DataFrame to a Markdown table.

    Args:
        df (pd.DataFrame): The DataFrame to convert.

    Returns:
        str: Markdown-formatted table.
    """
    if df.empty:
        return ""
    # Generate the header row
    markdown = "| " + " | ".join(df.columns) + " |\n"
    # Generate the separator row
    markdown += "| " + " | ".join(["---"] * len(df.columns)) + " |\n"

    # Generate each data row
    for _, row in df.iterrows():
        row_values = [str(cell) if pd.notnull(cell) else "" for cell in row]
        markdown += "| " + " | ".join(row_values) + " |\n"

    return markdown
