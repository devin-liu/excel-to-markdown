# excel_to_markdown/main.py

import pandas as pd
import sys
from pathlib import Path

from .detector import get_table_region
from .markdown_generator import dataframe_to_markdown
from .utils import create_output_filename

def excel_to_markdown(excel_file, sheet_name=0):
    """
    Convert a specific sheet in an Excel file to a Markdown table interactively.

    Args:
        excel_file (Path): The path to the Excel file.
        sheet_name (str or int): The name or index of the sheet to convert.

    Returns:
        str: The Markdown representation of the sheet.
    """
    # Read the entire sheet without specifying headers or columns
    df_full = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine='openpyxl')

    # Detect table region
    headers_row, usecols = get_table_region(df_full)

    # Read the table with detected or user-specified parameters
    df = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=headers_row,
        usecols=usecols,
        engine='openpyxl'
    )

    # Drop completely empty rows
    df.dropna(how='all', inplace=True)

    # Reset index after dropping rows
    df.reset_index(drop=True, inplace=True)

    # Generate the markdown table
    markdown = dataframe_to_markdown(df)

    return markdown

def process_file(input_file, output_dir):
    """
    Process an Excel file by converting each of its sheets to separate Markdown files interactively.

    Args:
        input_file (Path): The path to the Excel file.
        output_dir (Path): The directory where Markdown files will be saved.
    """
    try:
        # Load the Excel file to get all sheet names
        excel = pd.ExcelFile(input_file, engine='openpyxl')
        sheet_names = excel.sheet_names

        # Iterate through each sheet in the Excel file
        for sheet in sheet_names:
            print(f"\nProcessing sheet: '{sheet}' in file '{input_file.name}'")
            # Convert the current sheet to Markdown interactively
            markdown = excel_to_markdown(input_file, sheet)

            # Create the output filename using utility function
            output_file = create_output_filename(input_file, sheet, output_dir)

            # Write the Markdown content to the output file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(markdown)

            print(f"Markdown file '{output_file}' for sheet '{sheet}' has been created successfully.")

    except Exception as e:
        print(f"An error occurred processing {input_file}: {e}")

def main():
    """
    The main function to execute the script. It parses command-line arguments and processes Excel files interactively.
    """
    if len(sys.argv) < 2:
        print("Usage: python -m excel_to_markdown.main <input_directory> <output_directory>")
        sys.exit(1)

    # Define input and output directories from command-line arguments
    input_dir = Path(sys.argv[1])
    output_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else input_dir / 'output'

    # Create the output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    # Iterate through all Excel files in the input directory
    excel_files = list(input_dir.glob('*.xlsx')) + list(input_dir.glob('*.xls'))
    if not excel_files:
        print(f"No Excel files found in {input_dir}.")
        sys.exit(1)

    for excel_file in excel_files:
        process_file(excel_file, output_dir)

if __name__ == "__main__":
    main()
