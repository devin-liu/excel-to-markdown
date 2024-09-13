import pandas as pd
import sys
from pathlib import Path

def excel_to_markdown(excel_file, sheet_name=0):
    """
    Convert a specific sheet in an Excel file to a Markdown table.

    Args:
        excel_file (Path): The path to the Excel file.
        sheet_name (str or int): The name or index of the sheet to convert.

    Returns:
        str: The Markdown representation of the sheet.
    """
    # Read the specified sheet from the Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Start building the Markdown table
    markdown = "| " + " | ".join(df.columns) + " |\n"
    markdown += "| " + " | ".join(["---"] * len(df.columns)) + " |\n"

    # Add each row of the DataFrame to the Markdown table
    for _, row in df.iterrows():
        markdown += "| " + " | ".join(str(cell) for cell in row) + " |\n"

    return markdown

def process_file(input_file, output_dir):
    """
    Process an Excel file by converting each of its sheets to separate Markdown files.

    Args:
        input_file (Path): The path to the Excel file.
        output_dir (Path): The directory where Markdown files will be saved.
    """
    try:
        # Load the Excel file to get all sheet names
        excel = pd.ExcelFile(input_file)
        sheet_names = excel.sheet_names

        # Iterate through each sheet in the Excel file
        for sheet in sheet_names:
            # Convert the current sheet to Markdown
            markdown = excel_to_markdown(input_file, sheet)

            # Sanitize the sheet name to create a valid filename
            safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '_') else "_" for c in sheet).strip().replace(" ", "_")

            # Create the output filename by appending the sheet name
            output_file = output_dir / f"{input_file.stem}_{safe_sheet_name}.md"

            # Write the Markdown content to the output file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(markdown)

            print(f"Markdown file '{output_file}' for sheet '{sheet}' has been created successfully.")

    except Exception as e:
        print(f"An error occurred processing {input_file}: {e}")

def main():
    """
    The main function to execute the script. It parses command-line arguments and processes Excel files.
    """
    if len(sys.argv) < 2:
        print("Usage: python script.py <input_directory> <output_directory>")
        sys.exit(1)

    # Define input and output directories from command-line arguments
    input_dir = Path(sys.argv[1])
    output_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else input_dir / 'output'

    # Create the output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    # Iterate through all Excel files in the input directory
    for excel_file in input_dir.glob('*.xlsx'):
        process_file(excel_file, output_dir)

if __name__ == "__main__":
    main()
