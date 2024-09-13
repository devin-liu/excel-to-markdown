import string
import re
from pathlib import Path


def column_letter_to_index(letter):
    """
    Convert Excel column letter to zero-based index.

    Args:
        letter (str): Column letter (e.g., 'A', 'AA').

    Returns:
        int: Zero-based column index.
    """
    letter = letter.upper()
    result = 0
    for char in letter:
        if char in string.ascii_uppercase:
            result = result * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid column letter: {char}")
    return result - 1


def sanitize_sheet_name(sheet_name):
    """
    Sanitize sheet name to create a valid filename.

    Args:
        sheet_name (str): Original sheet name.

    Returns:
        str: Sanitized sheet name.
    """
    sanitized = re.sub(r'[^\w\s]', '_', sheet_name).strip().replace(" ", "_")
    return sanitized


def create_output_filename(input_file, sheet_name, output_dir):
    """
    Create a sanitized output filename based on input file and sheet name.

    Args:
        input_file (Path): Path to the input Excel file.
        sheet_name (str): Name of the sheet.
        output_dir (Path): Directory to save the output.

    Returns:
        Path: Full path to the output Markdown file.
    """
    safe_sheet_name = sanitize_sheet_name(sheet_name)
    output_filename = f"{input_file.stem}_{safe_sheet_name}.md"
    return output_dir / output_filename
