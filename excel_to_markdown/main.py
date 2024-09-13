import pandas as pd
import sys
import os
from pathlib import Path

def excel_to_markdown(excel_file, sheet_name=0):
    # Read the Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Generate the markdown table
    markdown = "| " + " | ".join(df.columns) + " |\n"
    markdown += "| " + " | ".join(["---"] * len(df.columns)) + " |\n"
    
    for _, row in df.iterrows():
        markdown += "| " + " | ".join(str(cell) for cell in row) + " |\n"
    
    return markdown

def process_file(input_file, output_dir):
    try:
        markdown = excel_to_markdown(input_file)
        
        # Create output filename
        output_file = output_dir / (input_file.stem + '.md')
        
        # Write the markdown to a file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown)
        
        print(f"Markdown file '{output_file}' has been created successfully.")
    except Exception as e:
        print(f"An error occurred processing {input_file}: {e}")

def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <input_directory> <output_directory>")
        sys.exit(1)
    
    input_dir = Path(sys.argv[1])
    output_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else input_dir / 'output'
    
    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Process all Excel files in the input directory
    for excel_file in input_dir.glob('*.xlsx'):
        process_file(excel_file, output_dir)

if __name__ == "__main__":
    main()