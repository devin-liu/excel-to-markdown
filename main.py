import pandas as pd
import sys

def excel_to_markdown(excel_file, sheet_name=0):
    # Read the Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Generate the markdown table
    markdown = "| " + " | ".join(df.columns) + " |\n"
    markdown += "| " + " | ".join(["---"] * len(df.columns)) + " |\n"
    
    for _, row in df.iterrows():
        markdown += "| " + " | ".join(str(cell) for cell in row) + " |\n"
    
    return markdown

def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <excel_file> [sheet_name]")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    sheet_name = sys.argv[2] if len(sys.argv) > 2 else 0
    
    try:
        markdown = excel_to_markdown(excel_file, sheet_name)
        
        # Write the markdown to a file
        output_file = excel_file.rsplit('.', 1)[0] + '.md'
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown)
        
        print(f"Markdown file '{output_file}' has been created successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()