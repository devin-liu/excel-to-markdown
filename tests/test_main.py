import pytest
import pandas as pd
import tempfile
from pathlib import Path
from unittest import mock
from excel_to_markdown.main import process_file, excel_to_markdown
from excel_to_markdown.utils import create_output_filename

@pytest.fixture
def sample_excel_file():
    # Create a temporary Excel file with multiple sheets
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        excel_path = tmpdir_path / "test_excel.xlsx"
        
        # Create sample data for two sheets
        sheet1_data = {
            'A': [None, 'Name', 'Alice', 'Bob'],
            'B': [None, 'Age', 30, 25],
            'C': [None, 'City', 'New York', 'Los Angeles']
        }
        sheet2_data = {
            'A': [None, None, 'Product', 'Item1'],
            'B': [None, None, 'Price', 99.99],
            'C': [None, None, 'Stock', 50]
        }
        
        # Create DataFrames
        df1 = pd.DataFrame(sheet1_data)
        df2 = pd.DataFrame(sheet2_data)
        
        # Write to Excel with two sheets
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
            df2.to_excel(writer, sheet_name='Sheet2', index=False, header=False)
        
        yield excel_path  # Provide the path to the test

@pytest.fixture
def output_directory():
    # Create a temporary directory for output
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)

def test_process_file_auto_detect(sample_excel_file, output_directory, mocker):
    # Mock the input function to avoid interactive prompts
    # Since automatic detection should succeed for Sheet1
    mocker.patch('builtins.input')
    
    # Run the process_file function
    process_file(sample_excel_file, output_directory)
    
    # Check if Markdown files are created
    expected_files = [
        output_directory / "test_excel_Sheet1.md",
        output_directory / "test_excel_Sheet2.md"
    ]
    
    for file in expected_files:
        assert file.exists()
    
    # Verify content of Sheet1's Markdown file
    sheet1_md = output_directory / "test_excel_Sheet1.md"
    with open(sheet1_md, 'r', encoding='utf-8') as f:
        content = f.read()
    
    expected_content_sheet1 = (
        "| Name | Age | City |\n"
        "| --- | --- | --- |\n"
        "| Alice | 30 | New York |\n"
        "| Bob | 25 | Los Angeles |\n"
    )
    assert content == expected_content_sheet1
    
    # Since Sheet2 lacks a fully populated header row, it should prompt for input
    # However, since we mocked 'input' without side effects, it may result in incomplete processing
    # Adjust the test accordingly or split into separate tests

def test_excel_to_markdown_manual_input(sample_excel_file, mocker):
    # Mock user input for Sheet2
    inputs = iter(['3', 'A:C'])  # Header row 3, columns A to C
    mocker.patch('builtins.input', lambda _: next(inputs))
    
    # Process Sheet2 manually
    markdown = excel_to_markdown(sample_excel_file, 'Sheet2')
    
    expected_markdown = (
        "| Product | Price | Stock |\n"
        "| --- | --- | --- |\n"
        "| Item1 | 99.99 | 50 |\n"
    )
    assert markdown == expected_markdown

def test_create_output_filename():
    # Test the utility function for creating output filenames
    from excel_to_markdown.utils import sanitize_sheet_name, create_output_filename
    
    input_file = Path("/path/to/report.xlsx")
    sheet_name = "Sales Data"
    output_dir = Path("/path/to/output")
    
    expected_filename = output_dir / "report_Sales_Data.md"
    result = create_output_filename(input_file, sheet_name, output_dir)
    assert result == expected_filename

def test_main_function(sample_excel_file, output_directory, mocker):
    # Mock command-line arguments and user inputs
    mocker.patch('sys.argv', ['excel_to_markdown.main', str(sample_excel_file.parent), str(output_directory)])
    mocker.patch('builtins.input', side_effect=['3', 'A:C'])  # For Sheet2
    
    # Run the main function
    with mock.patch('excel_to_markdown.main.process_file') as mock_process_file:
        from excel_to_markdown.main import main
        main()
    
    # Ensure process_file was called for the Excel file
    mock_process_file.assert_called_once_with(sample_excel_file, output_directory)

