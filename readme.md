# Excel to Markdown Converter

This Python script converts Excel files to Markdown tables.

## Installation

This project uses [Poetry](https://python-poetry.org/) for dependency management. To set up the project, follow these steps:

1. Ensure you have Poetry installed on your system.
2. Clone this repository.
3. Navigate to the project directory.
4. Run `poetry install` to install the dependencies.

## Dependencies

- Python ^3.12
- pandas ^2.2.2

## Usage

To use the script, run the following command:

```
python script.py <excel_file> [sheet_name]
```

- `<excel_file>`: Path to the Excel file you want to convert (required)
- `[sheet_name]`: Name or index of the sheet to convert (optional, defaults to the first sheet)

The script will generate a Markdown file with the same name as the input Excel file, but with a `.md` extension.

## How it works

1. The script reads the specified Excel file using pandas.
2. It converts the data into a Markdown table format.
3. The resulting Markdown is written to a new file.

## Example

```
poetry run excel-to-markdown data/input data/output
```

This command will create a file named `example.md` containing the Markdown table representation of the data in `Sheet1` of `example.xlsx`.

## Error Handling

The script includes basic error handling:
- It will display a usage message if no Excel file is specified.
- It will print an error message if any exception occurs during the conversion process.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.