import unittest
import pandas as pd
from excel_to_markdown.markdown_generator import dataframe_to_markdown


class TestMarkdownGenerator(unittest.TestCase):
    def test_dataframe_to_markdown_basic(self):
        # Test with a basic DataFrame
        data = {
            'Name': ['Alice', 'Bob'],
            'Age': [30, 25],
            'City': ['New York', 'Los Angeles']
        }
        df = pd.DataFrame(data)
        expected_markdown = (
            "| Name | Age | City |\n"
            "| --- | --- | --- |\n"
            "| Alice | 30 | New York |\n"
            "| Bob | 25 | Los Angeles |\n"
        )
        result = dataframe_to_markdown(df)
        self.assertEqual(result, expected_markdown)

    def test_dataframe_to_markdown_with_missing_values(self):
        # Test DataFrame with missing values
        data = {
            'Name': ['Alice', 'Bob', 'Charlie'],
            'Age': [30, None, 35],
            'City': ['New York', 'Los Angeles', None]
        }
        df = pd.DataFrame(data)
        expected_markdown = (
            "| Name | Age | City |\n"
            "| --- | --- | --- |\n"
            "| Alice | 30.0 | New York |\n"
            "| Bob |  | Los Angeles |\n"
            "| Charlie | 35.0 |  |\n"
        )
        result = dataframe_to_markdown(df)
        self.assertEqual(result, expected_markdown)

    def test_dataframe_to_markdown_empty_dataframe(self):
        # Test with an empty DataFrame
        df = pd.DataFrame()
        expected_markdown = ""
        result = dataframe_to_markdown(df)
        self.assertEqual(result, expected_markdown)

    def test_dataframe_to_markdown_no_columns(self):
        # Test DataFrame with no columns but with rows
        df = pd.DataFrame([[] for _ in range(3)])
        expected_markdown = ""
        result = dataframe_to_markdown(df)
        self.assertEqual(result, expected_markdown)

    def test_dataframe_to_markdown_single_row(self):
        # Test DataFrame with a single row
        data = {
            'Product': ['Laptop'],
            'Price': [999.99],
            'Stock': [50]
        }
        df = pd.DataFrame(data)
        expected_markdown = (
            "| Product | Price | Stock |\n"
            "| --- | --- | --- |\n"
            "| Laptop | 999.99 | 50 |\n"
        )
        result = dataframe_to_markdown(df)
        self.assertEqual(result, expected_markdown)

    def test_dataframe_to_markdown_single_column(self):
        # Test DataFrame with a single column
        data = {
            'Username': ['user1', 'user2', 'user3']
        }
        df = pd.DataFrame(data)
        expected_markdown = (
            "| Username |\n"
            "| --- |\n"
            "| user1 |\n"
            "| user2 |\n"
            "| user3 |\n"
        )
        result = dataframe_to_markdown(df)
        self.assertEqual(result, expected_markdown)

    def test_dataframe_to_markdown_numeric_data(self):
        # Test DataFrame with numeric data
        data = {
            'ID': [1, 2, 3],
            'Score': [85.5, 92.0, 78.25]
        }
        df = pd.DataFrame(data)
        expected_markdown = (
            "| ID | Score |\n"
            "| --- | --- |\n"
            "| 1.0 | 85.5 |\n"
            "| 2.0 | 92.0 |\n"
            "| 3.0 | 78.25 |\n"
        )
        result = dataframe_to_markdown(df)
        print(result)
        self.assertEqual(result, expected_markdown)


if __name__ == '__main__':
    unittest.main()
