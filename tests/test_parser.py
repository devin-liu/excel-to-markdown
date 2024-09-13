import unittest
import pandas as pd
from excel_to_markdown.parser import parse_columns
from excel_to_markdown.utils import column_letter_to_index

class TestParser(unittest.TestCase):
    def setUp(self):
        # Create a sample DataFrame with 10 columns labeled A to J
        columns = [chr(i) for i in range(65, 75)]  # ['A', 'B', ..., 'J']
        data = [[f'Data{i}{j}' for j in range(10)] for i in range(5)]
        self.df = pd.DataFrame(data, columns=columns)
    
    def test_parse_columns_letter_range(self):
        # Test parsing a letter-based column range
        cols_input = "A:D"
        expected = ['A', 'B', 'C', 'D']
        result = parse_columns(cols_input, self.df)
        self.assertEqual(result, expected)
    
    def test_parse_columns_number_range(self):
        # Test parsing a number-based column range
        cols_input = "1-4"
        expected = ['A', 'B', 'C', 'D']
        result = parse_columns(cols_input, self.df)
        self.assertEqual(result, expected)
    
    def test_parse_columns_single_letter(self):
        # Test parsing a single letter
        cols_input = "E"
        expected = ['E']
        result = parse_columns(cols_input, self.df)
        self.assertEqual(result, expected)
    
    def test_parse_columns_single_number(self):
        # Test parsing a single number
        cols_input = "5"
        expected = ['E']
        result = parse_columns(cols_input, self.df)
        self.assertEqual(result, expected)
    
    def test_parse_columns_mixed_letters(self):
        # Test parsing a mixed letter-based range
        cols_input = "G:I"
        expected = ['G', 'H', 'I']
        result = parse_columns(cols_input, self.df)
        self.assertEqual(result, expected)
    
    def test_parse_columns_invalid_letters(self):
        # Test parsing an invalid letter-based range
        cols_input = "K:M"  # Columns K to M do not exist
        with self.assertRaises(IndexError):
            parse_columns(cols_input, self.df)
    
    def test_parse_columns_invalid_numbers(self):
        # Test parsing an invalid number-based range
        cols_input = "11-13"  # Columns 11 to 13 do not exist
        with self.assertRaises(IndexError):
            parse_columns(cols_input, self.df)
    
    def test_parse_columns_start_after_end_letters(self):
        # Test when start letter comes after end letter
        cols_input = "D:A"
        with self.assertRaises(ValueError):
            parse_columns(cols_input, self.df)
    
    def test_parse_columns_start_after_end_numbers(self):
        # Test when start number comes after end number
        cols_input = "4-1"
        with self.assertRaises(ValueError):
            parse_columns(cols_input, self.df)
    
    def test_parse_columns_invalid_format(self):
        # Test invalid format input
        cols_input = "A-4"
        with self.assertRaises(ValueError):
            parse_columns(cols_input, self.df)
    
    def test_parse_columns_non_alphanumeric(self):
        # Test input with non-alphanumeric characters
        cols_input = "A$-D#"
        with self.assertRaises(ValueError):
            parse_columns(cols_input, self.df)

if __name__ == '__main__':
    unittest.main()
