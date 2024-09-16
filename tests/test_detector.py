import unittest
import pandas as pd
from excel_to_markdown.detector import detect_table_start


class TestDetector(unittest.TestCase):
    def test_detect_table_start_success(self):
        data = {
            'A': [None, 'Header1', 'Data1'],
            'B': [None, 'Header2', 'Data2'],
            'C': [None, 'Header3', 'Data3']
        }
        df = pd.DataFrame(data)
        expected_row = 1
        result = detect_table_start(df)
        self.assertEqual(result, expected_row)

    def test_detect_table_start_failure(self):
        data = {
            'A': [None, None, None],
            'B': [None, None, None],
            'C': [None, None, None],
        }
        df = pd.DataFrame(data)
        result = detect_table_start(df)
        self.assertIsNone(result)

    def test_detect_table_start_partial_fill(self):
        data = {
            'A': [None, None, None],
            'B': [None, 'Header1', 'Row1'],
            'C': [None, 'Header2', 'Row2'],
            'D': [None, 'Header3', 'Row3'],
            'E': [None, None, None]
        }
        expected_row = 1
        df = pd.DataFrame(data)
        result = detect_table_start(df)
        self.assertEqual(result, expected_row)

    def test_detect_table_start_with_none_first_row(self):
        data = {
            'A': [None, 'Name', 'Alice', 'Bob'],
            'B': [None, 'Age', 30, 25],
            'C': [None, 'City', 'New York', 'Los Angeles']
        }
        df = pd.DataFrame(data)
        expected_row = 1
        result = detect_table_start(df)
        self.assertEqual(result, expected_row)

    def test_detect_table_start_with_second_row_start(self):
        data = {
            'A': [None, None, 'Name', 'Alice', 'Bob'],
            'B': [None, None, 'Age', 30, 25],
            'C': [None, None, 'City', 'New York', 'Los Angeles']
        }
        df = pd.DataFrame(data)
        expected_row = 2
        result = detect_table_start(df)
        self.assertEqual(result, expected_row)

# test for this case

#    # Create sample data for two sheets
#         sheet1_data = {
#             'A': [None, 'Name', 'Alice', 'Bob'],
#             'B': [None, 'Age', 30, 25],
#             'C': [None, 'City', 'New York', 'Los Angeles']
#         }

# finds Name, Age, City


if __name__ == '__main__':
    unittest.main()
