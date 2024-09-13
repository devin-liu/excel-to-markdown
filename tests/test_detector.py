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
            'A': [None, 'Header1', 'Data1'],
            'B': [None, None, 'Data2'],
            'C': [None, 'Header3', 'Data3']
        }
        df = pd.DataFrame(data)
        result = detect_table_start(df)
        self.assertIsNone(result)

    def test_detect_table_start_partial_fill(self):
        data = {
            'A': [None, 'Header1', 'Data1'],
            'B': [None, 'Header2', None],
            'C': [None, 'Header3', 'Data3']
        }
        df = pd.DataFrame(data)
        result = detect_table_start(df)
        self.assertIsNone(result)

if __name__ == '__main__':
    unittest.main()
