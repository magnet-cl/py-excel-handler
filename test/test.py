from excel_handler import ExcelHandler

import unittest


class TestSequenceFunctions(unittest.TestCase):

    def test_read_rows(self):

        excel_file = open('test.xls', 'r')
        eh = ExcelHandler(excel_file=excel_file)

        column_structure = {
            'first': 0,
            'second': 1,
            'third': 2,
            'forth': 3
        }

        data = eh.read_rows(column_structure)

        self.assertEqual(len(data), 2)
        self.assertEqual(len(data[0]), 4)
        self.assertEqual(len(data[1]), 4)

        first_row = eh.read_rows(column_structure, max_rows=1)

        self.assertEqual(len(first_row), 1)
        self.assertEqual(len(first_row[0]), 4)

        second_row = eh.read_rows(column_structure, starting_row=1)

        self.assertEqual(len(second_row), 1)
        self.assertEqual(len(second_row[0]), 4)

        for key, value in data[0].items():
            self.assertEqual(first_row[0][key], value)

        for key, value in data[1].items():
            self.assertEqual(second_row[0][key], value)

if __name__ == '__main__':
    unittest.main()
