from excel_handler import ExcelHandler
from excel_handler import fields

import unittest


class MyExcelHandler(ExcelHandler):
    first = fields.IntegerField(col=0)
    second = fields.IntegerField(col=1)
    third = fields.CharField(col=2)
    forth = fields.CharField(col=3)


class TestExcelHandlerCase(unittest.TestCase):

    def test_read_rows(self):

        excel_file = open('test/test.xls', 'r')
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

    def test_write_rows(self):

        rows = [[0, 1, 2], [3, 4, 5], [6, 7, 8]]

        workbook = ExcelHandler(path='test_write.xls', mode='w')
        workbook.add_sheet(name='Attendance')
        workbook.write_rows(rows)
        workbook.save()


class TestCustomExcelHandler(unittest.TestCase):
    def test_read(self):
        eh = MyExcelHandler(path='test/test.xls', mode='r')

        data = eh.read()

        self.assertEqual(len(data), 2)
        self.assertEqual(len(data[0]), 4)
        self.assertEqual(len(data[1]), 4)

    def test_write(self):
        eh = MyExcelHandler(path='test/test_out.xls', mode='w')
        eh.add_sheet(name='Data')

        data = [{
            'first': 1,
            'second': 1,
            'third': "a",
            'forth': "a",
        }, {
            'first': 2,
            'second': 2,
            'third': "b",
            'forth': "c",
        }]
        eh.write(data)
        eh.save()

        eh = MyExcelHandler(path='test/test_out.xls', mode='r')

        in_data = eh.read()
        self.assertEqual(len(data), 2)

        for i, obj in enumerate(data):
            for k, v in obj.items():
                self.assertEqual(in_data[i][k], v)


if __name__ == '__main__':
    unittest.main()
