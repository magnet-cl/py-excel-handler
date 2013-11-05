from excel_handler import ExcelHandler
from excel_handler import fields

import unittest
import datetime


class MyExcelHandler(ExcelHandler):
    CHOICES = ((1, 'one'), (2, 'two'), (3, 'three'), (4, 'four'),
               (5, 'five'), (6, 'six'), (7, 'seven'), (8, 'eight'))
    first = fields.IntegerField(
        col=0, default=100, verbose_name="First"
    )
    second = fields.IntegerField(
        col=1, choices=CHOICES, default=3, verbose_name="Second"
    )
    third = fields.CharField(
        col=2, default="hello", verbose_name="Third"
    )
    fourth = fields.CharField(
        col=3, width=6, verbose_name="Fourth"
    )
    date_time = fields.DateTimeField(
        col=4, default=datetime.datetime.now, verbose_name="Date Time"
    )
    date = fields.DateField(
        col=5, default=datetime.date.today, verbose_name="Date"
    )
    boolean = fields.BooleanField(
        col=6, default=False, verbose_name="Boolean"
    )
    empty_last_fields = fields.CharField(
        col=7, default="", verbose_name="Empty"
    )


class InheritedExcelHandler(MyExcelHandler):
    pass


class TestChangeSheetCase(unittest.TestCase):

    def test_change_sheet_by_index(self):
        eh = MyExcelHandler(path='test/test.xls', mode='r')
        eh.set_sheet(0)

        data = eh.read()
        self.assertEqual(len(data), 3)

    def test_change_sheet_by_name(self):
        eh = MyExcelHandler(path='test/test.xls', mode='r')
        eh.set_sheet_by_name('Sheet2')

        data = eh.read()
        self.assertEqual(len(data), 2)


class TestExcelHandlerCase(unittest.TestCase):

    def test_read_rows(self):

        excel_file = open('test/test.xls', 'r')
        eh = ExcelHandler(excel_file=excel_file)

        column_structure = {
            'first': 0,
            'second': 1,
            'third': 2,
            'fourth': 3
        }

        data = eh.read_rows(column_structure)

        self.assertEqual(len(data), 3)
        self.assertEqual(len(data[0]), 4)
        self.assertEqual(len(data[1]), 4)

        first_row = eh.read_rows(column_structure, max_rows=1)

        self.assertEqual(len(first_row), 1)
        self.assertEqual(len(first_row[0]), 4)

        second_row = eh.read_rows(column_structure, starting_row=1)

        self.assertEqual(len(second_row), 2)
        self.assertEqual(len(second_row[0]), 4)

        for key, value in data[0].items():
            self.assertEqual(first_row[0][key], value)

        for key, value in data[1].items():
            self.assertEqual(second_row[0][key], value)

    def test_write_rows(self):

        rows = [[0, 1, 2], [3, 4, 5], [6, 7, 8]]

        workbook = ExcelHandler(path='test/test_write.xlsx', mode='w')
        workbook.add_sheet(name='Attendance')
        workbook.write_rows(rows)
        workbook.save()


class TestCustomExcelHandler(unittest.TestCase):
    def setUp(self):
        super(TestCustomExcelHandler, self).setUp()
        self.excel_handler_cls = MyExcelHandler

    def test_read(self):
        eh = self.excel_handler_cls(path='test/test.xls', mode='r')

        data = eh.read()

        expected_data = [{
            "first": 1,
            "second": 2,
            "third": "3.0",
            "fourth": "4.0",
            "date_time": datetime.datetime(2012, 10, 1, 12, 30, 47),
            "date": datetime.date(2013, 10, 1),
            "boolean": True,
            "empty_last_fields": "",
        }, {
            "first": 5,
            "second": 6,
            "third": "7.0",
            "fourth": "8.0",
            "date_time": datetime.datetime(2012, 10, 1, 12, 33, 56),
            "date": datetime.date(2013, 10, 2),
            "boolean": False,
            "empty_last_fields": "",
        }, {
            "first": 100,
            "second": 3,
            "third": "hello",
            "fourth": "12.0",
            "date": datetime.date.today(),
            "boolean": False,
            "empty_last_fields": "",
        }]

        for i, obj in enumerate(expected_data):
            for k, expected_value in obj.items():
                read_value = data[i][k]
                self.assertEqual(read_value, expected_value)

    def test_write(self):
        eh = self.excel_handler_cls(path='test/test_out.xlsx', mode='w')
        eh.add_sheet(name='Data')

        data = [{
            'first': 1,
            'second': 1,
            'third': "a",
            'fourth': "a",
            "date_time": datetime.datetime(2012, 10, 1, 12, 30, 47),
            "date": datetime.date(2013, 10, 1),
            "boolean": False,
            "empty_last_fields": "",
        }, {
            'first': 2,
            'second': 2,
            'third': "b",
            'fourth': "c",
            "date_time": datetime.datetime(2012, 10, 1, 12, 30, 47),
            "date": datetime.date(2013, 10, 1),
            "boolean": True,
            "empty_last_fields": "",
        }, {
            'fourth': "c",
        }]
        eh.write(data, set_titles=True)
        eh.save()

        eh = self.excel_handler_cls(path='test/test_out.xlsx', mode='r')

        in_data = eh.read(skip_titles=True)
        self.assertEqual(len(in_data), 3)

        for i, obj in enumerate(data):
            for k, v in obj.items():
                self.assertEqual(in_data[i][k], v)

        # testing the choices valus
        choices = dict((x, y) for x, y in self.excel_handler_cls.CHOICES)

        self.assertEqual(
            eh.sheet.cell(colx=1, rowx=1).value,
            choices[data[0]['second']]
        )

        self.assertEqual(
            eh.sheet.cell(colx=1, rowx=2).value,
            choices[data[1]['second']]
        )

        # testing the default valus
        self.assertEqual(in_data[2]['first'], 100)
        self.assertEqual(in_data[2]['second'], 3)
        self.assertEqual(in_data[2]['third'], 'hello')
        self.assertEqual(in_data[2]['boolean'], False)


class TestExcelHandlerInheritance(TestCustomExcelHandler):
    def setUp(self):
        super(TestExcelHandlerInheritance, self).setUp()
        self.excel_handler_cls = InheritedExcelHandler

if __name__ == '__main__':
    unittest.main()
