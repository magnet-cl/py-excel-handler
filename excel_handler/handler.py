""" This document defines the excel_handler module """
import xlrd
import xlsxwriter
from fields import Field


class FieldNotFound(Exception):
    pass


class ExcelHandlerMetaClass(type):
    def __new__(cls, name, bases, attrs):
        fieldname_to_field = {}

        for base in bases[::-1]:
            if hasattr(base, 'fieldname_to_field'):
                fieldname_to_field.update(base.fieldname_to_field)

        for k, v in attrs.items():
            if isinstance(v, Field):
                field = attrs.pop(k)
                field.name = k
                if field.verbose_name == "":
                    field.verbose_name = name
                fieldname_to_field[k] = field

        attrs['fieldname_to_field'] = fieldname_to_field
        sup = super(ExcelHandlerMetaClass, cls)
        return sup.__new__(cls, name, bases, attrs)


class ExcelHandler():
    """ ExcelHandler is a class that is used to wrap common operations in
    excel files """

    __metaclass__ = ExcelHandlerMetaClass

    def __init__(self, path=None, excel_file=None, mode='r'):
        if path is None and excel_file is None:
            raise Exception("path or excel_file requried")
        if path is not None and excel_file is not None:
            raise Exception("Only specify path or excel_file, not both")
        if mode == 'r':
            if path:
                excel_file = open(path, mode)

            self.workbook = xlrd.open_workbook(file_contents=excel_file.read())
            self.sheet = self.workbook.sheet_by_index(0)
        else:
            self.path = path

            # xlwt woorkbook
            # self.workbook = xlwt.Workbook()

            self.workbook = xlsxwriter.Workbook(self.path)

        self.parser = None

    def add_sheet(self, name):
        # xlwt
        # self.sheet = self.workbook.add_sheet(name)

        self.sheet = self.workbook.add_worksheet(name)

    def set_sheet(self, sheet_index):
        """ sets the current sheet with the given sheet_index """
        self.sheet = self.workbook.sheet_by_index(sheet_index)

    def read_rows(self, column_structure, starting_row=0, max_rows=-1):
        """ Reads the current sheet from the starting row to the last row or up
        to a max of max_rows if greater than 0

        returns an array with the data

        """
        data = []
        row = starting_row

        while max_rows != 0:
            column_data = {}

            for column_name in column_structure:
                try:
                    value = self.sheet.cell(
                        colx=column_structure[column_name],
                        rowx=row
                    ).value
                    column_data[column_name] = value
                except:
                    return data

            row += 1
            max_rows -= 1

            data.append(column_data)

        return data

    def read(self, skip_titles=False):
        """
        Using the structure defined with the Field attributes, reads the excel
        and returns the data in an array of dicts
        """
        data = []
        row = 0
        if skip_titles:
            row = 1

        while True:
            column_data = {}
            data_read = False

            for field_name, field in self.fieldname_to_field.items():
                try:
                    value = self.sheet.cell(
                        colx=field.col,
                        rowx=row
                    ).value
                except:
                    if hasattr(field, 'default'):
                        column_data[field.name] = field.default
                else:
                    column_data[field.name] = field.cast(value, self.workbook)
                    data_read = True

            row += 1

            if not data_read:
                return data

            data.append(column_data)

        return data

    def save(self):
        """ Save document """

        # xlwt save
        # self.workbook.save(self.path)
        self.workbook.close()

    def write_rows(self, rows):
        """ Write rows in the current sheet """

        for x, row in enumerate(rows):
            for y, value in enumerate(row):
                self.sheet.write(x, y, value)

    def write(self, data, set_titles=False):
        row = 0

        # set titles
        if set_titles:
            for field_name, field in self.fieldname_to_field.items():
                self.sheet.write(0, field.col,  field.verbose_name)
            row = 1

        # set format
        for field_name, field in self.fieldname_to_field.items():
            field.set_format(self.workbook, self.sheet)

        for row_data in data:
            for field_name, value in row_data.items():
                try:
                    field = self.fieldname_to_field[field_name]
                except KeyError:
                    pass
                else:
                    field.write(self.workbook, self.sheet, row, value)
            row += 1
