""" This document defines the excel_handler module """
import xlrd
import xlwt


class ExcelHandler():
    """ ExcelHandler is a class that is used to wrap common operations in
    excel files """

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
            self.workbook = xlwt.Workbook()

    def add_sheet(self, name):
        self.sheet = self.workbook.add_sheet(name)

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

    def save(self):
        """ Save document """

        self.workbook.save(self.path)

    def write_rows(self, rows):
        """ Write rows in the current sheet """

        for x, row in enumerate(rows):
            for y, value in enumerate(row):
                self.sheet.write(x, y, value)
