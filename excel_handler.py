""" This document defines the excel_handler module """
import xlrd


class ExcelHandler():
    """ ExcelHandler is a class that is used to wrap common operations in
    excel files """

    def __init__(self, path=None, excel_file=None):
        if path is None and excel_file is None:
            raise Exception("path or excel_file requried")
        if path is not None and excel_file is not None:
            raise Exception("Only specify path or excel_file, not both")

        if path:
            excel_file = open(path, 'r')

        self.workbook = xlrd.open_workbook(file_contents=excel_file.read())
        self.sheet = self.workbook.sheet_by_index(0)

    def set_sheet(self, sheet_index):
        self.sheet = self.workbook.sheet_by_index(sheet_index)

    def read_columns(self, column_structure, starting_row=0, max_rows=-1):
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
