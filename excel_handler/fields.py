import xlrd
import datetime


class Field(object):
    def __init__(self, col, **kwargs):
        self.col = col

        if 'choices' in kwargs:
            self.choices_inv = dict((y, x) for x, y in kwargs['choices'])
            self.choices = dict((x, y) for x, y in kwargs['choices'])
        else:
            self.choices = None

        if 'default' in kwargs:
            self.default = kwargs['default']
        else:
            self.default = None

        if 'width' in kwargs:
            self.width = (kwargs['width'] + 1) * 256
        else:
            self.width = None

        if 'verbose_name' in kwargs:
            self.verbose_name = kwargs['verbose_name']
        else:
            self.verbose_name = ""

    def cast(self, value, book):
        if value == '' and self.default:
            return self.default

        if self.choices:
            try:
                return self.cast_method(self.choices_inv[value])
            except ValueError, error:
                error.args += (self.name,)
                raise ValueError(error)
            except KeyError, error:
                error.args += (self.name,)
                raise KeyError(error)

        try:
            return self.cast_method(value)
        except ValueError, error:
            error.args += (self.name,)
            raise ValueError(error)

    def write(self, workbook, sheet, row, value):
        if self.choices:
            value = self.choices[value]

        sheet.write(row, self.col,  value)

        if self.width:
            # xlwt format size
            # sheet.col(self.col).width = self.width
            sheet.set_column(self.col, self.col, self.width)


class IntegerField(Field):
    def __init__(self, col, *args, **kwargs):
        super(IntegerField, self).__init__(col, *args, **kwargs)
        self.cast_method = int


class BooleanField(Field):
    def __init__(self, col, *args, **kwargs):
        super(BooleanField, self).__init__(col, *args, **kwargs)
        self.cast_method = bool


class CharField(Field):
    def __init__(self, col, *args, **kwargs):
        super(CharField, self).__init__(col, *args, **kwargs)
        self.cast_method = str


class DateTimeField(Field):
    def cast(self, value, workbook):
        if value == '' and self.default:
            if callable(self.default):
                return self.default()
            return self.default

        date_tuple = xlrd.xldate_as_tuple(value, datemode=workbook.datemode)
        return datetime.datetime(*date_tuple[:6])

    def write(self, workbook, sheet, row, value):
        # xlwt date format
        # xf = xlswriter.easyxf(num_format_str='MM/DD/YYYY HH:MM:SS')
        # sheet.write(row, self.col,  value, xf)

        date_format = workbook.add_format(
            {'num_format': 'MM/DD/YYYY HH:MM:SS'}
        )
        sheet.write(row, self.col,  value, date_format)
        # xlwt format size
        # sheet.col(self.col).width = 4864  # 19 * 256
        sheet.set_column(self.col, self.col, 15)


class DateField(DateTimeField):
    def cast(self, value, workbook):
        if value == '' and self.default:
            if callable(self.default):
                return self.default()
            return self.default

        date_tuple = xlrd.xldate_as_tuple(value, datemode=workbook.datemode)
        return datetime.date(*date_tuple[:3])

    def write(self, workbook, sheet, row, value):
        # xlwt date format
        # xf = xlswriter.easyxf(num_format_str='MM/DD/YYYY')
        # sheet.write(row, self.col,  value, xf)
        date_format = workbook.add_format(
            {'num_format': 'MM/DD/YYYY'}
        )
        sheet.write(row, self.col,  value, date_format)

        # xlwt format size
        # sheet.col(self.col).width = 2560  # 10 * 256
        sheet.set_column(self.col, self.col, 15)
