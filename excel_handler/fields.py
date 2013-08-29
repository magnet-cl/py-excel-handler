import xlrd
import xlwt
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

    def cast(self, value, book):
        if value == '' and self.default:
            return self.default

        if self.choices:
            try:
                return self.cast_method(self.choices_inv[value])
            except ValueError, error:
                error.args += (self.name,)
                raise ValueError(error)

        try:
            return self.cast_method(value)
        except ValueError, error:
            error.args += (self.name,)
            raise ValueError(error)

    def write(self, sheet, row, value):
        if self.choices:
            value = self.choices[value]

        sheet.write(row, self.col,  value)

        if self.width:
            sheet.col(self.col).width = self.width


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

    def write(self, sheet, row, value):
        xf = xlwt.easyxf(num_format_str='MM/DD/YYYY HH:MM:SS')
        sheet.write(row, self.col,  value, xf)
        sheet.col(self.col).width = 4864  # 19 * 256


class DateField(DateTimeField):
    def cast(self, value, workbook):
        if value == '' and self.default:
            if callable(self.default):
                return self.default()
            return self.default

        date_tuple = xlrd.xldate_as_tuple(value, datemode=workbook.datemode)
        return datetime.date(*date_tuple[:3])

    def write(self, sheet, row, value):
        xf = xlwt.easyxf(num_format_str='MM/DD/YYYY')
        sheet.write(row, self.col,  value, xf)
        sheet.col(self.col).width = 2560  # 10 * 256
