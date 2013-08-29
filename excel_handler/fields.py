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

    def cast(self, value, book):
        if value == '' and self.default:
            return self.default

        if self.choices:
            return self.cast_method(self.choices_inv[value])

        return self.cast_method(value)

    def decode(self, value):
        if self.choices:
            return self.choices[value]

        return value


class IntegerField(Field):
    def __init__(self, col, *args, **kwargs):
        super(IntegerField, self).__init__(col, *args, **kwargs)
        self.cast_method = int


class CharField(Field):
    def __init__(self, col, *args, **kwargs):
        super(CharField, self).__init__(col, *args, **kwargs)
        self.cast_method = str


class DateField(Field):
    def __init__(self, col, *args, **kwargs):
        super(DateField, self).__init__(col, *args, **kwargs)

    def cast(self, value, workbook):
        if value == '' and self.default:
            if callable(self.default):
                return self.default()
            return self.default

        date_tuple = xlrd.xldate_as_tuple(value, datemode=workbook.datemode)
        return datetime.date(*date_tuple[:3])
