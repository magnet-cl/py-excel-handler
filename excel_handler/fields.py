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

        if 'width' in kwargs:
            # xlwt format size
            # self.width = (kwargs['width'] + 1) * 256
            self.width = kwargs['width']
        else:
            self.width = None

        if 'verbose_name' in kwargs:
            self.verbose_name = kwargs['verbose_name']
        else:
            self.verbose_name = ""

        self.format = None

    def cast(self, value, book, row_data):
        if isinstance(value, basestring):
            if value.strip() == '' and hasattr(self, 'default'):
                return self.default

        if self.choices:
            try:
                return self.cast_method(self.choices_inv[value])
            except:
                try:
                    return self.choices_inv[value]
                except ValueError, error:
                    error.args += (self.name,)
                    raise ValueError(error)
                except KeyError, error:
                    error.args += (self.name,)
                    raise KeyError(error)

        try:
            return self.cast_method(value)
        except ValueError, error:
            error.args += (self.name, value)
            raise ValueError(error)

    def prepare_read(self):
        pass

    def prepare_write(self):
        pass

    def write(self, workbook, sheet, row, value):
        if self.choices:
            try:
                value = self.choices[value]
            except KeyError, error:
                if value is not None:
                    raise KeyError(error)

        if hasattr(value, 'translate'):
            value = unicode(value)

        sheet.write(row, self.col,  value)

    def set_column_format(self, handler):
        """
        Sets the format of the column this field, by setting the width
        """
        if self.width:
            handler.sheet.set_column(
                self.col,
                self.col,
                self.width,
                cell_format=self.format
            )

    def set_format(self, workbook, sheet):
        if self.width:
            sheet.set_column(self.col, self.col, self.width)


class BooleanField(Field):
    def __init__(self, col, *args, **kwargs):
        super(BooleanField, self).__init__(col, *args, **kwargs)
        self.cast_method = bool


class CharField(Field):
    def __init__(self, col, *args, **kwargs):
        super(CharField, self).__init__(col, *args, **kwargs)
        self.cast_method = unicode


class DateTimeField(Field):
    def __init__(self, *args, **kwargs):
        self.tzinfo = kwargs.pop('tzinfo', None)

        super(DateTimeField, self).__init__(*args, **kwargs)

    def cast(self, value, workbook, row_data):
        if value == '' and hasattr(self, 'default'):
            if callable(self.default):
                return self.default()
            return self.default

        date_tuple = xlrd.xldate_as_tuple(value, datemode=workbook.datemode)
        date = datetime.datetime(*date_tuple[:6])
        return date.replace(tzinfo=self.tzinfo)

    def write(self, workbook, sheet, row, value):
        if value:
            value = value.replace(tzinfo=None)
        sheet.write(row, self.col,  value)

    def set_column_format(self, handler):
        """
        DateTimeField Sets the format of the column this field is in using the
        handler's date format
        """
        handler.sheet.set_column(
            self.col,
            self.col,
            18,
            cell_format=handler.datetime_format
        )

    def set_format(self, workbook, sheet):
        date_format = workbook.add_format(
            {'num_format': 'MM/DD/YYYY HH:MM:SS'}
        )
        sheet.set_column(self.col, self.col, 18, date_format)


class TimeField(Field):
    def __init__(self, *args, **kwargs):
        self.tzinfo = kwargs.pop('tzinfo', None)

        super(TimeField, self).__init__(*args, **kwargs)

    def cast(self, value, workbook, row_data):
        if value == '' and hasattr(self, 'default'):
            if callable(self.default):
                return self.default()
            return self.default

        value = value - int(value)
        date_tuple = xlrd.xldate_as_tuple(value, datemode=workbook.datemode)
        time = datetime.time(*date_tuple[3:])
        return time.replace(tzinfo=self.tzinfo)

    def write(self, workbook, sheet, row, value):
        if value:
            # xslx writer does not handle timezone aware values
            value = value.replace(tzinfo=None)
        sheet.write(row, self.col, value)

    def set_column_format(self, handler):
        """
        TimeField Sets the format of the column this field is in using the
        handler's time format
        """
        handler.sheet.set_column(
            self.col,
            self.col,
            18,
            cell_format=handler.time_format
        )

    def set_format(self, workbook, sheet):
        date_format = workbook.add_format(
            {'num_format': 'HH:MM:SS'}
        )
        sheet.set_column(self.col, self.col, 18, date_format)


class DateField(DateTimeField):
    def cast(self, value, workbook, row_data):
        if value == '' and hasattr(self, 'default'):
            if callable(self.default):
                return self.default()
            return self.default

        try:
            # try to parse date in excel format
            date_tuple = xlrd.xldate_as_tuple(
                value,
                datemode=workbook.datemode
            )
        except ValueError:
            # try to parse dates like dd/mm/YYYY
            if len(value) <= 10:

                if '/' in value:
                    date_tuple = value.split('/')
                elif '-' in value:
                    date_tuple = value.split('-')

                date_tuple = [int(x) for x in date_tuple]

                if date_tuple[2] > 1000:
                    date_tuple.reverse()
            else:
                raise

        return datetime.date(*date_tuple[:3])

    def write(self, workbook, sheet, row, value):
        sheet.write(row, self.col,  value)

    def set_column_format(self, handler):
        """
        Sets the format of the column this field is in using the
        handler's time format
        """
        handler.sheet.set_column(
            self.col,
            self.col,
            18,
            cell_format=handler.date_format
        )

    def set_format(self, workbook, sheet):
        date_format = workbook.add_format(
            {'num_format': 'MM/DD/YYYY'}
        )
        sheet.set_column(self.col, self.col, 15, date_format)


class DjangoModelField(Field):
    """
    This field translates excel values to django models and viceversa
    """

    def __init__(self, col, model, lookup='pk', *args, **kwargs):
        super(DjangoModelField, self).__init__(col, *args, **kwargs)

        self.lookup = lookup

        self.model = model

    def cast(self, value, workbook, row_data):
        return self.model.objects.get(**{self.lookup: value})

    def write(self, workbook, sheet, row, value):
        value = getattr(value, self.lookup)
        super(DjangoModelField, self).write(self, workbook, sheet, row, value)


class ForeignKeyField(Field):
    """
    This field translates excel values to django foreign keys
    """
    def __init__(self, col, model, lookup='pk', default_on_lookup_fail=False,
                 case_insensitive=False, on_lookup_fail=None, *args, **kwargs):

        super(ForeignKeyField, self).__init__(col, *args, **kwargs)

        self.lookup = lookup
        self.model = model
        self.default_on_lookup_fail = default_on_lookup_fail
        self.case_insensitive = case_insensitive
        self.on_lookup_fail = on_lookup_fail

    def cast(self, value, workbook, row_data):
        if value == '' and hasattr(self, 'default'):
            return self.default

        value = self.lookup_type(value)
        if self.case_insensitive:
            value = value.lower()

        try:
            return self.lookup_to_pk[value]
        except KeyError:
            msg = ("%s matching query does not exist. "
                   "Lookup parameters were %s" %
                   (self.model._meta.object_name, {self.lookup: value}))
            if self.on_lookup_fail:
                return self.on_lookup_fail(row_data, value)

            if self.default_on_lookup_fail:
                return self.default

            raise self.model.DoesNotExist(msg)

    def write(self, workbook, sheet, row, value):
        if self.lookup != 'pk' and self.lookup != 'id' and value is not None:
            value = self.pk_to_lookup[value]

        super(ForeignKeyField, self).write(workbook, sheet, row, value)

    def prepare_read(self):
        self.objects = self.model.objects.all().values_list('id', self.lookup)

        try:
            self.lookup_type = type(self.objects[0][1])
        except:
            self.lookup_type = str

        self.pk_to_lookup = dict(self.objects)

        if self.case_insensitive:
            self.lookup_to_pk = dict((y.lower(), x) for x, y in self.objects)
        else:
            self.lookup_to_pk = dict((y, x) for x, y in self.objects)

    def prepare_write(self):
        self.prepare_read()


class IntegerField(Field):
    def __init__(self, col, *args, **kwargs):
        super(IntegerField, self).__init__(col, *args, **kwargs)
        self.cast_method = int


class FloatField(Field):
    def __init__(self, col, *args, **kwargs):
        super(FloatField, self).__init__(col, *args, **kwargs)
        self.cast_method = float
