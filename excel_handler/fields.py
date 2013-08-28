class Field(object):
    def __init__(self, col):
        self.col = col

    def cast(self, value):
        return value


class IntegerField(Field):
    def __init__(self, col):
        super(IntegerField, self).__init__(col)
        self.cast = int


class CharField(Field):
    def __init__(self, col):
        super(CharField, self).__init__(col)
        self.cast = str
