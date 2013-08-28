from distutils.core import setup
setup(
    name='py-excel-handler',
    version='0.1',
    py_modules=['excel_handler', 'fit_sheet_wrapper'],
    install_requires=[
        'xlutils >= 1.6.0',
    ],
)
