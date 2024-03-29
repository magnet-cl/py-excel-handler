try:
    from setuptools import setup
except ImportError:
    from ez_setup import use_setuptools

    use_setuptools()
    from setuptools import setup

setup(
    name="py-excel-handler",
    version="0.5.0",
    description="A set of tools over xlutils to read and write excel files",
    author="Ignacio Munizaga",
    author_email="muni@magnet.cl",
    url="http://github.com/magnet-cl/py-excel-handler/",
    packages=[
        "excel_handler",
    ],
    requires=[
        # 'mimeparse',
        "xlutils(>=1.6.0)",
        "XlsxWriter(>=0.5.7)",
        "future(>=0.18.2)",
        "openpyxl(==3.0.9)",
    ],
    install_requires=[
        "xlutils >= 1.6.0",
        "XlsxWriter >= 0.5.7",
        "future >= 0.18.2",
        "openpyxl == 3.0.9",
    ],
    package_data={},
    zip_safe=False,
)
