"""
groundwork-spreadsheets
-----------------------

Groundwork patterns to read and write spreadsheet documents. Excel 2010 (xlsx, xlsm) is supported at the moment.
The full documentation is available at https://groundwork-spreadsheets.readthedocs.io/

For more information regarding groundwork, see `here <https://groundwork.readthedocs.io.>`_.

**ExcelValidationPattern**

*   Uses the library `openpyxl <https://openpyxl.readthedocs.io/en/default/>`_
*   Can read Excel 2010 files (xlsx, xlsm)
*   Configure your sheet using a json file
*   Auto detect columns by names
*   Layout can be

    *   column based: headers are in a single *row* and data is below
    *   row based: headers are in a single *column* and data is right of the headers

*   Define column types and verify cells against it

    *   Date
    *   Enums (e.g. only  the values yes and no are allowed)
    *   Floating point numbers (+minimum/maximum check)
    *   Integer numbers (+minimum/maximum check)
    *   String (+RegEx pattern check)

*   Get a dictionary as output with the following form ``row number -> header name -> cell value``

Here is how an example json config file looks like::

    {
        "sheet_config": "last",
        "orientation": "column_based",
        "headers_index_config": {
            "row_index": {
                "first": 1,
                "last": "automatic"
            },
            "column_index": {
                "first": "automatic",
                "last": "severalEmptyCells:3"
            }
        },
        "data_index_config": {
            "row_index": {
                "first": 2,
                "last": "automatic"
            },
            "column_index": {
                "first": "automatic",
                "last": "automatic"
            }
        },
        "data_type_config": [
            {
                "header": "hex number",
                "fail_on_type_error": true,
                "fail_on_empty_cell": false,
                "fail_on_header_not_found": true,
                "type": {
                    "base": "string",
                    "pattern": "^0x[A-F0-9]{6}$"
                }
            },
            {
                "header": "int number",
                "type": {
                    "base": "integer",
                    "minimum": 2
                }
            }
        ]
    }

"""
from setuptools import setup, find_packages
import re
import ast

_version_re = re.compile(r'__version__\s+=\s+(.*)')
with open('groundwork_spreadsheets/version.py', 'rb') as f:
    version = str(ast.literal_eval(_version_re.search(
        f.read().decode('utf-8')).group(1)))

setup(
    name='groundwork_spreadsheets',
    version=version,
    url='http://groundwork-spreadsheets.readthedocs.io',
    license='MIT license',
    author='team useblocks',
    author_email='groundwork@useblocks.com',
    description="Patterns for reading writing spreadsheet documents",
    long_description=__doc__,
    packages=find_packages(exclude=['examples', 'tests']),
    include_package_data=True,
    platforms='any',
    setup_requires=[],
    tests_require=[],
    install_requires=['groundwork>=0.1.10', 'openpyxl', 'jsonschema'],
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Environment :: Console',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
    ],
    entry_points={
    }
)
