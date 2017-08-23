groundwork-spreadsheets
=======================

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

*   Define column types and verify cell values against them

    *   Date
    *   Enums (e.g. only  the values 'yes' and 'no' are allowed)
    *   Floating point numbers with optional min/max check
    *   Integer numbers with optional min/max check
    *   String with optional regular expression pattern check

*   Exclude data row/columns based on filter criteria
*   Output is a dictionary of the following form ``row or column number`` -> ``header name`` -> ``cell value``
*   Extensive logging of problems

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
