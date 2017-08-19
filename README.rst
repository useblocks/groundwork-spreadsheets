groundwork-spreadsheets
=======================

Groundwork patterns to read and write spreadsheet documents. Excel 2010 (xlsx, xlsm) is supported at the moment.

For more information regarding groundwork, see `here <https://groundwork.readthedocs.io.>`_.

*   **ExcelValidationPattern**

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

The full documentation is available at https://groundwork-spreadsheets.readthedocs.io/
