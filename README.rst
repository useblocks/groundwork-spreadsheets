groundwork-spreadsheets
-----------------------
Groundwork patterns to read and write spreadsheet documents. Excel 2010 is supported at the moment.

For more information regarding groundwork, see `here <https://groundwork.readthedocs.io.>`_.

*   **GwSpreadsheetsPattern**

    *   Basic read and write operations
    *   Uses the library `openpyxl <https://openpyxl.readthedocs.io/en/default/>`_
    *   Can read and write Excel 2010 files (xlsx/xlsm)

*   **GwSpreadsheetsColumnPattern**

    *   Based on GwSpreadsheetsPattern
    *   Configure sheet schemas using a json file
    *   Auto detect columns by names
    *   Define column types
    *   Verify columns against the defined schema

        *   Integer numbers
        *   Floating point numbers
        *   Text, RegEx patterns
        *   Enums (e.g. only  the values yes and no are allowed)

The full documentation is available at https://groundwork-spreadsheets.readthedocs.io/
