groundwork-spreadsheet
----------------------
Groundwork patterns to read and write Excel documents.
For more information regarding groundwork, see `here <https://groundwork.readthedocs.io.>`_.

*   **GwSpreadsheetPattern**

    *   Basic read and write operations
    *   Uses the library `openpyxl <https://openpyxl.readthedocs.io/en/default/>`_
    *   Can read and write Excel 2010 files (xlsx/xlsm)
        
*   **GwSpreadsheetColumnPattern**

    *   Based on GwSpreadsheetPattern
    *   Configure Excel sheet schemas using a json file
    *   Auto detect columns by names
    *   Define column types
    *   Verify columns against the defined schema
    
        *   Integer numbers
        *   Floating point numbers
        *   Text, RegEx patterns
        *   Enums (e.g. only  the values yes and no are allowed)

The full documentation is available at https://groundwork-spreadsheet.readthedocs.io/
