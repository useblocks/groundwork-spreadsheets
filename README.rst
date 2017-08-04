groundwork-utilities
----------
Groundwork patterns to read and write Excel documents.
For more information regarding groundwork, see `here <https://groundwork.readthedocs.io.>`_.

*   **GwSpreadsheetPattern**

    *   Basic read and write operations
    *   Uses the library `openpyxl <https://openpyxl.readthedocs.io/en/default/>`_.
        Can read and write Excel 2010 files (xlsx/xlsm)
        
*   **GwSpreadsheetColumnPattern**

    *   Based on GwSpreadsheetPattern
    *   Configure your Excel sheets using a json file
    *   Auto detect columns by names
    *   Define column types
    *   Verify columns against the defined schema
