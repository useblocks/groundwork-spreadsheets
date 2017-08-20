Configuration
=============

This is the base structure of the configuration .json file::

    {
        "sheet_config": "active",
        "orientation": "column_based",
        "headers_index_config": {
            "row_index": {
                "first": "automatic",
                "last": "automatic"
            },
            "column_index": {
                "first": "automatic",
                "last": "automatic"
            }
        },
        "data_index_config": {
            "row_index": {
                "first": "automatic",
                "last": "automatic"
            },
            "column_index": {
                "first": "automatic",
                "last": "automatic"
            }
        },
        "data_type_config": [
            {
                "header": "Text",
                "type": {
                    "base": "string"
                }
            }
        ]
    }

sheet_config
------------

Possible values are:

=================   ======= =============================   =======
Value               Type    Example                         Meaning
=================   ======= =============================   =======
active              string  "sheet_config": "active"        Chooses the active worksheet, that is the one that was
                                                            active when last saving the workbook
first               string  "sheet_config": "first"         Chooses the first worksheet
last                string  "sheet_config": "last"          Chooses the last worksheet
name:<sheet_name>   string  "sheet_config": "name:sheet2"   Chooses the worksheet with the name <sheet_name>
<index>             integer "sheet_config": 2               The index of the worksheet. The first sheet gets index 1.
=================   ======= =============================   =======
