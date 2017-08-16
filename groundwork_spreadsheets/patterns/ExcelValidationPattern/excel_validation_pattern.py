"""
Groundwork Excel read/write routines using openpyxl
"""

import json
import os

import openpyxl
from groundwork.patterns.gw_base_pattern import GwBasePattern
from jsonschema import validate, ValidationError, SchemaError

json_schema_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'excel_config_schema.json')


class ExcelValidationPattern(GwBasePattern):
    """
    Groundwork Excel read/write routines using openpyxl
    """

    def __init__(self, *args, **kwargs):
        super(ExcelValidationPattern, self).__init__(*args, **kwargs)
        self.excel_validation = ExcelValidationPlugin(plugin=self)


class ExcelValidationPlugin:
    """
    Plugin level class of the Excel validation pattern.
    Note: There are no instances of the ExcelValidationPattern on the app.
    """

    def __init__(self, plugin):
        """
        :param plugin: The plugin, which wants to use documents
        :type plugin: GwBasePattern
        """
        self._plugin = plugin
        self._app = plugin.app
        self.excel_config = None

    def read_excel(self, excel_config_json_path, excel_workbook_path):

        # The exceptions raised in this method shall be raised to the plugin level
        self.excel_config = self._validate_json(excel_config_json_path)

        # Check config: headers_index_config and data_index_config
        headers_index_config_row = self.excel_config['headers_index_config']['row_index']
        headers_index_config_column = self.excel_config['headers_index_config']['column_index']
        data_index_config_row = self.excel_config['data_index_config']['row_index']
        data_index_config_column = self.excel_config['data_index_config']['column_index']

        # Check if exactly one of row_index or column_index is of type byIndex
        error_msg = "One of headers_index_config (row_index, column_index) must be of type 'byIndex'."
        if headers_index_config_row in ['automatic', 'severalEmptyCells'] and headers_index_config_column != 'byIndex':
            self._plugin.log.error(error_msg)
            raise ValueError(error_msg)

        if headers_index_config_column in ['automatic', 'severalEmptyCells'] and headers_index_config_row != 'byIndex':
            self._plugin.log.error(error_msg)
            raise ValueError(error_msg)

        error_msg = "One of data_index_config (row_index, column_index) must be of type 'byIndex'."
        if data_index_config_row in ['automatic', 'severalEmptyCells'] and data_index_config_column != 'byIndex':
            self._plugin.log.error(error_msg)
            raise ValueError(error_msg)

        if data_index_config_column in ['automatic', 'severalEmptyCells'] and data_index_config_row != 'byIndex':
            self._plugin.log.error(error_msg)
            raise ValueError(error_msg)

        # Check if users put in a 1x1 matrix for headers
        if headers_index_config_row == 'byIndex' and headers_index_config_column == 'byIndex':
            self._plugin.log.warn("Both headers_index_config (row_index, column_index) are set to byIndex. That means"
                                  "only one row or column is chosen (depending on the data_index_config).")

        # Check if headers and data both have byIndex in the same direction (row or column)
        if headers_index_config_row == 'byIndex' and data_index_config_row != 'byIndex':
            error_msg = "If headers_index_config (row_index) is type 'byIndex', then data_index_config (row_index) " \
                        "must also be of type 'byIndex'."
            self._plugin.log.error(error_msg)
            raise ValueError(error_msg)

        if headers_index_config_column == 'byIndex' and data_index_config_column != 'byIndex':
            error_msg = "If headers_index_config (column_index) is type 'byIndex', then data_index_config " \
                        "(column_index) must also be of type 'byIndex'."
            self._plugin.log.error(error_msg)
            raise ValueError(error_msg)

        # Check if data index is larger than header index

        wb = openpyxl.load_workbook(excel_workbook_path, data_only=True)

        ws = self._get_sheet(wb)

        print(ws.max_row)
        print(ws.max_column)
        print(ws.dimensions)

        print(ws['A2'].value)
        print(type(ws['A2'].value))
        print(ws['B2'].value)
        print(type(ws['B2'].value))
        print(wb.get_sheet_names())

        return {}

    def _validate_json(self, excel_config_json_path):

        try:
            with open(excel_config_json_path) as f:
                json_obj = json.load(f)
        # the file is not deserializable as a json object
        except ValueError as e:
            self._plugin.log.error('Malformed JSON file: {0} \n {1}'.format(excel_config_json_path, e))
            raise e
        # some os error occured (e.g file not found or malformed path string)
        # have to catch two exception classes: in py2 : IOError; py3: OSError
        except (IOError, OSError) as e:
            self._plugin.log.error(e)
            # raise only OSError to make error handling in caller easier
            raise OSError()

        # validate json object if schema file path is there; otherwise throw warning
        try:
            with open(json_schema_file_path) as f:
                schema_obj = json.load(f)
        # the file is not deserializable as a json object
        except ValueError as e:
            self._plugin.log.error('Malformed JSON schema file: {0} \n {1}'.format(json_schema_file_path, e))
            raise e
        # some os error occured (e.g file not found or malformed path string)
        # have to catch two exception classes:  in py2 : IOError; py3: OSError
        except (IOError, OSError) as e:
            self._plugin.log.error(e)
            # raise only OSError to make error handling in caller easier
            raise OSError()

        # do the validation
        try:
            validate(json_obj, schema_obj)
        except ValidationError as error:
            self._plugin.log.error("Validation failed: {0}".format(error.message))
            raise
        except SchemaError:
            self._plugin.log.error("Invalid schema file: {0}".format(json_schema_file_path))
            raise

        return json_obj

    def _get_sheet(self, wb):

        # get sheet
        ws = None
        if self.excel_config['sheet_config']['search_type'] == 'active':
            ws = wb.active
        elif self.excel_config['sheet_config']['search_type'] == 'byIndex':
            ws = wb.worksheets[self.excel_config['sheet_config']['index'] - 1]
        elif self.excel_config['sheet_config']['search_type'] == 'byName':
            ws = wb[self.excel_config['sheet_config']['name']]
        elif self.excel_config['sheet_config']['search_type'] == 'first':
            ws = wb.worksheets[0]
        elif self.excel_config['sheet_config']['search_type'] == 'last':
            ws = wb.worksheets[len(wb.get_sheet_names())-1]
        else:
            # This cannot happen if json validation was ok
            pass
        return ws
