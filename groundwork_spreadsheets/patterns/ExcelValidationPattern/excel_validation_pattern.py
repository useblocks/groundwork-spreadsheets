"""
Groundwork Excel read/write routines using openpyxl
"""

import json
import os

import openpyxl
from openpyxl.utils import get_column_letter
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
        orientation = self.excel_config['orientation']
        headers_index_config_row_first = self.excel_config['headers_index_config']['row_index']['first']
        headers_index_config_row_last = self.excel_config['headers_index_config']['row_index']['last']
        headers_index_config_column_first = self.excel_config['headers_index_config']['column_index']['first']
        headers_index_config_column_last = self.excel_config['headers_index_config']['column_index']['last']
        data_index_config_row_first = self.excel_config['data_index_config']['row_index']['first']
        data_index_config_row_last = self.excel_config['data_index_config']['row_index']['last']
        data_index_config_column_first = self.excel_config['data_index_config']['column_index']['first']
        data_index_config_column_last = self.excel_config['data_index_config']['column_index']['last']

        # rotate coordinates so we can work with virtually column_based all the time
        if orientation == 'column_based':
            oriented_headers_index_config_row_first = headers_index_config_row_first
            oriented_headers_index_config_row_last = headers_index_config_row_last
            oriented_headers_index_config_column_first = headers_index_config_column_first
            oriented_headers_index_config_column_last = headers_index_config_column_last
            oriented_data_index_config_row_first = data_index_config_row_first
            oriented_data_index_config_row_last = data_index_config_row_last
            oriented_data_index_config_column_first = data_index_config_column_first
            oriented_data_index_config_column_last = data_index_config_column_last
        else:
            # row_based layout
            oriented_headers_index_config_row_first = headers_index_config_column_first
            oriented_headers_index_config_row_last = headers_index_config_column_last
            oriented_headers_index_config_column_first = headers_index_config_row_first
            oriented_headers_index_config_column_last = headers_index_config_row_last
            oriented_data_index_config_row_first = data_index_config_column_first
            oriented_data_index_config_row_last = data_index_config_column_last
            oriented_data_index_config_column_first = data_index_config_row_first
            oriented_data_index_config_column_last = data_index_config_row_last

        oriented_row_text = "row" if orientation == 'column_based' else "column"
        oriented_column_text = "column" if orientation == 'column_based' else "row"

        # set defaults of headers_index_config
        if type(oriented_headers_index_config_row_first) is not int:
            # Assume the user wants to use the first row as header row
            oriented_headers_index_config_row_first = 1
            self._plugin.log.debug("Config update: Setting headers_index_config -> {0}_index -> first to 1".format(
                oriented_row_text))

        if type(oriented_headers_index_config_row_last) is not int:
            # Only 1 header row is supported currently, set last header row to first header row
            oriented_headers_index_config_row_last = oriented_headers_index_config_row_first
            self._plugin.log.debug("Config update: Setting headers_index_config -> {0}_index -> last to {1}".format(
                oriented_row_text, oriented_headers_index_config_row_first))

        if type(oriented_headers_index_config_column_first) is not int:
            # Assume the user wants to start at the first column
            oriented_headers_index_config_column_first = 1
            self._plugin.log.debug("Config update: Setting headers_index_config -> {0}_index -> first to 1".format(
                oriented_column_text))

        if type(oriented_headers_index_config_column_last) is not int:
            if type(oriented_data_index_config_column_last) is int:
                # We don't have a last column in the header config,
                # so we use what we have in the data config
                oriented_headers_index_config_column_last = oriented_data_index_config_column_last
                self._plugin.log.debug("Config update: Setting headers_index_config -> {0}_index -> first to "
                                       "{1}".format(oriented_column_text, oriented_headers_index_config_column_last))

        # set defaults of data_index_config
        if type(oriented_data_index_config_row_first) is not int:
            # Assume the first data row is the next after oriented_headers_index_config_row_last
            oriented_data_index_config_row_first = oriented_headers_index_config_row_last + 1
            self._plugin.log.debug("Config update: Setting data_index_config -> {0}_index -> first to {1}".format(
                oriented_row_text, oriented_data_index_config_row_first))

        # oriented_data_index_config_row_last has no defaults - the user input is master

        if type(oriented_data_index_config_column_first) is not int:
            # Assume the first data column is equal to the first header column
            oriented_data_index_config_column_first = oriented_headers_index_config_column_first
            self._plugin.log.debug("Config update: Setting data_index_config -> {0}_index -> first to {1}".format(
                oriented_column_text, oriented_data_index_config_column_first))

        if type(oriented_data_index_config_column_last) is not int:
            if type(oriented_headers_index_config_column_last) is int:
                # We don't have a last column in the data config,
                # so we use what we have in the header config
                oriented_data_index_config_column_last = oriented_headers_index_config_column_last
                self._plugin.log.debug("Config update: Setting data_index_config -> {0}_index -> last to "
                                       "{1}".format(oriented_column_text, oriented_data_index_config_column_last))

        # Some more logic checks on rows
        if oriented_headers_index_config_row_last != oriented_headers_index_config_row_first:
            # We can compare both because at this point they have to be integer
            # Multi line headers given
            self._raise_value_error("Config error: Multi line (grouped) headers are not yet supported. "
                                    "First and last header {0} must be equal.".format(oriented_row_text))

        if oriented_data_index_config_row_first <= oriented_headers_index_config_row_last:
            # We can compare both because at this point they have to be integer
            # The data section is above the header section
            self._raise_value_error("Config error: headers_index_config -> {0}_index -> last is greater than "
                                    "data_index_config -> {0}_index -> first.".format(oriented_row_text))

        if type(oriented_data_index_config_row_last) is int:
            if oriented_data_index_config_row_last < oriented_data_index_config_row_first:
                # The last data row is smaller than the first
                self._raise_value_error("Config error: data_index_config -> {0}_index -> first is greater than "
                                        "data_index_config -> {0}_index -> last.".format(oriented_row_text))

        # Some more logic checks on columns
        if oriented_headers_index_config_column_first != oriented_data_index_config_column_first:
            # We can compare both because at this point they have to be integer
            # First column mismatch
            self._raise_value_error("Config error: header_index_config -> {0}_index -> first is not equal to "
                                    "data_index_config -> {0}_index -> first.".format(oriented_column_text))

        if type(oriented_headers_index_config_column_last) is int:
            if type(oriented_data_index_config_column_last) is int:
                if oriented_headers_index_config_column_last != oriented_data_index_config_column_last:
                    # Last columns are given but do not match
                    self._raise_value_error(
                        "Config error: header_index_config -> {0}_index -> last ({1}) is not equal to "
                        "data_index_config -> {0}_index -> last ({2}).".format(
                            oriented_column_text,
                            oriented_headers_index_config_column_last,
                            oriented_data_index_config_column_last
                        ))

        wb = openpyxl.load_workbook(excel_workbook_path, data_only=True)

        ws = self._get_sheet(wb)

        # Read header row
        # Determine header row length
        if type(oriented_headers_index_config_column_last) == int:
            header_column_last = oriented_headers_index_config_column_last
        elif oriented_headers_index_config_column_last == 'automatic':
            header_column_last = len(ws[self._transform_coordinates(oriented_headers_index_config_row_first)])
        else:
            # severalEmptyCells chosen
            target_empty_cell_count = int(oriented_headers_index_config_column_last.split(':')[1])
            empty_cell_count = 0
            header_column_last = oriented_headers_index_config_column_first
            while empty_cell_count < target_empty_cell_count:
                value = ws[self._transform_coordinates(oriented_headers_index_config_row_first,
                                                       header_column_last)].value
                if value is None:
                    empty_cell_count += 1
                header_column_last += 1
            header_column_last -= target_empty_cell_count + 1

        # Determine header column locations

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
        if type(self.excel_config['sheet_config']) == int:
            ws = wb.worksheets[self.excel_config['sheet_config'] - 1]
        elif self.excel_config['sheet_config'] == 'active':
            ws = wb.active
        elif self.excel_config['sheet_config'].startswith('name'):
            ws = wb[self.excel_config['sheet_config'].split(':')[1]]
        elif self.excel_config['sheet_config'] == 'first':
            ws = wb.worksheets[0]
        elif self.excel_config['sheet_config'] == 'last':
            ws = wb.worksheets[len(wb.get_sheet_names())-1]
        else:
            # This cannot happen if json validation was ok
            pass
        return ws

    def _raise_value_error(self, msg):
        self._plugin.log.error(msg)
        raise ValueError(msg)

    def _transform_coordinates(self, row=None, column=None):
        if row is None and column is None:
            raise ValueError("_transform_coordinates: row and column cannot both be None.")
        target_str = ''
        if self.excel_config['orientation'] == 'column_based':
            if column is not None:
                target_str = get_column_letter(column)
            if row is not None:
                target_str += str(row)
        else:
            if row is not None:
                target_str = get_column_letter(row)
            if column is not None:
                target_str += str(column)
        return target_str
