"""
Groundwork Excel read/write routines using openpyxl
"""
import datetime
import json
import os

import openpyxl
import re

import sys
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

        #########################
        # Get index configuration
        #########################
        orientation = self.excel_config['orientation']
        headers_index_config_row_first = self.excel_config['headers_index_config']['row_index']['first']
        headers_index_config_row_last = self.excel_config['headers_index_config']['row_index']['last']
        headers_index_config_column_first = self.excel_config['headers_index_config']['column_index']['first']
        headers_index_config_column_last = self.excel_config['headers_index_config']['column_index']['last']
        data_index_config_row_first = self.excel_config['data_index_config']['row_index']['first']
        data_index_config_row_last = self.excel_config['data_index_config']['row_index']['last']
        data_index_config_column_first = self.excel_config['data_index_config']['column_index']['first']
        data_index_config_column_last = self.excel_config['data_index_config']['column_index']['last']

        ############################################################################
        # Rotate coordinates so we can work with virtually column based all the time
        # Just before addressing the cells a coordinate transformation is done again
        ############################################################################
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

        # Used for log and exception messages
        oriented_row_text = "row" if orientation == 'column_based' else "column"
        oriented_column_text = "column" if orientation == 'column_based' else "row"

        ######################################
        # Set defaults of headers_index_config
        ######################################
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

        ###################################
        # Set defaults of data_index_config
        ###################################
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

        ################################
        # Some more logic checks on rows
        # (matrix size and row order)
        ################################
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

        ###################################
        # Some more logic checks on columns
        # (mismatches)
        ###################################
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

        if type(oriented_headers_index_config_column_last) is not int:
            if type(oriented_data_index_config_column_last) is not int:
                # The column count search is done on the header row, not on data
                if oriented_data_index_config_column_last != 'automatic':
                    self._raise_value_error(
                        "Config error: data_index_config -> {0}_index -> last ({1}) may only be an integer or "
                        "contain the value 'automatic'.".format(oriented_column_text,
                                                                oriented_data_index_config_column_last))

        #########################################
        # Set defaults for optional config values
        #########################################
        if 'sheet_config' not in self.excel_config:
            self.excel_config['sheet_config'] = 'active'

        for data_type_config in self.excel_config['data_type_config']:
            # default for possible problems should be strict if user tells nothing
            if 'fail_on_type_error' not in data_type_config:
                data_type_config['fail_on_type_error'] = True
            if 'fail_on_empty_cell' not in data_type_config:
                data_type_config['fail_on_empty_cell'] = True
            if 'fail_on_header_not_found' not in data_type_config:
                data_type_config['fail_on_header_not_found'] = True

            # default type is automatic
            if 'type' not in data_type_config:
                data_type_config['type'] = {'base': 'automatic'}

        ############################
        # Get the workbook and sheet
        ############################
        wb = openpyxl.load_workbook(excel_workbook_path, data_only=True)
        ws = self._get_sheet(wb)

        #############################
        # Determine header row length
        #############################
        if type(oriented_headers_index_config_column_last) == int:
            # oriented_headers_index_config_column_last already has the final value
            pass
        elif oriented_headers_index_config_column_last == 'automatic':
            # automatic: use the length of the header row
            oriented_headers_index_config_column_last = len(
                ws[self._transform_coordinates(row=oriented_headers_index_config_row_first)])
        else:
            # severalEmptyCells chosen
            target_empty_cell_count = int(oriented_headers_index_config_column_last.split(':')[1])
            empty_cell_count = 0
            curr_column = oriented_headers_index_config_column_first
            while empty_cell_count < target_empty_cell_count:
                value = ws[self._transform_coordinates(oriented_headers_index_config_row_first,
                                                       curr_column)].value
                if value is None:
                    empty_cell_count += 1
                curr_column += 1
            oriented_headers_index_config_column_last = curr_column - target_empty_cell_count - 1

        ###################################
        # Determine header column locations
        ###################################
        spreadsheet_headers2columns = {}
        for column in range(oriented_headers_index_config_column_first, oriented_headers_index_config_column_last):
            value = ws[self._transform_coordinates(oriented_headers_index_config_row_first, column)].value
            if value is not None:
                spreadsheet_headers2columns[value] = column
            else:
                # if the value is None we have either
                # - one or more empty header cell in between 2 filled header cells.
                # - some empty header cells at the end of the row.
                #   That might happen because when choosing 'automatic' header row detection, openpyxl functionality
                #   is used. It always returns the length of the longest row in the whole sheet.
                #   If that is not the header row, we have empty cells at the end of the header row.
                #   However, that is not a problem as header values 'None' are not added.
                pass
        spreadsheet_headers = spreadsheet_headers2columns.keys()

        ###############################################################
        # Check for not existing headers on spreadsheet and config side
        ###############################################################
        # Build a data_type_config dictionary with header as key
        config_header_dict = {x['header']: x for x in self.excel_config['data_type_config']}

        # Check: Are config data_type_config headers unique?
        if len(config_header_dict.keys()) != len(set(config_header_dict.keys())):
            self._raise_value_error("Config error: data_type_config -> header duplicate entries found.")

        # Check for configured headers not found in spreadsheet
        missing_headers_in_spreadsheet = list(set(config_header_dict.keys()) - set(spreadsheet_headers))
        for header in missing_headers_in_spreadsheet:
            # Check if the fail_on_header_not_found is true
            data_type = [x for x in self.excel_config['data_type_config'] if x['header'] == header][0]
            msg = "Config error: The header '{0}' could not be found in the spreadsheet.".format(header)
            if data_type['fail_on_header_not_found']:
                self._raise_value_error(msg)
            else:
                self._plugin.log.error(msg)

        # Check for spreadsheet headers not found in config
        missing_headers_in_config = list(set(spreadsheet_headers) - set(config_header_dict.keys()))
        self._plugin.log.debug("The following spreadsheet headers are not configured for reading: {0}".format(
            ', '.join(missing_headers_in_config)))
        for header in missing_headers_in_config:
            del spreadsheet_headers2columns[header]

        #########################
        # Determine last data row
        #########################
        if type(oriented_data_index_config_row_last) == int:
            # do nothing, the value is already set
            pass
        elif oriented_data_index_config_row_last == 'automatic':
            # Take data from library
            oriented_data_index_config_row_last = oriented_data_index_config_row_first
            for curr_column in range(oriented_headers_index_config_column_first,
                                     oriented_headers_index_config_column_last):
                len_curr_column = len(ws[self._transform_coordinates(column=curr_column)])
                if len_curr_column > oriented_data_index_config_row_last:
                    oriented_data_index_config_row_last = len_curr_column
        else:
            # severalEmptyCells is chosen
            target_empty_rows_count = int(oriented_data_index_config_row_last.split(':')[1])
            last_row_detected = False
            curr_row = oriented_data_index_config_row_first
            empty_rows_count = 0
            while not last_row_detected:
                # go through rows
                all_columns_empty = True
                for header in spreadsheet_headers2columns:
                    curr_column = spreadsheet_headers2columns[header]
                    value = ws[self._transform_coordinates(curr_row, curr_column)].value
                    if value is not None:
                        all_columns_empty = False
                        break
                if all_columns_empty:
                    empty_rows_count += 1
                if empty_rows_count >= target_empty_rows_count:
                    last_row_detected = True
                else:
                    curr_row += 1
            oriented_data_index_config_row_last = curr_row - target_empty_rows_count

        #################################################
        # Go through the rows, read and validate the data
        #################################################
        if sys.version.startswith('2.7'):
            str_type = 'unicode'
        elif sys.version.startswith('3'):
            str_type = 'str'
        else:
            raise RuntimeError('The enum type specification does only support Python 2.7 and 3.x')

        final_dict = {}
        for curr_row in range(oriented_data_index_config_row_first, oriented_data_index_config_row_last + 1):
            # Go through rows
            final_dict[curr_row] = {}

            for header in spreadsheet_headers2columns:
                # Go through columns
                curr_column = spreadsheet_headers2columns[header]
                cell_index_str = self._transform_coordinates(curr_row, curr_column)
                value = ws[cell_index_str].value

                # Start the validation
                config_header = config_header_dict[header]

                if value is None:
                    msg = "The '{0}' in cell {1} is empty".format(header, cell_index_str)
                    if config_header['fail_on_empty_cell']:
                        self._raise_value_error(msg)
                    else:
                        self._plugin.log.warning(msg)
                else:
                    if config_header['type']['base'] == 'automatic':
                        pass
                    elif config_header['type']['base'] == 'date':
                        if type(value) != datetime.datetime:
                            msg = 'The value {0} in cell {1} is of type {2}; required by ' \
                                  'specification is datetime'.format(value, cell_index_str, type(value))
                            if config_header['fail_on_type_error']:
                                self._raise_value_error(msg)
                            else:
                                self._plugin.log.warning(msg)
                    elif config_header['type']['base'] == 'enum':
                        if type(value).__name__ != str_type:
                            msg = 'The value {0} in cell {1} is of type {2}; required by ' \
                                  'specification is str (enum)'.format(value, cell_index_str, type(value))
                            if config_header['fail_on_type_error']:
                                self._raise_value_error(msg)
                            else:
                                self._plugin.log.warning(msg)
                        else:
                            valid_values = config_header['type']['enum_values']
                            if value not in valid_values:
                                msg = 'The value {0} in cell {1} is not contained in the given enum ' \
                                      '[{2}]'.format(value, cell_index_str, ', '.join(valid_values))
                                if config_header['fail_on_type_error']:
                                    self._raise_value_error(msg)
                                else:
                                    self._plugin.log.warning(msg)
                    elif config_header['type']['base'] == 'float':
                        if type(value) != float:
                            msg = 'The value {0} in cell {1} is of type {2}; required by ' \
                                  'specification is float'.format(value, cell_index_str, type(value))
                            if config_header['fail_on_type_error']:
                                self._raise_value_error(msg)
                            else:
                                self._plugin.log.warning(msg)
                        else:
                            if 'minimum' in config_header['type']:
                                if value < config_header['type']['minimum']:
                                    msg = 'The value {0} in cell {1} is smaller than the given minimum ' \
                                          'of {2}'.format(value, cell_index_str, config_header['type']['minimum'])
                                    if config_header['fail_on_type_error']:
                                        self._raise_value_error(msg)
                                    else:
                                        self._plugin.log.warning(msg)
                            if 'maximum' in config_header['type']:
                                if value > config_header['type']['maximum']:
                                    msg = 'The value {0} in cell {1} is greater than the given maximum ' \
                                          'of {2}'.format(value, cell_index_str, config_header['type']['maximum'])
                                    if config_header['fail_on_type_error']:
                                        self._raise_value_error(msg)
                                    else:
                                        self._plugin.log.warning(msg)
                    elif config_header['type']['base'] == 'integer':
                        # Integer values stored by Excel are returned as float (e.g. 3465.0)
                        # So we have to check if the float can be converted to int without precision loss
                        if type(value) == float:
                            if value.is_integer():
                                value = int(value)
                        if type(value) == int:
                            if 'minimum' in config_header['type']:
                                if value < config_header['type']['minimum']:
                                    msg = 'The value {0} in cell {1} is smaller than the given minimum ' \
                                          'of {2}'.format(value, cell_index_str, config_header['type']['minimum'])
                                    if config_header['fail_on_type_error']:
                                        self._raise_value_error(msg)
                                    else:
                                        self._plugin.log.warning(msg)
                            if 'maximum' in config_header['type']:
                                if value > config_header['type']['maximum']:
                                    msg = 'The value {0} in cell {1} is greater than the given maximum ' \
                                          'of {2}'.format(value, cell_index_str, config_header['type']['maximum'])
                                    if config_header['fail_on_type_error']:
                                        self._raise_value_error(msg)
                                    else:
                                        self._plugin.log.warning(msg)
                        else:
                            msg = 'The value {0} in cell {1} is of type {2}; required by ' \
                                  'specification is int'.format(value, cell_index_str, type(value))
                            if config_header['fail_on_type_error']:
                                self._raise_value_error(msg)
                            else:
                                self._plugin.log.warning(msg)

                    elif config_header['type']['base'] == 'string':
                        if type(value).__name__ != str_type:
                            msg = 'The value {0} in cell {1} is of type {2}; required by ' \
                                  'specification is string'.format(value, cell_index_str, type(value))
                            if config_header['fail_on_type_error']:
                                self._raise_value_error(msg)
                            else:
                                self._plugin.log.warning(msg)
                        else:
                            if 'pattern' in config_header['type']:
                                if re.search(config_header['type']['pattern'], value) is None:
                                    msg = 'The value {0} in cell {1} does not follow the ' \
                                          'given pattern {2}'.format(value, cell_index_str,
                                                                     config_header['type']['pattern'])
                                    if config_header['fail_on_type_error']:
                                        self._raise_value_error(msg)
                                    else:
                                        self._plugin.log.warning(msg)

                final_dict[curr_row][header] = value

        return final_dict

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
