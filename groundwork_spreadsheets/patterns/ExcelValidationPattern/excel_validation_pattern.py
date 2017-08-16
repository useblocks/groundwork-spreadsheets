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

    def read_excel(self, excel_config_json_path, excel_workbook_path):

        # The exceptions raised in this method shall be raised to the plugin level
        excel_config = self._validate_json(excel_config_json_path, json_schema_file_path)

        # TODO Workbook not found exception
        wb = openpyxl.load_workbook(excel_workbook_path, data_only=True)

        ws = self._get_sheet(excel_config, wb)

        # TODO Workbook not found exception

        print(ws.max_row)
        print(ws.max_column)
        print(ws.dimensions)

        print(ws['A2'].value)
        print(type(ws['A2'].value))
        print(ws['B2'].value)
        print(type(ws['B2'].value))
        print(wb.get_sheet_names())

        return {}

    def _validate_json(self, excel_config_json_path, json_schema_file_path):

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

    def _get_sheet(self, excel_config, wb):

        # get sheet
        if excel_config['sheet_config']['search_type'] == 'active':
            ws = wb.active
        elif excel_config['sheet_config']['search_type'] == 'byIndex':
            ws = wb.worksheets[excel_config['sheet_config']['index']]
        elif excel_config['sheet_config']['search_type'] == 'byName':
            ws = wb[excel_config['sheet_config']['name']]
        elif excel_config['sheet_config']['search_type'] == 'first':
            ws = wb.worksheets[0]
        elif excel_config['sheet_config']['search_type'] == 'last':
            ws = wb.worksheets[len(wb.get_sheet_names())-1]
        else:
            # This cannot happen if json validation was ok
            raise NotImplementedError("The sheet_config search type {0} is not implemented".format(
                excel_config['sheet_config']))
        return ws
