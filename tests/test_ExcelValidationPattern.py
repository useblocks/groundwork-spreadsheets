import os

from groundwork_spreadsheets import ExcelValidationPattern


# read tests
def test_read_valid_workbook(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = _get_test_data_path('test.xlsx')
    config_path = _get_test_data_path('test_config.json')
    workbook_data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert workbook_data is not None
    assert type(workbook_data) is dict


def _get_test_data_path(filename):
    return os.path.join(os.path.dirname(__file__), 'test_data', filename)


class EmptyPlugin(ExcelValidationPattern):
    def __init__(self, app, name=None, *args, **kwargs):
        self.name = name or self.__class__.__name__
        super(EmptyPlugin, self).__init__(app, *args, **kwargs)

    def activate(self):
        pass

    def deactivate(self):
        pass
