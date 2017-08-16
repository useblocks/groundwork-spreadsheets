
import pytest

from tests.conftest import EmptyPlugin, _get_test_data_path
import sys


def test_workbook_not_exist(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = 'not_exist.xlsx'
    config_path = _get_test_data_path('config_sheet_first.json')
    if sys.version.startswith('2.7'):
        with pytest.raises(IOError):
            plugin.excel_validation.read_excel(config_path, workbook_path)
    elif sys.version.startswith('3'):
        with pytest.raises(FileNotFoundError):
            plugin.excel_validation.read_excel(config_path, workbook_path)
    else:
        pytest.fail("Test does only support Python versions 2.7 and 3.x")


def test_read_valid_workbook(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = _get_test_data_path('test.xlsx')
    config_path = _get_test_data_path('config_sheet_first.json')
    workbook_data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert workbook_data is not None
    assert type(workbook_data) is dict
