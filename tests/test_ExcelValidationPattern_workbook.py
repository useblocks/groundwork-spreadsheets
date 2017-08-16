
import pytest

from tests.conftest import EmptyPlugin, _get_test_data_path


def test_workbook_not_exist(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = 'not_exist.xlsx'
    config_path = _get_test_data_path('config_sheet_first.json')
    with pytest.raises(FileNotFoundError):
        plugin.excel_validation.read_excel(config_path, workbook_path)


def test_read_valid_workbook(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = _get_test_data_path('test.xlsx')
    config_path = _get_test_data_path('config_sheet_first.json')
    workbook_data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert workbook_data is not None
    assert type(workbook_data) is dict
