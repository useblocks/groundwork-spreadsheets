import pytest

from tests.conftest import EmptyPlugin, get_test_data_path


def test_sheet_active(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('sheets.xlsx')
    config_path = get_test_data_path('config_sheet_active.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert data[2]['Text'] == "sheet_middle_active"

def test_sheet_byIndex(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('sheets.xlsx')
    config_path = get_test_data_path('config_sheet_index.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert data[2]['Text'] == "sheet_last"


def test_sheet_byIndex_out_of_range(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('sheets.xlsx')
    config_path = get_test_data_path('config_sheet_index_out_of_range.json')
    with pytest.raises(IndexError):
        plugin.excel_validation.read_excel(config_path, workbook_path)


def test_sheet_name(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('sheets.xlsx')
    config_path = get_test_data_path('config_sheet_name.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert data[2]['Text'] == "sheet_last"


def test_sheet_name_not_exist(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('sheets.xlsx')
    config_path = get_test_data_path('config_sheet_name_not_exist.json')
    with pytest.raises(KeyError):
        plugin.excel_validation.read_excel(config_path, workbook_path)


def test_sheet_first(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('sheets.xlsx')
    config_path = get_test_data_path('config_sheet_first.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert data[2]['Text'] == "sheet_first"


def test_sheet_last(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('sheets.xlsx')
    config_path = get_test_data_path('config_sheet_last.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert data[2]['Text'] == "sheet_last"
