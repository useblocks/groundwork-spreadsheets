from tests.conftest import EmptyPlugin, get_test_data_path
import datetime


def test_data_types(emptyApp):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('data_types_libreoffice.xlsx')
    config_path = get_test_data_path('config.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert data[2]['Date'] == datetime.datetime(year=2017, month=8, day=20)
    assert data[3]['Date'] == datetime.datetime(year=2018, month=9, day=21)
    assert data[2]['Enum'] == 'ape'
    assert data[3]['Enum'] == 'dog'
    assert data[2]['Float'] == 1.1
    assert data[3]['Float'] == 22.22
    assert data[2]['Integer'] == -2
    assert data[3]['Integer'] == 0
    assert data[2]['Text'] == 'Text 1'
    assert data[3]['Text'] == 'Text 2'
