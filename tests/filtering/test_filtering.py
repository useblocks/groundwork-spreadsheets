import pytest
import sys

from tests.conftest import EmptyPlugin, get_test_data_path


@pytest.mark.parametrize('orientation', ['column_based', 'row_based'])
def test_filtering(emptyApp, caplog, orientation):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path(orientation + '.xlsx')
    config_path = get_test_data_path(orientation + '.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert 2 in data
    assert data[2]['Enum'] == 'ape'
    assert 3 not in data
    assert 4 in data
    assert data[4]['Enum'] == 'cat'
    if orientation == 'column_based':
        assert 'The row 3 was excluded due to an exclude filter on cell B3 (dog not in [ape, cat])' in caplog.text()
    else:
        assert 'The column 3 was excluded due to an exclude filter on cell C2 (dog not in [ape, cat])' in caplog.text()


@pytest.mark.parametrize('path, config', [
    ('excluded_fail_on_type_error.xlsx', 'excluded_fail_on_type_error.json'),
    ('excluded_fail_on_empty_cell.xlsx', 'excluded_fail_on_empty_cell.json')
])
def test_excluded_fail_on_errors(emptyApp, path, config):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path(path)
    config_path = get_test_data_path(config)
    with pytest.raises(ValueError):
        plugin.excel_validation.read_excel(config_path, workbook_path)


@pytest.mark.parametrize('activation_state', ['enable', 'disable'])
def test_excluded_logging(emptyApp, caplog, activation_state):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = get_test_data_path('excluded_logging.xlsx')
    config_path = get_test_data_path('excluded_' + activation_state + '_logging.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert 2 in data
    assert data[2]['Enum'] == 'ape'
    assert 3 not in data
    assert 4 in data
    assert data[4]['Enum'] == 'cat'
    assert 'The row 3 was excluded due to an exclude filter on cell B3 (dog not in [ape, cat])' in caplog.text()

    if sys.version.startswith('2.7'):
        str_type = 'unicode'
        type_class = 'type'
    elif sys.version.startswith('3'):
        str_type = 'str'
        type_class = 'class'
    else:
        raise RuntimeError('This test case does only support Python 2.7 and 3.x')

    msg1 = "The value Text in cell A3 is of type <{0} '{1}'>; required by specification " \
           "is datetime".format(type_class, str_type)
    msg2 = "The 'Text' in cell E3 is empty"
    if activation_state == 'enable':
        assert msg1 in caplog.text()
        assert msg2 in caplog.text()
    else:
        assert msg1 not in caplog.text()
        assert msg2 not in caplog.text()
