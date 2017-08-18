import pytest

from tests.conftest import EmptyPlugin, _get_test_data_path


@pytest.mark.parametrize('config_json', [
    'config_matrix_errors_1.json',
    'config_matrix_errors_2.json',
    'config_matrix_errors_3.json',
    'config_matrix_errors_4.json'
])
def test_headers_data_indices_errors(emptyApp, config_json):
    plugin = EmptyPlugin(emptyApp)
    workbook_path = _get_test_data_path('test.xlsx')
    config_path = _get_test_data_path(config_json)
    with pytest.raises(ValueError):
        plugin.excel_validation.read_excel(config_path, workbook_path)
