import pytest

from tests.conftest import EmptyPlugin, get_test_data_path


@pytest.mark.parametrize('path', [
    'column_based_position_mix_1.xlsx',
    'column_based_position_mix_2.xlsx',
    'column_based_position_mix_3.xlsx',
    'column_based_position_mix_4.xlsx',
])
def test_matrix_column_based_positive(empty_app, path):
    plugin = EmptyPlugin(empty_app)
    workbook_path = get_test_data_path(path)
    config_path = get_test_data_path('config_column_based_position_mix.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert len(data) == 3
    assert data[2]['Text'] == 'Text 1'
    assert data[2]['Integer'] == -2
    assert data[4]['Text'] == 'Text 3'
    assert data[4]['Integer'] == 30


@pytest.mark.parametrize('config_json', [
    'config_column_based_errors_1.json',
    'config_column_based_errors_2.json',
    'config_column_based_errors_3.json',
    'config_column_based_errors_4.json'
])
def test_matrix_column_based_errors(empty_app, config_json):
    plugin = EmptyPlugin(empty_app)
    workbook_path = get_test_data_path('column_based.xlsx')
    config_path = get_test_data_path(config_json)
    with pytest.raises(ValueError):
        plugin.excel_validation.read_excel(config_path, workbook_path)


@pytest.mark.parametrize('path', [
    'row_based_position_mix_1.xlsx',
    'row_based_position_mix_2.xlsx',
    'row_based_position_mix_3.xlsx',
    'row_based_position_mix_4.xlsx',
])
def test_matrix_row_based_positive(empty_app, path):
    plugin = EmptyPlugin(empty_app)
    workbook_path = get_test_data_path(path)
    config_path = get_test_data_path('config_row_based_position_mix.json')
    data = plugin.excel_validation.read_excel(config_path, workbook_path)
    assert len(data) == 3
    assert data[2]['Text'] == 'Text 1'
    assert data[2]['Integer'] == -2
    assert data[4]['Text'] == 'Text 3'
    assert data[4]['Integer'] == 30
