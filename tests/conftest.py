import os
import pytest

from groundwork_spreadsheets import ExcelValidationPattern


@pytest.fixture
def emptyApp():
    """
    Loads an empty groundwork application and returns it.
    :return: app
    """
    from groundwork import App
    app = App(plugins=[], strict=True)
    return app


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