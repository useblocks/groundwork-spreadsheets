"""
Fixtures for all tests.
Mainly an empty groundwork app and a plugin that inherits from the ExcelValidationPattern.
"""

import inspect
import os
import pytest

from groundwork_spreadsheets import ExcelValidationPattern


@pytest.fixture
def empty_app():
    """
    Loads an empty groundwork application and returns it.
    :return: app
    """
    from groundwork import App
    app = App(config_files=[os.path.join(os.path.dirname(__file__), 'configuration.py')], plugins=[], strict=True)
    return app


def get_test_data_path(filename):
    """
    Returns the path of a file in test_data folders.
    It's called from the test type directories (data_types, filtering, matrix, ...).
    To make it re-usable it must know who called it. It does so by inspecting the call stack.

    :param filename: The name of the file to build the path for
    :return: The path to the file.
    """
    frame = inspect.stack()[1]
    py_module = inspect.getmodule(frame[0])
    return os.path.join(os.path.dirname(py_module.__file__), 'test_data', filename)


class EmptyPlugin(ExcelValidationPattern):
    """
    A plugin that inherits from the ExcelValidationPattern.
    It's used in all test cases
    """
    def __init__(self, app, name=None, *args, **kwargs):
        self.name = name or self.__class__.__name__
        super(EmptyPlugin, self).__init__(app, *args, **kwargs)

    def activate(self):
        """
        Activates the plugin. Nothing to activate because the plugin does not register anything.
        Needed to fulfill the groundwork API.
        :return: None
        """
        pass

    def deactivate(self):
        """
        Deactivates the plugin. As nothing gets registered, there is nothing to unregister.
        Needed to fulfill the groundwork API.
        :return: None
        """
        pass
