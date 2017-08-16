import pytest


@pytest.fixture
def emptyApp():
    """
    Loads an empty groundwork application and returns it.
    :return: app
    """
    from groundwork import App
    app = App(plugins=[], strict=True)
    return app
