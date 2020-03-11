"""
Makes ExcelValidationPattern available for the package on root level
"""
try:
    import importlib_metadata  # Python 3.6
except ModuleNotFoundError:
    import importlib.metadata as importlib_metadata  # Python 3.7 and 3.8

__version__ = importlib_metadata.version('groundwork_spreadsheets')
__summary__ = importlib_metadata.metadata('groundwork_spreadsheets')['Summary']

# F401 imported but unused - it's needed as an API
from .patterns.ExcelValidationPattern.excel_validation_pattern import ExcelValidationPattern  # noqa F401
# define importable objects
__all__ = ['ExcelValidationPattern']
