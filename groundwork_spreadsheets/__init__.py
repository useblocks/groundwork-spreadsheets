"""
Makes ExcelValidationPattern available for the package on root level
"""
# F401 imported but unused - it's needed as an API
from .patterns.ExcelValidationPattern.excel_validation_pattern import ExcelValidationPattern  # noqa F401
# define importable objects
__all__ = ['ExcelValidationPattern']
