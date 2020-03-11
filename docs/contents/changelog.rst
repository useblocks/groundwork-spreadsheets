Changelog
=========

0.4.0
-----

*   Fixed a encoding issue on Windows when reading the JSON config file (set to UTF-8)
*   Migrated from setuptools (setup.py) to poetry (pyproject.toml)
*   Removed py27 compatibility (only 3.6, 3.7 and 3.8 are supported now)
*   Fixed Travis CI

0.3.0
-----

*   Added 'convert_numbers' key to string type.
    This accepts numbers for string types too.
*   Added coveralls.io supprt
*   Added scrutinizer-ci.com support
*   Fixed pylint issues

0.2.0
-----

*   Added exclusion function for data row/columns based on filter criteria.
    Currently only enums whitelisting is supported.
*   Expanded test cases with documents saved by MS Excel 2013

0.1.2
-----

*   Added debug log messages for automatic last row/column detection
*   Fixed a small logging bug
*   Example: Added a configuration file to enable logging
*   Example: Added a readme
*   Description: Some rewriting

0.1.1
-----

*   Added description in setup.py for PyPi

0.1.0
-----

**Initial version**
