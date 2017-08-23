.. highlight:: python
    :linenothreshold: 5

.. image:: https://img.shields.io/pypi/l/groundwork-spreadsheets.svg
    :target: https://pypi.python.org/pypi/groundwork-spreadsheets
    :alt: License
.. image:: https://img.shields.io/pypi/pyversions/groundwork-spreadsheets.svg
    :target: https://pypi.python.org/pypi/groundwork-spreadsheets
    :alt: Supported versions
.. image:: https://readthedocs.org/projects/groundwork-spreadsheets/badge/?version=latest
    :target: https://readthedocs.org/projects/groundwork-spreadsheets/
.. image:: https://travis-ci.org/useblocks/groundwork-spreadsheets.svg?branch=master
    :target: https://travis-ci.org/useblocks/groundwork-spreadsheets
    :alt: Travis-CI Build Status
.. image:: https://coveralls.io/repos/github/useblocks/groundwork-spreadsheets/badge.svg?branch=master
    :target: https://coveralls.io/github/useblocks/groundwork-spreadsheets?branch=master
.. image:: https://img.shields.io/scrutinizer/g/useblocks/groundwork-spreadsheets.svg
    :target: https://scrutinizer-ci.com/g/useblocks/groundwork-spreadsheets/
    :alt: Code quality
.. image:: https://img.shields.io/pypi/v/groundwork-spreadsheets.svg
    :target: https://pypi.python.org/pypi/groundwork-spreadsheets
    :alt: PyPI Package latest release

.. _groundwork: https://groundwork.readthedocs.io

groundwork-spreadsheets
=======================

This is a `groundwork`_ extension package for reading and writing spreadsheet files.

`groundwork`_ is a plugin based Python application framework, which can be used to create various types of applications:
console scripts, desktop apps, dynamic websites and more.

Visit `groundwork.useblocks.com <http://groundwork.useblocks.com>`_
or read the `technical documentation <https://groundwork.readthedocs.io>`_ for more information.

Functions
---------

**ExcelValidationPattern**

Target audience are users who need to read well structured Excel documents with minimal overhead.
The Excel documents can be validated in various ways to detect input data problems.

*   Uses the library `openpyxl <https://openpyxl.readthedocs.io/en/default/>`_
*   Can read Excel 2010 files (xlsx, xlsm)
*   Configure your sheet using a json file
*   Auto detect columns by names. You can move columns without affecting the read routines.
*   The data layout can be

    *   column based: headers are in a single *row* and the data is below
    *   row based: headers are in a single *column* and the data is on the right

*   Define column types and verify cell values against them

    *   Date
    *   Enums (e.g. only  the values 'yes' and 'no' are allowed)
    *   Floating point numbers with optional min/max check
    *   Integer numbers with optional min/max check
    *   String with optional regular expression pattern check

*   Exclude data row/columns based on filter criteria
*   Output is a dictionary of the following form ``row or column number`` -> ``header name`` -> ``cell value``
*   Extensive logging of problems

Package content
---------------

.. toctree::
    :maxdepth: 4

    contents/usage
    contents/configuration
    contents/changelog
