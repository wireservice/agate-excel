0.2.4 - July 13, 2021
---------------------

* Add ``row_limit`` keyword argument to ``from_xls`` and ``from_xlsx``. (#40)
* Preserve column types from XLS files. (#36)
* Add support for Compound File Binary File (CFBF) XLS files. (#44)
* Close XLSX file before raising error for non-existent sheet. (#34)
* Use less memory and close XLS files. (#39)
* Drop support for Python 3.4 (end-of-life was March 18, 2019).

0.2.3 - March 16, 2019
----------------------

* Fix bug in accepting ``column_names`` as keyword argument.
* Add a ``reset_dimensions`` argument to :meth:`.Table.from_xlsx` to recalculate the data's dimensions, instead of trusting those in the file's properties.
* Include tests and examples in distribution.
* agate-excel is now tested against Python 3.6 and 3.7.
* Drop support for Python 3.3 (end-of-life was September 29, 2017).
* Add support for openpyxl 2.6.0.

0.2.2 - January 28, 2018
------------------------

* Add an ``encoding_override`` argument to :meth:`.Table.from_xls` to override the encoding of the input XLS file.
* Add a ``header`` argument to :meth:`.Table.from_xls` and :meth:`.Table.from_xlsx` to indicate the presence of a header row.
* Add a ``read_only`` argument to :meth:`.Table.from_xlsx` to allow disabling read-only mode for `some spreadsheets <https://openpyxl.readthedocs.io/en/stable/optimized.html#worksheet-dimensions>`_.

0.2.1 - February 28, 2017
-------------------------

* Overload :meth:`.Table.from_xls` and :meth:`.Table.from_xlsx` to accept and return multiple sheets.
* Add a ``skip_lines`` argument to :meth:`.Table.from_xls` and :meth:`.Table.from_xlsx` to skip rows from the top of the sheet.
* Fix bug in handling ambiguous dates in XLS. (#9)
* Fix bug in handling an empty XLS.
* Fix bug in handling non-string column names in XLSX.

0.2.0 - December 19, 2016
-------------------------

* Fix bug in handling of ``None`` in boolean columns for XLS. (#11)
* Removed usage of deprecated openpyxl method ``get_sheet_by_name``.
* Remove monkeypatching.
* Upgrade required agate version to ``1.5.0``.
* Ensure columns with numbers for names (e.g. years) are parsed as strings.

0.1.0 - February 5, 2016
------------------------

* Initial version.
