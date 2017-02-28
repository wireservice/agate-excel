0.2.1 - February 28, 2017
-------------------------

* Overload :meth:`.Table.from_xls` and :meth:`.Table.from_xlsx` to accept and return multiple sheets.
* Add a ``skip_lines`` argument to :meth:`.Table.from_xls` and :meth:`.Table.from_xlsx` to skip rows from the top of the sheet.
* Fix bug in handling ambiguous dates in XLS. (#9)
* Fix bug in handling an empty XLS.
* Fix bug in handling non-string column names in XLSX.

0.2.0
-----

* Fix bug in handling of ``None`` in boolean columns for XLS. (#11)
* Removed usage of deprecated openpyxl method ``get_sheet_by_name``.
* Remove monkeypatching.
* Upgrade required agate version to ``1.5.0``.
* Ensure columns with numbers for names (e.g. years) are parsed as strings.

0.1.0
-----

* Initial version.
