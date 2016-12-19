0.2.1
-----



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
