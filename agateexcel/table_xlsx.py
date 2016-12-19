#!/usr/bin/env python

"""
This module contains the XLSX extension to :class:`Table <agate.table.Table>`.
"""

import datetime

import agate
import openpyxl
import six

NULL_TIME = datetime.time(0, 0, 0)

def from_xlsx(cls, path, sheet=None):
    """
    Parse an XLSX file.

    :param path:
        Path to an XLSX file to load or a file or file-like object for one.
    :param sheet:
        The name or integer index of a worksheet to load. If not specified
        then the "active" sheet will be used.
    """
    if hasattr(path, 'read'):
        f = path
    else:
        f = open(path, 'rb')

    book = openpyxl.load_workbook(f, read_only=True, data_only=True)

    if isinstance(sheet, six.string_types):
        sheet = book[sheet]
    elif isinstance(sheet, int):
        sheet = book.worksheets[sheet]
    else:
        sheet = book.active

    column_names = []
    rows = []

    for i, row in enumerate(sheet.rows):
        if i == 0:
            column_names = [c.value for c in row]
            continue

        values = []

        for c in row:
            value = c.value

            if value.__class__ is datetime.datetime:
                # Handle default XLSX date as 00:00 time
                if value.date() == datetime.date(1904, 1, 1) and not has_date_elements(c):
                    value = value.time()

                    value = normalize_datetime(value)
                elif value.time() == NULL_TIME:
                    value = value.date()
                else:
                    value = normalize_datetime(value)

            values.append(value)

        rows.append(values)

    f.close()

    return agate.Table(rows, column_names)

def normalize_datetime(dt):
    if dt.microsecond == 0:
        return dt

    ms = dt.microsecond

    if ms < 1000:
        return dt.replace(microsecond=0)
    elif ms > 999000:
        return dt.replace(microsecond=0) + datetime.timedelta(seconds=1)

    return dt

def has_date_elements(cell):
    """
    Try to use formatting to determine if a cell contains only time info.

    See: http://office.microsoft.com/en-us/excel-help/number-format-codes-HP005198679.aspx
    """
    if 'd' in cell.number_format or \
        'y' in cell.number_format:

        return True

    return False

agate.Table.from_xlsx = classmethod(from_xlsx)
