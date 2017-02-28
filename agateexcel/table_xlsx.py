#!/usr/bin/env python

"""
This module contains the XLSX extension to :class:`Table <agate.table.Table>`.
"""

import datetime
from collections import OrderedDict

import agate
import openpyxl
import six

NULL_TIME = datetime.time(0, 0, 0)

def from_xlsx(cls, path, sheet=None, skip_lines=0, **kwargs):
    """
    Parse an XLSX file.

    :param path:
        Path to an XLSX file to load or a file-like object for one.
    :param sheet:
        The names or integer indices of the worksheets to load. If not specified
        then the "active" sheet will be used.
    :param skip_lines:
        The number of rows to skip from the top of the sheet.
    """
    if not isinstance(skip_lines, int):
        raise ValueError('skip_lines argument must be an int')

    if hasattr(path, 'read'):
        f = path
    else:
        f = open(path, 'rb')

    book = openpyxl.load_workbook(f, read_only=True, data_only=True)

    multiple = agate.utils.issequence(sheet)
    if multiple:
        sheets = sheet
    else:
        sheets = [sheet]

    tables = OrderedDict()

    for i, sheet in enumerate(sheets):
        if isinstance(sheet, six.string_types):
            sheet = book[sheet]
        elif isinstance(sheet, int):
            sheet = book.worksheets[sheet]
        else:
            sheet = book.active

        column_names = []
        rows = []

        for i, row in enumerate(sheet.iter_rows(row_offset=skip_lines)):
            if i == 0:
                column_names = [None if c.value is None else six.text_type(c.value) for c in row]
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

        tables[sheet.title] = agate.Table(rows, column_names, **kwargs)

    f.close()

    if multiple:
        return agate.MappedSequence(tables.values(), tables.keys())
    else:
        return tables.popitem()[1]

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
