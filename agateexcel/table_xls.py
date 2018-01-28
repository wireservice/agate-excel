#!/usr/bin/env python

"""
This module contains the XLS extension to :class:`Table <agate.table.Table>`.
"""

import datetime
from collections import OrderedDict

import agate
import six
import xlrd


def from_xls(cls, path, sheet=None, skip_lines=0, header=True, encoding_override=None, **kwargs):
    """
    Parse an XLS file.

    :param path:
        Path to an XLS file to load or a file-like object for one.
    :param sheet:
        The names or integer indices of the worksheets to load. If not specified
        then the first sheet will be used.
    :param skip_lines:
        The number of rows to skip from the top of the sheet.
    :param header:
        If :code:`True`, the first row is assumed to contain column names.
    """
    if not isinstance(skip_lines, int):
        raise ValueError('skip_lines argument must be an int')

    if hasattr(path, 'read'):
        book = xlrd.open_workbook(file_contents=path.read(), encoding_override=encoding_override)
    else:
        with open(path, 'rb') as f:
            book = xlrd.open_workbook(file_contents=f.read(), encoding_override=encoding_override)

    multiple = agate.utils.issequence(sheet)
    if multiple:
        sheets = sheet
    else:
        sheets = [sheet]

    tables = OrderedDict()

    for i, sheet in enumerate(sheets):
        if isinstance(sheet, six.string_types):
            sheet = book.sheet_by_name(sheet)
        elif isinstance(sheet, int):
            sheet = book.sheet_by_index(sheet)
        else:
            sheet = book.sheet_by_index(0)

        if header:
            offset = 1
            column_names = []
        else:
            offset = 0
            column_names = None

        columns = []

        for i in range(sheet.ncols):
            data = sheet.col_values(i)
            values = data[skip_lines + offset:]
            types = sheet.col_types(i)[skip_lines + offset:]
            excel_type = determine_excel_type(types)

            if excel_type == xlrd.biffh.XL_CELL_BOOLEAN:
                values = normalize_booleans(values)
            elif excel_type == xlrd.biffh.XL_CELL_DATE:
                values = normalize_dates(values, book.datemode)

            if header:
                name = six.text_type(data[skip_lines]) or None
                column_names.append(name)

            columns.append(values)

        rows = []

        if columns:
            for i in range(len(columns[0])):
                rows.append([c[i] for c in columns])

        tables[sheet.name] = agate.Table(rows, column_names, **kwargs)

    if multiple:
        return agate.MappedSequence(tables.values(), tables.keys())
    else:
        return tables.popitem()[1]


def determine_excel_type(types):
    """
    Determine the correct type for a column from a list of cell types.
    """
    types_set = set(types)
    types_set.discard(xlrd.biffh.XL_CELL_EMPTY)

    # Normalize mixed types to text
    if len(types_set) > 1:
        return xlrd.biffh.XL_CELL_TEXT

    try:
        return types_set.pop()
    except KeyError:
        return xlrd.biffh.XL_CELL_EMPTY


def normalize_booleans(values):
    normalized = []

    for value in values:
        if value is None or value == '':
            normalized.append(None)
        else:
            normalized.append(bool(value))

    return normalized


def normalize_dates(values, datemode=0):
    """
    Normalize a column of date cells.
    """
    normalized = []

    for v in values:
        if not v:
            normalized.append(None)
            continue

        v_tuple = xlrd.xldate.xldate_as_datetime(v, datemode).timetuple()

        if v_tuple[3:6] == (0, 0, 0):
            # Date only
            normalized.append(datetime.date(*v_tuple[:3]))
        elif v_tuple[:3] == (0, 0, 0):
            normalized.append(datetime.time(*v_tuple[3:6]))
        else:
            # Date and time
            normalized.append(datetime.datetime(*v_tuple[:6]))

    return normalized


agate.Table.from_xls = classmethod(from_xls)
