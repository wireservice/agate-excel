#!/usr/bin/env python

"""
This module contains the XLS extension to :class:`TableSet <agate.table.TableSet>`.
"""

from collections import OrderedDict

import agate
import xlrd

def from_xls(cls, path, sheets=None, **kwargs):
    """
    Parse an XLS file.

    :param path:
        Path to an XLS file to load or a file or file-like object for one.
    :param sheets:
        The name or integer indices of the worksheets to load. If not specified
        then all sheets will be loaded.
    """
    if hasattr(path, 'read'):
        book = xlrd.open_workbook(file_contents=path.read())
    else:
        with open(path, 'rb') as f:
            book = xlrd.open_workbook(file_contents=f.read())

    tables = OrderedDict()

    for i, sheet in enumerate(book.sheets()):
        name = sheet.name
        if not sheets or name in sheets or i in sheets:
            tables[name] = agate.Table.from_xls(None, sheet, **kwargs)

    return agate.MappedSequence(tables.values(), tables.keys())

agate.TableSet.from_xls = classmethod(from_xls)
