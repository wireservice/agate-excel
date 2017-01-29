#!/usr/bin/env python

"""
This module contains the XLSX extension to :class:`TableSet <agate.table.TableSet>`.
"""

from collections import OrderedDict

import agate
import openpyxl

def from_xlsx(cls, path, sheets=None, **kwargs):
    """
    Parse an XLSX file.

    :param path:
        Path to an XLSX file to load or a file or file-like object for one.
    :param sheets:
        The name or integer indices of the worksheets to load. If not specified
        then all sheets will be loaded.
    """
    if hasattr(path, 'read'):
        f = path
    else:
        f = open(path, 'rb')

    book = openpyxl.load_workbook(f, read_only=True, data_only=True)

    tables = OrderedDict()

    for i, sheet in enumerate(book.worksheets):
        name = sheet.title
        if not sheets or name in sheets or i in sheets:
            tables[name] = agate.Table.from_xlsx(None, sheet, **kwargs)

    f.close()

    return agate.MappedSequence(tables.values(), tables.keys())

agate.TableSet.from_xlsx = classmethod(from_xlsx)
