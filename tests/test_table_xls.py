#!/usr/bin/env python
# -*- coding: utf8 -*-

try:
    import unittest2 as unittest
except ImportError:
    import unittest

import agate
import agateexcel

agateexcel.patch()

class TestXLS(agate.AgateTestCase):
    def setUp(self):
        self.rows = (
            (1, 'a', True, '11/4/2015', '11/4/2015 12:22 PM'),
            (2, u'👍', False, '11/5/2015', '11/4/2015 12:45 PM'),
            (None, 'b', None, None, None)
        )

        self.column_names = [
            'number', 'text', 'boolean', 'date', 'datetime'
        ]

        self.column_types = [
            agate.Number(), agate.Text(), agate.Boolean(),
            agate.Date(), agate.DateTime()
        ]

        self.table = agate.Table(self.rows, self.column_names, self.column_types)

    def test_from_xls(self):
        table = agate.Table.from_xls('examples/test.xls')

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_file_like(self):
        with open('examples/test.xls', 'rb') as f:
            table = agate.Table.from_xls(f)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_sheet(self):
        table = agate.Table.from_xls('examples/test_sheets.xls', 'data')

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])
