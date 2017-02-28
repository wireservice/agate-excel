#!/usr/bin/env python
# -*- coding: utf8 -*-

import datetime

try:
    import unittest2 as unittest
except ImportError:
    import unittest

import agate
import agateexcel

class TestXLSX(agate.AgateTestCase):
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

    def test_from_xlsx(self):
        table = agate.Table.from_xlsx('examples/test.xlsx')

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_file_like(self):
        with open('examples/test.xlsx', 'rb') as f:
            table = agate.Table.from_xlsx(f)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_sheet_name(self):
        table = agate.Table.from_xlsx('examples/test_sheets.xlsx', 'data')

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_sheet_index(self):
        table = agate.Table.from_xlsx('examples/test_sheets.xlsx', 1)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_sheet_multiple(self):
        tables = agate.Table.from_xlsx('examples/test_sheets.xlsx', ['not this sheet', 1])

        self.assertEqual(len(tables), 2)

        table = tables['not this sheet']
        self.assertColumnNames(table, [])
        self.assertColumnTypes(table, [])
        self.assertRows(table, [])

        table = tables['data']
        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_skip_lines(self):
        table = agate.Table.from_xlsx('examples/test_skip_lines.xlsx', skip_lines=3)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_ambiguous_date(self):
        table = agate.Table.from_xlsx('examples/test_ambiguous_date.xlsx')

        self.assertColumnNames(table, ['s'])
        self.assertColumnTypes(table, [agate.Date])
        self.assertRows(table, [
            [datetime.date(1899, 12, 31)]
        ])

    def test_empty(self):
        table = agate.Table.from_xlsx('examples/test_empty.xlsx')

        self.assertColumnNames(table, [])
        self.assertColumnTypes(table, [])
        self.assertRows(table, [])

    def test_numeric_column_name(self):
        table = agate.Table.from_xlsx('examples/test_numeric_column_name.xlsx')

        self.assertColumnNames(table, ['Country', '2013', 'c'])
        self.assertColumnTypes(table, [agate.Text, agate.Number, agate.Text])
        self.assertRows(table, [
            ['Canada', 35160000, 'value']
        ])
