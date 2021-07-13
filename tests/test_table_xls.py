#!/usr/bin/env python
# -*- coding: utf8 -*-

import datetime

import agate

import agateexcel  # noqa: F401


class TestXLS(agate.AgateTestCase):
    def setUp(self):
        self.rows = (
            (1, 'a', True, '11/4/2015', '11/4/2015 12:22 PM'),
            (2, u'👍', False, '11/5/2015', '11/4/2015 12:45 PM'),
            (None, 'b', None, None, None),
        )

        self.column_names = [
            'number', 'text', 'boolean', 'date', 'datetime',
        ]

        self.user_provided_column_names = [
            'alt number', 'alt text', 'alt boolean', 'alt date', 'alt datetime',
        ]

        self.column_types = [
            agate.Number(), agate.Text(), agate.Boolean(),
            agate.Date(), agate.DateTime(),
        ]

        self.table = agate.Table(self.rows, self.column_names, self.column_types)

    def test_from_xls_with_column_names(self):
        table = agate.Table.from_xls('examples/test.xls', header=False, skip_lines=1,
                                     column_names=self.user_provided_column_names)

        self.assertColumnNames(table, self.user_provided_column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

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

    def test_sheet_name(self):
        table = agate.Table.from_xls('examples/test_sheets.xls', 'data')

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_sheet_index(self):
        table = agate.Table.from_xls('examples/test_sheets.xls', 1)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_sheet_multiple(self):
        tables = agate.Table.from_xls('examples/test_sheets.xls', ['not this sheet', 1])

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
        table = agate.Table.from_xls('examples/test_skip_lines.xls', skip_lines=3)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

    def test_header(self):
        table = agate.Table.from_xls('examples/test_zeros.xls', header=False)

        self.assertColumnNames(table, ('a', 'b', 'c'))
        self.assertColumnTypes(table, [agate.Text, agate.Text, agate.Text])
        self.assertRows(table, [
            ['ordinal', 'binary', 'all_zero'],
            ['0.0', '0.0', '0.0'],
            ['1.0', '1.0', '0.0'],
            ['2.0', '1.0', '0.0'],
        ])

    def test_zeros(self):
        table = agate.Table.from_xls('examples/test_zeros.xls')

        self.assertColumnNames(table, ['ordinal', 'binary', 'all_zero'])
        self.assertColumnTypes(table, [agate.Number, agate.Number, agate.Number])
        self.assertRows(table, [
            [0, 0, 0],
            [1, 1, 0],
            [2, 1, 0],
        ])

    def test_ambiguous_date(self):
        table = agate.Table.from_xls('examples/test_ambiguous_date.xls')

        self.assertColumnNames(table, ['s'])
        self.assertColumnTypes(table, [agate.Date])
        self.assertRows(table, [
            [datetime.date(1900, 1, 1)],
        ])

    def test_empty(self):
        table = agate.Table.from_xls('examples/test_empty.xls')

        self.assertColumnNames(table, [])
        self.assertColumnTypes(table, [])
        self.assertRows(table, [])

    def test_numeric_column_name(self):
        table = agate.Table.from_xls('examples/test_numeric_column_name.xls')

        self.assertColumnNames(table, ('Country', '2013.0', 'c'))
        self.assertColumnTypes(table, [agate.Text, agate.Number, agate.Text])
        self.assertRows(table, [
            ['Canada', 35160000, 'value'],
        ])

    def test_row_limit(self):
        table = agate.Table.from_xls('examples/test.xls', row_limit=2)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows][:2])

    def test_row_limit_too_high(self):
        table = agate.Table.from_xls('examples/test.xls', row_limit=200)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])
