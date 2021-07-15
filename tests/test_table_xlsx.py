#!/usr/bin/env python
# -*- coding: utf8 -*-

import datetime

import agate
import six

import agateexcel  # noqa: F401


class TestXLSX(agate.AgateTestCase):
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
            'number', 'text', 'boolean', 'date', 'datetime',
        ]

        self.column_types = [
            agate.Number(), agate.Text(), agate.Boolean(),
            agate.Date(), agate.DateTime(),
        ]

        self.table = agate.Table(self.rows, self.column_names, self.column_types)

    def test_from_xlsx_with_column_names(self):
        table = agate.Table.from_xlsx('examples/test.xlsx', header=False, skip_lines=1,
                                      column_names=self.user_provided_column_names)

        self.assertColumnNames(table, self.user_provided_column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])

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

    def test_ambiguous_date(self):
        table = agate.Table.from_xlsx('examples/test_ambiguous_date.xlsx')

        # openpyxl >= 3 fixes a bug, but Python 2 is constrained to openpyxl < 3.
        if six.PY2:
            expected = datetime.date(1899, 12, 31)
        else:
            expected = datetime.date(1900, 1, 1)

        self.assertColumnNames(table, ['s'])
        self.assertColumnTypes(table, [agate.Date])
        self.assertRows(table, [
            [expected],
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
            ['Canada', 35160000, 'value'],
        ])

    def test_row_limit(self):
        table = agate.Table.from_xlsx('examples/test.xlsx', row_limit=2)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows][:2])

    def test_row_limit_too_high(self):
        table = agate.Table.from_xlsx('examples/test.xlsx', row_limit=200)

        self.assertColumnNames(table, self.column_names)
        self.assertColumnTypes(table, [agate.Number, agate.Text, agate.Boolean, agate.Date, agate.DateTime])
        self.assertRows(table, [r.values() for r in self.table.rows])
