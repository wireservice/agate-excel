#!/usr/bin/env python

try:
    import unittest2 as unittest
except ImportError:
    import unittest

from collections import OrderedDict

import agate
import agateexcel

class TestSetXLS(agate.AgateTestCase):
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

        self.tables = OrderedDict([
            ('not this sheet', agate.Table((), [], [])),
            ('data', agate.Table(self.rows, self.column_names, self.column_types))
        ])

    def test_from_xls(self):
        tableset1 = agate.MappedSequence(self.tables.values(), self.tables.keys())
        tableset2 = agate.TableSet.from_xls('examples/test_sheets.xls')

        self.assertEqual(len(tableset1), len(tableset2))

        for name in ['not this sheet', 'data']:
            self.assertEqual(len(tableset1[name].columns), len(tableset2[name].columns))
            self.assertEqual(len(tableset1[name].rows), len(tableset2[name].rows))

            if len(tableset1[name].rows):
                for i in range(len(tableset1[name].rows)):
                    self.assertSequenceEqual(tableset1[name].rows[i], tableset2[name].rows[i])

    def test_file_like(self):
        tableset1 = agate.MappedSequence(self.tables.values(), self.tables.keys())
        with open('examples/test_sheets.xls', 'rb') as f:
            tableset2 = agate.TableSet.from_xls(f)

        self.assertEqual(len(tableset1), len(tableset2))

        for name in ['not this sheet', 'data']:
            self.assertEqual(len(tableset1[name].columns), len(tableset2[name].columns))
            self.assertEqual(len(tableset1[name].rows), len(tableset2[name].rows))

            if len(tableset1[name].rows):
                for i in range(len(tableset1[name].rows)):
                    self.assertSequenceEqual(tableset1[name].rows[i], tableset2[name].rows[i])

    def test_sheet_name(self):
        tables = OrderedDict([
            ('data', agate.Table(self.rows, self.column_names, self.column_types))
        ])

        tableset1 = agate.MappedSequence(tables.values(), tables.keys())
        tableset2 = agate.TableSet.from_xls('examples/test_sheets.xls', sheets=['data'])

        self.assertEqual(len(tableset1), len(tableset2))

        for name in ['data']:
            self.assertEqual(len(tableset1[name].columns), len(tableset2[name].columns))
            self.assertEqual(len(tableset1[name].rows), len(tableset2[name].rows))

            if len(tableset1[name].rows):
                for i in range(len(tableset1[name].rows)):
                    self.assertSequenceEqual(tableset1[name].rows[i], tableset2[name].rows[i])

    def test_sheet_index(self):
        tables = OrderedDict([
            ('data', agate.Table(self.rows, self.column_names, self.column_types))
        ])

        tableset1 = agate.MappedSequence(tables.values(), tables.keys())
        tableset2 = agate.TableSet.from_xls('examples/test_sheets.xls', sheets=[1])

        self.assertEqual(len(tableset1), len(tableset2))

        for name in ['data']:
            self.assertEqual(len(tableset1[name].columns), len(tableset2[name].columns))
            self.assertEqual(len(tableset1[name].rows), len(tableset2[name].rows))

            if len(tableset1[name].rows):
                for i in range(len(tableset1[name].rows)):
                    self.assertSequenceEqual(tableset1[name].rows[i], tableset2[name].rows[i])
