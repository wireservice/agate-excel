#!/usr/bin/env python
# -*- coding: utf8 -*-

try:
    import unittest2 as unittest
except ImportError:
    import unittest

import agate
import agateexcel
from sqlalchemy import create_engine

agateexcel.patch()

class TestTable(unittest.TestCase):
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
        self.connection_string = 'sqlite:///:memory:'

    def test_from_xls(self):
        pass

    def test_from_xlsx(self):
        pass
