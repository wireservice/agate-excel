#!/usr/bin/env python

import agate

import agateexcel

table = agate.Table.from_xls('examples/test.xls')

print(table)

table = agate.Table.from_xlsx('examples/test.xlsx')

print(table)
table.print_table()
