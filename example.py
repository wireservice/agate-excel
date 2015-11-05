#!/usr/bin/env python

import agate
import agateexcel

agateexcel.patch()

table = agate.Table.from_xls('filename.xls')

print(table.column_names)
print(table.column_types)
print(len(table.columns))
print(len(table.rows))

table = agate.Table.from_xlsx('filename.xlsx')

print(table.column_names)
print(table.column_types)
print(len(table.columns))
print(len(table.rows))
