#!/usr/bin/env python

import agate
import agateexcel

agateexcel.patch()

table = agate.Table.from_xls('examples/test.xls')

print(table)

table = agate.Table.from_xlsx('examples/test.xlsx')

print(table)
