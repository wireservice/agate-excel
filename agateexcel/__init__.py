#!/usr/bin/env python

def patch():
    """
    Patch the features of this library onto agate's core
    :class:`Table <agate.table.Table>` and
    :class:`TableSet <agate.tableset.TableSet>`.
    """
    import agate
    from agateexcel.table import TableExcel

    agate.Table.monkeypatch(TableExcel)
