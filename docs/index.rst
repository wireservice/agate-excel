=====================
agate-excel |release|
=====================

.. include:: ../README.rst

Install
=======

To install:

.. code-block:: bash

    pip install agate-excel

For details on development or supported platforms see the `agate documentation <http://agate.readthedocs.org>`_.

Usage
=====

agate-excel uses a monkey patching pattern to add read for xls and xlsx files support to all :class:`agate.Table <agate.table.Table>` and  :class:`agate.TableSet <agate.table.TableSet>` instances.

.. code-block:: python

  import agate
  import agateexcel

Importing agate-excel adds methods to :class:`agate.Table <agate.table.Table>` and :class:`agate.TableSet <agate.table.TableSet>`. Once you've imported it, you can create tables from both XLS and XLSX files.

.. code-block:: python

  table = agate.Table.from_xls('examples/test.xls')
  print(table)

  table = agate.Table.from_xlsx('examples/test.xlsx')
  print(table)

  tableset = agate.TableSet.from_xls('examples/test_sheets.xls')
  print(tableset)

  tableset = agate.TableSet.from_xlsx('examples/test_sheets.xlsx')
  print(tableset)

Both ``Table`` methods accept a :code:`sheet` argument to specify which sheet to create the table from. Both ``TableSet`` methods accept a :code:`sheets` argument to specify which sheets to create the table set from; otherwise, the table set will load all sheets.

===
API
===

.. autofunction:: agateexcel.table_xls.from_xls

.. autofunction:: agateexcel.table_xlsx.from_xlsx

.. autofunction:: agateexcel.tableset_xls.from_xls

.. autofunction:: agateexcel.tableset_xlsx.from_xlsx

Authors
=======

.. include:: ../AUTHORS.rst

Changelog
=========

.. include:: ../CHANGELOG.rst

License
=======

.. include:: ../COPYING

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
