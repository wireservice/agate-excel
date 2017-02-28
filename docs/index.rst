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

agate-excel uses a monkey patching pattern to add read for xls and xlsx files support to all :class:`agate.Table <agate.table.Table>` instances.

.. code-block:: python

  import agate
  import agateexcel

Importing agate-excel adds methods to :class:`agate.Table <agate.table.Table>`. Once you've imported it, you can create tables from both XLS and XLSX files.

.. code-block:: python

  table = agate.Table.from_xls('examples/test.xls')
  print(table)

  table = agate.Table.from_xlsx('examples/test.xlsx')
  print(table)

  table = agate.Table.from_xlsx('examples/test.xlsx', sheet=1)
  print(table)

  table = agate.Table.from_xlsx('examples/test.xlsx', sheet='dummy')
  print(table)

  table = agate.Table.from_xlsx('examples/test.xlsx', sheet=[1, 'dummy'])
  print(table)

Both ``Table`` methods accept a :code:`sheet` argument to specify which sheet to create the table from.

===
API
===

.. autofunction:: agateexcel.table_xls.from_xls

.. autofunction:: agateexcel.table_xlsx.from_xlsx

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
