=====================
agate-excel |release|
=====================

.. include:: ../README.rst

Install
=======

To install:

.. code-block:: bash

    pip install agateexcel

For details on development or supported platforms see the `agate documentation <http://agate.readthedocs.org>`_.

Usage
=====

agate-excel uses a monkey patching pattern to add read for xls and xlsx files support to all :class:`agate.Table <agate.table.Table>` instances.

.. code-block:: python

  import agate
  import agateexcel

  agateexcel.patch()

Calling :func:`.patch` attaches all the methods of :class:`.TableXLS` and :class:`.TableXLSX` to :class:`agate.Table <agate.table.Table>`. Once you've patched the module, you can create tables from both XLS and XLSX files.

.. code-block:: python

  table = agate.Table.from_xls('examples/test.xls')
  print(table)

  table = agate.Table.from_xlsx('examples/test.xlsx')
  print(table)

Both methods accept a :code:`sheet` argument to specify which sheet to create the table from.

===
API
===

.. autofunction:: agateexcel.patch

.. autoclass:: agateexcel.table_xls.TableXLS
    :members:

.. autoclass:: agateexcel.table_xlsx.TableXLSX
    :members:

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
