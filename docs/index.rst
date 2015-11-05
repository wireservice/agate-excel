===================
agate-sql |release|
===================

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

Calling :func:`.patch` attaches all the methods of :class:`.TableExcel` to :class:`agate.Table <agate.table.Table>`.

===
API
===

.. autofunction:: agatesql.patch

.. autoclass:: agatesql.table.TableExcel
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
