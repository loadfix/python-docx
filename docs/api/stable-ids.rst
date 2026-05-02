
.. _stable_ids_api:

Stable identifiers
==================

.. currentmodule:: docx.ids

The :mod:`docx.ids` module provides pragmatic mostly-stable identifiers for
paragraphs, runs, tables, and cells. The high-level API surface is the
``stable_id`` property on each of those proxy classes; the helper below is
exposed for advanced use-cases.


Functions
---------

.. autofunction:: compute_stable_id
