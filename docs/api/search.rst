
.. _search_api:

Search and replace
==================

.. currentmodule:: docx.search

High-level ``Document.search`` / ``Document.replace`` methods delegate to the
functions in this module. Use the functions directly when operating on a
list of paragraphs returned by another API.


|SearchMatch| objects
---------------------

.. autoclass:: SearchMatch()


Searching
---------

.. autofunction:: search_paragraphs

.. autofunction:: search_paragraphs_regex

.. autofunction:: search_all_paragraphs

.. autofunction:: search_all_paragraphs_regex


Replacing
---------

.. autofunction:: replace_in_paragraphs

.. autofunction:: replace_in_paragraphs_regex

.. autofunction:: replace_in_all_paragraphs

.. autofunction:: replace_in_all_paragraphs_regex
