.. _bookmarks:

Working with Bookmarks
======================

Word allows *bookmarks* to be defined on ranges of text in a document. A bookmark
names a specific range so that it can be navigated to (``Insert > Bookmark``) or
cross-referenced elsewhere — most notably by ``REF`` and ``PAGEREF`` fields, which
render the text or page number of the referenced bookmark.

A bookmark can be added to the main document, and bookmarks may also appear in
headers, footers, footnotes, and table cells. *python-docx* currently models
bookmarks that live in the main document body.

**Bookmark Anatomy.** Each bookmark is a *range* delimited by two empty marker
elements in the XML, ``<w:bookmarkStart/>`` and ``<w:bookmarkEnd/>``. Both
markers carry the same integer ``w:id`` attribute; the ``<w:bookmarkStart/>``
additionally carries the bookmark's ``w:name``. The start marker is placed
immediately before the first run in the range and the end marker immediately
after the last run in the range.

Like a comment reference, a bookmark range must begin and end at a *run*
boundary. A range can start in one paragraph and end in a later paragraph, but
it must always enclose a contiguous sequence of runs.

**Bookmark Names.** Bookmark names are strings. They must be unique within a
document — Word silently repairs duplicates on load. Names beginning with an
underscore (for example ``_Ref12345``) are conventionally *hidden* bookmarks
used internally by Word to back features such as automatic cross-references.
*python-docx* does not treat hidden bookmarks specially: they appear in the
``document.bookmarks`` collection alongside user-visible bookmarks.

**Bookmark IDs.** Each bookmark is identified by a non-negative integer ``id``
that is unique within the document. IDs are assigned automatically when a
bookmark is added via the *python-docx* API; the next available ID is chosen by
scanning existing ``w:bookmarkStart`` elements in the document body.


Adding a bookmark
-----------------

The simplest case is bookmarking a whole paragraph::

    >>> from docx import Document
    >>> document = Document()
    >>> paragraph = document.add_paragraph("Hello, world.")

    >>> bookmark = paragraph.add_bookmark("intro")
    >>> bookmark
    <docx.bookmarks.Bookmark object at 0x02468ACE>
    >>> bookmark.name
    'intro'
    >>> bookmark.bookmark_id
    0

To bookmark a specific range of runs within a paragraph, pass `start_run` and
`end_run`::

    >>> paragraph = document.add_paragraph("The ")
    >>> paragraph.add_run("middle").bold = True
    >>> paragraph.add_run(" run is special.")

    >>> bookmark = paragraph.add_bookmark(
    ...     "middle",
    ...     start_run=paragraph.runs[1],
    ...     end_run=paragraph.runs[1],
    ... )
    >>> bookmark.name
    'middle'

When only `start_run` is supplied, `end_run` defaults to `start_run`, so the
bookmark wraps that single run. When both `start_run` and `end_run` are
``None``, the bookmark wraps the whole paragraph.

.. note::
   The :meth:`.Paragraph.add_bookmark` method only bookmarks runs inside the
   paragraph it is called on. To create a bookmark that spans multiple
   paragraphs you currently need to drop down to the element level and insert
   the ``<w:bookmarkStart/>`` and ``<w:bookmarkEnd/>`` markers yourself.


Accessing the bookmarks collection
----------------------------------

The collection of bookmarks in a document is accessed via the
:attr:`.Document.bookmarks` property::

    >>> bookmarks = document.bookmarks
    >>> bookmarks
    <docx.bookmarks.Bookmarks object at 0x02468ACE>
    >>> len(bookmarks)
    2

The collection is iterable and yields |Bookmark| objects in document order::

    >>> for bm in bookmarks:
    ...     print(bm.name, bm.bookmark_id)
    intro 0
    middle 1


Looking up a bookmark by name
-----------------------------

Bookmarks are typically referenced by name. :meth:`.Bookmarks.get` returns the
matching bookmark, or ``None`` when no bookmark with that name is present::

    >>> bookmarks.get("intro")
    <docx.bookmarks.Bookmark object at 0x02468ACE>
    >>> bookmarks.get("not_a_bookmark") is None
    True

The collection also supports ``in`` for a quick presence check::

    >>> "intro" in bookmarks
    True
    >>> "not_a_bookmark" in bookmarks
    False


Deleting a bookmark
-------------------

A bookmark can be removed from the document by calling
:meth:`.Bookmark.delete`. This removes both the ``<w:bookmarkStart/>`` and the
matching ``<w:bookmarkEnd/>`` marker, leaving the bookmarked text in place::

    >>> bookmark = bookmarks.get("intro")
    >>> bookmark.delete()

    >>> "intro" in bookmarks
    False
    >>> len(bookmarks)
    1

Deleting a bookmark is safe even when its start and end markers live in
different paragraphs: both are found by ``w:id`` and removed.
