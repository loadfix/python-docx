.. _footnotes:

Working with Footnotes
======================

Word allows *footnotes* to be attached to running prose. A footnote appears as a small
superscript reference mark in the body (usually a number, asterisk, or Roman numeral)
paired with a separate block of text that Word renders at the bottom of the page, where
the reader can consult it without losing their place.

The procedure, from the Word UI, is simple:

- You place the insertion cursor where you want the reference mark to appear
- You press the *Insert Footnote* button (References toolbar)
- You type or paste in the footnote content

**Footnote Anatomy.** Each footnote has two parts, a *footnote-reference* and a
*footnote-content*:

The **footnote-reference** is an empty marker element (``<w:footnoteReference/>``)
inserted inside a run at the point in the body where the superscript number should
appear. The reference carries the numeric ``id`` of the footnote it points at, and
*python-docx* styles the containing run with the ``FootnoteReference`` character style
so Word displays the mark as a superscript number.

The **footnote-content** is the prose that appears at the bottom of the page. Each
footnote's content is stored in a separate ``w:footnote`` element in the *footnotes
part* (part-name ``word/footnotes.xml``), not in the main document body. The two halves
are tied together by the shared ``id`` attribute.

**Reserved Ids.** The footnotes part always contains at least two ids that are not real
footnotes: ``id=0`` is the *separator* (the horizontal line between body text and
footnotes) and ``id=1`` is the *continuation separator* used when a footnote overflows
to the next page. User-added footnotes start at ``id=2`` and are assigned sequentially.
*python-docx* hides these reserved ids from iteration and from the ``Document.footnotes``
length.

**Applicability.** Footnotes can be added only in the main document body. The
*python-docx* API does not currently support adding footnotes inside comments,
headers, or footers, and it does not support endnotes being anchored to footnote text.


Adding a footnote
-----------------

A simple example is anchoring a footnote to the first run of a paragraph::

    >>> from docx import Document
    >>> document = Document()
    >>> paragraph = document.add_paragraph("The rain in Spain.")

    >>> footnote = document.footnotes.add(
    ...     paragraph.runs[0],
    ...     "A common saying about Iberian weather.",
    ... )
    >>> footnote
    <docx.footnotes.Footnote object at 0x02468ACE>
    >>> footnote.footnote_id
    2
    >>> footnote.text
    'A common saying about Iberian weather.'

Note that :meth:`.Footnotes.add` takes a single |Run| (not a range), because a
footnote has a point of insertion rather than a range of selected text. The
``FootnoteReference`` marker is inserted at the end of that run. If you need the
reference to appear in the middle of a run, split the run first so that the
insertion point lies on a run boundary.


Reading footnotes from a document
---------------------------------

The footnotes collection is reached via the :attr:`.Document.footnotes` property::

    >>> document = Document("has-footnotes.docx")
    >>> footnotes = document.footnotes
    >>> footnotes
    <docx.footnotes.Footnotes object at 0x02468ACE>
    >>> len(footnotes)
    3

The collection is iterable and yields |Footnote| objects for every user footnote in
document order. The separator and continuation-separator entries are filtered out::

    >>> for footnote in document.footnotes:
    ...     print(footnote.footnote_id, footnote.text)
    2 A common saying about Iberian weather.
    3 As of the loadfix fork.
    4 Ids 0 and 1 are reserved for separators.


Inspecting a footnote
---------------------

A |Footnote| is a *block-item container*, just like a document body or a table
cell, so it exposes the same paragraph-access API. Each footnote contains at
least one paragraph, styled ``FootnoteText``, whose first run carries the
auto-number mark that Word renders in front of the footnote text::

    >>> footnote = document.footnotes[0]
    >>> footnote.footnote_id
    2
    >>> len(footnote.paragraphs)
    1
    >>> footnote.paragraphs[0].style.name
    'FootnoteText'
    >>> footnote.text
    'A common saying about Iberian weather.'

The :attr:`.Footnote.text` property concatenates the text of every paragraph in
the footnote, joined by newlines. All emphasis and character-level styling is
stripped; use ``Footnote.paragraphs`` to walk the runs yourself if you need
richer access.


Adding rich content to a footnote
---------------------------------

Because a footnote is a block-item container, you can add additional paragraphs
and runs to it just like you would to the document body::

    >>> footnote = document.footnotes.add(paragraph.runs[0], "First line.")
    >>> second_para = footnote.add_paragraph("Second line.")
    >>> second_para.style.name
    'FootnoteText'
    >>> footnote.paragraphs[0].add_run(" (emphasised)").italic = True

:meth:`.Footnote.add_paragraph` applies the ``FootnoteText`` paragraph style by
default so the added paragraph blends in visually with the footnote's existing
content.


Modifying and deleting footnotes
--------------------------------

:meth:`.Footnote.clear` drops every run after the initial auto-number mark,
leaving a single empty paragraph you can populate fresh::

    >>> footnote.clear()
    >>> footnote.text
    ''
    >>> len(footnote.paragraphs)
    1

:meth:`.Footnote.delete` removes the footnote outright. Both the
``w:footnote`` element in the footnotes part and every ``w:footnoteReference``
in the document body that targets it are removed; runs that contained only the
reference are cleaned up::

    >>> len(document.footnotes)
    3
    >>> document.footnotes[0].delete()
    >>> len(document.footnotes)
    2

After calling :meth:`~.Footnote.delete` the |Footnote| object is "defunct" and
should not be used further.


Footnote numbering properties
-----------------------------

Footnote numbering is configured at the document level via a |FootnoteProperties|
object. Access it via :attr:`.Document.footnote_properties`; the property is
|None| when no ``w:footnotePr`` element is present in the settings part, in
which case Word applies its defaults (Arabic numerals starting at 1, continuous
numbering, footnotes at the bottom of the page)::

    >>> document.footnote_properties is None
    True
    >>> props = document.add_footnote_properties()
    >>> props
    <docx.footnotes.FootnoteProperties object at 0x02468ACE>

|FootnoteProperties| exposes four read/write properties, each backed by a child
element of ``w:footnotePr``. Assigning |None| removes the underlying element,
restoring Word's default for that aspect::

    >>> from docx.enum.text import (
    ...     WD_FOOTNOTE_POSITION,
    ...     WD_FOOTNOTE_RESTART,
    ...     WD_NUMBER_FORMAT,
    ... )
    >>> props.number_format = WD_NUMBER_FORMAT.LOWER_ROMAN
    >>> props.start_number = 1
    >>> props.restart_rule = WD_FOOTNOTE_RESTART.EACH_SECTION
    >>> props.position = WD_FOOTNOTE_POSITION.BENEATH_TEXT

* :attr:`~.FootnoteProperties.number_format` — a :ref:`WdNumberFormat` member
  selecting the glyph family for the reference marks. Common choices are
  ``DECIMAL`` (1, 2, 3 ...), ``UPPER_ROMAN``, ``LOWER_ROMAN``, ``UPPER_LETTER``,
  and ``CHICAGO`` (the ``*``, ``†``, ``‡``, ``§`` cycle).
* :attr:`~.FootnoteProperties.start_number` — the integer at which numbering
  begins. Usually ``1``; set higher when continuing a numbering scheme across
  documents.
* :attr:`~.FootnoteProperties.restart_rule` — a :ref:`WdFootnoteRestart`
  member that controls whether numbering runs continuously
  (``CONTINUOUS``), restarts at each section (``EACH_SECTION``), or restarts
  on every page (``EACH_PAGE``).
* :attr:`~.FootnoteProperties.position` — a :ref:`WdFootnotePosition` member
  that places footnotes either at the page bottom (``BOTTOM_OF_PAGE``, the
  default) or immediately beneath the last line of body text
  (``BENEATH_TEXT``).

Section-level overrides are also supported via
:attr:`.Section.footnote_properties` and :meth:`.Section.add_footnote_properties`,
which accept the same |FootnoteProperties| API. When a section defines its own
``w:footnotePr`` it takes precedence over the document-level element for that
section.
