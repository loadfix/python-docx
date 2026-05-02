.. _endnotes:

Working with Endnotes
=====================

Word allows *endnotes* to be added to a document. An endnote is a piece of
reference material whose body appears at the end of the document (or end of a
section) while a numbered marker is inserted into the running text where the
citation occurs. Endnotes are generally used for citations or extended
remarks that would otherwise distract from the flow of the main document.

The procedure is simple:

- You place the cursor at the spot where you want the endnote reference mark
  to appear.
- You press the *Insert Endnote* button on the References toolbar.
- You type or paste in the endnote text, which is stored in a separate
  *endnotes part* at the bottom of the document.

**Endnote Anatomy.** Each endnote has two parts, the *endnote-reference* and
the *endnote-content*:

The **endnote-reference**, sometimes *endnote-anchor*, is the small
superscript mark placed into the main document where the endnote was
inserted. It is a single ``<w:endnoteReference>`` element carrying the *id*
of the endnote it anchors, wrapped in a run styled with the
"EndnoteReference" character style.

The **endnote-content**, sometimes just *endnote*, is whatever content was
typed or pasted in. The content for each endnote lives in a separate endnote
object, and these endnote objects are stored in a separate *endnotes part*
(part-name ``word/endnotes.xml``), not in the main document. Each endnote is
assigned a unique id when it is created, allowing the endnote reference to
be associated with its content and vice versa.

**Reserved Ids.** Endnote ids 0 and 1 are reserved — they identify the
*separator* and *continuation-separator* markers that Word uses to draw the
horizontal rule between the document body and the endnotes area. User
endnotes added by *python-docx* receive ids starting at 2. These reserved
entries are filtered out of iteration and are not counted by ``len()``.

**Endnote Content.** Although most endnotes contain a single paragraph of
plain text, an endnote is a *block-item container* — it can contain multiple
paragraphs and tables, and runs within paragraphs can carry character
emphasis such as bold or italic, embedded hyperlinks, and images.

**Endnote Properties.** Document-level endnote numbering is controlled by a
``<w:endnotePr>`` element which lives inside the settings part. Through the
|EndnoteProperties| proxy you can configure:

- *number_format* — the numeral style used for endnote marks (Arabic,
  Roman, lowercase letters, Chicago-style marks, etc.)
- *start_number* — the first number used for automatic numbering.
- *restart_rule* — when numbering resets (continuous or at each section).
- *position* — where the endnote body appears (end of document or end of
  section).

**Applicability.** Endnotes can only be added to the main document body.
An endnote cannot be added to a header, a footer, a footnote, a comment, or
nested inside another endnote. In general the *python-docx* API will not
allow these operations, but if you outsmart it the resulting endnote will
either be silently removed or trigger a repair error when the document is
loaded by Word.


Adding an endnote
-----------------

A simple example is adding an endnote anchored to a run::

    >>> from docx import Document
    >>> document = Document()
    >>> paragraph = document.add_paragraph("Hello, world.")
    >>> run = paragraph.runs[-1]

    >>> endnote = document.endnotes.add(run, text="See the appendix for details.")
    >>> endnote
    <docx.endnotes.Endnote object at 0x02468ACE>
    >>> endnote.endnote_id
    2
    >>> endnote.text
    'See the appendix for details.'

The :meth:`.Endnotes.add` call inserts a ``<w:endnoteReference>`` into the
supplied run, styled with the "EndnoteReference" character style, and
creates a new ``<w:endnote>`` element in the endnotes part whose first
paragraph carries the ``EndnoteText`` paragraph style. If ``text`` is
provided, it is added as a run in that paragraph following the auto-number
mark.


Accessing and iterating the Endnotes collection
-----------------------------------------------

The endnotes collection is accessed via the :attr:`.Document.endnotes`
property::

    >>> endnotes = document.endnotes
    >>> endnotes
    <docx.endnotes.Endnotes object at 0x02468ACE>
    >>> len(endnotes)
    1

The |Endnotes| object is iterable over user endnotes; the reserved
separator entries are skipped::

    >>> for endnote in document.endnotes:
    ...     print(endnote.endnote_id, endnote.text)
    2 See the appendix for details.


Adding rich content to an endnote
---------------------------------

An endnote is a *block-item container*, just like the document body or a
table cell, so it can contain any content those places can. The methods for
adding this content are the same as those used for the document and table
cells::

    >>> endnote = document.endnotes.add(run, text="")
    >>> endnote.add_paragraph("A longer citation follows.")
    >>> end_para = endnote.paragraphs[0]
    >>> end_para.add_run(" See p. 42.").italic = True


Deleting an endnote
-------------------

To remove an endnote from the document, call :meth:`.Endnote.delete`. This
removes both the ``<w:endnote>`` element from the endnotes part and the
``<w:endnoteReference>`` run from the main document body::

    >>> endnote = document.endnotes.add(paragraph.runs[-1], text="Temporary note.")
    >>> endnote.delete()

After calling :meth:`.Endnote.delete` the |Endnote| object is *defunct* and
should not be used further.


Configuring endnote numbering and position
------------------------------------------

Document-level endnote properties are accessed via
:attr:`.Document.endnote_properties`. The property returns |None| when no
``w:endnotePr`` element exists in the document settings; use
:meth:`.Document.add_endnote_properties` to add one and configure it::

    >>> from docx.enum.text import (
    ...     WD_ENDNOTE_POSITION, WD_FOOTNOTE_RESTART, WD_NUMBER_FORMAT,
    ... )
    >>> props = document.add_endnote_properties()
    >>> props.number_format = WD_NUMBER_FORMAT.LOWER_ROMAN
    >>> props.restart_rule = WD_FOOTNOTE_RESTART.EACH_SECTION
    >>> props.position = WD_ENDNOTE_POSITION.END_OF_SECTION
    >>> props.start_number = 1

All four properties are read/write. Assigning |None| removes the
corresponding child element from the ``w:endnotePr`` so that Word falls
back to its default behaviour.
