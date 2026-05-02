.. _toc:

Working with a Table of Contents
================================

Word documents commonly begin with a *table of contents* (TOC) — a navigable
list of the document's headings with their page numbers. In Word, a TOC is
not a static piece of prose: it is a **field**, a small instruction that Word
evaluates every time the document is opened (or when the user presses *F9*
to update fields). The field scans the document for headings and renders
itself as an up-to-date list.

*python-docx* lets you write that field into a document. It cannot evaluate
the field — it has no layout engine and therefore no real page numbers — so
what it writes is the instruction Word needs plus a *cached preview* that
consumers who do not evaluate fields (raw-XML tools, a Word session where the
user declined the "update fields?" prompt) can still display sensibly.


Adding a TOC at the end of the document
---------------------------------------

The simplest way to add a TOC is to append one to the end of the document::

    >>> from docx import Document
    >>> document = Document()
    >>> document.add_heading("Chapter One", level=1)
    >>> document.add_paragraph("Body of chapter one.")
    >>> document.add_heading("Section 1.1", level=2)
    >>> document.add_paragraph("Body of section 1.1.")

    >>> toc = document.add_table_of_contents()
    >>> toc
    <docx.text.paragraph.Paragraph object at 0x7f0...>

The returned object is the newly-appended |Paragraph|. It now carries one
complex field of type ``TOC``::

    >>> field = toc.fields[0]
    >>> field.type
    'TOC'
    >>> field.is_complex
    True

When Word opens the file it offers to update fields. Accepting rebuilds the
TOC against the current document state and inserts the real page numbers.
Declining leaves the cached preview visible.


Choosing which heading levels appear
------------------------------------

``add_table_of_contents`` accepts a ``levels`` keyword — a
``(min_level, max_level)`` tuple that selects which ``"Heading N"`` styles
feed into the TOC. The default ``(1, 3)`` matches Word's own default and
includes H1 through H3::

    >>> # H1 only
    >>> document.add_table_of_contents(levels=(1, 1))

    >>> # H2 and H3, skipping top-level chapter titles
    >>> document.add_table_of_contents(levels=(2, 3))

    >>> # every heading level Word supports
    >>> document.add_table_of_contents(levels=(1, 9))

The range is validated. ``1 <= min_level <= max_level <= 9`` must hold;
otherwise a |ValueError| is raised. A paragraph is treated as a heading only
when its style name matches ``"Heading N"`` (case-insensitive) for ``N`` in
1..9. Paragraphs styled *Title*, *Subtitle*, or custom heading styles do
not contribute.


Inserting a TOC at a specific position
--------------------------------------

A TOC is often placed near the start of the document rather than at the end.
Use the paragraph-level insertion methods to place a TOC relative to an
existing paragraph::

    >>> anchor = document.paragraphs[0]
    >>> toc = anchor.insert_table_of_contents_before()
    >>> document.paragraphs[0] is toc
    True

    >>> toc_after = anchor.insert_table_of_contents_after()

Both methods accept the same ``levels`` keyword as
``Document.add_table_of_contents``. The preview text scans the entire
document body regardless of where the TOC is inserted — Word itself rebuilds
the list on open, so the cached preview covers *all* headings even if some
appear after the TOC paragraph.


What ends up in the XML
-----------------------

A TOC is a *complex field*: three ``<w:fldChar>`` markers (``begin``,
``separate``, ``end``) wrap an ``<w:instrText>`` instruction and a cached
result. The generated paragraph looks approximately like:

.. code-block:: xml

    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r>
        <w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>
      </w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t xml:space="preserve">Chapter One&#9;1</w:t></w:r>
      <w:r><w:br/></w:r>
      <w:r><w:t xml:space="preserve">Section 1.1&#9;2</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>

The instruction switches match what Word writes when a TOC is inserted from
the *References* ribbon:

* ``\o "min-max"`` — build from outline levels ``min..max``
* ``\h`` — render entries as clickable hyperlinks
* ``\z`` — hide the tab leader and page number in web view
* ``\u`` — include paragraphs with an applied outline level, not only those
  using the built-in heading styles

You can read the instruction and result back via the |Field| proxy::

    >>> field = toc.fields[0]
    >>> field.instruction
    ' TOC \\o "1-3" \\h \\z \\u '
    >>> field.result_text
    'Chapter One\t1\nSection 1.1\t2'

Each line of ``result_text`` has the form ``"{heading text}\t{index}"``. The
trailing integer is a **1-based position** in the filtered heading list, not
a page number. *python-docx* does not paginate, so it cannot compute page
numbers; Word discards the cached numbers and recomputes real ones when it
next updates the field.


Verifying the result in Word
----------------------------

Because the TOC is a field, what you see in Word depends on whether fields
are up to date:

1. Open the document in Word. Word prompts *"This document contains fields
   that may refer to other files. Do you want to update the fields?"*.
2. Click **Yes**. Word scans the document, rebuilds the TOC entries, and
   inserts the real page numbers. The result matches what you would see if
   you inserted a TOC via *References > Table of Contents* in the Word UI.
3. Click **No** and the cached preview written by *python-docx* is shown
   instead — heading text is correct, but the trailing integers are heading
   indexes rather than page numbers.

You can also force an update at any time: click inside the TOC and press
*F9*, or right-click and choose *Update Field*.


API reference
-------------

* :meth:`docx.document.Document.add_table_of_contents` — append a TOC to the
  body.
* :meth:`docx.text.paragraph.Paragraph.insert_table_of_contents_before` and
  :meth:`~docx.text.paragraph.Paragraph.insert_table_of_contents_after` —
  insert a TOC relative to an existing paragraph.
* :mod:`docx.toc` — lower-level helpers
  (:func:`~docx.toc.build_toc_instruction`,
  :func:`~docx.toc.populate_toc_paragraph`) exposed for advanced use.
