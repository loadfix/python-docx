.. _fields:

Working with Fields
===================

Word supports so-called *field codes* — short instructions such as ``PAGE``,
``DATE``, ``AUTHOR``, or ``REF bookmark1 \h`` that Word evaluates at display
time to produce some rendered text. The rendered result is cached in the
document alongside the instruction; Word refreshes the cache when you press
*F9* or when the document is reopened.

A field therefore has two observable pieces:

* the **instruction** — what the field will evaluate, e.g. ``"PAGE"`` or
  ``"REF FavouriteValue \\h"``
* the **result text** — the cached rendered value, e.g. ``"7"`` or
  ``"The quoted value is forty-two."``

WordprocessingML represents fields two different ways and *python-docx*
exposes both forms behind a single :class:`~docx.fields.Field` proxy.


Simple vs. complex fields
-------------------------

**Simple fields** use a single ``<w:fldSimple>`` block. The instruction is
stored in the ``w:instr`` attribute and the rendered result is held in one or
more ``<w:r>`` run children::

    <w:fldSimple w:instr="DATE">
      <w:r><w:t>2025-01-02</w:t></w:r>
    </w:fldSimple>

**Complex fields** split the same information across a sequence of runs
delimited by three ``<w:fldChar>`` markers — ``begin``, ``separate``, and
``end``. The instruction lives in an ``<w:instrText>`` element between
``begin`` and ``separate``; the rendered result is the plain text of the runs
between ``separate`` and ``end``::

    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText>PAGE</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>7</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>

Word prefers the complex form for anything non-trivial (fields with switches,
nested fields, form-field controls), but consumer software must handle both.
*python-docx* reads either and you can choose which to write.

**Applicability.** Fields occur inside paragraphs. A simple field is a
block-level child of ``w:p``; a complex field is a sequence of ``w:r`` runs
inside a ``w:p``. Fields are supported in the document body, inside table
cells, and inside headers, footers, and comments.


The Field proxy
---------------

Every field — simple or complex — is wrapped by the same
:class:`~docx.fields.Field` class. The three read-only properties that matter
day-to-day are :attr:`~docx.fields.Field.instruction`,
:attr:`~docx.fields.Field.type`, and :attr:`~docx.fields.Field.result_text`::

    >>> from docx import Document
    >>> document = Document("my-doc.docx")
    >>> paragraph = document.paragraphs[2]

    >>> field = paragraph.fields[0]
    >>> field
    <docx.fields.Field object at 0x02468ACE>
    >>> field.is_complex
    False
    >>> field.instruction
    'DATE'
    >>> field.type
    'DATE'
    >>> field.result_text
    '2025-01-02'

:attr:`~docx.fields.Field.type` is the convenient shorthand — it is just the
first whitespace-delimited token of :attr:`~docx.fields.Field.instruction`.


Field-type constants
--------------------

The :class:`~docx.fields.WD_FIELD_TYPE` class collects common field-type
tokens as string constants, purely to help with autocomplete and typo
avoidance::

    >>> from docx.fields import WD_FIELD_TYPE
    >>> WD_FIELD_TYPE.PAGE
    'PAGE'
    >>> WD_FIELD_TYPE.REF
    'REF'

These are plain strings (not a real :class:`enum.Enum`) because the set of
field-type tokens used in real-world documents is open-ended; custom field
types simply round-trip as whatever string appears in the document.


Adding a field to a paragraph
-----------------------------

You can append either a simple or a complex field to an existing paragraph::

    >>> paragraph = document.add_paragraph("Today is ")

    >>> # -- simple form: one <w:fldSimple> element --
    >>> field = paragraph.add_simple_field(WD_FIELD_TYPE.DATE, "2025-01-02")
    >>> field.is_complex
    False
    >>> field.result_text
    '2025-01-02'

    >>> # -- complex form: begin/separate/end run sequence --
    >>> field = paragraph.add_complex_field(WD_FIELD_TYPE.PAGE, "7")
    >>> field.is_complex
    True
    >>> field.result_text
    '7'

The `text` / `result_text` parameter is optional. When omitted the field is
written with no cached result — Word (or another consumer) will populate it
the first time the field is evaluated.


Iterating fields in a document
------------------------------

Every paragraph exposes its fields in document order via
:attr:`Paragraph.fields <docx.text.paragraph.Paragraph.fields>`. To walk
every field in the body, iterate the paragraphs::

    >>> for paragraph in document.paragraphs:
    ...     for field in paragraph.fields:
    ...         print(field.type, repr(field.result_text))
    DATE '2025-01-02'
    PAGE '7'
    REF 'The quoted value is forty-two.'

:attr:`Paragraph.fields` returns both simple and complex fields in a single
flat list ordered by XML position. Fields inside tables, headers, footers, or
comments are reached by iterating the paragraphs of those containers directly.


Updating the rendered result
----------------------------

:meth:`Field.update_result_text() <docx.fields.Field.update_result_text>`
replaces the cached result text in place without disturbing the instruction::

    >>> field = paragraph.fields[0]
    >>> field.update_result_text("42")
    >>> field.result_text
    '42'

For a simple field this rewrites the run(s) inside the ``<w:fldSimple>``
element. For a complex field it replaces the runs between the ``separate``
and ``end`` markers with a single new run. If a complex field has no
``separate`` marker the call is a no-op — there is nowhere to write the
rendered text.


Cross-reference resolution (REF / PAGEREF)
------------------------------------------

``REF`` fields point at a bookmark elsewhere in the document; ``PAGEREF``
fields reference the page number that a bookmark falls on. *python-docx*
cannot compute real page numbers — it has no layout engine — but it can
resolve ``REF`` fields against the bookmark's current text using
:meth:`Field.resolve() <docx.fields.Field.resolve>`::

    >>> paragraph = document.add_paragraph("The quoted value is forty-two.")
    >>> paragraph.add_bookmark(
    ...     "FavouriteValue",
    ...     start_run=paragraph.runs[0],
    ...     end_run=paragraph.runs[0],
    ... )

    >>> ref_para = document.add_paragraph("As noted earlier: ")
    >>> ref_field = ref_para.add_complex_field("REF FavouriteValue \\h")
    >>> ref_field.resolve(document)
    'The quoted value is forty-two.'

:meth:`Field.resolve` is best-effort and never raises. For field types it
does not understand (``PAGE``, ``DATE``, ``SEQ``, custom fields, …) it simply
returns the existing :attr:`~docx.fields.Field.result_text`. For a
``PAGEREF`` whose cached result is empty it returns ``"?"``.

To rewrite every ``REF`` and ``PAGEREF`` result in the document body in one
go, use :meth:`Document.resolve_cross_references
<docx.document.Document.resolve_cross_references>`::

    >>> updated = document.resolve_cross_references()
    >>> updated
    1

The return value is the number of fields whose cached result was actually
changed. Fields whose cached result already matched the bookmark text — or
whose bookmark could not be found — are skipped.


A note about form fields
------------------------

A **form field** is a specific kind of complex field whose ``begin`` marker
carries a ``<w:ffData>`` child describing a text input, checkbox, or dropdown.
Form fields are presented through a dedicated
:class:`~docx.form_fields.FormField` proxy and are accessible via
:attr:`Document.form_fields <docx.document.Document.form_fields>`, not via the
``fields`` collection. Non-form complex fields (``PAGE``, ``REF``, …) appear
only in :attr:`Paragraph.fields` and the two collections are disjoint.
