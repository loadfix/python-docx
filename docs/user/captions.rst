.. _captions:

Captions for figures, tables, and equations
===========================================

Word's **References > Insert Caption** menu produces a caption paragraph of
the form

.. parsed-literal::

   Figure 1: A diagram of the system

where ``Figure`` is a label, ``1`` is an auto-number maintained by Word, and
``A diagram of the system`` is the caption text. The auto-number is driven
by a :ref:`SEQ <fields>` field so that adding, deleting, or reordering
captions automatically renumbers them when the document is next opened.

|docx| provides three entry points for authoring captions:

- :meth:`.Document.add_caption` — append a caption at the end of the document.
- :meth:`.Paragraph.add_caption_before` — insert a caption directly above an
  existing paragraph (typical for tables, where the caption sits above the
  table).
- :meth:`.Paragraph.add_caption_after` — insert a caption directly below an
  existing paragraph (typical for figures).


Anatomy of a caption
--------------------

Every caption emitted by |docx| has the same structure::

    <w:p>
      <w:pPr><w:pStyle w:val="Caption"/></w:pPr>
      <w:r><w:t xml:space="preserve">Figure </w:t></w:r>
      <w:fldSimple w:instr=' SEQ Figure \* ARABIC '>
        <w:r><w:t>1</w:t></w:r>
      </w:fldSimple>
      <w:r><w:t xml:space="preserve">: </w:t></w:r>
      <w:r><w:t>A diagram of the system</w:t></w:r>
    </w:p>

The ``1`` inside the ``w:fldSimple`` is a *cached* field result: when Word
reopens the document it will replace that value with the correct
auto-number. For non-Word consumers (spell-checkers, text extractors) the
cached result gives a reasonable fallback.


Adding a caption at the end of the document
-------------------------------------------

::

    >>> from docx import Document
    >>> document = Document()
    >>> caption = document.add_caption("A diagram of the system", label="Figure")
    >>> caption.style.name
    'Caption'
    >>> caption.text
    'Figure 1: A diagram of the system'

The returned |Paragraph| is the freshly-appended caption; modify its runs in
the usual way to add bold, italics, or other formatting.


Captioning tables and figures in place
--------------------------------------

Captions rarely belong at the end of the document. The more common pattern
is to add the figure or table first, then attach a caption immediately above
or below using the paragraph-level helpers::

    >>> p = document.add_paragraph()
    >>> p.add_run().add_picture("diagram.png")
    >>> p.add_caption_after("A diagram of the system", label="Figure")

    >>> heading = document.add_paragraph()
    >>> heading.add_caption_before("Quarterly results", label="Table")
    >>> heading.add_run().add_table(...)   # hypothetical

Both helpers return the inserted caption paragraph so the caller can chain
additional mutations.


Label grouping
--------------

Each distinct `label` argument defines an independent numbering sequence.
Word maintains one counter per SEQ identifier; so the first ``Figure``
caption is numbered ``1``, the first ``Table`` caption is also ``1``, and
the second ``Figure`` caption is ``2``. ``label`` is what controls the
counter, not the paragraph style.

Common labels are ``"Figure"``, ``"Table"``, and ``"Equation"``, but any
string Word will accept as a SEQ identifier is permitted — callers adding
localised captions can pass ``"Figura"``, ``"Таблица"``, or similar.


Customising the paragraph style
-------------------------------

The `style` parameter defaults to ``"Caption"``, which is the built-in
Word style for captions. A different style can be supplied when a document
uses a custom caption style::

    >>> document.add_caption(
    ...     "Performance metrics",
    ...     label="Figure",
    ...     style="FigureCaption",
    ... )

The style must already be defined in the document; |docx| does not
synthesise it for you.


Limitations
-----------

- |docx| does not implement a layout engine, so the cached ``1`` inside the
  SEQ field is always emitted literally. Word will rewrite this to the
  correct auto-number on open.
- There is no ``paragraph.caption`` read-side accessor; to enumerate
  captions, filter paragraphs by their style name::

        captions = [p for p in document.paragraphs if p.style.name == "Caption"]

- Cross-references to a caption (``Figure 1`` as a clickable link elsewhere
  in the document) require a bookmark and a REF field; neither is created
  by :meth:`add_caption`. Use :meth:`Paragraph.add_bookmark` and the
  :mod:`docx.fields` module to wire those up manually.
