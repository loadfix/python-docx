.. _numbering:

Numbering and list formatting
=============================

Bulleted and numbered lists in WordprocessingML are not authored per-paragraph
the way they are in the UI. Instead, Word stores **numbering definitions**
in the ``word/numbering.xml`` part and each paragraph that participates in a
list *references* a definition by id. This indirection is what lets Word
renumber automatically when paragraphs are inserted, deleted, or moved.

|docx| provides a three-layer proxy API over the numbering part:

- :class:`.Numbering` — the top-level collection, available as
  :attr:`Document.numbering <docx.document.Document.numbering>`.
- :class:`.NumberingDefinition` — a single ``w:abstractNum`` element that
  describes the visual format of a list (its level text, indentation, number
  format, font).
- :class:`.Level` — one ``w:lvl`` child of a |NumberingDefinition|, one per
  indent level (levels 0 through 8 are permitted by the spec).


Anatomy of a list
-----------------

The numbering part holds two kinds of element:

- ``w:abstractNum`` describes formatting (indentation, number format, level
  text pattern, font).
- ``w:num`` is a concrete *instance* that points at an ``w:abstractNum`` and
  can optionally override its starting value.

A paragraph joins a list by carrying two attributes inside its ``w:numPr``:

- ``w:numId`` — the id of a ``w:num`` instance.
- ``w:ilvl`` — the integer indent level (``0`` through ``8``).

|docx| hides the abstract/instance distinction behind
:meth:`.NumberingDefinition.apply_to`: you describe the *formatting* you
want, and python-docx allocates (or reuses) a matching ``w:num`` instance
internally.


Reading existing lists
----------------------

::

    >>> from docx import Document
    >>> document = Document("report.docx")
    >>> numbering = document.numbering
    >>> len(numbering)
    10
    >>> for definition in numbering:
    ...     print(definition.abstract_num_id, [lvl.number_format for lvl in definition.levels])

Each |NumberingDefinition| exposes the set of levels it declares::

    >>> definition = numbering.definitions[-1]
    >>> for level in definition.levels:
    ...     print(level.ilvl, level.number_format, level.text, level.start, level.indent)
    0 WD_NUMBER_FORMAT.DECIMAL %1. 5 228600
    1 WD_NUMBER_FORMAT.LOWER_LETTER %2) 1 457200
    2 WD_NUMBER_FORMAT.BULLET • 1 685800

:attr:`.Level.indent` is a :class:`.Length` (EMU); the other accessors are
straight strings or enumeration members.


Building a new numbering definition
-----------------------------------

Use :meth:`.Numbering.add_numbering_definition` to create a definition from a
sequence of per-level specifications. Each spec can be either a mapping or a
positional tuple::

    >>> from docx.enum.text import WD_NUMBER_FORMAT
    >>> from docx.shared import Inches
    >>> definition = document.numbering.add_numbering_definition([
    ...     {
    ...         "format": WD_NUMBER_FORMAT.DECIMAL,
    ...         "text": "%1.",
    ...         "indent": Inches(0.25),
    ...         "start": 1,
    ...     },
    ...     {
    ...         "format": "lowerLetter",        # string forms are accepted
    ...         "text": "%2)",
    ...         "indent": Inches(0.5),
    ...     },
    ...     {
    ...         "format": "bullet",
    ...         "text": "•",
    ...         "indent": Inches(0.75),
    ...         "font": "Symbol",               # required for bullet glyphs
    ...     },
    ... ])

``format`` accepts a :class:`.WD_NUMBER_FORMAT` member or a raw OOXML token string
(``"decimal"``, ``"bullet"``, ``"lowerLetter"``, ``"upperRoman"``, etc.).
``text`` is a :class:`str` template using ``%N`` placeholders where ``N`` is
1-based — ``%1.%2`` on level 1 produces ``"1.a"``, ``"1.b"``, and so on.
``indent`` accepts either a |Length| or a bare integer count of twips.
``font`` sets the ``w:rFonts`` on the level's run properties; it is usually
required for bullet lists that reference ``"•"`` or other non-Latin glyphs,
since Word's default body font often does not ship those shapes.


Applying a definition to paragraphs
-----------------------------------

:meth:`.NumberingDefinition.apply_to` joins a paragraph to the list and
selects its indent level::

    >>> p1 = document.add_paragraph("First point")
    >>> p2 = document.add_paragraph("Sub-point")
    >>> p3 = document.add_paragraph("Another sub-point")
    >>> definition.apply_to(p1, level=0)
    >>> definition.apply_to(p2, level=1)
    >>> definition.apply_to(p3, level=1)

Levels run 0 through 8; :meth:`apply_to` raises :class:`ValueError` for any
other value. The same definition can be applied to any number of paragraphs;
Word's numbering engine automatically renumbers them when the document is
opened.


Restart numbering
-----------------

The ``start`` key in the level spec sets the first number the list emits.
This is persisted on the ``w:start`` child of the ``w:lvl`` element::

    >>> definition = document.numbering.add_numbering_definition([
    ...     {"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%1.", "start": 5},
    ... ])
    >>> definition.levels[0].start
    5


Nested definitions
------------------

A "nested" list is simply a definition that declares more than one level.
Paragraphs at different levels within the same document reference the same
``w:num`` instance but with different ``w:ilvl`` values. The three-level
example above demonstrates the common pattern of decimal → lower-letter →
bullet for a technical outline.


Reading a paragraph's list membership
-------------------------------------

:attr:`Paragraph.list_format <docx.text.paragraph.Paragraph.list_format>`
returns a :class:`.ListFormat` named tuple of
``(numbering_definition, level)``. The definition is |None| for paragraphs
outside any list.


Limitations
-----------

- |docx| does not compute the *rendered* number for a paragraph — that is
  the job of Word's numbering engine, which runs at layout time.
- Per-instance starting-number overrides on ``w:num`` (the ``w:lvlOverride``
  mechanism) are not exposed by the proxy API; use the
  :attr:`~.NumberingDefinition.element` escape hatch for direct XML access.
- Modifying a level's formatting on an existing definition is not supported
  — create a new definition instead.
