.. _text_advanced:

Advanced Text Formatting
========================

This guide covers text-formatting features that go beyond the ``bold`` /
``italic`` / ``underline`` basics described in :doc:`text`.
Everything here lives on the |Font| object exposed by ``run.font`` or on the
|ParagraphFormat| exposed by ``paragraph.paragraph_format``.

All properties described here follow the same tri-state convention used
elsewhere in *python-docx*: reading |None| means the attribute is absent and
the effective value is inherited from the style hierarchy. Assigning |None|
removes the direct setting so the style default re-applies.


Run shading (background color)
------------------------------

``Font.shading_color`` provides a read/write |RGBColor| for the *run-level*
background fill. It maps to ``w:rPr/w:shd@w:fill`` with ``w:val="clear"``. It
is different from ``Font.highlight_color``, which selects from a fixed palette
of highlighter colors (``WD_COLOR_INDEX.YELLOW`` etc.) and maps to
``w:rPr/w:highlight``::

    >>> from docx import Document
    >>> from docx.shared import RGBColor
    >>> from docx.enum.text import WD_COLOR_INDEX
    >>> run = Document().add_paragraph().add_run("shaded text")

    >>> run.font.shading_color is None
    True
    >>> run.font.shading_color = RGBColor(0xFF, 0xFF, 0x00)
    >>> run.font.shading_color
    RGBColor(0xff, 0xff, 0x00)

    >>> run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
    >>> run.font.shading_color, run.font.highlight_color
    (RGBColor(0xff, 0xff, 0x00), BRIGHT_GREEN (4))

Setting ``shading_color`` to |None| removes the ``w:shd`` element entirely.
The two properties are independent, so both may be set on the same run.


Run borders
-----------

Word can draw a box around a single run. The border is controlled by the
``w:rPr/w:bdr`` element and is exposed on |Font| as four symmetrical
properties plus a convenience :meth:`~docx.text.font.Font.remove_border`
method::

    >>> from docx.enum.text import WD_BORDER_STYLE
    >>> from docx.shared import Pt, RGBColor
    >>> font = run.font

    >>> font.border_style = WD_BORDER_STYLE.SINGLE
    >>> font.border_color = RGBColor(0xFF, 0x00, 0x00)
    >>> font.border_width = Pt(1.5)
    >>> font.border_space = Pt(4)

* ``border_style`` — a :ref:`WdBorderStyle` member (``SINGLE``, ``DOUBLE``,
  ``DASHED``, ``DOTTED``, and more).
* ``border_color`` — an |RGBColor| or |None|. Reading returns |None| when the
  XML stores ``w:color="auto"`` so assigning a real color is distinguishable
  from inheritance.
* ``border_width`` — a |Length|. Word stores this as eighth-points; use
  |Pt| to get the right units (``Pt(0.5)``, ``Pt(1)``, ``Pt(1.5)`` etc.).
* ``border_space`` — a |Length| controlling the padding between the border
  and the text, typically entered in points.

Assigning |None| to any individual property clears just that attribute
while leaving the others intact. To clear the whole border in one call use
:meth:`Font.remove_border`::

    >>> font.remove_border()
    >>> font.border_style, font.border_color, font.border_width, font.border_space
    (None, None, None, None)


Kerning and character spacing
-----------------------------

Two closely related Font properties expose character-metric adjustments:

* ``Font.kerning`` is the *minimum* font size, in points, for which the Word
  rendering engine will perform automatic kerning. Set it with |Pt|::

      >>> font.kerning = Pt(10)
      >>> font.kerning.pt
      10.0

  Assigning |None| removes the ``w:kern`` element.

* ``Font.character_spacing`` is a fixed horizontal offset between characters
  in the run. Positive values expand the tracking, negative values condense
  it::

      >>> font.character_spacing = Pt(1)      # wider
      >>> font.character_spacing = Pt(-0.5)   # tighter
      >>> font.character_spacing = None       # back to inheritance


Language tags and East Asian fonts
----------------------------------

A run's ``w:rPr/w:lang`` element carries up to three BCP-47 language tags,
each surfaced as an independent property on |Font|:

* ``Font.language`` — primary (Latin-script) language, e.g. ``"en-US"``.
* ``Font.east_asian_language`` — East Asian language, e.g. ``"ja-JP"``.
* ``Font.bidi_language`` — complex-script (right-to-left) language, e.g.
  ``"ar-SA"``.

Because all three attributes share the same element, assigning |None| to an
individual property clears only the corresponding attribute. To drop the
entire ``w:lang`` element use :meth:`Font.remove_language`::

    >>> font.language = "en-US"
    >>> font.east_asian_language = "ja-JP"
    >>> font.bidi_language = "ar-SA"
    >>> font.remove_language()
    >>> font.language, font.east_asian_language, font.bidi_language
    (None, None, None)

Each script can also use a different typeface. ``Font.name`` drives the
primary (ASCII / high-ANSI) face, and ``Font.name_far_east`` sets the East
Asian face that appears in CJK (Chinese / Japanese / Korean) runs. A legacy
alias ``Font.name_east_asia`` is kept for symmetry with ECMA-376
terminology; both spellings read and write the same attribute::

    >>> font.name_far_east = "MS Mincho"
    >>> font.name_east_asia
    'MS Mincho'


East Asian typography
---------------------

``Font.east_asian_layout`` returns an |EastAsianLayout| proxy when the run
carries a ``w:rPr/w:eastAsianLayout`` child, or |None| otherwise. The proxy
exposes three booleans and an integer id:

* ``two_lines_in_one`` — collapses two adjacent characters into a single
  double-glyph (``w:combine``).
* ``vertical_alignment`` — rotates the run for vertical layout
  (``w:vert``).
* ``compressed`` — when vertical, compress the run (``w:vertCompress``).
* ``id`` — numeric id Word uses to group related layout runs.

Create or update the element with :meth:`Font.set_east_asian_layout`; drop it
with :meth:`Font.remove_east_asian_layout`::

    >>> font.set_east_asian_layout(two_lines_in_one=True, id=1)
    <docx.text.font.EastAsianLayout object at 0x7f...>
    >>> font.east_asian_layout.two_lines_in_one
    True

    >>> font.remove_east_asian_layout()
    >>> font.east_asian_layout is None
    True

Two paragraph-level toggles complete the East Asian story. Both live on
|ParagraphFormat| and are tri-state (|True| / |False| / |None|):

* ``kinsoku`` (``w:kinsoku``) — apply kinsoku shori line-break rules so that
  certain punctuation characters may not begin or end a line.
* ``word_wrap`` (``w:wordWrap``) — |True| wraps Latin text on word
  boundaries (the default); |False| allows breaks inside a word to keep a
  tight right edge, which is typical in Japanese layout.


Ruby (phonetic) annotations
---------------------------

A *ruby annotation* pairs a run of base text with a smaller above-the-line
annotation. Japanese furigana is the most common example.

*python-docx* exposes existing ruby annotations as read-only
|RubyAnnotation| objects via :attr:`Run.ruby_annotations`. The API does not
yet *create* ruby runs::

    >>> document = Document("sample-with-ruby.docx")
    >>> run = document.paragraphs[0].runs[0]
    >>> for ruby in run.ruby_annotations:
    ...     print(f"{ruby.base_text!r}  ↑  {ruby.ruby_text!r}")
    '日本'  ↑  'にほん'
    '東京'  ↑  'とうきょう'

    >>> ruby = run.ruby_annotations[0]
    >>> ruby.alignment, ruby.language
    ('distributeSpace', 'ja-JP')

``alignment`` is the raw value of ``w:rubyPr/w:rubyAlign@w:val`` (typical
values include ``distributeLetter``, ``distributeSpace``, ``center``,
``left``, ``right``, ``rightVertical``). ``language`` is the value of
``w:rubyPr/w:lid@w:val``, usually a BCP-47 tag.

The base text of a ``w:ruby`` also contributes to ``Run.text``, so
``paragraph.text`` stays readable even for paragraphs that contain ruby
markup.


Right-to-left (bidi) layout
---------------------------

Right-to-left rendering is controlled at *two* independent scopes; do not
confuse them with the section-level ``w:bidi`` that flips an entire page
layout.

**Run-level RTL.** ``Font.right_to_left`` (boolean) corresponds to
``w:rPr/w:rtl``. Setting it to |True| causes the run to be rendered
right-to-left using the complex-script (CS) font::

    >>> run = document.add_paragraph().add_run("שלום")
    >>> run.font.right_to_left = True

``Font.rtl`` exposes the same element as a *tri-state* (|True| / |False| /
|None|), following the style-inheritance convention used by other boolean
Font properties.

**Paragraph-level RTL.** ``ParagraphFormat.right_to_left`` controls
``w:pPr/w:bidi``. Flipping it reverses the visual order of any runs the
paragraph contains and mirrors paragraph-level indents::

    >>> p = document.add_paragraph("مرحبا")
    >>> p.paragraph_format.right_to_left = True

Assigning |False| or |None| removes the ``w:bidi`` element.


Symbols (glyphs from a named font)
----------------------------------

``Run.add_symbol`` appends a ``w:sym`` element that draws its glyph from a
named font rather than from the run's main typeface. Word uses this to
render Wingdings, bullet glyphs that are not standard Unicode, and similar
special characters::

    >>> run = document.add_paragraph().add_run()
    >>> sym = run.add_symbol(0xF0E0, "Wingdings")
    >>> sym.char_hex, sym.font
    ('F0E0', 'Wingdings')

The ``char_code`` argument accepts either an integer (``0xF0E0``) or a hex
string (``"F0E0"``, ``"0xf0e0"``). The XML always stores it as a 4-character
uppercase hex string; ``Symbol.char_hex`` returns that canonical form.

All symbols in a run are iterable via :attr:`Run.symbols`::

    >>> run.add_symbol(0xF0E1, "Wingdings")
    >>> [s.char_hex for s in run.symbols]
    ['F0E0', 'F0E1']
    >>> sym.delete()


Paragraph text frames
---------------------

A *text frame* is an absolutely-positioned text container, the legacy
predecessor of the modern text box. A frame is attached to a paragraph via a
``w:pPr/w:framePr`` element and carries a dozen size, position, and layout
attributes.

``ParagraphFormat.frame`` returns a read-only |TextFrame| proxy when the
element is present, or |None| otherwise. Use
:meth:`ParagraphFormat.set_frame` to create or update a frame, and
:meth:`ParagraphFormat.remove_frame` to detach it::

    >>> from docx.enum.text import (
    ...     WD_FRAME_H_ANCHOR, WD_FRAME_V_ANCHOR, WD_FRAME_WRAP
    ... )
    >>> from docx.shared import Inches

    >>> p = document.add_paragraph("Floating paragraph.")
    >>> frame = p.paragraph_format.set_frame(
    ...     width=Inches(3),
    ...     height=Inches(1),
    ...     horizontal_position=Inches(0.5),
    ...     vertical_position=Inches(0.75),
    ...     horizontal_anchor=WD_FRAME_H_ANCHOR.PAGE,
    ...     vertical_anchor=WD_FRAME_V_ANCHOR.MARGIN,
    ...     wrap=WD_FRAME_WRAP.AROUND,
    ... )
    >>> frame.width, frame.height
    (2743200, 914400)

Any keyword argument left at its default of |None| is left unchanged when the
frame already exists, so :meth:`set_frame` doubles as an in-place update::

    >>> p.paragraph_format.set_frame(width=Inches(4))
    >>> p.paragraph_format.frame.width
    3657600

The TextFrame properties cover the full attribute surface of ``w:framePr``:

* ``width`` / ``height`` — |Length|.
* ``horizontal_position`` / ``vertical_position`` — |Length|.
* ``horizontal_anchor`` / ``vertical_anchor`` — members of
  :ref:`WdFrameHAnchor` and :ref:`WdFrameVAnchor`.
* ``wrap`` — a :ref:`WdFrameWrap` member.
* ``drop_cap`` and ``lines`` — for drop-cap frames, via
  :ref:`WdFrameDropCap`.
* ``horizontal_alignment`` / ``vertical_alignment`` — :ref:`WdFrameHAlign`
  and :ref:`WdFrameVAlign`.

Assigning |None| to any individual attribute clears just that attribute on
the existing ``w:framePr`` element. To drop the element entirely use
:meth:`~docx.text.parfmt.ParagraphFormat.remove_frame`::

    >>> p.paragraph_format.remove_frame()
    >>> p.paragraph_format.frame is None
    True
