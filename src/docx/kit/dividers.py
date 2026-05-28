"""Section dividers and chapter ornaments — fleurons and decorative breaks.

Closes #89.

Long-form documents (novels, books of essays, reports with distinct
sub-sections) routinely insert *section dividers* between paragraphs to
signal a scene change, a topic shift, or the end of a chapter without
forcing a full chapter break. Convention varies: a row of asterisks, a
trio of stars, a single fleuron glyph (``❦``, ``❧``,
``⁂``), or a plain horizontal rule. Each is a short, centred,
single-line paragraph between two body paragraphs.

This module exposes four composition helpers built entirely on
python-docx's public API (``Document.add_paragraph``, ``Run`` / ``Font``
attributes, ``ParagraphFormat.space_before`` / ``space_after``)::

    from docx.kit.dividers import (
        add_divider,
        add_fleuron,
        add_three_stars,
        add_chapter_break,
    )

    add_divider(doc, kind="line")    # plain horizontal line
    add_divider(doc, kind="dashed")  # dashed line
    add_divider(doc, kind="dots")    # row of dots
    add_divider(doc, kind="wave")    # row of wave glyphs

    add_fleuron(doc, glyph="❦")  # centred decorative glyph

    add_three_stars(doc)              # three centred ✦ glyphs

    add_chapter_break(doc, ornament="line", spacing=Pt(36))

Every helper:

* appends a *single* paragraph (or, for ``add_chapter_break``, three
  paragraphs — leading whitespace, ornament, trailing whitespace);
* centres the resulting paragraph(s);
* returns the appended paragraph (or, for ``add_chapter_break``, the
  list of three appended paragraphs in document order).

Implementation notes:

* The ``line`` divider style draws an underline run rather than emitting
  a ``w:pBdr`` border element. A bordered paragraph requires an empty
  paragraph that Word renders as a horizontal rule across the page
  width; the underline approach keeps the helper purely on the public
  ``Run.font`` surface, matches the visual width of a typical text line
  (which most authors actually want for a section divider — full-width
  rules look heavy), and avoids reaching down into the ``oxml``
  layer.
* The ``dashed`` / ``dots`` / ``wave`` divider styles emit a fixed-width
  row of repeating glyphs (``—``, ``·``, ``∼``); Word
  renders these inline so they centre with the surrounding paragraph
  alignment.
* Glyph defaults follow common print-typography conventions: the
  fleuron default ``❦`` (Floral Heart) is one of the three Unicode
  fleurons most often seen in modern editions; ``add_three_stars`` uses
  ``✦`` (Black Four-Pointed Star) separated by em-spaces.
* ``add_chapter_break`` is a thin convenience over ``add_divider`` /
  ``add_fleuron`` / ``add_three_stars``: it adds vertical whitespace
  before and after the ornament so the resulting break visually
  separates the chapters above and below. The ornament defaults to
  ``"line"`` and the spacing to 36 points; both are overridable.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

if TYPE_CHECKING:
    from docx.document import Document
    from docx.shared import Length
    from docx.text.paragraph import Paragraph


# -- Default glyphs.  Values picked to render in Word's default body
# -- font (Calibri 11pt) without falling back to a system glyph; every
# -- one of these characters is in the BMP and present in the standard
# -- Calibri / Cambria / Times New Roman glyph sets shipped with Word.
_DEFAULT_FLEURON = "❦"      # FLORAL HEART
_DEFAULT_STAR = "✦"          # BLACK FOUR-POINTED STAR
_EM_SPACE = " "              # EM SPACE — separates the three stars

# -- Repeating-glyph row widths.  Picked so the row visually fills a
# -- single body line at typical 11pt body size without wrapping; the
# -- helper centres the paragraph so any width that fits a body line
# -- looks right.  These are deliberately *not* tied to the active
# -- section width — the public API doesn't expose section width to
# -- a kit helper without reaching down — so the row is a fixed glyph
# -- count rather than a true full-width rule.
_DASHED_GLYPH = "—"          # EM DASH
_DASHED_COUNT = 9                  # nine em-dashes ~ 1/3 of a line
_DOTS_GLYPH = "·"            # MIDDLE DOT
_DOTS_COUNT = 7                    # seven dots, em-spaced
_WAVE_GLYPH = "∼"            # TILDE OPERATOR
_WAVE_COUNT = 9                    # nine waves

# -- ``line`` divider underline length.  A run holding this many
# -- non-breaking spaces, given an underline run-property, paints a
# -- short solid horizontal rule centred in the paragraph.  Word
# -- renders an underlined `` `` (NBSP) run as a visible line; an
# -- ordinary ``" "`` (regular space) is *also* underlined but Word
# -- collapses trailing spaces during layout, so we use NBSP.
_LINE_NBSP_COUNT = 24              # ~24 nbsp = ~2.5cm at 11pt Calibri

# -- Default chapter-break vertical whitespace either side of the
# -- ornament.  36pt (1/2 inch) is the conventional choice in book
# -- typography; halfway between a paragraph break and a section break.
_DEFAULT_CHAPTER_BREAK_SPACING = Pt(36)

# -- Valid ``kind`` values for ``add_divider``. Keep the tuple
# -- alphabetised so error messages are deterministic.
_DIVIDER_KINDS = ("dashed", "dots", "line", "wave")


def _new_centered_paragraph(document):
    # type: (Document) -> Paragraph
    """Append a fresh empty paragraph to ``document``, centred.

    The kit's divider / fleuron / three-stars / chapter-break helpers
    all share this shape: a single appended paragraph, centred, into
    which the caller adds the ornament run(s). Factor the boilerplate
    into one place so a future change (e.g. honouring the document's
    body alignment style) lands once.
    """
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return paragraph


def _add_line_divider(document):
    # type: (Document) -> Paragraph
    """Emit a short, centred underline-run "horizontal rule" paragraph."""
    paragraph = _new_centered_paragraph(document)
    # -- An underlined run of NBSPs renders as a solid underline; `` ``
    # -- is preserved by Word's layout engine where ``" "`` is collapsed.
    run = paragraph.add_run(" " * _LINE_NBSP_COUNT)
    run.underline = True
    return paragraph


def _add_repeating_glyph_divider(document, glyph, count, separator=""):
    # type: (Document, str, int, str) -> Paragraph
    """Emit a centred row of `count` `glyph` characters."""
    paragraph = _new_centered_paragraph(document)
    text = separator.join([glyph] * count)
    paragraph.add_run(text)
    return paragraph


def add_divider(document, kind="line"):
    # type: (Document, str) -> Paragraph
    """Append a centred section-divider paragraph to ``document``.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    kind
        One of ``"line"`` (short underline rule, the default),
        ``"dashed"`` (row of em-dashes), ``"dots"`` (row of middle
        dots), or ``"wave"`` (row of tildes).

    Returns
    -------
    Paragraph
        The newly-appended paragraph holding the divider.

    Raises
    ------
    ValueError
        If ``kind`` is not one of the four supported divider styles.
    """
    if kind not in _DIVIDER_KINDS:
        raise ValueError(
            "kind must be one of %s; got %r" % (_DIVIDER_KINDS, kind)
        )
    if kind == "line":
        return _add_line_divider(document)
    if kind == "dashed":
        return _add_repeating_glyph_divider(
            document, _DASHED_GLYPH, _DASHED_COUNT
        )
    if kind == "dots":
        # -- Em-space-separated dots read as a row, not a run-on word.
        return _add_repeating_glyph_divider(
            document, _DOTS_GLYPH, _DOTS_COUNT, separator=_EM_SPACE
        )
    # -- ``wave`` — only remaining option after the membership check.
    return _add_repeating_glyph_divider(document, _WAVE_GLYPH, _WAVE_COUNT)


def add_fleuron(document, glyph=_DEFAULT_FLEURON):
    # type: (Document, str) -> Paragraph
    """Append a centred decorative-glyph paragraph (a "fleuron") to ``document``.

    A *fleuron* is a single ornamental glyph used as a section
    separator. The default ``glyph`` is :unicode:`U+2766` (FLORAL
    HEART); callers may supply any other Unicode character (or short
    string) — common alternatives include :unicode:`U+2767`
    (ROTATED FLORAL HEART BULLET), :unicode:`U+2042` (ASTERISM), or
    :unicode:`U+273D` (HEAVY TEARDROP-SPOKED ASTERISK).

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    glyph
        The Unicode glyph (or short string) to render. Defaults to
        the FLORAL HEART fleuron.

    Returns
    -------
    Paragraph
        The newly-appended paragraph holding the fleuron.

    Raises
    ------
    ValueError
        If ``glyph`` is empty.
    """
    if not glyph:
        raise ValueError("glyph must be a non-empty string")
    paragraph = _new_centered_paragraph(document)
    paragraph.add_run(glyph)
    return paragraph


def add_three_stars(document, glyph=_DEFAULT_STAR):
    # type: (Document, str) -> Paragraph
    """Append a centred ``✦ ✦ ✦`` (three-stars) divider paragraph.

    The traditional "asterism" section break, rendered as three
    em-space-separated copies of ``glyph``. Defaults to
    :unicode:`U+2726` (BLACK FOUR-POINTED STAR); callers may pass an
    asterisk (``"*"``), a six-pointed star (``"✶"``), or any
    other glyph as a substitute.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    glyph
        The Unicode glyph to repeat three times. Defaults to
        BLACK FOUR-POINTED STAR.

    Returns
    -------
    Paragraph
        The newly-appended paragraph holding the three-stars row.

    Raises
    ------
    ValueError
        If ``glyph`` is empty.
    """
    if not glyph:
        raise ValueError("glyph must be a non-empty string")
    paragraph = _new_centered_paragraph(document)
    text = _EM_SPACE.join([glyph] * 3)
    paragraph.add_run(text)
    return paragraph


def add_chapter_break(
    document,
    ornament="line",
    spacing=None,
    glyph=None,
):
    # type: (Document, str, Optional[Length], Optional[str]) -> List[Paragraph]
    """Append a vertical-whitespace + ornament + vertical-whitespace break.

    A "chapter break" — sometimes called a *thought break* in
    typography — is the visual cousin of a full chapter opener: it
    separates two adjacent body sections by a gap of vertical
    whitespace bracketing a small ornament. The helper appends three
    paragraphs in order:

    1. an empty paragraph with ``space_after = spacing`` — the leading
       gap;
    2. the ornament paragraph itself, produced by dispatching to
       :func:`add_divider`, :func:`add_fleuron`, or
       :func:`add_three_stars` based on ``ornament``;
    3. an empty paragraph with ``space_before = spacing`` — the
       trailing gap.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    ornament
        The visual style of the ornament. One of ``"line"``,
        ``"dashed"``, ``"dots"``, ``"wave"`` (any
        :func:`add_divider` ``kind``), ``"fleuron"``, or
        ``"stars"``. Defaults to ``"line"``.
    spacing
        Vertical whitespace either side of the ornament. Defaults to
        :class:`Pt(36)` — 1/2 inch, the conventional book-typography
        choice.
    glyph
        Optional override for the ornament glyph. Forwarded to
        :func:`add_fleuron` when ``ornament="fleuron"`` or
        :func:`add_three_stars` when ``ornament="stars"``. Ignored
        for the divider variants.

    Returns
    -------
    list of Paragraph
        The three newly-appended paragraphs (leading-gap, ornament,
        trailing-gap), in document order.

    Raises
    ------
    ValueError
        If ``ornament`` is not one of the recognised values.
    """
    if spacing is None:
        spacing = _DEFAULT_CHAPTER_BREAK_SPACING

    # -- Leading gap.  An empty centred paragraph with space_after
    # -- equal to ``spacing`` produces the visual gap above the
    # -- ornament without leaving a stray run at the start of the
    # -- next chapter.
    leading = _new_centered_paragraph(document)
    leading.paragraph_format.space_after = spacing

    # -- The ornament itself.  Dispatch on ``ornament`` to one of the
    # -- three sibling helpers; surface a clean error for unrecognised
    # -- values so callers don't silently get a "line" fallback.
    if ornament == "fleuron":
        ornament_para = add_fleuron(
            document, glyph=glyph if glyph is not None else _DEFAULT_FLEURON
        )
    elif ornament == "stars":
        ornament_para = add_three_stars(
            document, glyph=glyph if glyph is not None else _DEFAULT_STAR
        )
    elif ornament in _DIVIDER_KINDS:
        ornament_para = add_divider(document, kind=ornament)
    else:
        valid = sorted(set(_DIVIDER_KINDS) | {"fleuron", "stars"})
        raise ValueError(
            "ornament must be one of %s; got %r" % (valid, ornament)
        )

    # -- Trailing gap.  Mirror of the leading gap so the ornament sits
    # -- centred between equal whitespace above and below.
    trailing = _new_centered_paragraph(document)
    trailing.paragraph_format.space_before = spacing

    return [leading, ornament_para, trailing]


__all__ = [
    "add_divider",
    "add_fleuron",
    "add_three_stars",
    "add_chapter_break",
]
