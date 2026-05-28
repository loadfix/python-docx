"""Chapter opener pages — large title, optional decorative image, drop cap.

Closes #87.

This module composes existing python-docx primitives (sections, paragraphs,
runs, fonts, framePr drop caps, inline pictures) into a single high-level
helper, :func:`add_chapter_opener`, that emits the multi-element
"chapter-start" layout common to long-form documents (novels,
reports, theses)::

    from docx.kit import chapter

    chapter.add_chapter_opener(
        doc,
        chapter_number="Chapter 1",
        title="The First Light",
        epigraph='"In the beginning, there was..." -- Genesis 1:1',
        drop_cap=True,
        image="chapter1-opener.png",
        color="primary",
    )

    # Page break is automatic; the next add_paragraph after this
    # starts the chapter body.
    doc.add_paragraph("It was a dark and stormy night...")

The helper applies a section break before the opener so the new
chapter starts on a fresh page, then emits up to five logical
elements in canonical order:

1. ``chapter_number`` — small caps / bold heading line (e.g. "Chapter 1")
2. ``title`` — large, accent-coloured chapter title (Heading 1 style)
3. ``epigraph`` — italic, centered quotation block (optional)
4. ``image`` — decorative inline picture (optional)
5. drop-cap setup paragraph — when ``drop_cap=True``, the helper records
   that the *next* paragraph added to ``doc`` should be styled with a
   3-line ``w:framePr`` drop cap on its first letter.

Drop-cap framePr semantics follow ECMA-376 Part 1 ``17.3.1.11`` —
``w:dropCap="drop"``, ``w:lines="3"``, ``w:wrap="around"``,
``w:vAnchor="text"``, ``w:hAnchor="text"``. The helper splits the
target paragraph into two paragraphs: one paragraph containing only
the leading character of the body text with the framePr applied,
followed by a sibling paragraph containing the remainder. This is
the layout Word writes when the user enables Insert -> Drop Cap.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Optional, Union

from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_FRAME_DROP_CAP, WD_FRAME_WRAP
from docx.shared import Pt, RGBColor

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# -- Named accent colours.  Resolves the keyword strings exposed in
# -- the public API (``color="primary"``) to concrete RGB triples.
# -- Values picked to play nicely against Word's default body text on
# -- both light and dark themes; they intentionally do *not* depend on
# -- the underlying theme part so that the helper works on any document.
_NAMED_COLORS = {
    "primary": RGBColor(0x1F, 0x4E, 0x79),    # deep blue
    "secondary": RGBColor(0x70, 0x30, 0xA0),  # purple
    "accent": RGBColor(0xC0, 0x50, 0x4D),     # warm red
    "muted": RGBColor(0x59, 0x59, 0x59),      # neutral grey
    "black": RGBColor(0x00, 0x00, 0x00),
}

# -- Default sizing.  ``Pt`` is a typed Length — keep the integer
# -- literals here so the visual identity of an opener stays
# -- consistent across users.
_CHAPTER_NUMBER_SIZE = Pt(14)
_TITLE_SIZE = Pt(36)
_EPIGRAPH_SIZE = Pt(11)
_DROP_CAP_LINES = 3


def _resolve_color(color):  # type: (Union[str, RGBColor, None]) -> Optional[RGBColor]
    """Return an :class:`RGBColor` for ``color`` or |None|.

    ``color`` may be:

    - |None| — return |None| (caller leaves text colour at the style default);
    - one of the named keys in :data:`_NAMED_COLORS`;
    - an :class:`RGBColor` instance — returned unchanged;
    - a 6-character hex string (with or without leading ``#``).

    Any other value raises :class:`ValueError`.
    """
    if color is None:
        return None
    if isinstance(color, RGBColor):
        return color
    if isinstance(color, str):
        key = color.lower()
        if key in _NAMED_COLORS:
            return _NAMED_COLORS[key]
        # -- treat as hex string (strip leading '#' if present)
        return RGBColor.from_string(color.lstrip("#"))
    raise ValueError(
        "color must be None, a named preset (%s), an RGBColor, or a hex "
        "string; got %r" % (", ".join(sorted(_NAMED_COLORS)), color)
    )


def _add_chapter_number(doc, text, color):
    # type: (Document, str, Optional[RGBColor]) -> Paragraph
    """Append the small "Chapter N" line above the title."""
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = _CHAPTER_NUMBER_SIZE
    run.font.all_caps = True
    if color is not None:
        run.font.color.rgb = color
    return paragraph


def _add_title(doc, text, color):
    # type: (Document, str, Optional[RGBColor]) -> Paragraph
    """Append the large chapter title (Heading 1)."""
    # -- Heading 1 is the canonical style for chapter titles; keep the
    # -- style assignment so navigation panes and TOCs pick the title up.
    paragraph = doc.add_paragraph(style="Heading 1")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.font.size = _TITLE_SIZE
    run.bold = True
    if color is not None:
        run.font.color.rgb = color
    return paragraph


def _add_epigraph(doc, text):
    # type: (Document, str) -> Paragraph
    """Append the italic, centered epigraph quotation."""
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.italic = True
    run.font.size = _EPIGRAPH_SIZE
    return paragraph


def _add_decorative_image(doc, image):
    # type: (Document, Union[str, os.PathLike]) -> Paragraph
    """Append the decorative inline image, centered.

    The image is added as an inline picture in its own paragraph so
    callers can replace, resize, or delete it without disturbing the
    surrounding text. ``image`` may be any value :meth:`Document.add_picture`
    accepts (path string, :class:`os.PathLike`, or binary file-like).
    """
    # -- ``add_picture`` lays the image in its own newly-appended
    # -- paragraph; we then centre that paragraph.
    doc.add_picture(image)
    paragraph = doc.paragraphs[-1]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return paragraph


def _apply_drop_cap_marker(doc):
    # type: (Document) -> None
    """Mark the document so the next ``add_paragraph`` gets a drop cap.

    Word's drop-cap layout splits the body paragraph into two adjacent
    paragraphs: the first contains only the leading character with a
    ``w:framePr`` (``w:dropCap="drop"``) wrapping it; the second
    contains the remainder of the text. Splitting eagerly here would
    require us to know the body text up-front, which we don't — the
    caller supplies the body in a follow-up ``doc.add_paragraph(...)``
    call. Instead we register a one-shot hook that intercepts the next
    ``add_paragraph`` and rewrites it.

    The hook is monkey-patched onto the ``Document`` *instance*
    (not the class) and is removed after a single firing so unrelated
    later calls behave normally.

    Implementation notes:

    - Stashing on the instance keeps the helper stateless across
      documents and across invocations on different threads.
    - ``object.__setattr__`` is unnecessary; ``Document`` inherits from
      ``ElementProxy`` which doesn't use slots for the attributes we
      add here.
    - We only need to fire once. After the body paragraph has been
      rewritten, ``add_paragraph`` is restored to its original
      bound-method form.
    """
    original_add_paragraph = doc.add_paragraph

    def add_paragraph_with_drop_cap(text="", style=None, **kwargs):
        # -- Restore *before* doing work so a re-entrant call (eg.
        # -- inside a callback) doesn't re-trigger this branch.
        doc.add_paragraph = original_add_paragraph  # type: ignore[method-assign]
        if not text:
            # -- Nothing to drop-cap; emit the paragraph as-is.
            return original_add_paragraph(text=text, style=style, **kwargs)
        first_char, remainder = text[0], text[1:]
        # -- 1. The drop-cap paragraph itself: a 1-character paragraph
        # --    whose pPr/framePr has dropCap="drop", lines="3", wrap="around".
        drop_para = original_add_paragraph(text="", style=style, **kwargs)
        drop_para.paragraph_format.set_frame(
            drop_cap=WD_FRAME_DROP_CAP.DROP,
            lines=_DROP_CAP_LINES,
            wrap=WD_FRAME_WRAP.AROUND,
        )
        drop_run = drop_para.add_run(first_char)
        # -- Word renders the drop cap large by default but the lines
        # -- attribute alone doesn't size the glyph in WordprocessingML;
        # -- Word emits an explicit run-level font size as well.
        drop_run.font.size = Pt(48)
        drop_run.bold = True
        # -- 2. The body paragraph: same style, holds the remainder text.
        body_para = original_add_paragraph(text=remainder, style=style, **kwargs)
        return body_para

    doc.add_paragraph = add_paragraph_with_drop_cap  # type: ignore[method-assign]


def add_chapter_opener(
    doc,
    chapter_number=None,
    title=None,
    epigraph=None,
    drop_cap=False,
    image=None,
    color=None,
):
    # type: (Document, Optional[str], Optional[str], Optional[str], bool, Optional[Union[str, os.PathLike]], Union[str, RGBColor, None]) -> dict
    """Append a chapter-opener layout to ``doc``.

    Emits, in order:

    1. A ``WD_SECTION.NEW_PAGE`` section break, so the chapter starts
       on a fresh page.
    2. A centered, bold, all-caps "Chapter N" line — when
       ``chapter_number`` is given.
    3. The chapter ``title`` styled as ``Heading 1``, large and
       accent-coloured. The Heading 1 style means the title appears
       in the document's navigation pane and any auto-generated TOC.
    4. An italic, centered ``epigraph`` quotation — when given.
    5. A centered decorative ``image`` — when given. Accepts any value
       :meth:`Document.add_picture` accepts.
    6. When ``drop_cap=True``, registers a one-shot hook that
       intercepts the *next* :meth:`Document.add_paragraph` call and
       rewrites it as a 1-character framePr drop-cap paragraph
       followed by a body paragraph containing the remainder. This
       matches Word's "Insert -> Drop Cap (Dropped)" output.

    Returns a dict mapping each emitted element name to the
    :class:`Paragraph` that holds it (or |None| for omitted elements).
    The dict is intended for tests / fluent callers that need to apply
    further formatting; ordinary callers can ignore the return value.

    Parameters
    ----------
    doc
        The :class:`Document` to mutate. The opener is appended at the
        end of the document; callers wanting an opener at the start of
        a fresh document should construct ``Document()`` first and
        pass it in.
    chapter_number
        Short label like ``"Chapter 1"`` or ``"Part II"``. Rendered
        small, bold, all-caps, centered above the title.
    title
        The chapter title proper. Rendered large, bold, centered,
        styled ``Heading 1`` so navigation / TOC pick it up. Required
        in practice (an opener with no title is not a chapter opener)
        — passing |None| omits the title and is supported only for
        composability.
    epigraph
        Optional italic quotation rendered between the title and any
        decorative image.
    drop_cap
        When |True|, registers a one-shot hook on ``doc`` that
        rewrites the *next* :meth:`Document.add_paragraph` call as a
        framePr drop-cap paragraph plus body paragraph.
    image
        Optional decorative image. Any value accepted by
        :meth:`Document.add_picture` works.
    color
        Accent colour for the chapter number and title. Accepts a
        named preset (``"primary"``, ``"secondary"``, ``"accent"``,
        ``"muted"``, ``"black"``), an :class:`RGBColor`, or a
        6-character hex string. |None| leaves text at the style
        default.

    Returns
    -------
    dict
        Keys: ``"section"``, ``"chapter_number"``, ``"title"``,
        ``"epigraph"``, ``"image"``. Each value is the corresponding
        :class:`Paragraph` (or :class:`Section` for ``"section"``), or
        |None| when the element was omitted.

    .. versionadded:: 2026.05.29
    """
    rgb = _resolve_color(color)

    # -- 1. Section break.  ``add_section`` appends a new section break
    # --    and returns the new ``Section`` object — same shape as
    # --    every other section helper in python-docx.
    section = doc.add_section(WD_SECTION.NEW_PAGE)

    result = {
        "section": section,
        "chapter_number": None,
        "title": None,
        "epigraph": None,
        "image": None,
    }

    # -- 2. Chapter number line (small, bold, all-caps).
    if chapter_number:
        result["chapter_number"] = _add_chapter_number(doc, chapter_number, rgb)

    # -- 3. Chapter title (Heading 1).
    if title:
        result["title"] = _add_title(doc, title, rgb)

    # -- 4. Epigraph (italic, centered).
    if epigraph:
        result["epigraph"] = _add_epigraph(doc, epigraph)

    # -- 5. Decorative image (centered).
    if image is not None:
        result["image"] = _add_decorative_image(doc, image)

    # -- 6. Drop-cap one-shot hook.  Apply *last* so any paragraphs
    # --    we emit above don't accidentally trigger it.
    if drop_cap:
        _apply_drop_cap_marker(doc)

    return result


__all__ = ["add_chapter_opener"]
