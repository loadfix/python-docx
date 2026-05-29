"""Branded header / cover-page helpers (cover page, first-page banner, running header).

Closes #288.

This module composes existing python-docx primitives (``Document.add_paragraph``,
``Document.add_page_break``, ``Run.add_picture``, ``Section.first_page_header``,
``Section.different_first_page_header_footer``, three-cell tables in headers /
footers) into three high-level helpers that produce the conventional
"branded document" shape used by reports, proposals, and white papers::

    from docx import Document
    from docx.kit import headers

    doc = Document()
    headers.cover_page(
        doc,
        title="Annual Report",
        subtitle="FY2026",
        logo="logo.png",
        date="2026-05-29",
        author="Jane Smith",
    )
    headers.first_page_banner(doc, title="Annual Report", logo="logo.png")
    headers.running_header(doc, left="Annual Report", right="Confidential")
    doc.save("out.docx")

Each helper writes to / appends to the *current* (first) section. They are
composable: callers typically invoke :func:`cover_page` first to lay down a
title page, then :func:`first_page_banner` to brand the first page header,
then :func:`running_header` to brand the per-page running header that the
rest of the document inherits.

- :func:`cover_page` — appends a styled cover page (centered title,
  optional logo, subtitle, date, author) to the document body. Returns the
  list of newly-appended |Paragraph| objects, in document order, including
  the trailing page break paragraph when ``page_break=True`` (the default).
- :func:`first_page_banner` — sets the *first-page* header of the current
  section to a banner (logo + title with a horizontal rule). Toggles
  :attr:`Section.different_first_page_header_footer` to |True| so the
  banner only appears on page one.
- :func:`running_header` — sets the *primary* (running) header — or footer
  when ``footer=True`` — of the current section, with an optional 3-cell
  layout (left / center / right). Each cell is independent so a caller
  can populate any subset.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, List, Optional, Union

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

if TYPE_CHECKING:
    from docx.document import Document
    from docx.section import _Footer, _Header
    from docx.text.paragraph import Paragraph


# -- Visual sizing constants. Pt() is a typed Length; keep integer
# -- literals here so the visual identity stays consistent across uses.
_TITLE_SIZE = Pt(36)
_SUBTITLE_SIZE = Pt(20)
_AUTHOR_SIZE = Pt(14)
_DATE_SIZE = Pt(12)
_BANNER_TITLE_SIZE = Pt(16)
_RUNNING_HEADER_SIZE = Pt(10)
_DEFAULT_COVER_LOGO_HEIGHT = Pt(96)   # ~1.3 inch — prominent on cover
_DEFAULT_BANNER_LOGO_HEIGHT = Pt(40)  # ~half inch — readable banner
_RULE = "—" * 30  # 30 em-dashes — banner / cover horizontal rule


def _coerce_color(line_color):
    # type: (Union[str, RGBColor, None]) -> Optional[RGBColor]
    """Return an :class:`RGBColor` for ``line_color`` or |None|.

    Accepts an :class:`RGBColor` instance (returned unchanged), a 6-character
    hex string (with or without leading ``#``), or |None| (returned as |None|
    so the caller leaves text colour at the style default).
    """
    if line_color is None:
        return None
    if isinstance(line_color, RGBColor):
        return line_color
    if isinstance(line_color, str):
        return RGBColor.from_string(line_color.lstrip("#"))
    raise ValueError(
        "line_color must be None, an RGBColor, or a hex string; got %r"
        % (line_color,)
    )


def _apply_color(run, rgb):
    # type: (object, Optional[RGBColor]) -> None
    """Set the run's font colour to `rgb` when supplied; no-op otherwise."""
    if rgb is not None:
        run.font.color.rgb = rgb


def _clear_existing_paragraphs(container):
    # type: (Union[_Header, _Footer]) -> None
    """Empty ``container`` so the helper writes from a clean slate.

    Re-using the canonical idiom from :mod:`docx.kit.letterhead` — Word
    requires a header/footer to contain at least one paragraph, so the
    first paragraph is wiped to empty and the surplus paragraphs are
    removed via the public ``_p`` proxy attribute.
    """
    paragraphs = list(container.paragraphs)
    if not paragraphs:
        return
    first = paragraphs[0]
    first.text = ""
    for para in paragraphs[1:]:
        p_elm = para._p
        parent = p_elm.getparent()
        if parent is not None:
            parent.remove(p_elm)


def cover_page(
    document,
    title,
    subtitle=None,
    logo=None,
    date=None,
    author=None,
    page_break=True,
):
    # type: (Document, str, Optional[str], Union[str, os.PathLike, None], Optional[str], Optional[str], bool) -> List[Paragraph]
    """Append a styled cover page to ``document`` and return the new paragraphs.

    Produces (in order) an optional centred logo paragraph, a centred
    large-type title, an optional centred subtitle, an optional centred
    horizontal rule, an optional centred author paragraph, an optional
    centred date paragraph, and (when ``page_break`` is true, the
    default) a trailing page break so the next content lands on a fresh
    page.

    `title` is required; every other field is optional and skipped when
    |None|. `date` is rendered verbatim — pass a pre-formatted string
    like ``"2026-05-29"``; the kit does not impose a date format.

    Parameters
    ----------
    document
        The :class:`Document` to brand.
    title
        Cover-page title, rendered in a large bold centred run. Required.
    subtitle
        Optional subtitle rendered below the title in a medium run.
    logo
        Optional logo image. Any value accepted by
        :meth:`Run.add_picture` works (string path, :class:`os.PathLike`,
        binary file-like object). Logo is centred above the title.
    date
        Optional date string rendered verbatim near the foot of the cover.
    author
        Optional author byline rendered above the date.
    page_break
        When |True| (the default), append a trailing page break so the
        next content starts on a fresh page.

    Returns
    -------
    list[Paragraph]
        Newly-appended paragraphs in document order, including the
        trailing page-break paragraph when emitted.

    Raises
    ------
    ValueError
        When ``title`` is empty.

    .. versionadded:: 2026.05.29
    """
    if not title:
        raise ValueError("title must be a non-empty string")

    paragraphs: List[Paragraph] = []

    if logo is not None:
        logo_para = document.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_run = logo_para.add_run()
        logo_run.add_picture(logo, height=_DEFAULT_COVER_LOGO_HEIGHT)
        paragraphs.append(logo_para)

    title_para = document.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(title)
    title_run.bold = True
    title_run.font.size = _TITLE_SIZE
    paragraphs.append(title_para)

    if subtitle:
        sub_para = document.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_run = sub_para.add_run(subtitle)
        sub_run.italic = True
        sub_run.font.size = _SUBTITLE_SIZE
        paragraphs.append(sub_para)

    # -- decorative rule sits below the title block when any byline /
    # -- date follows; it gives the cover its visual centre. --
    if author or date:
        rule_para = document.add_paragraph()
        rule_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rule_run = rule_para.add_run(_RULE)
        rule_run.font.size = _SUBTITLE_SIZE
        paragraphs.append(rule_para)

    if author:
        author_para = document.add_paragraph()
        author_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        author_run = author_para.add_run(author)
        author_run.font.size = _AUTHOR_SIZE
        paragraphs.append(author_para)

    if date:
        date_para = document.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_run = date_para.add_run(date)
        date_run.font.size = _DATE_SIZE
        paragraphs.append(date_para)

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def first_page_banner(
    document,
    title,
    logo=None,
    line_color="#000000",
):
    # type: (Document, str, Union[str, os.PathLike, None], Union[str, RGBColor, None]) -> List[Paragraph]
    """Set the first-page header of the current section to a branded banner.

    Toggles :attr:`Section.different_first_page_header_footer` to |True|
    so the banner only renders on the first page of the section, then
    writes (in order) a centred logo paragraph (when supplied), a
    centred title paragraph, and a centred horizontal rule paragraph
    coloured with ``line_color``.

    Any pre-existing first-page header content on the section is cleared
    first, so calling :func:`first_page_banner` is idempotent across
    re-runs.

    Parameters
    ----------
    document
        The :class:`Document` to brand.
    title
        Banner title, rendered in a centred medium-bold run. Required.
    logo
        Optional logo image. Any value accepted by
        :meth:`Run.add_picture` works.
    line_color
        Colour for the horizontal rule. Accepts an :class:`RGBColor`,
        a 6-character hex string (with or without leading ``#``), or
        |None| (rule keeps the style default colour). Defaults to
        ``"#000000"``.

    Returns
    -------
    list[Paragraph]
        Newly-written first-page-header paragraphs in document order.

    Raises
    ------
    ValueError
        When ``title`` is empty, or when ``line_color`` is neither
        |None|, an :class:`RGBColor`, nor a recognisable hex string.

    .. versionadded:: 2026.05.29
    """
    if not title:
        raise ValueError("title must be a non-empty string")
    if not document.sections:  # pragma: no cover - defensive
        raise ValueError("document has no sections; cannot apply banner")

    rgb = _coerce_color(line_color)
    section = document.sections[0]

    # -- enable the distinct first-page header so the banner only shows
    # -- on page one; the running header (set by ``running_header``) is
    # -- the per-page default. --
    section.different_first_page_header_footer = True

    header = section.first_page_header
    if header.is_linked_to_previous:
        header.is_linked_to_previous = False

    _clear_existing_paragraphs(header)

    paragraphs: List[Paragraph] = []

    first = header.paragraphs[0]
    first.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo is not None:
        logo_run = first.add_run()
        logo_run.add_picture(logo, height=_DEFAULT_BANNER_LOGO_HEIGHT)
    paragraphs.append(first)

    title_para = header.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(title)
    title_run.bold = True
    title_run.font.size = _BANNER_TITLE_SIZE
    paragraphs.append(title_para)

    rule_para = header.add_paragraph()
    rule_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rule_run = rule_para.add_run(_RULE)
    rule_run.font.size = _BANNER_TITLE_SIZE
    _apply_color(rule_run, rgb)
    paragraphs.append(rule_para)

    return paragraphs


def running_header(
    document,
    left=None,
    center=None,
    right=None,
    footer=False,
):
    # type: (Document, Optional[str], Optional[str], Optional[str], bool) -> List[Paragraph]
    """Set the running (primary) header — or footer — of the current section.

    The three cells are independent: pass any subset of ``left`` /
    ``center`` / ``right``. When at least two are supplied, the helper
    emits a borderless 3-column 1-row table so each cell sits at its
    edge of the page (left-aligned, centred, right-aligned). When only
    one cell is supplied, the helper falls back to a single paragraph
    aligned to the matching edge — simpler XML for a simpler intent.

    Any pre-existing primary-header (or footer) content on the section
    is cleared first, so calling :func:`running_header` is idempotent
    across re-runs.

    Parameters
    ----------
    document
        The :class:`Document` to brand.
    left
        Text for the left-aligned cell. Pass |None| to leave the cell
        empty.
    center
        Text for the centred cell. Pass |None| to leave the cell empty.
    right
        Text for the right-aligned cell. Pass |None| to leave the cell
        empty.
    footer
        When |True|, write to the section's primary footer rather than
        its primary header. Defaults to |False|.

    Returns
    -------
    list[Paragraph]
        Newly-written running-header (or footer) paragraphs, in document
        order. For the multi-cell layout, returns the three cell
        paragraphs.

    Raises
    ------
    ValueError
        When every cell is |None| (nothing to write).

    .. versionadded:: 2026.05.29
    """
    if left is None and center is None and right is None:
        raise ValueError(
            "running_header requires at least one of left, center, or right"
        )
    if not document.sections:  # pragma: no cover - defensive
        raise ValueError("document has no sections; cannot apply running header")

    section = document.sections[0]
    container = section.footer if footer else section.header
    if container.is_linked_to_previous:
        container.is_linked_to_previous = False

    _clear_existing_paragraphs(container)

    populated = [(name, value) for name, value in (
        ("left", left), ("center", center), ("right", right)
    ) if value is not None]

    paragraphs: List[Paragraph] = []

    if len(populated) == 1:
        # -- Single-cell shortcut: one paragraph, aligned to the cell's
        # -- edge of the page. Avoids dragging in a table just to hold
        # -- one piece of text. --
        name, value = populated[0]
        para = container.paragraphs[0]
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
        }
        para.alignment = align_map[name]
        run = para.add_run(value)
        run.font.size = _RUNNING_HEADER_SIZE
        paragraphs.append(para)
        return paragraphs

    # -- Multi-cell layout: 3-column borderless table. Each cell is
    # -- aligned to its own edge of the page. The table's first row
    # -- replaces the container's empty placeholder paragraph. --
    # -- 6 inch is the canonical body width on US-Letter / A4 with the
    # -- default 1.25" left/right margins; the cells then sit at the
    # -- usable-page edges. ``autofit = True`` lets Word re-flow the
    # -- columns when the page geometry differs. --
    table = container.add_table(rows=1, cols=3, width=Inches(6))
    table.autofit = True
    cells = table.rows[0].cells
    cell_alignments = (
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.RIGHT,
    )
    cell_values = (left, center, right)
    for cell, alignment, value in zip(cells, cell_alignments, cell_values):
        cell_para = cell.paragraphs[0]
        cell_para.alignment = alignment
        if value is not None:
            run = cell_para.add_run(value)
            run.font.size = _RUNNING_HEADER_SIZE
        paragraphs.append(cell_para)

    return paragraphs


__all__ = ["cover_page", "first_page_banner", "running_header"]
