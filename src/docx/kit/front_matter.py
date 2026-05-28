"""Front-matter helpers — title / copyright / dedication / preface / TOC.

A long-form document (book, report, proposal, white paper) typically
opens with a sequence of *front-matter* sections that precede the body
copy. The exact order varies by house style, but the canonical
sequence is::

    Title page         (title + subtitle + author + date)
    Copyright page     (holder + year + edition + rights)
    Dedication         (one short centred paragraph)
    Preface / Foreword (heading + a few paragraphs)
    Table of contents
    List of figures    (TOC of Figure-labelled captions)
    List of tables     (TOC of Table-labelled captions)

This module exposes one helper per section. Each helper is a thin
composition of existing python-docx public API — :meth:`Document.add_paragraph`,
:meth:`Document.add_heading`, :meth:`Document.add_page_break`,
:meth:`Document.add_table_of_contents`,
:meth:`Paragraph.add_complex_field` — and returns the list of
paragraphs it appended, in document order, so callers can
post-process them (attach bookmarks, tweak alignment, set run-level
formatting) without having to rediscover them via
``document.paragraphs[-N:]``.

All seven helpers append at the end of the body. Callers who need to
insert front matter ahead of existing content should build a fresh
|Document| with the helpers, then merge their existing body in via
:meth:`Document.add_table_copy` and friends; the kit deliberately
does not own a "move-to-front" primitive.

The helpers prefer the conventional Word built-in styles
(``Title``, ``Subtitle``, ``Heading 1``, ``Quote``, ``BookTitle``)
that ship with python-docx's default template. When a caller has
loaded a template that lacks one of those styles, the helpers fall
back to ``Normal`` rather than raising — the spirit of a *kit* is
"works out of the box, customise as you like".

.. versionadded:: 2026.05.0
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional, Sequence, Union

from docx.enum.text import WD_ALIGN_PARAGRAPH

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# -- Word built-in styles the kit reaches for. These are the styleIds
# -- shipped in python-docx's default styles template; consult
# -- ``src/docx/templates/default-docx-template/word/styles.xml``. When a
# -- caller supplies a custom template that is missing a particular
# -- style, the kit silently falls back to "Normal" rather than raising —
# -- the helpers are best-effort cosmetic, not strict validators. --
_STYLE_TITLE = "Title"
_STYLE_SUBTITLE = "Subtitle"
_STYLE_HEADING_1 = "Heading 1"
_STYLE_QUOTE = "Quote"
_STYLE_NORMAL = "Normal"


def _has_style(document: Document, style_name: str) -> bool:
    """Return |True| when `document` defines a paragraph style named `style_name`.

    The kit prefers Word's built-in styles (``Title``, ``Subtitle``, etc.)
    but cannot assume every document loaded by the caller has them — a
    custom corporate template may have been stripped down. This helper
    lets each kit function fall back to ``Normal`` rather than raise.
    """
    try:
        styles = document.styles
    except Exception:  # pragma: no cover - defensive
        return False
    try:
        styles[style_name]
        return True
    except KeyError:
        return False


def _resolve_style(document: Document, preferred: str) -> str:
    """Return `preferred` if it exists on `document`, else ``"Normal"``."""
    return preferred if _has_style(document, preferred) else _STYLE_NORMAL


def _coerce_body(body: Union[str, Sequence[str]]) -> List[str]:
    """Return `body` as a list of paragraph strings.

    A string `body` is split on blank lines (``"\\n\\n"``) so callers
    can pass a single multi-paragraph string and get sensible
    paragraph breaks. A pre-split sequence is returned as a list
    unchanged. Empty strings are dropped to avoid producing empty
    paragraphs.
    """
    if isinstance(body, str):
        chunks = [chunk.strip() for chunk in body.split("\n\n")]
        return [chunk for chunk in chunks if chunk]
    return [chunk for chunk in body if chunk]


def add_title_page(
    document: Document,
    title: str,
    subtitle: Optional[str] = None,
    author: Optional[str] = None,
    date: Optional[str] = None,
    page_break: bool = True,
) -> List[Paragraph]:
    """Append a title page to `document` and return the new paragraphs.

    Produces (in order) a centred ``Title`` paragraph, an optional
    centred ``Subtitle`` paragraph, an optional centred author
    paragraph, an optional centred date paragraph, and (when
    ``page_break`` is true, the default) a trailing page break so the
    next content lands on a fresh page.

    `title` is required; `subtitle`, `author`, and `date` are each
    optional and skipped when |None|. `date` is rendered verbatim —
    pass a pre-formatted string like ``"March 2026"``; the kit does
    not impose a date format.

    Returns the list of newly-appended |Paragraph| objects, in document
    order, including the trailing page-break paragraph when emitted.

    .. versionadded:: 2026.05.0
    """
    if not title:
        raise ValueError("title must be a non-empty string")

    paragraphs: List[Paragraph] = []
    title_style = _resolve_style(document, _STYLE_TITLE)
    subtitle_style = _resolve_style(document, _STYLE_SUBTITLE)

    title_para = document.add_paragraph(title, style=title_style)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraphs.append(title_para)

    if subtitle:
        sub_para = document.add_paragraph(subtitle, style=subtitle_style)
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraphs.append(sub_para)

    if author:
        author_para = document.add_paragraph(author)
        author_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraphs.append(author_para)

    if date:
        date_para = document.add_paragraph(date)
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraphs.append(date_para)

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def add_copyright_page(
    document: Document,
    holder: str,
    year: Union[int, str],
    edition: Optional[str] = None,
    rights: Optional[str] = None,
    page_break: bool = True,
) -> List[Paragraph]:
    """Append a copyright page to `document` and return the new paragraphs.

    The page is rendered as four (or fewer) centred paragraphs:

    1. ``"Copyright © {year} {holder}"`` — copyright notice with the
       Unicode copyright sign.
    2. The `edition` string (e.g. ``"First Edition"``), when supplied.
    3. The `rights` notice (e.g. ``"All rights reserved."``); defaults
       to ``"All rights reserved."`` when |None|, suppressed when the
       caller passes the empty string explicitly.
    4. A trailing page break, when ``page_break`` is true (the default).

    `year` may be an ``int`` (rendered as decimal) or a string
    (rendered verbatim — useful for ranges such as ``"2024–2026"``).

    Returns the list of newly-appended |Paragraph| objects in document
    order.

    .. versionadded:: 2026.05.0
    """
    if not holder:
        raise ValueError("holder must be a non-empty string")

    paragraphs: List[Paragraph] = []
    notice = f"Copyright © {year} {holder}"
    notice_para = document.add_paragraph(notice)
    notice_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraphs.append(notice_para)

    if edition:
        edition_para = document.add_paragraph(edition)
        edition_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraphs.append(edition_para)

    # -- rights: default to the conventional notice, but allow the
    # -- caller to suppress entirely with an explicit empty string. --
    if rights is None:
        rights = "All rights reserved."
    if rights:
        rights_para = document.add_paragraph(rights)
        rights_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraphs.append(rights_para)

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def add_dedication(
    document: Document,
    text: str,
    page_break: bool = True,
) -> List[Paragraph]:
    """Append a dedication paragraph to `document` and return the new paragraphs.

    A dedication is rendered as a single centred italic paragraph
    using Word's built-in ``Quote`` style (which is itself italic when
    the default Word style template is in use). Long dedications are
    fine — Word will wrap as normal — but the kit deliberately does
    not split on blank lines: the convention is "one short
    sentence".

    Returns the list of newly-appended |Paragraph| objects, including
    the trailing page-break paragraph when ``page_break`` is true (the
    default).

    .. versionadded:: 2026.05.0
    """
    if not text:
        raise ValueError("text must be a non-empty string")

    paragraphs: List[Paragraph] = []
    style = _resolve_style(document, _STYLE_QUOTE)
    para = document.add_paragraph(text, style=style)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # -- belt-and-braces italic: Quote is italic in Word's default
    # -- style template, but custom templates may have stripped it.
    # -- Set italic on every run so the dedication renders as expected
    # -- even when the style is "Normal" via the fallback path. --
    for run in para.runs:
        run.italic = True
    paragraphs.append(para)

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def add_preface(
    document: Document,
    title: str = "Preface",
    body: Union[str, Sequence[str]] = "",
    heading_level: int = 1,
    page_break: bool = True,
) -> List[Paragraph]:
    """Append a preface to `document` and return the new paragraphs.

    Renders a heading paragraph (``Heading {heading_level}``, default
    ``Heading 1``) followed by one paragraph per item in `body`. When
    `body` is a string, it is split on blank lines (``"\\n\\n"``) so
    multi-paragraph prose can be passed as a single triple-quoted
    string. When `body` is a sequence, each item becomes one paragraph.

    `title` defaults to ``"Preface"`` but accepts any string —
    ``"Foreword"`` and ``"Introduction"`` are common alternatives that
    use the same shape.

    Returns the list of newly-appended |Paragraph| objects, including
    the trailing page-break paragraph when ``page_break`` is true (the
    default).

    .. versionadded:: 2026.05.0
    """
    if not title:
        raise ValueError("title must be a non-empty string")
    if not 0 <= heading_level <= 9:
        raise ValueError(
            "heading_level must be in 0..9 (matching Document.add_heading), "
            "got %d" % heading_level
        )

    paragraphs: List[Paragraph] = [document.add_heading(title, level=heading_level)]
    for chunk in _coerce_body(body):
        paragraphs.append(document.add_paragraph(chunk))

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def add_table_of_contents(
    document: Document,
    title: Optional[str] = "Table of Contents",
    levels: tuple = (1, 3),
    heading_level: int = 1,
    page_break: bool = True,
) -> List[Paragraph]:
    """Append a table-of-contents section to `document`.

    Wraps :meth:`docx.document.Document.add_table_of_contents` with an
    optional preceding heading. When `title` is non-empty, a heading
    paragraph (``Heading {heading_level}``) is appended first and the
    TOC paragraph follows. Pass ``title=None`` (or the empty string)
    to skip the heading and append only the TOC field paragraph.

    `levels` is forwarded to
    :meth:`Document.add_table_of_contents`; valid ``(min, max)`` tuples
    satisfy ``1 <= min <= max <= 9``.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true (the default).

    .. versionadded:: 2026.05.0
    """
    paragraphs: List[Paragraph] = []
    if title:
        paragraphs.append(document.add_heading(title, level=heading_level))
    paragraphs.append(document.add_table_of_contents(levels=levels))

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def _add_caption_toc(
    document: Document,
    title: str,
    label: str,
    heading_level: int,
    page_break: bool,
) -> List[Paragraph]:
    """Shared implementation for ``add_list_of_figures`` / ``add_list_of_tables``.

    Both helpers emit a ``TOC`` field that filters to a single SEQ
    label — Word's conventional list-of-figures / list-of-tables
    output. The instruction shape is::

        TOC \\h \\z \\c "{label}"

    where ``\\c`` selects entries by SEQ identifier (``"Figure"``,
    ``"Table"``, etc.). The cached result text is left empty —
    python-docx has no layout engine and Word rebuilds the field on
    open or field-update anyway. See ``docx/captions.py`` for how
    those SEQ entries are emitted by ``Document.add_caption``.
    """
    paragraphs: List[Paragraph] = []
    if title:
        paragraphs.append(document.add_heading(title, level=heading_level))

    # -- TOC field with \c switch filters by SEQ identifier. The cached
    # -- result text is None because we have no fixtures from which to
    # -- build a preview without scanning every caption in the document
    # -- — Word recomputes on first field-update anyway. --
    toc_para = document.add_paragraph()
    instr = f' TOC \\h \\z \\c "{label}" '
    toc_para.add_complex_field(instr, result_text=None)
    paragraphs.append(toc_para)

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def add_list_of_figures(
    document: Document,
    title: Optional[str] = "List of Figures",
    label: str = "Figure",
    heading_level: int = 1,
    page_break: bool = True,
) -> List[Paragraph]:
    """Append a "List of Figures" TOC section to `document`.

    Emits an optional heading (``Heading {heading_level}``) followed
    by a ``TOC`` field whose ``\\c`` switch filters to entries
    captioned with the given SEQ `label` (default ``"Figure"`` —
    matches the default of :meth:`Document.add_caption`). Word
    rebuilds the listing on open or field-update; the cached result is
    intentionally empty.

    Pass ``title=None`` (or the empty string) to suppress the heading
    and append only the TOC paragraph.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true (the default).

    .. versionadded:: 2026.05.0
    """
    return _add_caption_toc(document, title or "", label, heading_level, page_break)


def add_list_of_tables(
    document: Document,
    title: Optional[str] = "List of Tables",
    label: str = "Table",
    heading_level: int = 1,
    page_break: bool = True,
) -> List[Paragraph]:
    """Append a "List of Tables" TOC section to `document`.

    Emits an optional heading (``Heading {heading_level}``) followed
    by a ``TOC`` field whose ``\\c`` switch filters to entries
    captioned with the given SEQ `label` (default ``"Table"`` —
    matches what callers pass to :meth:`Document.add_caption` for
    table captions). Word rebuilds the listing on open or field-update;
    the cached result is intentionally empty.

    Pass ``title=None`` (or the empty string) to suppress the heading
    and append only the TOC paragraph.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true (the default).

    .. versionadded:: 2026.05.0
    """
    return _add_caption_toc(document, title or "", label, heading_level, page_break)


__all__ = [
    "add_title_page",
    "add_copyright_page",
    "add_dedication",
    "add_preface",
    "add_table_of_contents",
    "add_list_of_figures",
    "add_list_of_tables",
]
