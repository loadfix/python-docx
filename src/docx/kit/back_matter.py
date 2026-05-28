"""Back-matter helpers — appendix / glossary / index / bibliography.

Closes #86.

A long-form document closes with a sequence of *back-matter* sections
that follow the body copy. The exact order varies by house style, but
the canonical sequence is::

    Appendix A, B, C ...     (additional reference material)
    Glossary                 (term -> definition pairs)
    Index                    (auto-built A-Z entry list, INDEX field)
    Bibliography             (APA-formatted source list)

This module exposes one helper per section. Each helper is a thin
composition of existing python-docx public API
(:meth:`Document.add_paragraph`, :meth:`Document.add_heading`,
:meth:`Document.add_page_break`, :meth:`Document.add_table`,
:meth:`Paragraph.add_complex_field`) and returns the list of
paragraphs (or, in the appendix case, paragraphs *and* the kept
section break) it appended, in document order, so callers can
post-process them without rediscovering them via
``document.paragraphs[-N:]``.

All four helpers append at the end of the body. Callers wanting to
splice back matter ahead of an existing tail should compose a fresh
|Document| and append their body via :meth:`Document.add_table_copy`
and friends; the kit deliberately does not own a "move-to-end"
primitive.

The helpers prefer Word's conventional built-in styles
(``Heading 1``, ``Normal``) and fall back to ``Normal`` when a custom
template lacks a style — the spirit of a *kit* is "works out of the
box, customise as you like".

Design choices
--------------

**Glossary as 2-column table.** ECMA-376 idiomatic glossaries use one
of two shapes: a 2-column table (term | definition) or a definition
list (indented paragraphs). The table form is picked because it is
the cleaner shape across both Word and LibreOffice — the columns
auto-align, long definitions wrap inside their cell rather than
visually flowing under the term, and the table gets picked up by
screen-readers as a structural element. The alternative (an indented
paragraph list) is *also* supported by :func:`add_glossary` via
``layout="list"`` for callers who prefer the running-text shape.

**Bibliography as APA-ish.** Sources are formatted in a lightly
APA-flavoured style — the kit is opinionated about output rather
than a strict APA validator, since house styles differ on minutiae
(italics on titles, "Inc." / "Ltd." dropping, DOI vs URL, ampersand
in author lists). The format renders consistently for the four
canonical kinds (book, article, web, report) and prints raw key=value
pairs for unknown kinds so callers see exactly what was passed.

**Index as INDEX field.** The body of an index is an ``INDEX`` Word
complex field that Word rebuilds from the document's ``XE`` (index
entry) markers on first open or field-update. python-docx has no
layout engine, so the cached result is intentionally empty.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, List, Mapping, Optional, Sequence, Union

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph


# -- Word built-in styles the kit reaches for. The default python-docx
# -- template ships ``Normal`` and the ``Heading 1..9`` styles. When a
# -- caller-supplied template is missing a particular style, we fall
# -- back to ``Normal`` rather than raise — same policy as
# -- ``front_matter``.
_STYLE_HEADING_1 = "Heading 1"
_STYLE_NORMAL = "Normal"


# -- ``add_glossary`` accepts these layout keywords. "table" emits a
# -- 2-column borderless table with one row per term; "list" emits one
# -- bold-term paragraph followed by its definition paragraph.
_GLOSSARY_LAYOUTS = ("table", "list")


def _has_style(document: "Document", style_name: str) -> bool:
    """Return |True| when `document` defines a paragraph style named `style_name`.

    Mirrors ``front_matter._has_style`` so back matter helpers fall
    back to ``Normal`` rather than raise when a stripped-down corporate
    template lacks an expected style.
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


def _resolve_style(document: "Document", preferred: str) -> str:
    """Return `preferred` if it exists on `document`, else ``"Normal"``."""
    return preferred if _has_style(document, preferred) else _STYLE_NORMAL


def _coerce_body(body: Union[str, Sequence[str]]) -> List[str]:
    """Return `body` as a list of paragraph strings.

    A string `body` is split on blank lines (``"\\n\\n"``) so callers can
    pass a single multi-paragraph block of prose and get sensible
    paragraph breaks. A pre-split sequence is returned as a list
    unchanged. Empty chunks are dropped.
    """
    if isinstance(body, str):
        chunks = [chunk.strip() for chunk in body.split("\n\n")]
        return [chunk for chunk in chunks if chunk]
    return [chunk for chunk in body if chunk]


def add_appendix(
    document: "Document",
    label: str,
    title: str,
    body: Union[str, Sequence[str]] = "",
    heading_level: int = 1,
    page_break: bool = True,
) -> List["Paragraph"]:
    """Append an appendix section to `document` and return the new paragraphs.

    Renders a single heading paragraph of the form
    ``"{label}: {title}"`` (e.g. ``"Appendix A: Data Tables"``) at
    ``Heading {heading_level}`` (default ``Heading 1``), followed by
    one paragraph per `body` chunk.

    `label` is the conventional appendix prefix (``"Appendix A"``,
    ``"Annex 1"``); `title` is the appendix's own descriptive name.
    Both are required — an appendix without a label or title is rare
    enough to be a mistake. To suppress the colon and render a single
    line ``"{label}"``, pass ``title=""`` — that branch is supported
    so callers can author Annex pages with just a number.

    `body` may be a single string (split on blank lines) or a sequence
    of strings (one paragraph each). Empty bodies are valid — useful
    for callers who want the heading and intend to follow up with
    tables / images via the regular Document API.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true (the default).

    .. versionadded:: 2026.05.29
    """
    if not label:
        raise ValueError("label must be a non-empty string")
    if not 0 <= heading_level <= 9:
        raise ValueError(
            "heading_level must be in 0..9 (matching Document.add_heading), "
            "got %d" % heading_level
        )

    heading_text = f"{label}: {title}" if title else label
    paragraphs: List["Paragraph"] = [
        document.add_heading(heading_text, level=heading_level)
    ]
    for chunk in _coerce_body(body):
        paragraphs.append(document.add_paragraph(chunk))

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def add_glossary(
    document: "Document",
    entries: Mapping[str, str],
    title: Optional[str] = "Glossary",
    heading_level: int = 1,
    layout: str = "table",
    page_break: bool = True,
) -> List["Paragraph"]:
    """Append a glossary section to `document` and return the new paragraphs.

    Renders an optional heading (``Heading {heading_level}``) followed
    by the term/definition pairs in `entries`. Two layouts are
    supported:

    * ``layout="table"`` (default) — a 2-column table, one row per
      entry, term in the left column and definition in the right.
      The table inherits the document's default table style; the
      caller can restyle via ``glossary_table.style = "Light Grid"``
      after the call (the table is *not* in the returned paragraph
      list — it is a |Table|, not a |Paragraph|; reach it via
      ``document.tables[-1]``).
    * ``layout="list"`` — one bold-term paragraph immediately
      followed by its definition paragraph (indented one tab).

    Entries are emitted in iteration order. Pass an ``OrderedDict``
    (or a plain ``dict`` on Python 3.7+) to control the rendering
    order; pass a sorted-by-key dict if alphabetical order is wanted.

    Pass ``title=None`` (or the empty string) to suppress the heading
    and append only the entries.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true (the default). When ``layout="table"``, the
    entry rows live in a |Table| rather than paragraphs; the heading
    and trailing page-break paragraph are still returned, but the
    table itself is reached via ``document.tables[-1]``.

    .. versionadded:: 2026.05.29
    """
    if layout not in _GLOSSARY_LAYOUTS:
        raise ValueError(
            "layout must be one of %r, got %r" % (_GLOSSARY_LAYOUTS, layout)
        )
    if not 0 <= heading_level <= 9:
        raise ValueError(
            "heading_level must be in 0..9 (matching Document.add_heading), "
            "got %d" % heading_level
        )

    paragraphs: List["Paragraph"] = []
    if title:
        paragraphs.append(document.add_heading(title, level=heading_level))

    if entries:
        if layout == "table":
            _emit_glossary_table(document, entries)
        else:
            paragraphs.extend(_emit_glossary_list(document, entries))

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


def _emit_glossary_table(
    document: "Document", entries: Mapping[str, str]
) -> "Table":
    """Append a 2-column table holding `entries` and return it.

    One row per (term, definition) pair, term in column 0 and
    definition in column 1. The term is emitted bold to give it
    visual weight without a custom paragraph style — this works
    even when the document has no ``Strong`` character style.
    """
    table = document.add_table(rows=len(entries), cols=2)
    for row_idx, (term, definition) in enumerate(entries.items()):
        term_cell = table.cell(row_idx, 0)
        term_para = term_cell.paragraphs[0]
        term_run = term_para.add_run(str(term))
        term_run.bold = True

        def_cell = table.cell(row_idx, 1)
        def_cell.paragraphs[0].add_run(str(definition))
    return table


def _emit_glossary_list(
    document: "Document", entries: Mapping[str, str]
) -> List["Paragraph"]:
    """Append paragraphs in the "definition list" layout.

    One bold-term paragraph followed by an indented definition
    paragraph per entry. Returns the appended paragraphs in
    document order.
    """
    paragraphs: List["Paragraph"] = []
    for term, definition in entries.items():
        term_para = document.add_paragraph()
        term_run = term_para.add_run(str(term))
        term_run.bold = True
        paragraphs.append(term_para)

        def_para = document.add_paragraph(str(definition))
        # -- one-tab indent on the definition paragraph; using the
        # -- typed ParagraphFormat surface keeps the helper out of
        # -- the oxml layer per kit conventions.
        from docx.shared import Inches

        def_para.paragraph_format.left_indent = Inches(0.25)
        paragraphs.append(def_para)
    return paragraphs


def add_index(
    document: "Document",
    title: Optional[str] = "Index",
    columns: int = 2,
    heading_level: int = 1,
    page_break: bool = True,
) -> List["Paragraph"]:
    """Append an index section to `document` and return the new paragraphs.

    Renders an optional heading (``Heading {heading_level}``) followed
    by an ``INDEX`` complex field. Word builds the index from any
    ``XE`` (index-entry) markers in the document on first open or
    field-update; python-docx has no layout engine, so the cached
    result is intentionally empty.

    `columns` controls the index's column count via the ``\\c`` switch
    (``\\c "2"`` is the canonical 2-column index Word produces by
    default). Pass ``columns=1`` for a single-column index.

    Pass ``title=None`` (or the empty string) to suppress the heading
    and append only the INDEX field paragraph.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true (the default).

    .. versionadded:: 2026.05.29
    """
    if columns < 1 or columns > 4:
        raise ValueError(
            "columns must be in 1..4 (Word's supported range), got %d" % columns
        )
    if not 0 <= heading_level <= 9:
        raise ValueError(
            "heading_level must be in 0..9 (matching Document.add_heading), "
            "got %d" % heading_level
        )

    paragraphs: List["Paragraph"] = []
    if title:
        paragraphs.append(document.add_heading(title, level=heading_level))

    # -- INDEX field with \h "A" emits a heading letter between blocks
    # -- (matches Word's "Insert -> Index" default), \c "{n}" sets the
    # -- column count. Word recomputes the body on first open. --
    instr = f' INDEX \\h "A" \\c "{columns}" '
    index_para = document.add_paragraph()
    index_para.add_complex_field(instr, result_text=None)
    paragraphs.append(index_para)

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


# -- Bibliography source rendering ----------------------------------------
# A bibliography source is a dict with a ``kind`` key plus per-kind fields.
# The kit accepts four canonical kinds and emits APA-flavoured one-line
# entries; unknown kinds are rendered as raw "key=value" pairs so the
# author sees what they passed without a silent drop.


def _format_authors(authors: Any) -> str:
    """Return a single-line author string suitable for an APA-ish bibliography.

    Accepts a string (returned trimmed) or a sequence of strings (joined
    with ", " and "& " before the last). |None| yields the empty
    string. Each item is ``str()``-coerced so callers can pass typed
    Author objects with a sensible ``__str__`` without a custom
    serialiser.
    """
    if authors is None:
        return ""
    if isinstance(authors, str):
        return authors.strip()
    items = [str(a).strip() for a in authors if str(a).strip()]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + ", & " + items[-1]


def _format_book(source: Mapping[str, Any]) -> str:
    """APA-ish: ``Authors (Year). *Title*. Publisher.``"""
    authors = _format_authors(source.get("authors") or source.get("author"))
    year = source.get("year", "n.d.")
    title = source.get("title", "")
    publisher = source.get("publisher", "")

    parts: List[str] = []
    if authors:
        parts.append(f"{authors} ({year}).")
    elif year:
        parts.append(f"({year}).")
    if title:
        parts.append(f"{title}.")
    if publisher:
        parts.append(f"{publisher}.")
    return " ".join(parts)


def _format_article(source: Mapping[str, Any]) -> str:
    """APA-ish: ``Authors (Year). Title. *Journal*, Vol(Issue), Pages.``"""
    authors = _format_authors(source.get("authors") or source.get("author"))
    year = source.get("year", "n.d.")
    title = source.get("title", "")
    journal = source.get("journal", "")
    volume = source.get("volume")
    issue = source.get("issue")
    pages = source.get("pages")

    parts: List[str] = []
    if authors:
        parts.append(f"{authors} ({year}).")
    elif year:
        parts.append(f"({year}).")
    if title:
        parts.append(f"{title}.")
    if journal:
        loc = journal
        if volume is not None:
            loc += f", {volume}"
            if issue is not None:
                loc += f"({issue})"
        if pages is not None:
            loc += f", {pages}"
        parts.append(f"{loc}.")
    return " ".join(parts)


def _format_web(source: Mapping[str, Any]) -> str:
    """APA-ish: ``Authors (Year). Title. Site. Retrieved from URL``"""
    authors = _format_authors(source.get("authors") or source.get("author"))
    year = source.get("year", "n.d.")
    title = source.get("title", "")
    site = source.get("site") or source.get("publisher") or ""
    url = source.get("url", "")

    parts: List[str] = []
    if authors:
        parts.append(f"{authors} ({year}).")
    elif year:
        parts.append(f"({year}).")
    if title:
        parts.append(f"{title}.")
    if site:
        parts.append(f"{site}.")
    if url:
        parts.append(f"Retrieved from {url}")
    return " ".join(parts)


def _format_report(source: Mapping[str, Any]) -> str:
    """APA-ish: ``Authors (Year). *Title* (Report No. N). Publisher.``"""
    authors = _format_authors(source.get("authors") or source.get("author"))
    year = source.get("year", "n.d.")
    title = source.get("title", "")
    number = source.get("number") or source.get("report_number")
    publisher = source.get("publisher") or source.get("institution") or ""

    parts: List[str] = []
    if authors:
        parts.append(f"{authors} ({year}).")
    elif year:
        parts.append(f"({year}).")
    if title:
        if number:
            parts.append(f"{title} (Report No. {number}).")
        else:
            parts.append(f"{title}.")
    if publisher:
        parts.append(f"{publisher}.")
    return " ".join(parts)


def _format_unknown(source: Mapping[str, Any]) -> str:
    """Fallback: emit raw ``key=value`` pairs.

    Better to surface the raw source than to silently drop a kind the
    kit doesn't yet know — the caller spots a typo or an unsupported
    kind on first preview.
    """
    pairs = [
        f"{k}={v}"
        for k, v in source.items()
        if k != "kind" and v is not None and v != ""
    ]
    return "; ".join(pairs)


_FORMATTERS = {
    "book": _format_book,
    "article": _format_article,
    "journal": _format_article,
    "web": _format_web,
    "website": _format_web,
    "report": _format_report,
}


def _format_source(source: Mapping[str, Any]) -> str:
    """Dispatch on ``source['kind']``, defaulting to the raw-pair formatter."""
    kind = str(source.get("kind", "")).lower()
    formatter = _FORMATTERS.get(kind, _format_unknown)
    return formatter(source)


def add_bibliography(
    document: "Document",
    sources: Sequence[Mapping[str, Any]],
    title: Optional[str] = "Bibliography",
    heading_level: int = 1,
    page_break: bool = True,
) -> List["Paragraph"]:
    """Append a bibliography section to `document` and return the new paragraphs.

    Renders an optional heading (``Heading {heading_level}``) followed
    by one paragraph per source in `sources`. Each source is a mapping
    with a ``kind`` key (``"book"``, ``"article"``, ``"web"``,
    ``"report"``) plus the per-kind fields the formatter expects.
    Unknown kinds are rendered as raw ``key=value`` pairs rather than
    silently dropped, so the author spots typos on first preview.

    The output is APA-flavoured rather than strictly APA — the kit is
    opinionated about output rather than acting as a citation
    validator, since house styles disagree on minutiae (italics on
    titles, "Inc." / "Ltd." dropping, DOI vs URL, ampersand in author
    lists). For full APA compliance run the output through a
    citation-style processor (the kit's contract is "presentable
    default", not "publishable APA").

    Per-kind fields recognised:

    * ``"book"`` — ``authors`` / ``author``, ``year``, ``title``,
      ``publisher``.
    * ``"article"`` / ``"journal"`` — ``authors``, ``year``, ``title``,
      ``journal``, ``volume``, ``issue``, ``pages``.
    * ``"web"`` / ``"website"`` — ``authors``, ``year``, ``title``,
      ``site`` (or ``publisher``), ``url``.
    * ``"report"`` — ``authors``, ``year``, ``title``, ``number`` (or
      ``report_number``), ``publisher`` (or ``institution``).

    ``authors`` may be a single string or a sequence of strings.
    Missing ``year`` is rendered as the APA-conventional ``"n.d."``
    (no date). Pass ``title=None`` (or the empty string) to suppress
    the heading and append only the source paragraphs.

    Returns the list of newly-appended |Paragraph| objects in document
    order, including the trailing page-break paragraph when
    ``page_break`` is true (the default).

    .. versionadded:: 2026.05.29
    """
    if not 0 <= heading_level <= 9:
        raise ValueError(
            "heading_level must be in 0..9 (matching Document.add_heading), "
            "got %d" % heading_level
        )

    paragraphs: List["Paragraph"] = []
    if title:
        paragraphs.append(document.add_heading(title, level=heading_level))

    style = _resolve_style(document, _STYLE_NORMAL)
    for source in sources:
        rendered = _format_source(source)
        if rendered:
            paragraphs.append(document.add_paragraph(rendered, style=style))

    if page_break:
        paragraphs.append(document.add_page_break())

    return paragraphs


__all__ = [
    "add_appendix",
    "add_bibliography",
    "add_glossary",
    "add_index",
]
