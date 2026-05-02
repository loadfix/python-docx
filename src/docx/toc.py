"""Table-of-contents (TOC) building helpers.

A Word table of contents is a field — a paragraph that contains a ``TOC``
instruction which Word evaluates at display time to produce a list of
headings with page numbers. The field uses the "complex" XML shape
(``w:fldChar`` begin / separate / end markers around a ``w:instrText``) and
between the *separate* and *end* markers Word caches a preview of the
rendered TOC. Word rebuilds the result when the document is opened or when
the user asks to update fields, so the cached result is purely a preview
used when the document is viewed by a consumer that does not itself
evaluate fields (e.g. a raw-XML tool or Word in an "update fields?"
prompt-declined state).

The XML shape produced by this module looks approximately like::

    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r>
        <w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>
      </w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t xml:space="preserve">Heading one\\t1</w:t></w:r>
      <w:r><w:br/></w:r>
      <w:r><w:t xml:space="preserve">Heading two\\t2</w:t></w:r>
      ...
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>

The per-entry ``\\t`` and "page hint" are cosmetic: python-docx has no
layout engine, so the cached page number is a 1-based heading index rather
than a true page number. Word discards it and recomputes the real page
numbers on open.

This module exposes two helpers — :func:`build_toc_instruction` (builds
the ``TOC`` instruction string for a level range) and
:func:`populate_toc_paragraph` (populates a freshly-created empty
paragraph with the TOC field). The public API is surfaced via
:meth:`docx.document.Document.add_table_of_contents`,
:meth:`docx.text.paragraph.Paragraph.insert_table_of_contents_before`,
and :meth:`docx.text.paragraph.Paragraph.insert_table_of_contents_after`.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Iterable

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph


# -- heading style names look like "Heading 1" .. "Heading 9" (case-insensitive).
#    The regex is shared with `docx.accessibility` but duplicated here to avoid
#    importing that module just for a regex. --
_HEADING_RE = re.compile(r"^heading\s+([1-9])$", re.IGNORECASE)


def _paragraph_heading_level(paragraph: Paragraph) -> int | None:
    """Return the integer heading level for `paragraph`, or |None| if not a heading.

    A paragraph is considered a heading when its style name matches ``"Heading N"``
    (case-insensitively) for ``N`` in 1..9.
    """
    style = paragraph.style
    if style is None:
        return None
    name = style.name
    if name is None:
        return None
    match = _HEADING_RE.match(name.strip())
    if match is None:
        return None
    return int(match.group(1))


def _validate_levels(levels: tuple[int, int]) -> tuple[int, int]:
    """Return `levels` after validating shape and contents.

    `levels` must be a 2-tuple of integers ``(min_level, max_level)`` where
    ``1 <= min_level <= max_level <= 9``. ``levels`` is also accepted as a
    list of two integers for caller convenience.
    """
    try:
        min_level, max_level = levels
    except (TypeError, ValueError):
        raise ValueError(
            "levels must be a 2-tuple of ints (min_level, max_level), got %r"
            % (levels,)
        )
    if not (isinstance(min_level, int) and isinstance(max_level, int)):  # pyright: ignore[reportUnnecessaryIsInstance]
        raise ValueError(
            "levels must be a 2-tuple of ints (min_level, max_level), got %r"
            % (levels,)
        )
    if not 1 <= min_level <= max_level <= 9:
        raise ValueError(
            "levels must satisfy 1 <= min_level <= max_level <= 9, got %r"
            % (levels,)
        )
    return (min_level, max_level)


def build_toc_instruction(levels: tuple[int, int] = (1, 3)) -> str:
    """Return the ``TOC`` field instruction string for a heading-level range.

    `levels` is a ``(min_level, max_level)`` tuple. The produced instruction
    uses Word's conventional switches:

    * ``\\o "min-max"`` — build from outline levels ``min..max``
    * ``\\h`` — render entries as hyperlinks
    * ``\\z`` — hide tab-leader and page numbers in web view
    * ``\\u`` — use applied paragraph outline levels (not just headings)

    The returned string is wrapped in single spaces, matching the form Word
    writes when it inserts a TOC via the Ribbon.

    .. versionadded:: 1.3.0.dev0
    """
    min_level, max_level = _validate_levels(levels)
    return f' TOC \\o "{min_level}-{max_level}" \\h \\z \\u '


def _collect_entries(
    paragraphs: Iterable[Paragraph], levels: tuple[int, int]
) -> list[tuple[int, str]]:
    """Return ``(level, text)`` pairs for each heading in `paragraphs` matching `levels`."""
    min_level, max_level = levels
    entries: list[tuple[int, str]] = []
    for paragraph in paragraphs:
        level = _paragraph_heading_level(paragraph)
        if level is None:
            continue
        if level < min_level or level > max_level:
            continue
        entries.append((level, paragraph.text))
    return entries


def _render_result_text(entries: list[tuple[int, str]]) -> str:
    """Return a newline-joined cached TOC preview built from `entries`.

    Each entry becomes ``"{text}\\t{index}"`` where ``index`` is the 1-based
    position of the heading in the filtered list — a stand-in for the page
    number python-docx cannot compute. An empty `entries` list produces an
    empty string.
    """
    lines: list[str] = []
    for idx, (_, text) in enumerate(entries, start=1):
        lines.append(f"{text}\t{idx}")
    return "\n".join(lines)


def populate_toc_paragraph(
    paragraph: Paragraph,
    source_paragraphs: Iterable[Paragraph],
    levels: tuple[int, int] = (1, 3),
) -> Paragraph:
    """Populate `paragraph` with a TOC complex field and return it.

    `paragraph` must be an empty, freshly-created |Paragraph|.
    `source_paragraphs` is the iterable of paragraphs to scan for headings
    (typically ``document.paragraphs``). `levels` selects the heading-level
    range to include in the TOC (default H1..H3).

    The paragraph's style is set to ``"TOC Heading"`` would be conventional,
    but since that style is not guaranteed to exist we leave the style
    untouched; callers can assign a style explicitly if they have one
    defined. Word rebuilds the TOC on open or field-update, so the cached
    result added here is intended only as a preview for consumers that do
    not themselves evaluate fields.

    .. versionadded:: 1.3.0.dev0
    """
    levels = _validate_levels(levels)
    entries = _collect_entries(source_paragraphs, levels)
    instr = build_toc_instruction(levels)
    result_text = _render_result_text(entries)
    # -- pass None when result is empty so add_complex_field emits no
    #    separator run; Word still happily renders the empty TOC. --
    paragraph.add_complex_field(instr, result_text if result_text else None)
    return paragraph
