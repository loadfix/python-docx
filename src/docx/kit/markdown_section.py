"""Drop-in Markdown -> docx renderer.

Closes #54 (companion of ``Document.to_markdown()``).

The :func:`add` helper takes a Markdown blob and appends fully-formatted
content to a |Document| in a single call: headings, paragraphs with
inline ``**bold**`` / ``*italic*`` / ``` `code` ``` / ``[link](url)``
runs, bulleted and numbered lists, GFM pipe tables, blockquotes, fenced
code blocks, horizontal rules, and inline ``![alt](path)`` images. It is
the *inverse* of :meth:`docx.Document.to_markdown` — feed the output of
``to_markdown()`` back through :func:`add` and you get a doc that
renders the same body content (cosmetic fidelity only — Markdown is
strictly a subset of Word's expressiveness).

::

    from docx import Document
    from docx.kit import markdown_section

    doc = Document()
    markdown_section.add(doc, '''
    # Section title

    A paragraph with **bold** and *italic* text.

    - bullet 1
    - bullet 2

    | Col | Other |
    |-----|-------|
    | 1   | 2     |
    ''')
    doc.save("out.docx")

The implementation vendors a *minimal* stdlib-only Markdown subset
parser (no PyPI ``markdown`` / ``commonmark`` dependency) tracking the
shape of ``python-ooxml-compile``'s parser. The kit cannot import from
``python-ooxml-compile`` directly without coupling python-docx's runtime
graph to a sibling package, and the parser is small enough to vendor
cleanly. Round-tripping spec-perfect Markdown is a non-goal; what's
covered is the set of constructs ``Document.to_markdown()`` emits, plus
the common authoring shorthands callers actually paste in (fenced code
blocks, tables, images).

Every helper composes only python-docx's *public* API
(``Document.add_paragraph`` / ``add_heading`` / ``add_table`` /
``add_picture``, ``Paragraph.add_run`` / ``add_hyperlink``,
``Run.add_picture``, ``Run.font.name``). No ``_element`` / ``oxml`` /
``etree`` access. When a referenced built-in style (e.g. ``List
Bullet``, ``Intense Quote``) is missing from the loaded template, the
helper degrades to ``Normal`` rather than raising — the spirit of a kit
is "works out of the box".

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import os
import re
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Sequence, Tuple, Union

from docx.shared import Pt

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph


# ---------------------------------------------------------------------------
# Type aliases (kept loose — block / inline nodes are plain ``dict``).

# A block / inline AST node is a plain dict; aliasing keeps signatures
# readable without a TypedDict (which would inflate the parser for an
# internal data shape).
Node = Dict[str, Any]
# A returned object is either a Paragraph or a Table — both are public
# python-docx classes.
Returned = Union["Paragraph", "Table"]


# ---------------------------------------------------------------------------
# Inline parsing
# ---------------------------------------------------------------------------

# Inline regex priority: image > link (image is link-with-leading-bang)
# > code (backticks beat emphasis since `` `*not bold*` `` should stay
# code) > bold (``**``) > italic (``*``). Bold has to win over italic
# because the bold pattern is the longer match at the same start.
_RE_BOLD = re.compile(r"\*\*([^*]+)\*\*")
_RE_ITALIC = re.compile(r"(?<!\*)\*([^*]+)\*(?!\*)")
_RE_CODE = re.compile(r"`([^`]+)`")
_RE_IMAGE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
_RE_LINK = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")


def _parse_inlines(text: str) -> List[Node]:
    """Parse a single line of inline Markdown into inline AST nodes.

    The parser walks the string with the regexes above. At each
    position it picks the *earliest* match across all patterns
    (resolving ties by priority order — the order in the candidates
    tuple below). Anything not matched lands in ``text`` nodes.
    """
    nodes: List[Node] = []
    pos = 0
    n = len(text)

    while pos < n:
        candidates: List[Tuple[int, str, "re.Match[str]"]] = []
        for kind, pat in (
            ("image", _RE_IMAGE),
            ("link", _RE_LINK),
            ("code", _RE_CODE),
            ("bold", _RE_BOLD),
            ("italic", _RE_ITALIC),
        ):
            m = pat.search(text, pos)
            if m is not None:
                candidates.append((m.start(), kind, m))

        if not candidates:
            nodes.append({"type": "text", "text": text[pos:]})
            break

        candidates.sort(key=lambda c: c[0])
        start, kind, m = candidates[0]
        if start > pos:
            nodes.append({"type": "text", "text": text[pos:start]})

        if kind == "image":
            nodes.append({"type": "image", "alt": m.group(1), "src": m.group(2)})
        elif kind == "link":
            nodes.append(
                {
                    "type": "link",
                    "href": m.group(2),
                    "inlines": _parse_inlines(m.group(1)),
                }
            )
        elif kind == "code":
            nodes.append({"type": "code", "text": m.group(1)})
        elif kind == "bold":
            nodes.append({"type": "bold", "inlines": _parse_inlines(m.group(1))})
        else:  # italic
            nodes.append({"type": "italic", "inlines": _parse_inlines(m.group(1))})

        pos = m.end()

    return nodes


# ---------------------------------------------------------------------------
# Block parsing
# ---------------------------------------------------------------------------

_RE_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
_RE_FENCE = re.compile(r"^```\s*(\S*)\s*$")
_RE_HR = re.compile(r"^(---+|\*\*\*+|___+)\s*$")
_RE_BULLET = re.compile(r"^[-*+]\s+(.*)$")
_RE_NUMBER = re.compile(r"^\d+\.\s+(.*)$")
_RE_BLOCKQUOTE = re.compile(r"^>\s?(.*)$")
# Table divider must contain at least one ``-``; allows the
# single-column ``|---|`` shape as well as the multi-column
# ``|---|---|`` shape (the original compile parser required two
# columns; we relax that).
_RE_TABLE_DIVIDER = re.compile(
    r"^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$"
)


def _split_pipe_row(line: str) -> List[str]:
    """Split a ``| a | b |`` row into trimmed cells.

    Drops the empty edge tokens produced by the leading / trailing
    pipe so ``| a | b |`` returns ``["a", "b"]`` rather than
    ``["", "a", "b", ""]``.
    """
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [cell.strip() for cell in s.split("|")]


def _parse_blocks(source: str) -> List[Node]:
    """Parse a Markdown document into a list of block AST nodes."""
    text = source.replace("\r\n", "\n").replace("\r", "\n")
    if text.startswith("﻿"):
        text = text[1:]
    lines = text.split("\n")

    blocks: List[Node] = []
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # -- Blank line — consume.
        if not stripped:
            i += 1
            continue

        # -- Fenced code block.
        m_fence = _RE_FENCE.match(stripped)
        if m_fence is not None:
            lang = m_fence.group(1) or ""
            body: List[str] = []
            i += 1
            while i < n and not _RE_FENCE.match(lines[i].strip()):
                body.append(lines[i])
                i += 1
            if i < n:  # skip closing fence
                i += 1
            blocks.append({"type": "code_block", "lang": lang, "text": "\n".join(body)})
            continue

        # -- Horizontal rule.
        if _RE_HR.match(stripped):
            blocks.append({"type": "hr"})
            i += 1
            continue

        # -- ATX heading (``# Title`` … ``###### Title``).
        m_h = _RE_HEADING.match(stripped)
        if m_h is not None:
            level = len(m_h.group(1))
            blocks.append(
                {
                    "type": "heading",
                    "level": level,
                    "inlines": _parse_inlines(m_h.group(2).strip()),
                }
            )
            i += 1
            continue

        # -- Pipe table — header line followed by a divider line.
        if "|" in stripped and i + 1 < n and _RE_TABLE_DIVIDER.match(lines[i + 1]):
            header = [_parse_inlines(c) for c in _split_pipe_row(line)]
            i += 2  # past header + divider
            rows: List[List[List[Node]]] = []
            while i < n and "|" in lines[i] and lines[i].strip():
                rows.append([_parse_inlines(c) for c in _split_pipe_row(lines[i])])
                i += 1
            blocks.append({"type": "table", "header": header, "rows": rows})
            continue

        # -- Bulleted list.
        if _RE_BULLET.match(stripped):
            items: List[List[Node]] = []
            while i < n:
                m_b = _RE_BULLET.match(lines[i].strip())
                if m_b is None:
                    break
                items.append(_parse_inlines(m_b.group(1)))
                i += 1
            blocks.append({"type": "list", "ordered": False, "items": items})
            continue

        # -- Numbered list.
        if _RE_NUMBER.match(stripped):
            items = []
            while i < n:
                m_o = _RE_NUMBER.match(lines[i].strip())
                if m_o is None:
                    break
                items.append(_parse_inlines(m_o.group(1)))
                i += 1
            blocks.append({"type": "list", "ordered": True, "items": items})
            continue

        # -- Blockquote — consume contiguous ``>`` lines into a single
        # -- collapsed paragraph (multi-line blockquotes flatten to one
        # -- paragraph; round-trippable enough for the kit).
        if _RE_BLOCKQUOTE.match(stripped):
            body_lines: List[str] = []
            while i < n:
                m_q = _RE_BLOCKQUOTE.match(lines[i].strip())
                if m_q is None:
                    break
                body_lines.append(m_q.group(1))
                i += 1
            blocks.append(
                {
                    "type": "blockquote",
                    "inlines": _parse_inlines(" ".join(body_lines).strip()),
                }
            )
            continue

        # -- Paragraph — consume until blank line or another block opener.
        para_lines: List[str] = [line]
        i += 1
        while i < n and lines[i].strip() and not _is_block_opener(lines[i], lines, i):
            para_lines.append(lines[i])
            i += 1
        blocks.append(
            {
                "type": "paragraph",
                "inlines": _parse_inlines(" ".join(para_lines).strip()),
            }
        )

    return blocks


def _is_block_opener(line: str, all_lines: Sequence[str], i: int) -> bool:
    """Heuristic — would ``line`` start a new non-paragraph block?

    Used to terminate a paragraph mid-stream when the *next* line
    really starts e.g. a heading or a list. Without this, a paragraph
    would greedily swallow the heading on the next line.
    """
    s = line.strip()
    if not s:
        return True
    if _RE_HEADING.match(s):
        return True
    if _RE_FENCE.match(s):
        return True
    if _RE_HR.match(s):
        return True
    if _RE_BULLET.match(s):
        return True
    if _RE_NUMBER.match(s):
        return True
    if _RE_BLOCKQUOTE.match(s):
        return True
    if "|" in s and i + 1 < len(all_lines) and _RE_TABLE_DIVIDER.match(all_lines[i + 1]):
        return True
    return False


# ---------------------------------------------------------------------------
# Style resolution
# ---------------------------------------------------------------------------

# Style names we look up under ``style_prefix`` first, then fall back to
# python-docx's built-in defaults, then finally to ``Normal``. Each tuple
# is (prefixed_lookup_suffix, builtin_fallback).
_STYLE_NORMAL = "Normal"


def _has_style(document: Any, style_name: str) -> bool:
    """Return True when ``document`` defines a paragraph style named ``style_name``."""
    try:
        styles = document.styles
    except Exception:  # pragma: no cover - defensive
        return False
    try:
        styles[style_name]
        return True
    except KeyError:
        return False


def _resolve(document: Any, prefixed: str, builtin: str) -> str:
    """Resolve a style name preferring ``prefixed`` over ``builtin``.

    Looks up ``prefixed`` (e.g. ``"MD Heading 1"``) first; if absent,
    falls back to ``builtin`` (``"Heading 1"``); if that's also missing,
    returns ``"Normal"``. Lets corporate templates override the kit's
    rendered styles wholesale (define ``MD Heading 1`` and the kit will
    use it) without forcing every caller to register the prefixed
    variants.
    """
    if _has_style(document, prefixed):
        return prefixed
    if _has_style(document, builtin):
        return builtin
    return _STYLE_NORMAL


# ---------------------------------------------------------------------------
# Rendering — inline trees -> Paragraph runs
# ---------------------------------------------------------------------------


def _flatten_inlines(
    inlines: Sequence[Node],
    *,
    bold: bool,
    italic: bool,
    code: bool,
    link: Optional[str],
) -> List[Dict[str, Any]]:
    """Flatten a nested inline tree to a flat list of styled runs.

    Each output dict carries the resolved style flags (``bold`` /
    ``italic`` / ``code``), an optional ``link`` URL, and either a
    ``text`` string or an ``image`` path (the renderer dispatches on
    whichever is present). Bold + italic compose: ``**bold *and italic***``
    becomes a single run with both flags set.
    """
    out: List[Dict[str, Any]] = []
    for node in inlines:
        kind = node["type"]
        if kind == "text":
            out.append(
                {
                    "kind": "text",
                    "text": node["text"],
                    "bold": bold,
                    "italic": italic,
                    "code": code,
                    "link": link,
                }
            )
        elif kind == "code":
            out.append(
                {
                    "kind": "text",
                    "text": node["text"],
                    "bold": bold,
                    "italic": italic,
                    "code": True,
                    "link": link,
                }
            )
        elif kind == "image":
            out.append(
                {
                    "kind": "image",
                    "src": node.get("src", ""),
                    "alt": node.get("alt", ""),
                    "bold": bold,
                    "italic": italic,
                    "code": code,
                    "link": link,
                }
            )
        elif kind == "bold":
            out.extend(
                _flatten_inlines(
                    node["inlines"], bold=True, italic=italic, code=code, link=link
                )
            )
        elif kind == "italic":
            out.extend(
                _flatten_inlines(
                    node["inlines"], bold=bold, italic=True, code=code, link=link
                )
            )
        elif kind == "link":
            out.extend(
                _flatten_inlines(
                    node["inlines"],
                    bold=bold,
                    italic=italic,
                    code=code,
                    link=node.get("href", ""),
                )
            )
    return out


def _emit_inlines(
    paragraph: Any,
    inlines: Sequence[Node],
    *,
    inline_images: bool,
) -> None:
    """Walk an inline tree and append runs / hyperlinks / pictures to ``paragraph``."""
    runs = _flatten_inlines(
        inlines, bold=False, italic=False, code=False, link=None
    )
    for spec in runs:
        if spec["kind"] == "image":
            _emit_image_run(paragraph, spec, inline_images=inline_images)
            continue

        text = spec["text"]
        if not text:
            continue

        if spec["link"]:
            # -- Public hyperlink helper handles the rId + relationship.
            try:
                paragraph.add_hyperlink(url=spec["link"], text=text)
            except Exception:
                # -- Some templates lack the ``Hyperlink`` character style;
                # -- fall back to a plain run with the URL appended so the
                # -- target isn't lost.
                run = paragraph.add_run(text + " (" + spec["link"] + ")")
                _apply_run_format(run, spec)
            continue

        run = paragraph.add_run(text)
        _apply_run_format(run, spec)


def _apply_run_format(run: Any, spec: Dict[str, Any]) -> None:
    """Apply the resolved bold / italic / code formatting to ``run``."""
    if spec["bold"]:
        run.bold = True
    if spec["italic"]:
        run.italic = True
    if spec["code"]:
        run.font.name = "Courier New"


def _emit_image_run(
    paragraph: Any, spec: Dict[str, Any], *, inline_images: bool
) -> None:
    """Append an inline picture to ``paragraph`` (or fall back to alt text).

    The kit attempts an in-package picture embed via ``Run.add_picture``
    when ``inline_images`` is true *and* the ``src`` resolves to a
    readable local file. Anything else (remote URLs, missing files,
    ``inline_images=False``) falls back to a textual ``[image: alt]``
    placeholder so the information isn't silently dropped.
    """
    src = spec.get("src", "") or ""
    alt = spec.get("alt", "") or ""
    if inline_images and src and not _is_remote(src) and os.path.isfile(src):
        try:
            run = paragraph.add_run()
            run.add_picture(src)
            return
        except Exception:
            # -- Fall through to the placeholder run on any embed error
            # -- (corrupt image, unsupported format, EMU overflow).
            pass
    placeholder = "[image: " + alt + "]" if alt else "[image]"
    run = paragraph.add_run(placeholder)
    _apply_run_format(run, spec)


def _is_remote(src: str) -> bool:
    """Return True for ``http(s)://`` / ``file://`` / ``data:`` URIs."""
    lower = src.strip().lower()
    return (
        lower.startswith("http://")
        or lower.startswith("https://")
        or lower.startswith("ftp://")
        or lower.startswith("data:")
        or lower.startswith("file://")
    )


# ---------------------------------------------------------------------------
# Rendering — block AST -> Document mutations
# ---------------------------------------------------------------------------


def _render_heading(
    document: Any, block: Node, style_prefix: str, *, inline_images: bool
) -> "Paragraph":
    """Render a ``# … ######`` heading."""
    # -- Clamp level to 1..9 (python-docx caps add_heading at 0..9 with
    # -- 0 == Title; Markdown headings start at level 1).
    level = max(1, min(9, int(block.get("level", 1))))
    builtin = "Heading %d" % level
    prefixed = style_prefix + builtin
    style = _resolve(document, prefixed, builtin)
    # -- ``add_heading`` would re-derive the style from level; we want the
    # -- prefix-aware override so call ``add_paragraph`` directly instead.
    para = document.add_paragraph("", style=style)
    _emit_inlines(para, block.get("inlines", []), inline_images=inline_images)
    return para


def _render_paragraph(
    document: Any, block: Node, style_prefix: str, *, inline_images: bool
) -> "Paragraph":
    """Render a regular body paragraph."""
    builtin = _STYLE_NORMAL
    prefixed = style_prefix + "Body"
    style = _resolve(document, prefixed, builtin)
    para = document.add_paragraph(style=style)
    _emit_inlines(para, block.get("inlines", []), inline_images=inline_images)
    return para


def _render_blockquote(
    document: Any, block: Node, style_prefix: str, *, inline_images: bool
) -> "Paragraph":
    """Render a single-line collapsed ``> blockquote``."""
    prefixed = style_prefix + "Quote"
    style = _resolve(document, prefixed, "Intense Quote")
    if style == _STYLE_NORMAL:
        # -- Try the lighter "Quote" style as a second fallback before
        # -- giving up to Normal — many templates ship Quote but not
        # -- Intense Quote.
        style = _resolve(document, prefixed, "Quote")
    para = document.add_paragraph(style=style)
    _emit_inlines(para, block.get("inlines", []), inline_images=inline_images)
    return para


def _render_code_block(document: Any, block: Node, style_prefix: str) -> "Paragraph":
    """Render a fenced code block as a Courier-styled paragraph.

    Multi-line code blocks emit a single paragraph containing one run
    per line separated by line-breaks (the run's text contains
    embedded ``\\n`` characters which Word renders as soft breaks
    inside the run on save).
    """
    prefixed = style_prefix + "Code"
    style = _resolve(document, prefixed, _STYLE_NORMAL)
    para = document.add_paragraph(style=style)
    text = block.get("text", "") or ""
    run = para.add_run(text)
    run.font.name = "Courier New"
    # -- Slightly smaller for code-block readability; honoured by Word's
    # -- default rendering. Skipped when a prefixed ``MD Code`` style is
    # -- present (the template author has expressed an opinion).
    if style == _STYLE_NORMAL:
        run.font.size = Pt(10)
    return para


def _render_hr(document: Any, style_prefix: str) -> "Paragraph":
    """Render a horizontal rule.

    python-docx has no first-class ``<hr/>``; we emit a centred row of
    em-dashes which is what ``Document.to_markdown()`` interprets as
    ``---`` on the round-trip path. Avoids reaching down into the oxml
    layer for a ``w:pBdr`` border element.
    """
    style = _resolve(document, style_prefix + "HorizontalRule", _STYLE_NORMAL)
    para = document.add_paragraph(style=style)
    para.add_run("―" * 40)
    return para


def _render_list(
    document: Any, block: Node, style_prefix: str, *, inline_images: bool
) -> List["Paragraph"]:
    """Render a bulleted or numbered list as one paragraph per item."""
    ordered = bool(block.get("ordered", False))
    builtin = "List Number" if ordered else "List Bullet"
    prefixed = style_prefix + builtin
    style = _resolve(document, prefixed, builtin)
    paragraphs: List["Paragraph"] = []
    for item in block.get("items", []):
        para = document.add_paragraph(style=style)
        _emit_inlines(para, item, inline_images=inline_images)
        paragraphs.append(para)
    return paragraphs


def _render_table(
    document: Any, block: Node, style_prefix: str, *, inline_images: bool
) -> "Table":
    """Render a GFM pipe table.

    Header row is rendered as bold runs; each body cell is an
    inline-formatted paragraph. Falls back to ``Table Grid`` when the
    prefixed ``MD Table`` style is missing — ``Table Grid`` is in
    python-docx's default styles template so it's reliably present.
    """
    header = block.get("header", []) or []
    rows = block.get("rows", []) or []
    n_cols = len(header)
    if n_cols == 0:
        # -- Degenerate empty table — emit a 1x1 placeholder so the
        # -- caller still gets a Table back.
        n_cols = 1
    table = document.add_table(rows=1 + len(rows), cols=n_cols)

    prefixed_style = style_prefix + "Table"
    if _has_style(document, prefixed_style):
        try:
            table.style = prefixed_style
        except Exception:
            pass
    elif _has_style(document, "Table Grid"):
        try:
            table.style = "Table Grid"
        except Exception:
            pass

    # -- Header row (bold runs).
    for col_idx, cell_inlines in enumerate(header[:n_cols]):
        cell = table.rows[0].cells[col_idx]
        cell.text = ""
        para = cell.paragraphs[0]
        _emit_inlines(para, cell_inlines, inline_images=inline_images)
        for run in para.runs:
            run.bold = True

    # -- Body rows.
    for row_idx, row in enumerate(rows, start=1):
        for col_idx in range(n_cols):
            cell = table.rows[row_idx].cells[col_idx]
            cell.text = ""
            cell_inlines = row[col_idx] if col_idx < len(row) else []
            _emit_inlines(
                cell.paragraphs[0], cell_inlines, inline_images=inline_images
            )

    return table


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def add(
    document: "Document",
    markdown_text: str,
    *,
    style_prefix: str = "MD ",
    inline_images: bool = True,
) -> List[Returned]:
    """Append ``markdown_text`` to ``document`` as fully-formatted content.

    Parses ``markdown_text`` and renders each block to the end of the
    document body, preserving inline formatting. Returns the list of
    newly-appended |Paragraph| / |Table| objects (in document order)
    so callers can post-process them — apply additional styles, attach
    bookmarks, change alignment, etc.

    Supported Markdown constructs:

    * ATX headings ``#`` … ``######`` (``Heading 1`` … ``Heading 6``).
    * Paragraphs with inline ``**bold**``, ``*italic*``,
      ``` `inline code` ``` and ``[label](url)`` hyperlinks.
    * Inline images ``![alt](path)`` — embedded via
      :meth:`Run.add_picture` when ``path`` resolves to a readable
      local file and ``inline_images`` is true; otherwise a
      ``[image: alt]`` placeholder run.
    * Bulleted lists (``- ``, ``* ``, ``+ ``) and numbered lists
      (``1. ``) — each item becomes one ``List Bullet`` /
      ``List Number`` paragraph.
    * GFM pipe tables (header row + ``---`` divider + body rows) —
      rendered with bold header cells and the ``Table Grid`` style.
    * Single- and multi-line blockquotes (``> ``) — collapsed to one
      ``Intense Quote`` / ``Quote`` paragraph.
    * Fenced code blocks (```` ``` ````) — rendered as a single
      Courier-styled paragraph.
    * Horizontal rules (``---``, ``***``, ``___``) — rendered as a
      centred row of em-dashes.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    markdown_text
        The Markdown blob to render.
    style_prefix
        Style-name prefix the helper checks for *before* falling back
        to python-docx's built-in style names. Defaults to ``"MD "``,
        so the helper looks for ``"MD Heading 1"`` before ``"Heading
        1"``, ``"MD Body"`` before ``"Normal"``, ``"MD Code"`` before
        ``"Normal"``, ``"MD Table"`` before ``"Table Grid"``, etc. Lets
        a corporate template override the kit's appearance wholesale by
        defining the prefixed variants. Pass ``""`` to skip the prefix
        check and use built-ins directly.
    inline_images
        When |True| (the default), inline ``![alt](path)`` images are
        embedded via :meth:`Run.add_picture`. When |False| (or when the
        path is remote / missing), the image renders as a ``[image:
        alt]`` placeholder run.

    Returns
    -------
    list of Paragraph or Table
        The newly-appended top-level objects in document order. List
        items appear once per item paragraph; table cells are accessed
        through the returned :class:`Table` object.

    Notes
    -----
    Round-tripping is best-effort: feeding the output of
    :meth:`Document.to_markdown` back through :func:`add` produces a
    document that *renders the same body content*, but cosmetic
    fidelity (run-level fonts, paragraph spacing) is not preserved —
    Markdown is strictly a subset of Word's expressiveness.

    Inline content inside table cells, list items, and blockquotes
    supports the same bold / italic / code / link primitives as
    paragraphs. Block-level constructs nested inside lists or tables
    (a sub-list inside a list item, a table inside a blockquote) are
    rendered as if they were top-level — the parser does not track
    indentation depth.

    .. versionadded:: 2026.05.29
    """
    if markdown_text is None:
        raise TypeError("markdown_text must be a string, got None")

    blocks = _parse_blocks(markdown_text)

    appended: List[Returned] = []
    for block in blocks:
        kind = block.get("type")
        if kind == "heading":
            appended.append(
                _render_heading(
                    document, block, style_prefix, inline_images=inline_images
                )
            )
        elif kind == "paragraph":
            appended.append(
                _render_paragraph(
                    document, block, style_prefix, inline_images=inline_images
                )
            )
        elif kind == "blockquote":
            appended.append(
                _render_blockquote(
                    document, block, style_prefix, inline_images=inline_images
                )
            )
        elif kind == "code_block":
            appended.append(_render_code_block(document, block, style_prefix))
        elif kind == "hr":
            appended.append(_render_hr(document, style_prefix))
        elif kind == "list":
            appended.extend(
                _render_list(
                    document, block, style_prefix, inline_images=inline_images
                )
            )
        elif kind == "table":
            appended.append(
                _render_table(
                    document, block, style_prefix, inline_images=inline_images
                )
            )

    return appended


__all__ = ["add"]
