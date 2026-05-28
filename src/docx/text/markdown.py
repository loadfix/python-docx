"""Inline-Markdown rendering helpers for python-docx paragraphs and cells.

Used by :meth:`docx.text.paragraph.Paragraph.add_markdown` and
:meth:`docx.table._Cell.add_markdown` (issue #23). The implementation is
a small hand-rolled parser (no `markdown` / `mistune` runtime dependency
— project policy) that supports the subset called out in the issue:

* inline emphasis: ``**bold**`` / ``__bold__``, ``*italic*`` /
  ``_italic_``, inline code (``` `code` ```)
* inline links: ``[text](url)``
* block bullet lists (``-`` / ``*``) → ``style="List Bullet"`` paragraphs
* numbered lists (``1.`` …) → ``style="List Number"`` paragraphs
* ATX headings (``#`` … ``######``) — only at the top of the input;
  rendered as ``Heading 1`` … ``Heading 6`` paragraphs
* blank-line paragraph separator
* single ``\\n`` inside a block becomes a soft line-break (``w:br``)

Round-trip note: the markdown source is **not** preserved on the
paragraph. Once :meth:`add_markdown` returns, the paragraph (and any
added sibling paragraphs for lists / blank-line separators) contain the
equivalent OOXML and a subsequent read does not recover the original
markdown string. This matches the way every other "author-time helper"
in python-docx behaves.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, List, NamedTuple, Optional

if TYPE_CHECKING:
    from docx.table import _Cell
    from docx.text.paragraph import Paragraph


# ---------------------------------------------------------------------------
# Inline tokenizer
# ---------------------------------------------------------------------------


class InlineRun(NamedTuple):
    """A piece of inline-formatted text destined for a single ``w:r`` element.

    ``link`` is non-None when the run should be wrapped in a
    ``w:hyperlink`` (external relationship of type HYPERLINK).
    """

    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    link: Optional[str] = None


_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)\s]+)\)")
_CODE_RE = re.compile(r"`([^`\n]+)`")
_BOLD_RE = re.compile(r"\*\*([^*\n]+?)\*\*|__([^_\n]+?)__")
_ITAL_RE = re.compile(r"(?<![*\w])\*([^*\n]+?)\*(?!\*)|(?<![_\w])_([^_\n]+?)_(?!_)")


def tokenize_inline(text: str) -> List[InlineRun]:
    """Return a list of |InlineRun| covering `text` left-to-right."""
    runs: List[InlineRun] = []
    pos = 0
    n = len(text)
    plain: List[str] = []

    def flush_plain() -> None:
        if plain:
            runs.append(InlineRun("".join(plain)))
            plain.clear()

    while pos < n:
        ch = text[pos]
        if ch == "`":
            m = _CODE_RE.match(text, pos)
            if m is not None:
                flush_plain()
                runs.append(InlineRun(m.group(1), code=True))
                pos = m.end()
                continue
        if ch == "[":
            m = _LINK_RE.match(text, pos)
            if m is not None:
                flush_plain()
                inner = tokenize_inline(m.group(1))
                url = m.group(2)
                for r in inner:
                    runs.append(r._replace(link=url))
                pos = m.end()
                continue
        if ch in ("*", "_") and pos + 1 < n and text[pos + 1] == ch:
            m = _BOLD_RE.match(text, pos)
            if m is not None:
                flush_plain()
                inner = tokenize_inline(m.group(1) or m.group(2) or "")
                for r in inner:
                    runs.append(r._replace(bold=True))
                pos = m.end()
                continue
        if ch in ("*", "_"):
            m = _ITAL_RE.match(text, pos)
            if m is not None:
                flush_plain()
                inner = tokenize_inline(m.group(1) or m.group(2) or "")
                for r in inner:
                    runs.append(r._replace(italic=True))
                pos = m.end()
                continue
        plain.append(ch)
        pos += 1

    flush_plain()
    return runs


# ---------------------------------------------------------------------------
# Block tokenizer
# ---------------------------------------------------------------------------


class Block(NamedTuple):
    """A logical block-level item discovered by :func:`tokenize_blocks`."""

    kind: str  # "heading" | "bullet" | "number" | "para" | "blank"
    level: int  # heading depth (1..6) or 0 for non-headings
    text: str


_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_BULLET_RE = re.compile(r"^(\s*)([-*])\s+(.*)$")
_NUMBER_RE = re.compile(r"^(\s*)(\d+)\.\s+(.*)$")


def tokenize_blocks(md: str) -> List[Block]:
    """Split `md` into a flat list of block-level items in source order."""
    out: List[Block] = []
    lines = md.split("\n")
    i = 0
    n = len(lines)
    seen_content = False
    para_buf: List[str] = []

    def flush_para() -> None:
        if para_buf:
            out.append(Block("para", 0, "\n".join(para_buf)))
            para_buf.clear()

    while i < n:
        line = lines[i]
        if line.strip() == "":
            flush_para()
            out.append(Block("blank", 0, ""))
            i += 1
            continue

        if not seen_content:
            mh = _HEADING_RE.match(line)
            if mh is not None:
                out.append(Block("heading", len(mh.group(1)), mh.group(2).strip()))
                seen_content = True
                i += 1
                continue

        seen_content = True

        mb = _BULLET_RE.match(line)
        if mb is not None:
            flush_para()
            out.append(Block("bullet", 0, mb.group(3)))
            i += 1
            continue

        mn = _NUMBER_RE.match(line)
        if mn is not None:
            flush_para()
            out.append(Block("number", 0, mn.group(3)))
            i += 1
            continue

        para_buf.append(line)
        i += 1

    flush_para()
    while out and out[-1].kind == "blank":
        out.pop()
    return out


# ---------------------------------------------------------------------------
# Render helpers
# ---------------------------------------------------------------------------


def _emit_inline(paragraph: "Paragraph", runs: List[InlineRun]) -> None:
    """Emit `runs` into `paragraph` with the appropriate run formatting."""
    # -- Group consecutive runs that share the same link URL into a single
    # -- w:hyperlink so Word renders one click target. Plain (non-link)
    # -- runs are appended directly to the paragraph.
    i = 0
    while i < len(runs):
        r = runs[i]
        if r.link is None:
            if r.text:
                _add_inline_run(paragraph, r)
            i += 1
            continue
        # -- accumulate consecutive runs sharing this link url --
        url = r.link
        group: List[InlineRun] = []
        while i < len(runs) and runs[i].link == url:
            group.append(runs[i])
            i += 1
        if not any(g.text for g in group):
            continue
        # -- create one hyperlink with concatenated visible text, then
        # -- decorate the spawned run(s) with bold/italic/code as needed.
        _add_link_group(paragraph, url, group)


def _add_inline_run(paragraph: "Paragraph", r: InlineRun) -> None:
    run = paragraph.add_run(r.text)
    if r.bold:
        run.bold = True
    if r.italic:
        run.italic = True
    if r.code:
        run.font.name = "Consolas"


def _hyperlink_style_available(paragraph: "Paragraph") -> bool:
    """Return True when the document's style table defines a "Hyperlink" style.

    Word ships with a latent Hyperlink style that materialises on first
    use, but the bare ``Document()`` template doesn't pre-declare it,
    so :meth:`Paragraph.add_hyperlink` raises ``KeyError`` when asked
    to apply ``style="Hyperlink"`` to a fresh document. We probe the
    style table instead of try/except so the code path is symmetric for
    typed callers.
    """
    try:
        # -- paragraph.part is the StoryPart; the document part is its parent.
        document_part = paragraph.part._document_part  # type: ignore[attr-defined]  # pyright: ignore[reportPrivateUsage]
        styles = document_part.styles
    except AttributeError:
        return False
    try:
        styles["Hyperlink"]
        return True
    except KeyError:
        return False


def _add_link_group(paragraph: "Paragraph", url: str, group: List[InlineRun]) -> None:
    """Append a single ``w:hyperlink`` with `url` carrying every run in `group`.

    The first run is created via :meth:`Paragraph.add_hyperlink`, and any
    additional fragments are appended as siblings inside the same
    ``w:hyperlink`` element so they share the relationship and style.
    """
    from docx.oxml.ns import qn
    from docx.oxml.parser import OxmlElement

    first = group[0]
    style: Optional[str] = (
        "Hyperlink" if _hyperlink_style_available(paragraph) else None
    )
    hyperlink = paragraph.add_hyperlink(url=url, text=first.text, style=style)
    # -- decorate the first run with bold/italic/code on top of "Hyperlink"
    runs = hyperlink.runs
    if runs:
        if first.bold:
            runs[0].bold = True
        if first.italic:
            runs[0].italic = True
        if first.code:
            runs[0].font.name = "Consolas"

    # -- additional fragments share the rId; build raw w:r elements and
    # -- append them to the existing w:hyperlink so a single click target
    # -- spans them all.
    if len(group) > 1:
        h_elm = hyperlink._hyperlink  # pyright: ignore[reportPrivateUsage]
        for frag in group[1:]:
            if not frag.text:
                continue
            r_elm = OxmlElement("w:r")
            rPr = OxmlElement("w:rPr")
            if style is not None:
                rStyle = OxmlElement("w:rStyle")
                rStyle.set(qn("w:val"), style)
                rPr.append(rStyle)
            if frag.bold:
                rPr.append(OxmlElement("w:b"))
            if frag.italic:
                rPr.append(OxmlElement("w:i"))
            if frag.code:
                rFonts = OxmlElement("w:rFonts")
                rFonts.set(qn("w:ascii"), "Consolas")
                rFonts.set(qn("w:hAnsi"), "Consolas")
                rPr.append(rFonts)
            if len(rPr) > 0:
                r_elm.append(rPr)
            t = OxmlElement("w:t")
            t.text = frag.text
            if frag.text != frag.text.strip():
                t.set(qn("xml:space"), "preserve")
            r_elm.append(t)
            h_elm.append(r_elm)


def _try_set_style(paragraph: "Paragraph", style_name: str) -> None:
    """Apply `style_name` to `paragraph` when the style is in the table.

    Word's default template carries the *latent* List Bullet / List
    Number / Heading N styles in ``latentStyles``; the proxy materialises
    them on demand the first time they're requested. On a bare
    ``Document()`` they may not yet be materialised — and assigning a
    style name that isn't present in the styles part raises
    ``KeyError``. We probe first and fall back to leaving the paragraph
    in the default style so ``add_markdown`` always succeeds.
    """
    try:
        paragraph.style = style_name
    except KeyError:
        pass


def _emit_block_into(
    paragraph: "Paragraph", block: Block, has_runs: bool
) -> None:
    """Render `block` into `paragraph` (assumed empty of relevant content)."""
    if block.kind == "heading":
        _try_set_style(paragraph, "Heading %d" % block.level)
        _emit_inline(paragraph, tokenize_inline(block.text))
        return
    if block.kind == "bullet":
        _try_set_style(paragraph, "List Bullet")
        _emit_inline(paragraph, tokenize_inline(block.text))
        return
    if block.kind == "number":
        _try_set_style(paragraph, "List Number")
        _emit_inline(paragraph, tokenize_inline(block.text))
        return
    if block.kind == "blank":
        # -- caller arranged for an empty paragraph; nothing to add.
        return
    # -- "para" — emit inline runs with explicit soft-breaks between
    # -- the source lines (single \n inside a paragraph).
    lines = block.text.split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            paragraph.add_run().add_break()
        _emit_inline(paragraph, tokenize_inline(line))


def apply_markdown_to_paragraph(paragraph: "Paragraph", md: str) -> None:
    """Render `md` starting from `paragraph`, appending sibling paragraphs.

    The first parsed block populates `paragraph` itself. Any further
    blocks (subsequent list items, paragraph separators, …) are added
    via :meth:`Paragraph.insert_paragraph_after` on the previously
    emitted paragraph so the result lives in the paragraph's parent
    block container (body, cell, header / footer, …).
    """
    if not isinstance(md, str):
        raise TypeError(f"md must be str, got {type(md).__name__}")

    blocks = tokenize_blocks(md)
    if not blocks:
        return

    current = paragraph
    first = True
    for block in blocks:
        if first:
            _emit_block_into(current, block, has_runs=False)
            first = False
            continue
        # -- append a fresh sibling paragraph after the previous one.
        next_p = current.insert_paragraph_after()
        _emit_block_into(next_p, block, has_runs=False)
        current = next_p


def apply_markdown_to_cell(cell: "_Cell", md: str) -> "Paragraph":
    """Render `md` into `cell`, appending a fresh paragraph for the first block.

    Returns the first paragraph created. Subsequent blocks (list items,
    blank-line separators, …) become further paragraphs inside the same
    cell. Pre-existing cell content is preserved — markdown blocks are
    appended.
    """
    if not isinstance(md, str):
        raise TypeError(f"md must be str, got {type(md).__name__}")

    blocks = tokenize_blocks(md)
    if not blocks:
        # -- still return *a* paragraph for caller convenience (matches
        # -- _Cell.add_paragraph which returns a fresh empty paragraph).
        return cell.add_paragraph()

    first_p = cell.add_paragraph()
    _emit_block_into(first_p, blocks[0], has_runs=False)
    current = first_p
    for block in blocks[1:]:
        next_p = current.insert_paragraph_after()
        _emit_block_into(next_p, block, has_runs=False)
        current = next_p
    return first_p
