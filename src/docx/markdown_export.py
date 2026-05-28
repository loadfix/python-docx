# pyright: reportPrivateUsage=false

"""Minimal docx -> Markdown (GFM) exporter.

A read-only preview-grade renderer. Walks the document body and emits a
GitHub-Flavoured-Markdown string. Designed for handoff to PR comments,
issue trackers, static-site generators, and LLM ingestion pipelines --
*not* for round-tripping Markdown back into docx, and not a
full-fidelity converter.

Mapping:

* ``Heading 1`` .. ``Heading 6``  -> ``#`` .. ``######``
* Bold runs                       -> ``**text**``
* Italic runs                     -> ``_text_``
* Inline code (``Code``           -> ``` `text` ``` (when the run uses
  Word's built-in *Code* / ``HTMLCode`` / ``SourceCode`` character
  style, or its font name is one of the typical mono families:
  ``Consolas``, ``Courier New``, ``Source Code Pro``, ``Menlo``,
  ``Monaco``)
* Hyperlinks                      -> ``[text](url)``
* Bullet list items               -> ``- `` (with two-space indent
  per nesting level)
* Numbered list items             -> ``1. ``
* Tables                          -> GFM ``| col | col |`` with the
  separator row matching the first row's column count
* Block quotes (paragraphs        -> ``> `` prefix
  using *Quote* / *Intense Quote*
  styles)
* Inline pictures                 -> ``![alt](path)`` where ``path``
  is the relative archive path (``word/media/<name>``)
* Page breaks                     -> ``---`` on its own line
* Footnotes                       -> ``[^N]`` reference markers, with a
  ``[^N]: text`` block at the end of the document

Lossy conversions (intentional, since Markdown is a subset of Word's
expressiveness):

* All run-level fonts, sizes, and colours are dropped -- only bold /
  italic / inline-code / underline-as-bold-fallback survive.
* Paragraph alignment, indentation, and spacing collapse to plain
  paragraph breaks.
* Drawing anchors, text boxes, OMML equations, fields, and SmartArt
  are skipped entirely (not even an HTML comment, since GFM renders
  HTML comments verbatim in some viewers).
* Tables containing block-level content (nested tables, multi-paragraph
  cells) flatten to space-joined cell text -- GFM tables cannot carry
  block-level cell content.
* Images store their archive path verbatim; consumers must extract
  the bytes from the .docx zip if they want a self-contained render.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional, Tuple, Union

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.hyperlink import Hyperlink
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


_HEADING_STYLE_PREFIX = "Heading "

# -- character / paragraph style names that map to GFM constructs. Names are
# -- compared case-insensitively after stripping non-alphanumerics so that
# -- ``HTML Code``, ``html-code``, ``HTMLCode`` all match the ``htmlcode`` key.
_INLINE_CODE_STYLES = frozenset(
    {"code", "htmlcode", "sourcecode", "verbatimchar", "consoletext"}
)
_INLINE_CODE_FONTS = frozenset(
    {
        "consolas",
        "courier",
        "couriernew",
        "sourcecodepro",
        "menlo",
        "monaco",
        "dejavusansmono",
        "ubuntumono",
        "firacode",
    }
)
_QUOTE_STYLES = frozenset({"quote", "intensequote", "blockquote"})

# -- characters that need backslash-escaping when emitted as GFM body text.
# -- See <https://github.github.com/gfm/#backslash-escapes>. We keep the set
# -- conservative -- escaping ``#`` only at the start of a line, ``|`` only
# -- inside a table cell, etc., is handled at the call site. --
_GFM_INLINE_ESCAPE = "\\`*_{}[]<>"


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def document_to_markdown(document: "Document") -> str:
    """Return a GFM-Markdown rendering of `document`.

    See :meth:`docx.document.Document.to_markdown` for the public-API
    contract and a list of lossy conversions.
    """
    renderer = _MarkdownRenderer(document)
    return renderer.render()


# ---------------------------------------------------------------------------
# Renderer
# ---------------------------------------------------------------------------


class _MarkdownRenderer:
    """Walks the document body and emits Markdown fragments.

    Statefully tracks footnote references so the same footnote
    referenced multiple times shares a single ``[^N]`` block at the end.
    """

    def __init__(self, document: "Document"):
        self._document = document
        self._part = document.part
        # -- footnote_id -> (marker_index, text). marker_index starts at 1
        # -- so callers see ``[^1]`` for the first reference, regardless of
        # -- what Word numbered it internally. --
        self._footnotes: List[Tuple[int, str]] = []
        self._footnote_index_by_id: dict[int, int] = {}

    # -- public API ----------------------------------------------------------

    def render(self) -> str:
        """Render the whole document as a single Markdown string.

        The document body is walked once, then any collected footnotes
        are appended in reference order.
        """
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        blocks: List[str] = []
        # -- list_state: tuple (kind, level) for the *previous* paragraph,
        # -- or |None| when the previous block wasn't a list item. --
        prev_list_state: Optional[Tuple[str, int]] = None

        for block in self._document.iter_inner_content():
            if isinstance(block, Paragraph):
                rendered = self._render_paragraph(block)
                if rendered is None:
                    continue
                kind, level, text = rendered
                if kind == "list":
                    blocks.append(text)
                    prev_list_state = ("list", level)
                else:
                    if prev_list_state is not None:
                        # -- a blank line ends the open list block --
                        blocks.append("")
                        prev_list_state = None
                    blocks.append(text)
            elif isinstance(block, Table):
                if prev_list_state is not None:
                    blocks.append("")
                    prev_list_state = None
                blocks.append(self._render_table(block))
            else:  # pragma: no cover -- iter_inner_content yields only these
                continue

        body = "\n\n".join(b for b in blocks if b is not None)

        if self._footnotes:
            footnotes_block = "\n".join(
                f"[^{idx}]: {text}" for idx, text in self._footnotes
            )
            body = f"{body}\n\n{footnotes_block}" if body else footnotes_block

        # -- collapse runs of 3+ blank lines that the joiner above can
        # -- introduce when individual blocks contained trailing newlines. --
        while "\n\n\n\n" in body:
            body = body.replace("\n\n\n\n", "\n\n\n")
        return body.rstrip() + "\n" if body else ""

    # -- paragraph-level -----------------------------------------------------

    def _render_paragraph(
        self, paragraph: "Paragraph"
    ) -> Optional[Tuple[str, int, str]]:
        """Render `paragraph` and return ``(kind, level, text)``.

        ``kind`` is one of ``"heading"``, ``"list"``, ``"quote"``,
        ``"para"``. ``level`` carries list nesting (0 for non-lists).
        ``text`` is the rendered Markdown line(s). |None| is returned
        when the paragraph emits nothing (e.g. an empty separator).
        """
        # -- page break detection: a hard page break inserted via
        # -- ``run.add_break(WD_BREAK.PAGE)`` becomes a ``w:br w:type="page"``.
        # -- We split the paragraph into a thematic break and continue with
        # -- the rest of the paragraph rendered normally. --
        page_break_marker = _has_page_break(paragraph)

        heading_level = _heading_level_for(paragraph)
        if heading_level is not None:
            inline = self._render_inline(paragraph)
            if not inline.strip():
                return None
            text = f"{'#' * heading_level} {inline}"
            return ("heading", 0, _maybe_prefix_pagebreak(text, page_break_marker))

        # -- quote paragraphs --
        if _is_quote_style(paragraph):
            inline = self._render_inline(paragraph)
            if not inline.strip():
                return None
            quoted = "\n".join(f"> {line}" for line in inline.splitlines())
            return ("quote", 0, _maybe_prefix_pagebreak(quoted, page_break_marker))

        list_info = _list_kind_for(paragraph)
        if list_info is not None:
            kind, level = list_info
            inline = self._render_inline(paragraph)
            indent = "  " * level
            marker = "- " if kind == "ul" else "1. "
            text = f"{indent}{marker}{inline}"
            return ("list", level, _maybe_prefix_pagebreak(text, page_break_marker))

        inline = self._render_inline(paragraph)
        if not inline.strip() and not page_break_marker:
            # -- empty paragraphs collapse into the join-by-blank-line --
            return None
        if page_break_marker and not inline.strip():
            return ("para", 0, "---")
        return ("para", 0, _maybe_prefix_pagebreak(inline, page_break_marker))

    # -- inline-level --------------------------------------------------------

    def _render_inline(self, paragraph: "Paragraph") -> str:
        """Render the inline children of `paragraph` as a Markdown string."""
        from docx.text.hyperlink import Hyperlink
        from docx.text.run import Run

        parts: List[str] = []
        for child in paragraph.iter_inner_content():
            if isinstance(child, Run):
                parts.append(self._render_run(child))
            elif isinstance(child, Hyperlink):
                parts.append(self._render_hyperlink(child))
            else:  # pragma: no cover -- only Run / Hyperlink today
                continue

        # -- pick up images (which live inside ``w:r/w:drawing`` and so are
        # -- not surfaced by iter_inner_content above). We append them at
        # -- the end of the paragraph rather than weaving them in -- the
        # -- run-level interleaving would require descending raw XML and
        # -- duplicating the run-text we already collected. --
        for img in self._render_images(paragraph):
            parts.append(img)

        return "".join(parts)

    def _render_run(self, run: "Run") -> str:
        # -- emit footnote / endnote reference markers when the run carries
        # -- one. The marker replaces any text the run might otherwise have
        # -- since Word almost always puts the reference in its own run. --
        marker = self._footnote_marker_for(run)
        if marker:
            return marker

        text = run.text or ""
        if not text:
            return ""

        # -- escape Markdown-significant inline characters. Tabs and
        # -- newlines stay as-is so paragraph structure is preserved. --
        escaped = _escape_inline(text)

        # -- inline code beats other formatting: a code-styled run renders
        # -- as backtick-wrapped literal text, untouched by escapes since
        # -- code spans don't honour them. --
        if _is_inline_code(run):
            return _wrap_code(text)

        bold = bool(run.bold)
        italic = bool(run.italic)

        if bold and italic:
            return f"**_{escaped}_**"
        if bold:
            return f"**{escaped}**"
        if italic:
            return f"_{escaped}_"
        return escaped

    def _render_hyperlink(self, hyperlink: "Hyperlink") -> str:
        from docx.text.run import Run

        # -- collect run text honouring formatting --
        run_md_parts: List[str] = []
        for r_elm in hyperlink._hyperlink.r_lst:
            run_md_parts.append(self._render_run(Run(r_elm, hyperlink._parent)))
        text = "".join(run_md_parts).strip() or hyperlink.text or hyperlink.url
        url = hyperlink.url or ""
        if not url:
            anchor = hyperlink.fragment or ""
            url = f"#{anchor}" if anchor else ""
        # -- escape ``)`` in the URL since GFM uses it as the delimiter --
        url_escaped = url.replace(")", "%29").replace("(", "%28")
        if not url_escaped:
            return text
        return f"[{text}]({url_escaped})"

    def _render_images(self, paragraph: "Paragraph") -> List[str]:
        """Return a list of ``![alt](path)`` strings for every inline picture."""
        out: List[str] = []
        p_elm = paragraph._p
        for drawing in p_elm.iter(qn("w:drawing")):
            inline = drawing.find(qn("wp:inline"))
            if inline is None:
                # -- anchored picture: still try to extract, but flag in alt
                blip_iter = drawing.xpath(".//a:blip")
                if not blip_iter:
                    continue
                blip = blip_iter[0]
                anchored = True
            else:
                anchored = False
                blips = drawing.xpath(".//a:blip")
                if not blips:
                    continue
                blip = blips[0]
            rId = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
            if rId is None:
                continue
            try:
                image_part = self._part.related_parts[rId]
            except KeyError:
                continue
            partname = getattr(image_part, "partname", "")
            if not partname:
                continue
            # -- partname is like ``/word/media/image1.png``; drop the
            # -- leading slash so it reads as a zip-relative path. --
            archive_path = str(partname).lstrip("/")
            alt = (
                drawing.xpath("string(.//wp:docPr/@descr)")
                or drawing.xpath("string(.//wp:docPr/@title)")
                or ""
            )
            if anchored and not alt:
                alt = "anchored image"
            out.append(f"![{alt}]({archive_path})")
        return out

    # -- footnotes -----------------------------------------------------------

    def _footnote_marker_for(self, run: "Run") -> Optional[str]:
        """Return ``"[^N]"`` if `run` contains a footnote/endnote reference."""
        ref = run._r.find(qn("w:footnoteReference"))
        kind = "footnote"
        if ref is None:
            ref = run._r.find(qn("w:endnoteReference"))
            kind = "endnote"
        if ref is None:
            return None
        try:
            fn_id = int(ref.get(qn("w:id")) or "0")
        except (TypeError, ValueError):
            return None
        if fn_id in self._footnote_index_by_id:
            idx = self._footnote_index_by_id[fn_id]
        else:
            idx = len(self._footnotes) + 1
            self._footnote_index_by_id[fn_id] = idx
            text = self._lookup_footnote_text(fn_id, kind)
            self._footnotes.append((idx, text))
        return f"[^{idx}]"

    def _lookup_footnote_text(self, fn_id: int, kind: str) -> str:
        try:
            if kind == "footnote":
                store = self._document.footnotes
            else:
                store = self._document.endnotes  # pyright: ignore[reportAttributeAccessIssue]
        except Exception:  # pragma: no cover -- defensive
            return ""
        for fn in store:
            try:
                if fn.footnote_id == fn_id:  # pyright: ignore[reportAttributeAccessIssue]
                    return _flatten_footnote_text(fn)
            except AttributeError:
                # -- endnotes use ``endnote_id`` --
                if getattr(fn, "endnote_id", None) == fn_id:
                    return _flatten_footnote_text(fn)
        return ""

    # -- tables --------------------------------------------------------------

    def _render_table(self, table) -> str:
        """Render `table` as a GFM table.

        Cell content is flattened to a single line: paragraph breaks
        within a cell collapse to a single space, and inline formatting
        (bold/italic/code/links) is preserved. Returns the empty string
        for tables with zero rows.
        """
        if not table.rows:
            return ""
        rendered_rows: List[List[str]] = []
        max_cols = 0
        for row in table.rows:
            cells: List[str] = []
            for cell in row.cells:
                cells.append(self._render_cell_inline(cell))
            max_cols = max(max_cols, len(cells))
            rendered_rows.append(cells)

        # -- pad every row to the widest, in case of merged or ragged rows --
        for cells in rendered_rows:
            while len(cells) < max_cols:
                cells.append("")

        lines: List[str] = []
        # -- header row --
        header = rendered_rows[0]
        lines.append("| " + " | ".join(header) + " |")
        lines.append("| " + " | ".join(["---"] * max_cols) + " |")
        for cells in rendered_rows[1:]:
            lines.append("| " + " | ".join(cells) + " |")
        return "\n".join(lines)

    def _render_cell_inline(self, cell) -> str:
        """Flatten a cell's paragraphs into a single GFM-table-safe line."""
        bits: List[str] = []
        for p in cell.paragraphs:
            inline = self._render_inline(p)
            if inline.strip():
                bits.append(inline.strip())
        flat = " ".join(bits)
        # -- escape pipe so it doesn't terminate the cell, and collapse any
        # -- newlines that survived (e.g. from a soft line break) into a
        # -- single space. --
        return flat.replace("|", "\\|").replace("\n", " ").replace("\r", " ")


# ---------------------------------------------------------------------------
# Standalone helpers
# ---------------------------------------------------------------------------


def _heading_level_for(paragraph: "Paragraph") -> Optional[int]:
    """Return the 1..6 heading level for `paragraph`, or |None|.

    Heading 7+ collapses to 6 to stay within Markdown's vocabulary.
    """
    style = paragraph.style
    if style is None:
        return None
    name = style.name or ""
    if not name.startswith(_HEADING_STYLE_PREFIX):
        return None
    suffix = name[len(_HEADING_STYLE_PREFIX) :].strip()
    if not suffix.isdigit():
        return None
    level = int(suffix)
    if level < 1:
        return None
    return min(level, 6)


def _is_quote_style(paragraph: "Paragraph") -> bool:
    style = paragraph.style
    if style is None:
        return False
    return _normalise_style_name(style.name or "") in _QUOTE_STYLES


def _list_kind_for(
    paragraph: "Paragraph",
) -> Optional[Tuple[str, int]]:
    """Return ``("ul" | "ol", level)`` if `paragraph` is a list item, else |None|."""
    p_elm = paragraph._p
    pPr = p_elm.find(qn("w:pPr"))
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    numId_elm = numPr.find(qn("w:numId"))
    if numId_elm is None:
        return None
    try:
        num_id = int(numId_elm.get(qn("w:val")) or "0")
    except (TypeError, ValueError):
        num_id = 0
    if num_id == 0:
        return None

    ilvl_elm = numPr.find(qn("w:ilvl"))
    try:
        level = int(ilvl_elm.get(qn("w:val")) or "0") if ilvl_elm is not None else 0
    except (TypeError, ValueError):
        level = 0

    fmt = _resolve_num_format(paragraph, num_id, level)
    kind = "ul" if fmt == "bullet" else "ol"
    return (kind, level)


def _resolve_num_format(
    paragraph: "Paragraph", num_id: int, level: int
) -> Optional[str]:
    numbering_part = getattr(paragraph.part, "numbering_part", None)
    if numbering_part is None:
        return None
    numbering_elm = getattr(numbering_part, "numbering_element", None)
    if numbering_elm is None:
        return None

    num_elms = numbering_elm.xpath(f'./w:num[@w:numId="{num_id}"]')
    if not num_elms:
        return None
    abstract_id_elms = num_elms[0].xpath("./w:abstractNumId/@w:val")
    if not abstract_id_elms:
        return None
    abstract_id = abstract_id_elms[0]

    fmt_vals = numbering_elm.xpath(
        f'./w:abstractNum[@w:abstractNumId="{abstract_id}"]/'
        f'w:lvl[@w:ilvl="{level}"]/w:numFmt/@w:val'
    )
    if fmt_vals:
        return fmt_vals[0]
    return None


def _has_page_break(paragraph: "Paragraph") -> bool:
    """Return |True| when `paragraph` contains a hard ``w:br w:type='page'``."""
    p_elm = paragraph._p
    for br in p_elm.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _maybe_prefix_pagebreak(text: str, has_break: bool) -> str:
    if not has_break:
        return text
    return f"---\n\n{text}"


def _is_inline_code(run: "Run") -> bool:
    """Return |True| when `run` should render as inline code."""
    style = getattr(run, "style", None)
    if style is not None:
        name = getattr(style, "name", "") or ""
        if _normalise_style_name(name) in _INLINE_CODE_STYLES:
            return True
    font = getattr(run, "font", None)
    if font is not None:
        font_name = (font.name or "").lower().replace(" ", "").replace("-", "")
        if font_name in _INLINE_CODE_FONTS:
            return True
    return False


def _normalise_style_name(name: str) -> str:
    return "".join(ch for ch in name.lower() if ch.isalnum())


def _wrap_code(text: str) -> str:
    """Wrap `text` in backticks, escaping any embedded backticks per GFM rules.

    The GFM rule is: pick a backtick run length not present in the text.
    For the common case (no backticks in the source) a single backtick
    works; otherwise we widen the fence and pad with a leading/trailing
    space when the text starts or ends with a backtick.
    """
    if "`" not in text:
        return f"`{text}`"
    # -- pick the shortest run of backticks not appearing in `text` --
    longest = 0
    cur = 0
    for ch in text:
        if ch == "`":
            cur += 1
            longest = max(longest, cur)
        else:
            cur = 0
    fence = "`" * (longest + 1)
    pad_left = " " if text.startswith("`") else ""
    pad_right = " " if text.endswith("`") else ""
    return f"{fence}{pad_left}{text}{pad_right}{fence}"


def _escape_inline(text: str) -> str:
    """Backslash-escape Markdown-significant characters in `text`.

    Conservative -- escapes the backslash itself, plus the eight
    characters that GFM uses as inline structural delimiters.
    """
    out_chars: List[str] = []
    for ch in text:
        if ch in _GFM_INLINE_ESCAPE:
            out_chars.append("\\")
        out_chars.append(ch)
    return "".join(out_chars)


def _flatten_footnote_text(footnote) -> str:
    """Return a single-line representation of a footnote's text content."""
    raw = getattr(footnote, "text", "") or ""
    # -- collapse intra-footnote paragraph breaks to a single space; the
    # -- ``[^N]: `` prefix demands a one-line value to keep the footnote
    # -- block visually compact and unambiguous on parse. --
    return " ".join(line.strip() for line in raw.splitlines() if line.strip())
