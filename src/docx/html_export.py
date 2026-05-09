# pyright: reportPrivateUsage=false

"""Minimal docx → HTML exporter.

A read-only preview-grade renderer. Walks the document body and emits an
HTML5 document string. Designed for quick web previews and handoff to
other web platforms — *not* for round-tripping HTML back into docx, and
not a full-fidelity converter.

The mapping follows the scope declared in round 10 / ticket R10-11:

* ``w:p``           → ``<p>`` with inline style for alignment + margins
  (heading paragraph styles are promoted to ``<h1>``–``<h6>``).
* ``w:r``           → ``<span>`` (promoted to ``<strong>`` / ``<em>`` /
  ``<u>`` when the *only* direct formatting applied is bold / italic /
  underline, respectively).
* ``w:hyperlink``   → ``<a href="…">``.
* ``w:tbl``         → ``<table>`` with ``<tr>`` / ``<td>`` and inline-CSS
  borders.
* ``w:drawing``     (inline picture) → ``<img src="…">``. When
  ``embed_images=True``, the image bytes are emitted as a
  ``data:`` URL; when ``embed_images=False``, a ``cid:{rId}`` placeholder
  is used so downstream MIME assemblers can attach the parts.

Unsupported constructs (fields, shapes, text boxes, anchored pictures,
equations) yield an HTML comment like
``<!-- unsupported: w:fldSimple -->`` and the walker continues.

Text is escaped with :func:`html.escape` at every text node to guard
against XSS from document content.

.. versionadded:: 2026.05.10
"""

from __future__ import annotations

import base64
import html
from typing import TYPE_CHECKING, Iterable, cast

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.text.paragraph import Paragraph


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def document_to_html(
    document: "Document",
    include_styles: bool = True,
    embed_images: bool = True,
) -> str:
    """Return a complete HTML5 document rendering of `document`.

    See :meth:`docx.document.Document.to_html` for the public-API contract.
    """
    renderer = _HtmlRenderer(document, embed_images=embed_images)
    body_html = renderer.render_body()

    title = ""
    try:
        title_val = document.core_properties.title
        if isinstance(title_val, str):
            title = html.escape(title_val)
    except Exception:  # pragma: no cover — defensive; core props missing
        title = ""

    head_parts: list[str] = [
        "<!DOCTYPE html>",
        '<html lang="en">',
        "<head>",
        '<meta charset="utf-8">',
        f"<title>{title}</title>" if title else "<title>Document</title>",
    ]

    if include_styles:
        css = renderer.render_style_block()
        if css:
            head_parts.append("<style>\n" + css + "\n</style>")

    head_parts.append("</head>")
    head_parts.append("<body>")

    return (
        "\n".join(head_parts)
        + "\n"
        + body_html
        + "\n</body>\n</html>"
    )


# ---------------------------------------------------------------------------
# Renderer
# ---------------------------------------------------------------------------


_HEADING_STYLE_PREFIX = "Heading "


class _HtmlRenderer:
    """Walks the document body and emits HTML fragments."""

    def __init__(self, document: "Document", embed_images: bool = True):
        self._document = document
        self._part = document.part
        self._embed_images = embed_images
        # -- list-state for flushing consecutive <li> elements into a single
        # -- <ol>/<ul> wrapper. None when no list is open. --
        self._open_list: str | None = None

    # -- public API ----------------------------------------------------------

    def render_body(self) -> str:
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        fragments: list[str] = []
        for block in self._document.iter_inner_content():
            if isinstance(block, Paragraph):
                frag = self._render_paragraph(block)
            elif isinstance(block, Table):
                fragments.extend(self._flush_list())
                frag = self._render_table(block)
            else:  # pragma: no cover — iter_inner_content only yields these today
                frag = f"<!-- unsupported: {type(block).__name__} -->"
            if frag:
                fragments.append(frag)
        fragments.extend(self._flush_list())
        return "\n".join(fragments)

    def render_style_block(self) -> str:
        """Return a CSS string derived from the document's style definitions.

        Emits one rule per paragraph or character style: the rule's
        selector is ``.docx-style-<safe-id>`` so callers can tag runs
        with the matching class. Only ``font-family``, ``font-size``,
        ``color``, and a rough margin approximation are translated —
        everything else is intentionally dropped.
        """
        try:
            styles = list(self._document.styles)
        except Exception:  # pragma: no cover — defensive
            return ""

        rules: list[str] = []
        for style in styles:
            rule = _style_to_css_rule(style)
            if rule:
                rules.append(rule)
        return "\n".join(rules)

    # -- paragraph-level -----------------------------------------------------

    def _render_paragraph(self, paragraph: "Paragraph") -> str:
        from docx.text.paragraph import Paragraph as _ParagraphCls

        assert isinstance(paragraph, _ParagraphCls)

        # -- headings promote to <hN>, stepping out of any open list --
        heading_level = _heading_level_for(paragraph)
        if heading_level is not None:
            flushed = "\n".join(self._flush_list())
            inner = self._render_inline_children(paragraph)
            tag = f"h{heading_level}"
            out = f"<{tag}>{inner}</{tag}>"
            return f"{flushed}\n{out}" if flushed else out

        # -- numbered / bulleted list items --
        list_kind = _list_kind_for(paragraph)
        if list_kind is not None:
            prefix: list[str] = []
            if self._open_list != list_kind:
                prefix.extend(self._flush_list())
                prefix.append(f"<{list_kind}>")
                self._open_list = list_kind
            inner = self._render_inline_children(paragraph)
            return "\n".join(prefix + [f"<li>{inner}</li>"])

        # -- plain paragraph --
        flushed = self._flush_list()
        style_attr = _paragraph_style_attr(paragraph)
        inner = self._render_inline_children(paragraph)
        open_tag = "<p" + (f' style="{style_attr}"' if style_attr else "") + ">"
        out = f"{open_tag}{inner}</p>"
        return "\n".join(flushed + [out]) if flushed else out

    def _flush_list(self) -> list[str]:
        """Close any open ``<ol>`` / ``<ul>``, returning the closing tag(s)."""
        if self._open_list is None:
            return []
        closer = f"</{self._open_list}>"
        self._open_list = None
        return [closer]

    # -- inline-level --------------------------------------------------------

    def _render_inline_children(self, paragraph: "Paragraph") -> str:
        """Render every inline child (``w:r`` and ``w:hyperlink``) of `paragraph`."""
        from docx.text.hyperlink import Hyperlink
        from docx.text.run import Run

        parts: list[str] = []
        for child in paragraph.iter_inner_content():
            if isinstance(child, Run):
                frag = self._render_run(child)
            elif isinstance(child, Hyperlink):
                frag = self._render_hyperlink(child)
            else:  # pragma: no cover — only Run / Hyperlink are yielded today
                frag = f"<!-- unsupported: {type(child).__name__} -->"
            if frag:
                parts.append(frag)

        # -- inline unsupported-element comments for any drawings / fields /
        # -- other children that don't surface via `iter_inner_content` --
        parts.extend(self._render_raw_children_markers(paragraph))
        return "".join(parts)

    def _render_raw_children_markers(self, paragraph: "Paragraph") -> list[str]:
        """Emit images and ``<!-- unsupported -->`` comments from raw ``w:r`` descendants.

        Inline images live inside ``w:r/w:drawing`` and so are not surfaced
        by :meth:`Paragraph.iter_inner_content`, which only yields ``Run``
        and ``Hyperlink`` proxies. Walk the raw XML tree to pick up
        drawings (inline pictures only) and flag anything else
        (anchored pictures, fields, shapes, OMML, etc.) as unsupported.
        """
        markers: list[str] = []
        p_elm = paragraph._p  # pyright: ignore[reportPrivateUsage]

        # -- inline pictures --
        for drawing in p_elm.iter(qn("w:drawing")):
            inline = drawing.find(qn("wp:inline"))
            if inline is None:
                markers.append("<!-- unsupported: w:drawing (anchor) -->")
                continue
            img = self._render_inline_image(drawing)
            if img is not None:
                markers.append(img)
            else:
                markers.append("<!-- unsupported: w:drawing -->")

        # -- other paragraph-level unsupported constructs --
        unsupported_tags = {
            qn("w:fldSimple"): "w:fldSimple",
            qn("m:oMath"): "m:oMath",
            qn("m:oMathPara"): "m:oMathPara",
            qn("w:object"): "w:object",
            qn("w:pict"): "w:pict",
        }
        for el in p_elm.iter():
            marker = unsupported_tags.get(el.tag)
            if marker is not None:
                markers.append(f"<!-- unsupported: {marker} -->")
        return markers

    def _render_run(self, run) -> str:
        text = run.text or ""
        # -- preserve tabs and break runs of ``\n`` as ``<br>`` --
        escaped = html.escape(text).replace("\n", "<br>")

        if not escaped:
            return ""

        font = run.font
        bold = bool(font.bold)
        italic = bool(font.italic)
        underline = bool(font.underline)

        # -- Tag promotion: when exactly one of bold/italic/underline is
        # -- applied and nothing else is needed, wrap in a semantic tag. --
        style_attr = _run_style_attr(run)
        flags = (bold, italic, underline)
        if not style_attr and sum(flags) == 1:
            if bold:
                return f"<strong>{escaped}</strong>"
            if italic:
                return f"<em>{escaped}</em>"
            if underline:
                return f"<u>{escaped}</u>"

        if not (bold or italic or underline or style_attr):
            # -- no formatting at all: emit raw text, no wrapping span --
            return escaped

        inner = escaped
        if underline:
            inner = f"<u>{inner}</u>"
        if italic:
            inner = f"<em>{inner}</em>"
        if bold:
            inner = f"<strong>{inner}</strong>"

        if style_attr:
            return f'<span style="{style_attr}">{inner}</span>'
        return inner

    def _render_hyperlink(self, hyperlink) -> str:
        runs_html = "".join(self._render_run(r) for r in hyperlink.runs)
        url = hyperlink.url or ""
        if not url:
            # -- internal / anchor-only links: fall back to a fragment --
            anchor = hyperlink.fragment
            url = f"#{anchor}" if anchor else ""
        escaped_url = html.escape(url, quote=True)
        if url:
            return f'<a href="{escaped_url}">{runs_html}</a>'
        return f"<a>{runs_html}</a>"

    def _render_inline_image(self, drawing) -> str | None:
        """Return an ``<img>`` tag for an inline picture, or |None|."""
        blip = drawing.find(
            qn("wp:inline") + "/" + qn("a:graphic") + "/" + qn("a:graphicData")
            + "/" + qn("pic:pic") + "/" + qn("pic:blipFill") + "/" + qn("a:blip")
        )
        if blip is None:
            # -- blip may live deeper / under a different wrapper; xpath it --
            blips = drawing.xpath(".//a:blip")
            if not blips:
                return None
            blip = blips[0]
        rId = blip.get(qn("r:embed"))
        if rId is None:
            rId = blip.get(qn("r:link"))
        if rId is None:
            return None

        alt = html.escape(
            drawing.xpath("string(.//wp:docPr/@descr)")
            or drawing.xpath("string(.//wp:docPr/@title)")
            or "",
            quote=True,
        )

        if not self._embed_images:
            return f'<img src="cid:{html.escape(rId, quote=True)}" alt="{alt}">'

        try:
            image_part = self._part.related_parts[rId]
        except KeyError:
            return f"<!-- unsupported: w:drawing (missing rId {rId}) -->"
        blob = getattr(image_part, "blob", None)
        content_type = getattr(image_part, "content_type", "image/png")
        if not blob:
            return f"<!-- unsupported: w:drawing (no blob for rId {rId}) -->"
        encoded = base64.b64encode(blob).decode("ascii")
        return (
            f'<img src="data:{content_type};base64,{encoded}" alt="{alt}">'
        )

    # -- tables --------------------------------------------------------------

    def _render_table(self, table) -> str:
        border_css = "border-collapse:collapse;"
        rows_html: list[str] = []
        for row in table.rows:
            cells_html: list[str] = []
            for cell in row.cells:
                cell_inner = self._render_cell_content(cell)
                cells_html.append(
                    '<td style="border:1px solid #000;padding:4px;">'
                    f"{cell_inner}</td>"
                )
            rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
        return (
            f'<table style="{border_css}">\n' + "\n".join(rows_html) + "\n</table>"
        )

    def _render_cell_content(self, cell) -> str:
        # -- reuse paragraph rendering; tables inside cells are not descended
        # -- into for this minimal exporter.
        inner_renderer = _HtmlRenderer(self._document, embed_images=self._embed_images)
        fragments: list[str] = []
        for p in cell.paragraphs:
            fragments.append(inner_renderer._render_paragraph(p))
        fragments.extend(inner_renderer._flush_list())
        return "\n".join(f for f in fragments if f)


# ---------------------------------------------------------------------------
# Heading / list detection helpers
# ---------------------------------------------------------------------------


def _heading_level_for(paragraph: "Paragraph") -> int | None:
    """Return the ``1..6`` heading level for `paragraph`, or |None|.

    Detection is purely by paragraph style name: ``"Heading N"`` for
    ``N`` in 1..6 maps to ``hN``. Any level above 6 maps to ``h6`` to
    stay within the HTML vocabulary.
    """
    style = paragraph.style
    if style is None:
        return None
    name = style.name or ""
    if not name.startswith(_HEADING_STYLE_PREFIX):
        return None
    suffix = name[len(_HEADING_STYLE_PREFIX):].strip()
    if not suffix.isdigit():
        return None
    level = int(suffix)
    if level < 1:
        return None
    return min(level, 6)


def _list_kind_for(paragraph: "Paragraph") -> str | None:
    """Return ``"ol"`` or ``"ul"`` if `paragraph` is a list item, else |None|.

    Inspects ``w:numPr/w:numId`` and resolves the level's ``w:numFmt``.
    Falls back to ``<ul>`` when the format can't be resolved but the
    paragraph is otherwise marked as numbered.
    """
    p_elm = paragraph._p  # pyright: ignore[reportPrivateUsage]
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
        # -- 0 means "remove list formatting"; treat as not a list --
        return None

    ilvl_elm = numPr.find(qn("w:ilvl"))
    try:
        level = int(ilvl_elm.get(qn("w:val")) or "0") if ilvl_elm is not None else 0
    except (TypeError, ValueError):
        level = 0

    fmt = _resolve_num_format(paragraph, num_id, level)
    if fmt == "bullet":
        return "ul"
    return "ol"


def _resolve_num_format(
    paragraph: "Paragraph", num_id: int, level: int
) -> str | None:
    """Return the ``w:numFmt`` value for `num_id` / `level`, or |None|."""
    numbering_part = getattr(paragraph.part, "numbering_part", None)
    if numbering_part is None:
        return None
    numbering_elm = getattr(numbering_part, "numbering_element", None)
    if numbering_elm is None:
        return None

    # -- num → abstractNumId --
    num_elms = numbering_elm.xpath(f'./w:num[@w:numId="{num_id}"]')
    if not num_elms:
        return None
    abstract_id_elms = num_elms[0].xpath("./w:abstractNumId/@w:val")
    if not abstract_id_elms:
        return None
    abstract_id = abstract_id_elms[0]

    # -- abstractNum → lvl[@ilvl=N]/numFmt/@val --
    fmt_vals = numbering_elm.xpath(
        f'./w:abstractNum[@w:abstractNumId="{abstract_id}"]/'
        f'w:lvl[@w:ilvl="{level}"]/w:numFmt/@w:val'
    )
    if fmt_vals:
        return fmt_vals[0]
    return None


# ---------------------------------------------------------------------------
# Inline-style helpers
# ---------------------------------------------------------------------------


def _paragraph_style_attr(paragraph: "Paragraph") -> str:
    """Return the ``style=…`` contents for a ``<p>`` element (no quotes)."""
    pieces: list[str] = []

    alignment = paragraph.alignment
    if alignment is not None:
        align_name = getattr(alignment, "name", "") or ""
        # -- map enum names to CSS text-align values --
        css_align = {
            "LEFT": "left",
            "CENTER": "center",
            "RIGHT": "right",
            "JUSTIFY": "justify",
            "START": "left",
            "END": "right",
            "DISTRIBUTE": "justify",
        }.get(align_name)
        if css_align:
            pieces.append(f"text-align:{css_align}")

    fmt = paragraph.paragraph_format
    left = fmt.left_indent
    if left is not None:
        try:
            pieces.append(f"margin-left:{int(left.pt)}pt")
        except Exception:
            pass
    right = fmt.right_indent
    if right is not None:
        try:
            pieces.append(f"margin-right:{int(right.pt)}pt")
        except Exception:
            pass

    return ";".join(pieces)


def _run_style_attr(run) -> str:
    """Return the ``style=…`` contents for a run wrapper (no quotes)."""
    pieces: list[str] = []
    font = run.font
    name = font.name
    if name:
        safe = html.escape(name, quote=True)
        pieces.append(f"font-family:'{safe}'")
    size = font.size
    if size is not None:
        try:
            pieces.append(f"font-size:{int(size.pt)}pt")
        except Exception:
            pass
    try:
        rgb = font.color.rgb
        if rgb is not None:
            pieces.append(f"color:#{str(rgb)}")
    except Exception:
        pass
    return ";".join(pieces)


def _style_to_css_rule(style) -> str:
    """Return a single CSS rule for the given style, or ``""`` to skip it.

    Only styles with a name produce rules, and only ``font-family``,
    ``font-size``, ``color``, and a ``margin`` approximation are emitted.
    """
    name = getattr(style, "name", None)
    if not name:
        return ""
    font = getattr(style, "font", None)
    pieces: list[str] = []
    if font is not None:
        fname = getattr(font, "name", None)
        if fname:
            pieces.append(f"font-family:'{html.escape(fname, quote=True)}'")
        try:
            fsize = font.size
            if fsize is not None:
                pieces.append(f"font-size:{int(fsize.pt)}pt")
        except Exception:
            pass
        try:
            rgb = font.color.rgb
            if rgb is not None:
                pieces.append(f"color:#{str(rgb)}")
        except Exception:
            pass

    # -- paragraph margins: only available on ParagraphStyle.paragraph_format --
    paragraph_format = getattr(style, "paragraph_format", None)
    if paragraph_format is not None:
        try:
            left = paragraph_format.left_indent
            if left is not None:
                pieces.append(f"margin-left:{int(left.pt)}pt")
        except Exception:
            pass

    if not pieces:
        return ""
    selector = _style_selector(name)
    return f"{selector} {{ {'; '.join(pieces)} }}"


def _style_selector(name: str) -> str:
    """Return a CSS class selector derived from a style name."""
    safe = "".join(ch if ch.isalnum() else "-" for ch in name).strip("-").lower()
    return f".docx-style-{safe or 'unnamed'}"
