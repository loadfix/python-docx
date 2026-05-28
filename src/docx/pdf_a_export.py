# pyright: reportPrivateUsage=false
# pyright: reportMissingImports=false
# pyright: reportMissingTypeStubs=false
# pyright: reportUnknownMemberType=false
# pyright: reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportUnknownParameterType=false
# pyright: reportGeneralTypeIssues=false

"""Minimal docx -> PDF/A archival exporter.

A best-effort renderer that walks the document body and emits a PDF
that *aims* at the PDF/A (ISO 19005) archival flavour. The library
spec mandates a closed, self-describing rendition: all fonts embedded,
no JavaScript, no external references, with an XMP metadata packet
declaring conformance.

This implementation lives one rung below "spec-strict": it produces a
PDF that opens in any PDF reader, embeds its declared font
(Helvetica via the bundled ReportLab Type-1 face — see Gaps below),
and stamps the catalog's ``Metadata`` stream with a PDF/A XMP packet
naming the requested ``part`` (1 / 2 / 3) and ``conformance`` (A / B).

The rendering covers:

* Paragraphs (with ``Heading 1`` .. ``Heading 6`` styles promoted to
  larger font sizes that step from 22pt down to 11pt).
* Inline runs (``bold``, ``italic``, ``underline`` survive; font name,
  colour, and size collapse to the renderer defaults — see Gaps).
* Tables (a basic grid; cell text is space-joined; no nested-block
  cell content).
* Inline images (``w:drawing`` with a resolvable ``r:embed`` rId; the
  bytes are read from the related image part and placed at the
  paragraph's flow position).
* Hard page breaks (``run.add_break(WD_BREAK.PAGE)`` -> reportlab
  ``Spacer`` + ``PageBreak``).
* Bullet / numbered lists (rendered as text with a leading marker;
  GFM-style indentation by nesting level).

Gaps (documented as future work; see ``FEATURES.md``):

* **Font embedding fidelity.** The exporter uses ReportLab's stock
  Helvetica family. Helvetica is one of the 14 PDF "core" fonts and
  is technically *not* a PDF/A-compliant choice — true PDF/A demands
  a fully-embedded font program. A follow-up pass should ship a
  liberally-licensed TrueType face (DejaVu Sans / Liberation Sans)
  bundled or auto-discovered, with subset embedding via
  ReportLab's ``TTFont`` + ``UnicodeCIDFont`` machinery.
* **Colour space.** PDF/A requires a declared output intent; this
  exporter omits the ``OutputIntents`` array. Most validators flag
  this. A follow-up should embed a sRGB ICC profile.
* **Footnotes / endnotes / fields / equations / drawings / change
  tracking** are skipped silently (no on-page placeholder).
* **Section breaks, headers, footers, page numbers** collapse — every
  page uses the default A4 / Letter layout from the requested page
  size.
* **Hyperlinks** lose their target URL and render as plain text.
* **Validation.** This is *best-effort*; running output through a
  PDF/A validator (veraPDF, Acrobat Pro Preflight) will surface
  conformance errors. The module name carries the "_export" suffix
  rather than "_pdfa" or "_archival" to make the best-effort scope
  explicit.

Public entry point:

* :func:`document_to_pdf_a` — module-level callable; takes a
  :class:`~docx.document.Document` plus a destination path / file-like
  object plus a PDF/A level string.

Usage::

    from docx import Document

    doc = Document("report.docx")
    doc.save_as_pdf_a("report.pdf", level="3a")

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from io import BytesIO
from typing import IO, TYPE_CHECKING, List, Optional, Tuple, Union

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table
    from docx.table import _Cell as TableCell
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


# ---------------------------------------------------------------------------
# Public surface
# ---------------------------------------------------------------------------

# Accepted PDF/A conformance level strings. Each maps to a
# ``(part, conformance)`` tuple where ``part`` is 1/2/3 and
# ``conformance`` is "A" (accessible / tagged) or "B" (basic / visual
# fidelity only). PDF/A-3 additionally permits attached files; we
# don't author any.
_LEVELS: dict = {
    "1a": (1, "A"),
    "1b": (1, "B"),
    "2a": (2, "A"),
    "2b": (2, "B"),
    "3a": (3, "A"),
    "3b": (3, "B"),
}

# Heading-style -> point size. ``Heading 1`` is the largest; subsequent
# levels step down 2pt each, capping at 11pt for ``Heading 7+``. The
# floor matches the body-text default so deep headings don't render
# *smaller* than body text.
_HEADING_SIZES = {
    1: 22.0,
    2: 18.0,
    3: 16.0,
    4: 14.0,
    5: 13.0,
    6: 12.0,
}

_BODY_FONT_SIZE = 11.0
_LEADING_MULTIPLIER = 1.2

# The default PDF/A producer string, embedded into the XMP packet.
_PRODUCER = "python-docx (loadfix fork) PDF/A archival exporter"


def document_to_pdf_a(
    document: "Document",
    path_or_stream: Union[str, IO[bytes]],
    level: str = "3a",
) -> None:
    """Render `document` to a PDF/A file at `path_or_stream`.

    `level` is one of ``"1a"``, ``"1b"``, ``"2a"``, ``"2b"``, ``"3a"``
    (default) or ``"3b"``. Anything else raises :class:`ValueError`.

    Raises :class:`ImportError` when ``reportlab`` is not importable —
    the rendering backend is an opt-in dependency. Install via the
    ``[pdfa]`` extra::

        pip install 'python-docx[pdfa]'

    See :meth:`docx.document.Document.save_as_pdf_a` for the public
    contract; this function is exposed at module level so callers
    can reuse it without instantiating ``Document.save_as_pdf_a``.
    """
    if level not in _LEVELS:
        raise ValueError(
            f"level must be one of {sorted(_LEVELS.keys())}; got {level!r}"
        )
    part, conformance = _LEVELS[level]

    try:
        # -- imported here so the module is import-safe even when
        # -- reportlab is not installed; only the entry point requires it.
        from reportlab.lib.pagesizes import LETTER  # type: ignore[import-not-found]
        from reportlab.pdfbase import pdfdoc  # type: ignore[import-not-found]
        from reportlab.pdfgen import canvas as _canvas_mod  # type: ignore[import-not-found]
    except ImportError as exc:  # pragma: no cover -- exercised when reportlab missing
        raise ImportError(
            "reportlab is required for save_as_pdf_a(); install it via "
            "the optional extra: pip install 'python-docx[pdfa]'"
        ) from exc

    renderer = _PdfARenderer(
        document,
        canvas_mod=_canvas_mod,
        pdfdoc_mod=pdfdoc,
        page_size=LETTER,
        part=part,
        conformance=conformance,
    )
    pdf_bytes = renderer.render()

    if hasattr(path_or_stream, "write"):
        path_or_stream.write(pdf_bytes)  # type: ignore[union-attr]
    else:
        with open(path_or_stream, "wb") as fh:  # type: ignore[arg-type]
            fh.write(pdf_bytes)


# ---------------------------------------------------------------------------
# Renderer
# ---------------------------------------------------------------------------


_HEADING_STYLE_PREFIX = "Heading "


class _PdfARenderer:
    """Walks the document body and emits a PDF/A byte string.

    The renderer is single-shot; reuse a fresh instance per output.
    Page layout uses a fixed left/right margin of 72pt (1in) and a
    1in top/bottom margin. Content that overflows the bottom margin
    triggers an automatic page break.
    """

    def __init__(
        self,
        document: "Document",
        canvas_mod,
        pdfdoc_mod,
        page_size: Tuple[float, float],
        part: int,
        conformance: str,
        margin_pt: float = 72.0,
    ):
        self._document = document
        self._part = document.part
        self._canvas_mod = canvas_mod
        self._pdfdoc_mod = pdfdoc_mod
        self._page_w, self._page_h = page_size
        self._margin = margin_pt
        self._pdfa_part = part
        self._pdfa_conformance = conformance

    # -- public API ---------------------------------------------------------

    def render(self) -> bytes:
        """Render the document body and return the resulting PDF bytes."""
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        buf = BytesIO()
        c = self._canvas_mod.Canvas(
            buf, pagesize=(self._page_w, self._page_h)
        )

        # -- attach an XMP metadata stream with PDF/A conformance keys.
        # -- Done before any drawing so the catalog's Metadata reference
        # -- is registered when the doc is finally serialised. --
        self._attach_xmp(c)

        # -- baseline document metadata (title / producer) --
        self._set_doc_info(c)

        cursor_y = self._page_h - self._margin
        c.setFont("Helvetica", _BODY_FONT_SIZE)

        for block in self._document.iter_inner_content():
            if isinstance(block, Paragraph):
                cursor_y = self._render_paragraph(c, block, cursor_y)
            elif isinstance(block, Table):
                cursor_y = self._render_table(c, block, cursor_y)
            else:  # pragma: no cover -- iter_inner_content yields only these
                continue

        c.showPage()
        c.save()
        return buf.getvalue()

    # -- metadata -----------------------------------------------------------

    def _attach_xmp(self, c) -> None:
        """Inject a PDF/A XMP packet into the document catalog."""
        xmp_packet = self._build_xmp_packet()

        XMP = self._pdfdoc_mod.XMP
        # -- the creator callback is invoked at format-time and returns a
        # -- str. ReportLab encodes via its 'extpdfdoc' codec which does
        # -- not handle the U+FEFF BOM; we omit it from the packet. --
        c._doc.Catalog.Metadata = XMP(creator=lambda _doc: xmp_packet)

    def _build_xmp_packet(self) -> str:
        """Return the XMP metadata XML as a string.

        Includes the ``pdfaid:part`` / ``pdfaid:conformance`` keys
        plus ``dc:title`` / ``xmp:CreatorTool`` derived from the
        ``docProps/core.xml`` and ``docProps/app.xml`` parts when
        available.
        """
        title = ""
        creator = _PRODUCER
        try:
            cp = self._document.core_properties
            title = (cp.title or "").strip()
        except Exception:  # pragma: no cover -- defensive
            pass

        # -- escape XML metacharacters in user-provided strings --
        title_xml = _xml_escape(title)
        creator_xml = _xml_escape(creator)

        return (
            '<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?>\n'
            '<x:xmpmeta xmlns:x="adobe:ns:meta/">\n'
            '  <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">\n'
            '    <rdf:Description rdf:about=""\n'
            '        xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"\n'
            '        xmlns:dc="http://purl.org/dc/elements/1.1/"\n'
            '        xmlns:xmp="http://ns.adobe.com/xap/1.0/"\n'
            '        xmlns:pdf="http://ns.adobe.com/pdf/1.3/">\n'
            f'      <pdfaid:part>{self._pdfa_part}</pdfaid:part>\n'
            f'      <pdfaid:conformance>{self._pdfa_conformance}</pdfaid:conformance>\n'
            '      <dc:format>application/pdf</dc:format>\n'
            '      <dc:title>\n'
            '        <rdf:Alt>\n'
            f'          <rdf:li xml:lang="x-default">{title_xml}</rdf:li>\n'
            '        </rdf:Alt>\n'
            '      </dc:title>\n'
            f'      <xmp:CreatorTool>{creator_xml}</xmp:CreatorTool>\n'
            f'      <pdf:Producer>{creator_xml}</pdf:Producer>\n'
            '    </rdf:Description>\n'
            '  </rdf:RDF>\n'
            '</x:xmpmeta>\n'
            '<?xpacket end="w"?>'
        )

    def _set_doc_info(self, c) -> None:
        """Stamp the PDF Info dictionary."""
        try:
            cp = self._document.core_properties
            title = (cp.title or "").strip()
            author = (cp.author or "").strip()
        except Exception:  # pragma: no cover -- defensive
            title = ""
            author = ""
        if title:
            c.setTitle(title)
        if author:
            c.setAuthor(author)
        c.setProducer(_PRODUCER)
        c.setCreator(_PRODUCER)

    # -- paragraphs ---------------------------------------------------------

    def _render_paragraph(self, c, paragraph: "Paragraph", cursor_y: float) -> float:
        """Render `paragraph` and return the new cursor Y."""
        # -- inline images live inside w:r/w:drawing; harvest first so we
        # -- can flow them in after the paragraph text. --
        image_blobs = self._collect_inline_images(paragraph)

        # -- detect hard page break first so the new page starts under
        # -- the same paragraph. --
        if _has_page_break(paragraph):
            c.showPage()
            c.setFont("Helvetica", _BODY_FONT_SIZE)
            cursor_y = self._page_h - self._margin

        heading_level = _heading_level_for(paragraph)
        if heading_level is not None:
            font_size = _HEADING_SIZES.get(heading_level, _BODY_FONT_SIZE)
        else:
            font_size = _BODY_FONT_SIZE

        # -- list-item indentation: ``- `` for bullets, ``1. `` for
        # -- numbered. Indented by 18pt per nesting level. --
        list_info = _list_kind_for(paragraph)
        if list_info is not None:
            kind, level = list_info
            indent_pt = 18.0 * (level + 1)
            marker = "- " if kind == "ul" else "1. "
        else:
            indent_pt = 0.0
            marker = ""

        text_segments = self._collect_run_segments(paragraph)
        if marker and text_segments:
            text_segments = [(marker, False, False, False)] + text_segments

        cursor_y = self._draw_segments(
            c,
            text_segments,
            x=self._margin + indent_pt,
            y=cursor_y,
            font_size=font_size,
        )

        # -- emit any inline pictures collected from this paragraph. --
        for blob, content_type in image_blobs:
            cursor_y = self._draw_image(c, blob, content_type, cursor_y)

        # -- end-of-paragraph half-line spacing for readability --
        cursor_y -= font_size * 0.4
        return cursor_y

    def _collect_run_segments(
        self, paragraph: "Paragraph"
    ) -> List[Tuple[str, bool, bool, bool]]:
        """Return ``[(text, bold, italic, underline)]`` for `paragraph`.

        Hyperlinks lose their URL and are flattened to their text. Empty
        runs are dropped.
        """
        from docx.text.hyperlink import Hyperlink
        from docx.text.run import Run

        segments: List[Tuple[str, bool, bool, bool]] = []
        for child in paragraph.iter_inner_content():
            if isinstance(child, Run):
                seg = self._segment_for_run(child)
                if seg is not None:
                    segments.append(seg)
            elif isinstance(child, Hyperlink):
                for r in child.runs:
                    seg = self._segment_for_run(r)
                    if seg is not None:
                        segments.append(seg)
        return segments

    @staticmethod
    def _segment_for_run(
        run: "Run",
    ) -> Optional[Tuple[str, bool, bool, bool]]:
        text = run.text or ""
        if not text:
            return None
        font = run.font
        return (
            text,
            bool(font.bold),
            bool(font.italic),
            bool(font.underline),
        )

    # -- text drawing -------------------------------------------------------

    def _draw_segments(
        self,
        c,
        segments: List[Tuple[str, bool, bool, bool]],
        x: float,
        y: float,
        font_size: float,
    ) -> float:
        """Wrap and draw inline `segments` starting at (x, y).

        Returns the new cursor Y after the last drawn line.
        """
        if not segments:
            return y - font_size * _LEADING_MULTIPLIER

        max_width = self._page_w - self._margin - x
        leading = font_size * _LEADING_MULTIPLIER
        # -- soft margin so the bottom of the page doesn't clip text --
        bottom_limit = self._margin

        # Flatten segments into (word, bold, italic, underline) tuples,
        # word being a single token. We treat newlines inside a segment
        # as forced line breaks. Spaces between words are preserved as
        # padding via the typesetter.
        words: List[Tuple[str, bool, bool, bool]] = []
        for text, bold, italic, underline in segments:
            # -- explicit \n inside a run -> line break sentinel --
            for line_idx, line in enumerate(text.split("\n")):
                if line_idx > 0:
                    words.append(("\n", bold, italic, underline))
                if not line:
                    continue
                buf_word = ""
                for ch in line:
                    if ch == " ":
                        if buf_word:
                            words.append((buf_word, bold, italic, underline))
                            buf_word = ""
                        words.append((" ", bold, italic, underline))
                    elif ch == "\t":
                        if buf_word:
                            words.append((buf_word, bold, italic, underline))
                            buf_word = ""
                        words.append(("    ", bold, italic, underline))
                    else:
                        buf_word += ch
                if buf_word:
                    words.append((buf_word, bold, italic, underline))

        # Greedy line-breaking by accumulated width.
        line: List[Tuple[str, bool, bool, bool]] = []
        line_width = 0.0
        cursor = y

        def flush_line() -> float:
            nonlocal line, line_width, cursor
            if not line:
                return cursor
            if cursor < bottom_limit:
                c.showPage()
                c.setFont("Helvetica", font_size)
                cursor = self._page_h - self._margin
            self._render_line(c, line, x, cursor, font_size)
            line = []
            line_width = 0.0
            cursor -= leading
            return cursor

        for word, bold, italic, underline in words:
            if word == "\n":
                cursor = flush_line()
                continue
            font_name = _font_for(bold, italic)
            try:
                w = c.stringWidth(word, font_name, font_size)
            except Exception:
                w = font_size * 0.5 * len(word)
            if line_width + w > max_width and line:
                cursor = flush_line()
            line.append((word, bold, italic, underline))
            line_width += w

        cursor = flush_line()
        return cursor

    @staticmethod
    def _render_line(
        c,
        line: List[Tuple[str, bool, bool, bool]],
        x: float,
        y: float,
        font_size: float,
    ) -> None:
        """Render one assembled line at (x, y).

        Each word is set in its run-correct font; underlined words get
        an explicit ``c.line()`` underline drawn at the descender.
        """
        cursor_x = x
        for word, bold, italic, underline in line:
            font_name = _font_for(bold, italic)
            c.setFont(font_name, font_size)
            c.drawString(cursor_x, y, word)
            try:
                w = c.stringWidth(word, font_name, font_size)
            except Exception:
                w = font_size * 0.5 * len(word)
            if underline and word.strip():
                c.line(cursor_x, y - 1.5, cursor_x + w, y - 1.5)
            cursor_x += w

    # -- images -------------------------------------------------------------

    def _collect_inline_images(
        self, paragraph: "Paragraph"
    ) -> List[Tuple[bytes, str]]:
        """Return ``[(blob, content_type)]`` for inline pictures in `paragraph`.

        Anchored drawings, drawings missing an ``r:embed`` rId, and
        drawings whose related part has no blob are skipped silently.
        """
        out: List[Tuple[bytes, str]] = []
        p_elm = paragraph._p
        for drawing in p_elm.iter(qn("w:drawing")):
            inline = drawing.find(qn("wp:inline"))
            if inline is None:
                continue  # -- anchored picture, skipped --
            blip = drawing.xpath(".//a:blip")
            if not blip:
                continue
            r_id = blip[0].get(qn("r:embed"))
            if r_id is None:
                continue
            try:
                image_part = self._part.related_parts[r_id]
            except KeyError:
                continue
            blob = getattr(image_part, "blob", None)
            content_type = getattr(image_part, "content_type", "image/png")
            if not blob:
                continue
            out.append((blob, content_type))
        return out

    def _draw_image(
        self, c, blob: bytes, content_type: str, cursor_y: float
    ) -> float:
        """Draw `blob` inline; return the new cursor Y.

        Images that ReportLab can't decode (EMF / WMF / bare SVG) are
        skipped with no on-page placeholder. Images wider than the
        usable page width are scaled down proportionally.
        """
        try:
            from reportlab.lib.utils import ImageReader  # type: ignore[import-not-found]
        except ImportError:  # pragma: no cover -- reportlab guaranteed at this point
            return cursor_y

        try:
            reader = ImageReader(BytesIO(blob))
            iw, ih = reader.getSize()
        except Exception:
            # -- format reportlab can't handle (e.g. EMF / SVG) --
            return cursor_y

        max_w = self._page_w - 2 * self._margin
        max_h = (self._page_h - 2 * self._margin) * 0.5
        scale = 1.0
        if iw > max_w:
            scale = max_w / iw
        if ih * scale > max_h:
            scale = max_h / ih

        draw_w = iw * scale
        draw_h = ih * scale

        if cursor_y - draw_h < self._margin:
            c.showPage()
            c.setFont("Helvetica", _BODY_FONT_SIZE)
            cursor_y = self._page_h - self._margin

        try:
            c.drawImage(
                reader,
                self._margin,
                cursor_y - draw_h,
                width=draw_w,
                height=draw_h,
                mask="auto",
            )
        except Exception:
            return cursor_y

        return cursor_y - draw_h - 6

    # -- tables -------------------------------------------------------------

    def _render_table(self, c, table: "Table", cursor_y: float) -> float:
        """Render a basic grid for `table`.

        Cells flatten to space-joined paragraph text. Column widths
        are evenly distributed across the usable page width. Rows that
        wrap onto more than one line stretch the row height. Tables
        wider than two pages get truncated at their last fitting row
        (loud failure modes are out of scope for this best-effort
        exporter).
        """
        rows = list(table.rows)
        if not rows:
            return cursor_y

        col_count = max(len(row.cells) for row in rows) or 1
        usable_w = self._page_w - 2 * self._margin
        col_w = usable_w / col_count
        row_padding = 4.0

        font_size = _BODY_FONT_SIZE
        leading = font_size * _LEADING_MULTIPLIER

        for row in rows:
            cells = list(row.cells)
            # -- compute per-cell wrapped lines first to find row height --
            wrapped_per_cell: List[List[str]] = []
            for cell in cells:
                text = self._cell_text(cell)
                lines = _wrap_text(text, col_w - 2 * row_padding, font_size, c)
                wrapped_per_cell.append(lines or [""])
            row_h = (max(len(ws) for ws in wrapped_per_cell) * leading) + (
                2 * row_padding
            )
            # -- page-break the entire row when it won't fit --
            if cursor_y - row_h < self._margin:
                c.showPage()
                c.setFont("Helvetica", _BODY_FONT_SIZE)
                cursor_y = self._page_h - self._margin

            row_top = cursor_y
            for col_idx, lines in enumerate(wrapped_per_cell):
                cell_x = self._margin + col_idx * col_w
                # -- cell border --
                c.rect(cell_x, row_top - row_h, col_w, row_h, stroke=1, fill=0)
                # -- cell text --
                text_y = row_top - row_padding - font_size
                c.setFont("Helvetica", font_size)
                for line in lines:
                    c.drawString(cell_x + row_padding, text_y, line)
                    text_y -= leading

            cursor_y = row_top - row_h

        # -- trailing half-line below the table --
        cursor_y -= leading * 0.5
        return cursor_y

    @staticmethod
    def _cell_text(cell: "TableCell") -> str:
        """Return space-joined text for `cell` (multi-paragraph cells flatten)."""
        return " ".join(
            p.text for p in cell.paragraphs if (p.text or "").strip()
        ).strip()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _heading_level_for(paragraph: "Paragraph") -> Optional[int]:
    """Return the ``1..6`` heading level for `paragraph`, or |None|."""
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


def _list_kind_for(paragraph: "Paragraph") -> Optional[Tuple[str, int]]:
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
    """Return the ``w:numFmt`` value for `num_id` / `level`, or |None|."""
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
    return any(br.get(qn("w:type")) == "page" for br in p_elm.iter(qn("w:br")))


def _font_for(bold: bool, italic: bool) -> str:
    """Return the ReportLab Helvetica face name for the bold/italic combo."""
    if bold and italic:
        return "Helvetica-BoldOblique"
    if bold:
        return "Helvetica-Bold"
    if italic:
        return "Helvetica-Oblique"
    return "Helvetica"


def _wrap_text(text: str, max_width: float, font_size: float, c) -> List[str]:
    """Greedy word-wrap `text` to `max_width` points at `font_size`.

    Falls back to character-wise wrapping when a single word exceeds the
    column width.
    """
    if not text:
        return []
    words = text.split(" ")
    lines: List[str] = []
    current = ""
    for word in words:
        candidate = (current + " " + word).strip() if current else word
        try:
            w = c.stringWidth(candidate, "Helvetica", font_size)
        except Exception:
            w = font_size * 0.5 * len(candidate)
        if w <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            # -- word alone overflows: break by characters --
            if c.stringWidth(word, "Helvetica", font_size) > max_width:
                buf = ""
                for ch in word:
                    if c.stringWidth(buf + ch, "Helvetica", font_size) > max_width:
                        if buf:
                            lines.append(buf)
                        buf = ch
                    else:
                        buf += ch
                current = buf
            else:
                current = word
    if current:
        lines.append(current)
    return lines


def _xml_escape(s: str) -> str:
    """Escape XML metacharacters for the XMP packet."""
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
