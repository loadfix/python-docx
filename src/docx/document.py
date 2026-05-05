# pyright: reportImportCycles=false
# pyright: reportPrivateUsage=false

"""|Document| and closely related objects."""

from __future__ import annotations

import datetime as dt
import os
import re
from typing import IO, TYPE_CHECKING, cast
from collections.abc import Iterator, Sequence

from docx.blkcntnr import BlockItemContainer
from docx.enum.section import WD_SECTION
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.section import Section, Sections
from docx.shared import ElementProxy, Emu, Inches, Length, RGBColor
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.accessibility import HeadingIssue
    from docx.alt_chunk import AltChunk
    from docx.attachments import Attachment
    from docx.bibliography import Bibliography, Source
    from docx.bookmarks import Bookmark, Bookmarks
    from docx.chart import Chart, WD_CHART_TYPE
    from docx.comments import Comment, Comments
    from docx.content_controls import ContentControl, ContentControlType
    from docx.custom_properties import CustomProperties
    from docx.extended_properties import ExtendedProperties
    from docx.custom_xml import CustomXmlPart
    from docx.drawing import Canvas
    from docx.embedded_objects import EmbeddedObject
    from docx.endnotes import Endnotes, EndnoteProperties
    from docx.equations import Equation
    from docx.font_table import FontTable
    from docx.footnotes import FootnoteProperties, Footnotes
    from docx.form_fields import FormField
    from docx.glossary import Glossary
    from docx.ink import InkAnnotation
    from docx.oxml.content_controls import CT_Sdt
    from docx.oxml.document import CT_Body, CT_Document
    from docx.oxml.table import CT_Tbl
    from docx.parts.document import DocumentPart
    from docx.permissions import PermissionRange
    from docx.search import SearchMatch
    from docx.settings import Settings
    from docx.signatures import SignatureInfo
    from docx.smart_art import SmartArt
    from docx.statistics import DocumentStatistics
    from docx.styles.style import ParagraphStyle, _TableStyle
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.theme import Theme
    from docx.tracked_changes import _TrackedChangesCtx
    from docx.web_settings import WebSettings


class Document(ElementProxy):
    """WordprocessingML (WML) document.

    Not intended to be constructed directly. Use :func:`docx.Document` to open or create
    a document.
    """

    def __init__(self, element: CT_Document, part: DocumentPart):
        super().__init__(element)
        self._element = element
        self._part = part
        self.__body = None
        # -- Name of the style queued for the *next* ``add_paragraph`` call
        # -- that doesn't pass an explicit `style`. Set when the previous
        # -- paragraph was added with a style whose ``w:next`` points at
        # -- another style. See :meth:`add_paragraph`. upstream#888.
        self._pending_next_style: str | None = None
        # -- active `tracked_changes()` context, if any. Stored as a simple
        # -- LIFO stack so nested `with document.tracked_changes(...)` blocks
        # -- shadow outer ones.
        from docx.tracked_changes import _TrackedChangesCtx

        self._tracked_changes_stack: list[_TrackedChangesCtx] = []
        # -- expose this proxy on the part so deep children (Paragraph,
        # -- Run, etc.) can look up the active tracked-changes state via
        # -- `self.part._track_changes_doc_proxy`. Safe to set: the part
        # -- has exactly one Document proxy. The try/except tolerates
        # -- `spec_set` mocks used in unit tests. --
        try:
            setattr(part, "_track_changes_doc_proxy", self)
        except AttributeError:  # pragma: no cover -- test fixtures only
            pass

    def __enter__(self) -> Document:
        """Enter context-manager; returns `self`.

        Pairs with :meth:`__exit__` / :meth:`close` to provide lifecycle
        symmetry with other file-like resources. The document itself does
        not hold an OS file handle after opening — |docx| fully parses the
        underlying ``.docx`` zip on construction — so the context-manager
        protocol exists mainly to let callers use the familiar ``with``
        idiom (upstream#379)::

            with Document('example.docx') as document:
                document.add_paragraph('Added in context.')
                document.save('out.docx')

        .. versionadded:: 2026.05.0
        """
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Exit context-manager; delegates to :meth:`close`.

        .. versionadded:: 2026.05.0
        """
        self.close()

    def close(self) -> None:
        """Release resources associated with this document.

        python-docx reads the ``.docx`` package eagerly on construction and
        does not retain the source file handle, so :meth:`close` is a no-op
        today. It exists to give callers a symmetric lifecycle API — useful
        for code that treats a ``Document`` like any other closeable
        resource — and is safe to call multiple times.

        .. versionadded:: 2026.05.0
        """
        # -- drop any stale tracked-changes context state --
        self._tracked_changes_stack = []

    def add_comment(
        self,
        runs: Run | Sequence[Run],
        text: str | None = "",
        author: str = "",
        initials: str | None = "",
        date: dt.datetime | None = None,
    ) -> Comment:
        """Add a comment to the document, anchored to the specified runs.

        `runs` can be a single `Run` object or a non-empty sequence of `Run` objects. Only the
        first and last run of a sequence are used, it's just more convenient to pass a whole
        sequence when that's what you have handy, like `paragraph.runs` for example. When `runs`
        contains a single `Run` object, that run serves as both the first and last run.

        A comment can be anchored only on an even run boundary, meaning the text the comment
        "references" must be a non-zero integer number of consecutive runs. The runs need not be
        _contiguous_ per se, like the first can be in one paragraph and the last in the next
        paragraph, but all runs between the first and the last will be included in the reference.

        The comment reference range is delimited by placing a `w:commentRangeStart` element before
        the first run and a `w:commentRangeEnd` element after the last run. This is why only the
        first and last run are required and why a single run can serve as both first and last.
        Word works out which text to highlight in the UI based on these range markers.

        `text` allows the contents of a simple comment to be provided in the call, providing for
        the common case where a comment is a single phrase or sentence without special formatting
        such as bold or italics. More complex comments can be added using the returned `Comment`
        object in much the same way as a `Document` or (table) `Cell` object, using methods like
        `.add_paragraph()`, .add_run()`, etc.

        The `author` and `initials` parameters allow that metadata to be set for the comment.
        `author` is a required attribute on a comment and is the empty string by default.
        `initials` is optional on a comment and may be omitted by passing |None|, but Word adds an
        `initials` attribute by default and we follow that convention by using the empty string
        when no `initials` argument is provided.

        `date` is the timestamp recorded on the comment's ``w:date`` attribute. When
        omitted or |None|, ``datetime.now(timezone.utc)`` is used. Pass an explicit
        datetime for reproducible-save scenarios where an implicit wall-clock
        timestamp would defeat byte-identical output.

        .. versionchanged:: 2026.05.5
           Added the ``date`` parameter — previously only the lower-level
           ``Document.comments.add_comment(date=...)`` accepted it, so
           reproducible fixtures had to poke ``comment._element.date`` by hand.
        """
        # -- normalize `runs` to a sequence of runs --
        runs = [runs] if isinstance(runs, Run) else runs
        first_run = runs[0]
        last_run = runs[-1]

        # -- Note that comments can only appear in the document part --
        comment = self.comments.add_comment(
            text=text, author=author, initials=initials, date=date
        )

        # -- let the first run orchestrate placement of the comment range start and end --
        first_run.mark_comment_range(last_run, comment.comment_id)

        return comment

    def add_citation(
        self,
        tag: str,
        title: "str | None" = None,
        author: "str | None" = None,
        year: "str | int | None" = None,
        source_type: str = "Book",
        **extra: str,
    ) -> "Source":
        """Append a bibliographic source to the document's bibliography.

        Creates the backing ``/customXml/item{N}.xml`` part (and its
        ``itemProps{N}.xml`` sibling) on first use. `tag` is the unique
        citation key that :meth:`Paragraph.add_citation_reference` looks up;
        re-using an existing tag raises :class:`ValueError`.

        `source_type` defaults to ``"Book"``; pass e.g. ``"JournalArticle"``
        or ``"InternetSite"`` for richer fixtures. Any ``**extra`` kwargs
        become text-only ``<b:{Capitalized}>`` children of the source
        element (e.g. ``city="London"`` → ``<b:City>London</b:City>``).

        Returns the newly-added :class:`Source`.

        .. versionadded:: 2026.05.7
        """
        return self.bibliography.add_source(
            tag,
            title=title,
            author=author,
            year=year,
            source_type=source_type,
            **extra,
        )

    def add_bookmark(self, runs: Run | Sequence[Run], name: str) -> Bookmark:
        """Add a bookmark spanning `runs`, and return the |Bookmark| proxy.

        `runs` may be a single |Run| or a non-empty sequence of |Run| objects.
        Only the first and last run of a sequence are used — just as with
        :meth:`add_comment` — so the caller can pass a whole
        ``paragraph.runs`` or a selection that spans paragraphs.

        A ``w:bookmarkStart`` element is inserted immediately before the
        first run and a matching ``w:bookmarkEnd`` is inserted immediately
        after the last run. A fresh ``@w:id`` is allocated from the body's
        existing bookmark ids; `name` must be unique within the document
        (caller enforced — Word accepts duplicates silently but treats
        them as ambiguous cross-reference targets).

        This is the symmetric counterpart to :meth:`Paragraph.add_bookmark`
        — use that helper for a single-paragraph bookmark and this one when
        the range must span multiple paragraphs without dropping into the
        oxml layer.

        .. versionadded:: 2026.05.0
        """
        from docx.bookmarks import Bookmark
        from docx.text.paragraph import Paragraph

        runs = [runs] if isinstance(runs, Run) else list(runs)
        if not runs:
            raise ValueError("runs must be a non-empty sequence of Run objects")
        first_run = runs[0]
        last_run = runs[-1]

        body = self._element.body
        bookmark_id = Paragraph._next_bookmark_id(body)

        first_run._r.insert_bookmark_start_before(bookmark_id, name)
        last_run._r.insert_bookmark_end_after(bookmark_id)

        bookmarkStart = body.xpath(f".//w:bookmarkStart[@w:id='{bookmark_id}']")[0]
        return Bookmark(bookmarkStart, body)

    def add_caption(
        self,
        text: str,
        label: str = "Figure",
        style: str = "Caption",
    ) -> Paragraph:
        """Return a new caption |Paragraph| appended to the end of the document body.

        A Word caption is a paragraph styled with the ``Caption`` style that
        auto-numbers itself via a ``SEQ`` field. The resulting paragraph has
        the shape ``"{label} N: {text}"`` where ``N`` is produced by a
        ``SEQ {label} \\* ARABIC`` field; Word updates the number when the
        document is opened or fields are refreshed.

        `label` groups captions into a named sequence (e.g. ``"Figure"`` or
        ``"Table"``). `style` selects the paragraph style applied to the
        caption paragraph and defaults to the built-in ``"Caption"`` style.

        .. versionadded:: 2026.05.0
        """
        from docx.captions import new_caption_paragraph

        paragraph = self.add_paragraph()
        return new_caption_paragraph(paragraph, text, label=label, style=style)

    def add_heading(self, text: str = "", level: int = 1):
        """Return a heading paragraph newly added to the end of the document.

        The heading paragraph will contain `text` and have its paragraph style
        determined by `level`. If `level` is 0, the style is set to `Title`. If `level`
        is 1 (or omitted), `Heading 1` is used. Otherwise the style is set to `Heading
        {level}`. Raises |ValueError| if `level` is outside the range 0-9.
        """
        if not 0 <= level <= 9:
            raise ValueError("level must be in range 0-9, got %d" % level)
        style = "Title" if level == 0 else "Heading %d" % level
        return self.add_paragraph(text, style)

    def add_page_break(self) -> Paragraph:
        """Return newly |Paragraph| object containing only a page break."""
        paragraph = self.add_paragraph()
        paragraph.add_page_break()
        return paragraph

    def add_paragraph(
        self,
        text: str = "",
        style: str | ParagraphStyle | None = None,
        track_author: str | None = None,
    ) -> Paragraph:
        """Return paragraph newly added to the end of the document.

        The paragraph is populated with `text` and having paragraph style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break.

        If `track_author` is supplied, the inserted run is wrapped in a
        `w:ins` element with ``@w:author`` set to that string (closes
        upstream#1025). When the document has an active
        :meth:`tracked_changes` context, `track_author` defaults to the
        author of the innermost context; pass ``track_author=""`` explicitly
        if you need to opt out of the active context for one call.

        If the previously-added paragraph had a style whose ``w:next`` pointed
        at another style, and `style` is |None| on this call, that "next"
        style is applied automatically — mirroring the behaviour Word exhibits
        when the user presses Enter. An explicit `style` argument (including
        an explicit ``style=None`` passed positionally) always takes
        precedence. Closes upstream#888.

        .. versionadded:: 2026.05.0
           Added ``track_author`` keyword argument.
        """
        effective_style = style
        if effective_style is None and self._pending_next_style is not None:
            effective_style = self._pending_next_style
        # -- reset pending before recomputing to avoid runaway chains --
        self._pending_next_style = None

        if track_author is None:
            paragraph = self._body.add_paragraph(text, effective_style)
        else:
            paragraph = self._body.add_paragraph(
                text, effective_style, track_author=track_author
            )

        # -- queue the `w:next` style, if any, for the subsequent call --
        if effective_style is not None:
            next_style_id = self._resolve_next_style_id(effective_style)
            if next_style_id is not None:
                self._pending_next_style = next_style_id

        return paragraph

    def _resolve_next_style_id(self, style: str | ParagraphStyle) -> str | None:
        """Return the ``w:styleId`` of `style`'s ``w:next`` style, if any.

        `style` may be a style name, a style id, or a |ParagraphStyle| proxy.
        Returns |None| when the style is not found, has no ``w:next``
        element, or ``w:next`` points at an undefined style id.
        """
        from docx.styles.style import BaseStyle

        if isinstance(style, BaseStyle):
            style_elm = style._element
        else:
            # -- try name first, fall back to style_id lookup --
            styles = self._part.styles
            try:
                resolved = styles[style]
            except KeyError:
                resolved = None
            if resolved is None:
                return None
            style_elm = resolved._element
        return style_elm.next_val

    def tracked_changes(
        self, author: str, date: "dt.datetime | None" = None
    ) -> "_TrackedChangesCtx":
        """Return a context manager that wraps new content in `w:ins` markers.

        Every call to :meth:`add_paragraph` / :meth:`BlockItemContainer.add_paragraph`
        or :meth:`Paragraph.add_run` made while the returned context is
        active wraps the new ``w:r`` element in a ``w:ins`` element whose
        ``@w:author`` is `author` and whose ``@w:date`` is `date` (defaulting
        to the current UTC time). Contexts can be nested; the innermost
        author/date wins.

        Closes upstream#1025::

            with document.tracked_changes(author="Reviewer"):
                document.add_paragraph("A new paragraph under review.")

        .. versionadded:: 2026.05.0
        """
        from docx.tracked_changes import _TrackedChangesCtx

        return _TrackedChangesCtx(self, author, date)

    def add_content_control(
        self,
        type: ContentControlType,
        tag: str | None = None,
        title: str | None = None,
    ) -> ContentControl:
        """Add a block-level content control at the end of the document body.

        `type` selects the kind of content control (see :class:`ContentControlType`).
        `tag` and `title` map to `w:sdtPr/w:tag/@w:val` and `w:sdtPr/w:alias/@w:val`
        respectively.

        .. versionadded:: 2026.05.0
        """
        return self._body.add_content_control(type, tag=tag, title=title)

    def add_chart(
        self,
        chart_type: WD_CHART_TYPE,
        categories: list[str],
        series_data: dict[str, list[float]],
        width: Length | None = None,
        height: Length | None = None,
    ) -> Chart:
        """Append a new chart to the end of the document and return it.

        `chart_type` selects the kind of chart (see :class:`docx.chart.WD_CHART_TYPE`).
        Only ``BAR``, ``BAR_STACKED``, ``COLUMN``, ``COLUMN_STACKED``, ``LINE``,
        and ``PIE`` are supported for creation.

        `categories` is a list of category labels (for example x-axis labels).
        `series_data` maps each series name to its list of numeric values; every
        value list must be the same length as `categories`.

        `width` and `height` are |Length| instances controlling the display size
        of the inline chart. When omitted a 6" x 3" default is used (similar to
        Word's default inline chart size).

        .. versionadded:: 2026.05.0
        """
        from docx.chart import Chart
        from docx.opc.constants import RELATIONSHIP_TYPE as _RT
        from docx.oxml.shape import CT_Inline
        from docx.parts.chart import ChartPart
        from docx.shared import Emu, Inches

        cx = width if width is not None else Inches(6)
        cy = height if height is not None else Inches(3)
        cx = Emu(int(cx))
        cy = Emu(int(cy))

        # -- create the chart part and relate it to the document part --
        package = self._part.package
        assert package is not None
        chart_part = ChartPart.new(package, chart_type, categories, series_data)
        rId = self._part.relate_to(chart_part, _RT.CHART)

        # -- build the wp:inline drawing pointing at the chart part --
        shape_id = self._part.next_id
        inline = CT_Inline.new_chart_inline(shape_id, rId, cx, cy)

        # -- append a new paragraph and drawing to the body --
        paragraph = self.add_paragraph()
        run = paragraph.add_run()
        run._r.add_drawing(inline)

        return Chart(chart_part)

    def add_shape(
        self,
        shape_type,
        width: Length | None = None,
        height: Length | None = None,
        text: str | None = None,
    ):
        """Append an inline DrawingML preset shape in its own paragraph.

        `shape_type` is a :class:`docx.enum.shape.WD_SHAPE` member (e.g.
        ``WD_SHAPE.ROUNDED_RECTANGLE``). `width` and `height` are |Length|
        values; they default to 2" x 1" when omitted. When `text` is provided a
        minimal text-frame is attached so Word renders the string inside the
        shape. The shape is emitted as a ``wps:wsp`` with ``a:prstGeom`` inside
        a ``w:drawing/wp:inline``.

        Returns a :class:`docx.drawing.WordprocessingShape` proxy for the new
        shape. Closes upstream#1112 and upstream#517.

        .. versionadded:: 2026.05.0
        """
        paragraph = self.add_paragraph()
        return paragraph.add_shape(shape_type, width, height, text=text)

    def add_canvas(
        self,
        width: Length | None = None,
        height: Length | None = None,
    ) -> Canvas:
        """Append a DrawingML canvas (``wpc:wpc``) in its own paragraph.

        A canvas groups one or more shapes or pictures under a single
        ``w:drawing/wp:inline/a:graphic/a:graphicData`` whose URI is the
        WordprocessingCanvas namespace. `width` and `height` default to
        6" x 3".

        Returns a :class:`docx.drawing.Canvas` proxy. Callers can build up the
        canvas contents via :meth:`Canvas.add_shape`. Closes upstream#411.

        .. versionadded:: 2026.05.0
        """
        from docx.drawing import Canvas
        from docx.oxml.drawing import new_inline_canvas_drawing

        cx = int(width) if width is not None else int(Inches(6))
        cy = int(height) if height is not None else int(Inches(3))

        shape_id = self._part.next_id
        name = "Canvas %d" % shape_id
        drawing = new_inline_canvas_drawing(cx, cy, shape_id, name)

        paragraph = self.add_paragraph()
        run = paragraph.add_run()
        run._r.append(drawing)

        wpc = drawing.xpath(".//wp:inline/a:graphic/a:graphicData/wpc:wpc")[0]
        return Canvas(wpc, paragraph)

    def add_text_box(
        self,
        width: Length | None = None,
        height: Length | None = None,
        text: str | None = None,
    ):
        """Append an inline DrawingML text box in its own paragraph.

        A text box is a ``wps:wsp`` whose preset geometry is a simple rectangle
        carrying a ``wps:txbx/w:txbxContent`` text frame. `width` and `height`
        default to 3" x 1.5". When `text` is provided, the text box is
        initialised with a single paragraph containing that string; otherwise
        callers can append paragraphs via
        :meth:`~docx.drawing.WordprocessingShape.add_paragraph`.

        Returns a :class:`docx.drawing.WordprocessingShape` proxy. Closes
        upstream#524.

        .. versionadded:: 2026.05.0
        """
        from docx.enum.shape import WD_SHAPE

        cx = width if width is not None else Inches(3)
        cy = height if height is not None else Inches(1.5)

        paragraph = self.add_paragraph()
        shape = paragraph.add_shape(
            WD_SHAPE.RECTANGLE, cx, cy, text=text if text is not None else ""
        )
        return shape

    def add_picture(
        self,
        image_path_or_stream: "str | os.PathLike[str] | IO[bytes] | None" = None,
        width: int | Length | None = None,
        height: int | Length | None = None,
        link: bool = False,
        save_with_document: bool = True,
        url: str | None = None,
    ):
        """Return new picture shape added in its own paragraph at end of the document.

        The picture contains the image at `image_path_or_stream`, scaled based on
        `width` and `height`. If neither width nor height is specified, the picture
        appears at its native size. If only one is specified, it is used to compute a
        scaling factor that is then applied to the unspecified dimension, preserving the
        aspect ratio of the image. The native size of the picture is calculated using
        the dots-per-inch (dpi) value specified in the image file, defaulting to 72 dpi
        if no value is specified, as is often the case.

        `image_path_or_stream` may be a ``str`` path, an ``os.PathLike`` (e.g.
        :class:`pathlib.Path`), or a binary file-like object.

        When `link` is |True| and `save_with_document` is |False|, a linked
        (external) picture is inserted: the `a:blip` uses ``r:link`` and no
        image part is added to the package. `url` may be supplied to link
        a remote image rather than a local path. See
        :meth:`docx.text.run.Run.add_picture` for details.

        .. versionchanged:: 2026.05.0
           Accepts :class:`os.PathLike` path arguments.

        .. versionadded:: 2026.05.0
            ``link``, ``save_with_document``, and ``url`` parameters.
        """
        if isinstance(image_path_or_stream, os.PathLike):
            image_path_or_stream = os.fspath(image_path_or_stream)
        run = self.add_paragraph().add_run()
        return run.add_picture(
            image_path_or_stream,
            width,
            height,
            link=link,
            save_with_document=save_with_document,
            url=url,
        )

    def add_section(self, start_type: WD_SECTION = WD_SECTION.NEW_PAGE):
        """Return a |Section| object newly added at the end of the document.

        The optional `start_type` argument must be a member of the :ref:`WdSectionStart`
        enumeration, and defaults to ``WD_SECTION.NEW_PAGE`` if not provided.
        """
        new_sectPr = self._element.body.add_section_break()
        new_sectPr.start_type = start_type
        return Section(new_sectPr, self._part)

    def add_table_of_contents(
        self, levels: tuple[int, int] = (1, 3)
    ) -> Paragraph:
        """Append a table-of-contents paragraph at the end of the document body.

        Creates a new paragraph containing a ``TOC`` complex field whose
        cached *result text* is a preview built from the document's existing
        heading paragraphs. ``levels`` is a ``(min_level, max_level)`` tuple
        (default ``(1, 3)``) that selects which ``"Heading N"`` paragraphs
        contribute to the preview; headings outside the range are skipped.

        The preview lists one heading per line in the form
        ``"{text}\\t{index}"`` — the tab-separated trailing integer is a
        1-based position in the filtered heading list, not a page number
        (python-docx has no layout engine). Word rebuilds the TOC the next
        time the document is opened or when fields are refreshed, so the
        cached preview is purely a convenience for raw-XML consumers and
        non-Word viewers.

        Returns the newly-appended |Paragraph|.

        .. versionadded:: 2026.05.0
        """
        from docx.toc import populate_toc_paragraph

        # -- snapshot the current headings before appending the new paragraph
        #    so the empty TOC paragraph doesn't self-include if its style
        #    ever matched the heading regex. --
        source_paragraphs = list(self.paragraphs)
        paragraph = self.add_paragraph()
        return populate_toc_paragraph(paragraph, source_paragraphs, levels)

    def add_table_copy(self, other_table: Table) -> Table:
        """Append a deep copy of `other_table` (possibly from another document) to this body.

        The entire ``w:tbl`` element is deep-copied, then scanned for
        cross-document references that must be rewired into this document's
        package:

        - Embedded images: every ``a:blip/@r:embed`` (plus SVG sibling references
          via ``asvg:svgBlip/@r:embed``) is resolved against `other_table`'s
          document part. The image part is copied into this document's package
          (de-duplicated by SHA-1) and the ``r:embed`` attribute is rewritten to
          the freshly-minted ``rId`` in this part's relationships.
        - Table-style reference: when ``w:tblStyle/@w:val`` names a style that
          does not yet exist in this document, the source style's ``w:style``
          element is deep-copied into this document's styles part. *Advanced
          style cascades (``w:basedOn`` / ``w:link`` / ``w:next`` chains),
          numbering references, and conditional table-style formatting are
          not recursively imported; those remain TODO.*

        Returns the |Table| wrapping the inserted ``w:tbl``. When `other_table`
        belongs to this same document the copy is inserted without any rewiring
        (rIds are already valid in the current part).

        Closes upstream#612, #270.

        .. versionadded:: 2026.05.0
        """
        from copy import deepcopy

        from docx.table import Table

        src_tbl = other_table._tbl
        src_part = other_table.part
        dest_part = self._part

        new_tbl = deepcopy(src_tbl)

        # -- rewire image references (a:blip and asvg:svgBlip @r:embed) --
        if src_part is not dest_part:
            self._rewire_blip_refs(new_tbl, src_part)
            self._import_table_style(new_tbl, src_part)

        # -- insert the copied w:tbl at the end of the body (before sectPr) --
        body = self._element.body
        body._insert_tbl(new_tbl)  # pyright: ignore[reportPrivateUsage]

        return Table(new_tbl, self._body)

    def add_table_from(self, other_table: Table) -> Table:
        """Alias for :meth:`add_table_copy`.

        .. versionadded:: 2026.05.0
        """
        return self.add_table_copy(other_table)

    def _rewire_blip_refs(self, tbl: "CT_Tbl", src_part: DocumentPart) -> None:
        """Rewire ``a:blip/@r:embed`` refs on ``tbl`` to this document's rels.

        Copies the referenced image parts into this document's package
        (de-duplicating by SHA-1 via
        :meth:`docx.package.Package.get_or_add_image_part`) and rewrites every
        matching ``r:embed`` attribute on the copied tree. Unresolvable refs
        (the source rId is missing or points to a non-image part) are left
        untouched so Word still reads the file; the caller can flag them.

        Not a public API.
        """
        from docx.opc.constants import RELATIONSHIP_TYPE as _RT
        from docx.oxml.ns import qn
        from docx.parts.image import ImagePart

        dest_part = self._part
        package = dest_part.package

        # -- collect every blip-style element with @r:embed --
        embed_attr = qn("r:embed")
        blips = tbl.xpath(
            ".//a:blip[@r:embed] | .//asvg:svgBlip[@r:embed]"
        )
        rid_map: dict[str, str] = {}
        for blip in blips:
            old_rid = blip.get(embed_attr)
            if old_rid is None:
                continue
            new_rid = rid_map.get(old_rid)
            if new_rid is None:
                try:
                    src_img_part = src_part.related_parts[old_rid]
                except KeyError:
                    continue
                if not isinstance(src_img_part, ImagePart):
                    continue
                # -- copy blob into this package's image_parts (dedup by sha1) --
                import io as _io

                new_img_part = package.get_or_add_image_part(
                    _io.BytesIO(src_img_part.blob)
                )
                new_rid = dest_part.relate_to(new_img_part, _RT.IMAGE)
                rid_map[old_rid] = new_rid
            blip.set(embed_attr, new_rid)

    def _import_table_style(self, tbl: "CT_Tbl", src_part: DocumentPart) -> None:
        """Ensure the table's ``w:tblStyle/@w:val`` exists in this document.

        If the referenced styleId is not already defined in this document's
        styles part, deep-copy the source ``w:style`` element across. Styles
        that cascade through ``w:basedOn`` / ``w:link`` / ``w:next`` are *not*
        imported recursively — a V1 limitation noted in the method's caller.

        Not a public API.
        """
        from copy import deepcopy

        from docx.oxml.ns import qn

        styleId = tbl.xpath("string(./w:tblPr/w:tblStyle/@w:val)")
        if not styleId:
            return
        # -- resolve styles parts on both sides (create empty one on dest if needed) --
        dest_styles_part = self._part._styles_part  # pyright: ignore[reportPrivateUsage]
        dest_styles_elm = dest_styles_part.element
        if dest_styles_elm.get_by_id(styleId) is not None:
            return
        try:
            src_styles_part = src_part._styles_part  # pyright: ignore[reportPrivateUsage]
        except Exception:
            try:
                from docx.opc.constants import RELATIONSHIP_TYPE as _RT

                src_styles_part = src_part.part_related_by(_RT.STYLES)
            except Exception:
                return
        src_style = src_styles_part.element.get_by_id(styleId)
        if src_style is None:
            return
        dest_styles_elm.append(deepcopy(src_style))

    def add_alt_chunk(
        self,
        content: bytes | str,
        content_type: str = "text/html",
    ) -> AltChunk:
        """Append an ``altChunk`` import reference and return an |AltChunk| proxy.

        Creates a new alternate-format import part carrying `content` with
        the given `content_type` (``text/html`` by default), wires it to
        the main document through an ``aFChunk`` relationship, and appends
        a ``w:altChunk`` element at the end of the body that references
        the relationship. Word substitutes the payload when the document
        is opened — python-docx does not evaluate the content itself.

        `content` may be :class:`bytes` or a UTF-8-decodable :class:`str`.
        Closes upstream#1317, upstream#1103, and PR#649.

        .. versionadded:: 2026.05.0
        """
        from docx.alt_chunk import add_alt_chunk_to_document

        return add_alt_chunk_to_document(self._part, content, content_type)

    @property
    def alt_chunks(self) -> list[AltChunk]:
        """List of |AltChunk| proxies for every ``w:altChunk`` in this document.

        Returns an empty list when the document has no ``altChunk``
        imports. Order matches the document order of the ``w:altChunk``
        elements.

        .. versionadded:: 2026.05.0
        """
        from docx.alt_chunk import iter_alt_chunks

        return iter_alt_chunks(self._part)

    def add_list_of_figures(self, caption_label: str = "Figure") -> Paragraph:
        """Append a "List of Figures" field paragraph at the end of the document.

        Emits a ``TOC \\c "Figure"`` complex field which Word evaluates to
        build a table of items whose caption label matches `caption_label`.
        The cached result text is left empty; the field is marked *dirty*
        so Word rebuilds it on open. Closes upstream#723.

        `caption_label` defaults to ``"Figure"`` but may be set to any
        caption label in the document (e.g. ``"Illustration"``).

        Returns the newly-appended |Paragraph|.

        .. versionadded:: 2026.05.0
        """
        paragraph = self.add_paragraph()
        instr = f' TOC \\c "{caption_label}" '
        field = paragraph.add_complex_field(instr)
        field.mark_dirty()
        return paragraph

    def add_list_of_tables(self, caption_label: str = "Table") -> Paragraph:
        """Append a "List of Tables" field paragraph at the end of the document.

        Emits a ``TOC \\c "Table"`` complex field which Word evaluates to
        build a table of items whose caption label matches `caption_label`.
        The cached result text is left empty; the field is marked *dirty*
        so Word rebuilds it on open. Closes upstream#723.

        Returns the newly-appended |Paragraph|.

        .. versionadded:: 2026.05.0
        """
        paragraph = self.add_paragraph()
        instr = f' TOC \\c "{caption_label}" '
        field = paragraph.add_complex_field(instr)
        field.mark_dirty()
        return paragraph

    def add_table(self, rows: int, cols: int, style: str | _TableStyle | None = None):
        """Add a table having row and column counts of `rows` and `cols` respectively.

        `style` may be a table style object or a table style name. If `style` is |None|,
        the table inherits the default table style of the document.

        Raises |KeyError| (or whatever the style lookup surfaces) when `style`
        names a style that does not exist or is of the wrong type. In that case
        no ``w:tbl`` element is left behind in the body — the freshly-appended
        table is rolled back before the exception propagates. See upstream#563.
        """
        table = self._body.add_table(rows, cols, self._block_width)
        try:
            table.style = style
        except Exception:
            # -- rollback: remove the freshly-added w:tbl so a bad style name
            # -- doesn't leave an orphan table in the body. --
            table.delete()
            raise
        return table

    @property
    def background_color(self) -> RGBColor | None:
        """Document-wide page background color, or |None| if not set.

        Maps to the ``w:color`` attribute on the ``w:background`` child of the
        ``w:document`` root element. Assigning an |RGBColor| writes (or updates)
        the ``w:background`` element. Assigning |None| removes the element.

        .. versionadded:: 2026.05.0
        """
        background = self._element.background
        if background is None:
            return None
        color = background.color
        if not isinstance(color, RGBColor):
            return None
        return color

    @background_color.setter
    def background_color(self, value: RGBColor | None) -> None:
        if value is None:
            self._element._remove_background()
            return
        background = self._element.get_or_add_background()
        background.color = value

    @property
    def bibliography(self) -> "Bibliography":
        """A |Bibliography| of citation sources backing this document.

        Lazily materializes a ``/customXml/item{N}.xml`` part carrying an empty
        ``<b:Sources>`` root if no bibliography is already related to the
        document. Subsequent calls return a proxy over the same underlying
        part.

        .. versionadded:: 2026.05.7
        """
        return self._part.bibliography

    @property
    def bookmarks(self) -> Bookmarks:
        """A |Bookmarks| object providing access to the bookmarks in this document.

        .. versionadded:: 2026.05.0
        """
        from docx.bookmarks import Bookmarks

        return Bookmarks(self._element.body)

    @property
    def charts(self) -> list[Chart]:
        """List of |Chart| for each chart referenced from the document body.

        A chart is any ``c:chart`` reference element inside a drawing in the
        body (inline or floating). Each reference is resolved to its
        :class:`docx.parts.chart.ChartPart` via the document's relationship
        graph. References whose target part is missing or of the wrong type
        are skipped. Empty list when no charts are present.

        .. versionadded:: 2026.05.0
        """
        from docx.chart import Chart
        from docx.parts.chart import ChartPart

        result: list[Chart] = []
        rIds = self._element.body.xpath(
            ".//w:drawing/wp:inline/a:graphic/a:graphicData/c:chart/@r:id"
            " | .//w:drawing/wp:anchor/a:graphic/a:graphicData/c:chart/@r:id"
        )
        seen: set[str] = set()
        for rId in rIds:
            if rId in seen:
                continue
            seen.add(rId)
            try:
                chart_part = self._part.related_parts[rId]
            except KeyError:
                continue
            if isinstance(chart_part, ChartPart):
                result.append(Chart(chart_part))
        return result

    @property
    def comments(self) -> Comments:
        """A |Comments| object providing access to comments added to the document."""
        return self._part.comments

    @property
    def permission_ranges(self) -> list[PermissionRange]:
        """All rich-text permission ranges (`w:permStart`) in this document body.

        Returned list is ordered by document-order of the `w:permStart` elements
        in the body.

        .. versionadded:: 2026.05.0
        """
        from docx.oxml.permissions import CT_PermStart
        from docx.permissions import PermissionRange

        body = self._element.body
        return [
            PermissionRange(cast("CT_PermStart", ps), body)
            for ps in body.xpath(".//w:permStart")
        ]

    @property
    def content_controls(self) -> list[ContentControl]:
        """All block-level |ContentControl| objects in this document body, in order.

        Only block-level content controls (direct children of `w:body`) are returned.
        Inline content controls are accessible via :attr:`Paragraph.content_controls`.

        .. versionadded:: 2026.05.0
        """
        return self._body.content_controls

    @property
    def endnotes(self) -> Endnotes:
        """A |Endnotes| object providing access to endnotes in the document.

        .. versionadded:: 2026.05.0
        """
        return self._part.endnotes

    @property
    def equations(self) -> list[Equation]:
        """List of |Equation| for each OMML expression in the document body.

        Walks the body for both inline ``m:oMath`` elements and display-mode
        ``m:oMathPara`` wrappers. Each equation is returned once — an
        ``m:oMath`` nested inside an ``m:oMathPara`` is represented by its
        enclosing wrapper, not separately. Equations inside headers, footers,
        footnotes, endnotes, or comments are not included here; those stories
        are accessible via the corresponding container objects.

        .. versionadded:: 2026.05.0
        """
        from docx.equations import Equation

        body = self._element.body
        result: list[Equation] = []
        for el in body.xpath(
            ".//m:oMathPara | .//m:oMath[not(ancestor::m:oMathPara)]"
        ):
            result.append(Equation(el))
        return result

    @property
    def form_fields(self) -> list[FormField]:
        """All legacy form fields (``w:ffData``) found in the document body, in order.

        Walks top-level body paragraphs only. Form fields nested inside table
        cells, headers, footers, footnotes, or endnotes are not included in
        this collection — callers can access those via the ``form_fields``
        property on the enclosing paragraph.

        .. versionadded:: 2026.05.0
        """
        result: list[FormField] = []
        for paragraph in self.paragraphs:
            result.extend(paragraph.form_fields)
        return result

    @property
    def has_macros(self) -> bool:
        """True if this document contains a VBA project (macros).

        .. versionadded:: 2026.05.0
        """
        try:
            self._part.part_related_by(RT.VBA_PROJECT)
            return True
        except KeyError:
            return False

    @property
    def is_signed(self) -> bool:
        """True when this document's package contains digital-signature parts.

        python-docx does not verify signatures; this reports only whether a
        ``_xmlsignatures/origin.sigs`` or signature relationship is present at the
        package level.

        .. versionadded:: 2026.05.0
        """
        from docx.package import Package

        return cast("Package", self._part.package).is_signed

    @property
    def signatures(self) -> list[SignatureInfo]:
        """List of |SignatureInfo| for each digital signature in the package.

        Empty list when the document is unsigned. See :class:`docx.signatures.SignatureInfo`
        for the available metadata.

        .. versionadded:: 2026.05.0
        """
        from docx.package import Package

        return cast("Package", self._part.package).signatures

    @property
    def font_table(self) -> FontTable | None:
        """A |FontTable| collection, or |None| if no font-table part is related.

        Returns |None| when the document has no ``fontTable`` relationship. For
        authoring workflows that need to embed a font use
        :attr:`font_table_or_new` instead, which materialises an empty
        ``fontTable.xml`` on demand.

        .. versionadded:: 2026.05.0
        """
        return self._part.font_table

    @property
    def font_table_or_new(self) -> FontTable:
        """A |FontTable| collection, creating an empty ``fontTable.xml`` if needed.

        Use this when the caller intends to *add* to the font table (e.g. via
        :meth:`FontTable.add_embedded_font`) — unlike :attr:`font_table` it
        never returns |None|.

        .. versionadded:: 2026.05.0
        """
        return self._part.font_table_or_new

    @property
    def footnotes(self) -> Footnotes:
        """A |Footnotes| object providing access to footnotes in the document.

        .. versionadded:: 2026.05.0
        """
        return self._part.footnotes

    def accept_all_changes(self) -> int:
        """Accept every tracked change in the document body.

        Insertions are flattened into live content, deletions are removed, and any
        `w:rPrChange`, `w:pPrChange`, or `w:sectPrChange` elements are discarded
        (the current, post-edit formatting is retained).

        Returns the number of change elements resolved.

        .. versionadded:: 2026.05.0
        """
        from docx.tracked_changes import _resolve_all_changes

        return _resolve_all_changes(self._element.body, accept=True)

    def reject_all_changes(self) -> int:
        """Reject every tracked change in the document body.

        Insertions are removed, deletions are restored as live content, and any
        `w:rPrChange`, `w:pPrChange`, or `w:sectPrChange` elements are unwound so
        the prior formatting is restored.

        Returns the number of change elements resolved.

        .. versionadded:: 2026.05.0
        """
        from docx.tracked_changes import _resolve_all_changes

        return _resolve_all_changes(self._element.body, accept=False)

    @property
    def footnote_properties(self) -> FootnoteProperties | None:
        """Document-level |FootnoteProperties| or |None| if not configured.

        Returns |None| when no ``w:footnotePr`` element exists in the document settings.
        Use :meth:`add_footnote_properties` to add one and configure it.

        .. versionadded:: 2026.05.0
        """
        return self.settings.footnote_properties

    def add_footnote_properties(self) -> FootnoteProperties:
        """Return document-level |FootnoteProperties|, adding a ``w:footnotePr`` if needed.

        .. versionadded:: 2026.05.0
        """
        return self.settings.add_footnote_properties()

    @property
    def endnote_properties(self) -> EndnoteProperties | None:
        """Document-level |EndnoteProperties| or |None| if not configured.

        Returns |None| when no ``w:endnotePr`` element exists in the document settings.
        Use :meth:`add_endnote_properties` to add one and configure it.

        .. versionadded:: 2026.05.0
        """
        return self.settings.endnote_properties

    def add_endnote_properties(self) -> EndnoteProperties:
        """Return document-level |EndnoteProperties|, adding a ``w:endnotePr`` if needed.

        .. versionadded:: 2026.05.0
        """
        return self.settings.add_endnote_properties()

    @property
    def core_properties(self):
        """A |CoreProperties| object providing Dublin Core properties of document."""
        return self._part.core_properties

    @property
    def custom_properties(self) -> CustomProperties:
        """A |CustomProperties| collection providing access to custom document properties.

        Custom properties are user-defined, typed name/value pairs stored in the
        ``docProps/custom.xml`` part of the package. They are distinct from the fixed
        "core" Dublin-Core properties available via :attr:`core_properties`.

        .. versionadded:: 2026.05.0
        """
        return self._part.custom_properties

    def set_language(
        self,
        latin: str | None,
        east_asian: str | None = None,
        bidi: str | None = None,
    ) -> None:
        """Set the document-level theme-font language tags.

        Convenience wrapper for assigning
        :attr:`Settings.theme_font_language`. Pass ``latin`` as a BCP-47
        language tag (e.g. ``"en-US"``); the East-Asian and bidi tags are
        optional and default to |None| (meaning leave unset).

        .. versionadded:: 2026.05.0
        """
        self.settings.theme_font_language = (latin, east_asian, bidi)

    @property
    def extended_properties(self) -> ExtendedProperties:
        """An |ExtendedProperties| proxy for this document's ``docProps/app.xml``.

        Exposes application-written metadata such as ``Company``, ``Manager``,
        ``Application``, ``AppVersion``, ``TotalTime``, and the cached
        ``Pages`` / ``Words`` / ``Characters`` statistics. A default (empty)
        extended-properties part is created on demand when none is present.

        .. versionadded:: 2026.05.0
        """
        return self._part.extended_properties

    @property
    def custom_xml_parts(self) -> list[CustomXmlPart]:
        """List of |CustomXmlPart| proxies for the custom XML data parts in the package.

        Empty when the document has no ``customXml`` relationships. Each entry
        exposes the data part's :attr:`~docx.custom_xml.CustomXmlPart.item_id`,
        :attr:`~docx.custom_xml.CustomXmlPart.schema_refs`,
        :attr:`~docx.custom_xml.CustomXmlPart.root_element`, and
        :attr:`~docx.custom_xml.CustomXmlPart.blob` read-only.

        .. versionadded:: 2026.05.0
        """
        return self._part.custom_xml_parts

    @property
    def numbering(self):
        """A |Numbering| object providing read/write access to the list-style
        numbering definitions for this document.

        Creates a default (empty) numbering part if one is not already related to the
        document.

        .. versionadded:: 2026.05.0
        """
        return self._part.numbering_part.numbering

    def list_labels(self) -> dict[int, str]:
        """Return ``{id(p_element): label}`` for every numbered paragraph in the body.

        Walks the document body top-to-bottom exactly once and, for each
        paragraph that resolves to a list (via a direct ``w:numPr`` or a
        paragraph style that declares one), computes the Word-rendered label
        (``"1."``, ``"a)"``, ``"I."``, ``"•"``, ``"1.1."``, ...) using the
        level's ``w:lvlText`` pattern and ``w:numFmt`` value.
        Paragraphs that are not part of any list are omitted.
        The mapping key is ``id(paragraph._p)`` — stable for the lifetime
        of the underlying element. To look up by |Paragraph| use
        ``labels[id(paragraph._p)]`` or use :attr:`Paragraph.list_label`.
        Returns an empty mapping when the document has no numbering part or
        no paragraph resolves to a list.

        Supported ``numFmt`` values: ``decimal``, ``decimalZero``,
        ``upperRoman``, ``lowerRoman``, ``upperLetter``, ``lowerLetter``,
        ``bullet``. Other formats (``cardinalText``, ``ordinalText``, ...)
        fall back to decimal.

        .. versionadded:: 2026.05.0
        """
        from docx.numbering import ListLabelRenderer

        numbering_part = getattr(self._part, "numbering_part", None)
        numbering_elm = (
            numbering_part.numbering_element if numbering_part is not None else None
        )
        styles_elm = None
        try:
            styles_part = self._part.part_related_by(RT.STYLES)
        except (KeyError, AttributeError):
            styles_part = None
        if styles_part is not None:
            styles_elm = getattr(styles_part, "element", None)

        renderer = ListLabelRenderer(numbering_elm, styles_elm)
        return renderer.label_map(self._element.body.xpath(".//w:p"))

    @property
    def ink_annotations(self) -> list[InkAnnotation]:
        """List of |InkAnnotation| objects for each ink annotation in the body.

        An ink annotation is any ``w:contentPart`` element that targets an ink part
        (content type ``application/inkml+xml``). The list is empty when no ink
        annotations are present. Read-only — python-docx does not support creating
        or modifying ink annotations.

        .. versionadded:: 2026.05.0
        """
        from docx.text.paragraph import Paragraph

        result: list[InkAnnotation] = []
        for p in self._element.body.xpath(".//w:p[.//w:contentPart]"):
            paragraph = Paragraph(p, self._body)
            result.extend(paragraph.ink_annotations)
        return result

    @property
    def attachments(self) -> list[Attachment]:
        """List of |Attachment| for each ``w:altChunk`` in the document body.

        An ``altChunk`` is an arbitrary foreign payload (HTML, RTF, another
        docx, etc.) that Word merges into the document on open. python-docx
        exposes them read-only: callers can iterate the altChunk elements,
        inspect their content-type (via the related part), and retrieve the
        raw payload bytes for further processing.

        Returns an empty list when the document has no altChunks.

        .. versionadded:: 2026.05.0
        """
        from docx.attachments import Attachment

        result: list[Attachment] = []
        doc_part = self._part
        for alt_elm in self._element.body.xpath(".//w:altChunk"):
            rId = alt_elm.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            target_part = None
            if rId is not None:
                target_part = doc_part.related_parts.get(rId)
            result.append(Attachment(alt_elm, target_part))
        return result

    @property
    def embedded_objects(self) -> list[EmbeddedObject]:
        """List of |EmbeddedObject| for each embedded OLE object in the body.

        An embedded object is any ``w:object`` element containing an
        ``o:OLEObject`` descendant (content type
        ``application/vnd.openxmlformats-officedocument.oleObject`` for the
        related binary). The list is empty when no embedded objects are
        present. Read-only — python-docx does not support creating or
        modifying embedded objects, or extracting the ``w:pict`` image that
        Word displays in place of the OLE content.

        .. versionadded:: 2026.05.0
        """
        from docx.text.paragraph import Paragraph

        result: list[EmbeddedObject] = []
        for p in self._element.body.xpath(".//w:p[.//w:object/o:OLEObject]"):
            paragraph = Paragraph(p, self._body)
            result.extend(paragraph.embedded_objects)
        return result

    @property
    def smart_art(self) -> list[SmartArt]:
        """List of |SmartArt| proxies for every SmartArt diagram in the body.

        Walks top-level body paragraphs (including paragraphs nested inside
        body-level tables) and returns one entry for each ``w:drawing`` that
        references a SmartArt diagram (i.e. contains a ``dgm:relIds``
        element). Empty list when the document has no SmartArt. Read-only —
        python-docx does not support creating or modifying SmartArt.

        .. versionadded:: 2026.05.0
        """
        from docx.drawing import Drawing

        result: list[SmartArt] = []
        for d in self._element.body.xpath(".//w:drawing"):
            drawing = Drawing(d, self._body)
            sa = drawing.smart_art
            if sa is not None:
                result.append(sa)
        return result

    @property
    def inline_shapes(self):
        """The |InlineShapes| collection for this document.

        An inline shape is a graphical object, such as a picture, contained in a run of
        text and behaving like a character glyph, being flowed like other text in a
        paragraph.
        """
        return self._part.inline_shapes

    def iter_inner_content(
        self, include_sdt_flat: bool = False
    ) -> Iterator[Paragraph | Table]:
        """Generate each `Paragraph` or `Table` in this document in document order.

        When `include_sdt_flat` is |True|, any block-level ``w:sdt`` wrapper
        found in the body is flattened: the paragraphs and tables inside the
        sdt's ``w:sdtContent`` are yielded in place, as if the wrapper were
        transparent. This surfaces paragraphs that live inside content
        controls (including a TOC's own ``w:sdt`` wrapper), which the default
        iteration does not reach — closes upstream#1280.

        .. versionchanged:: 2026.05.0
            Added ``include_sdt_flat`` parameter.
        """
        if not include_sdt_flat:
            return self._body.iter_inner_content()
        return self._body._iter_inner_content_flat_sdt()

    @property
    def text(self) -> str:
        """Concatenated text of every top-level paragraph in the document body.

        Paragraphs are joined with a single ``"\\n"`` separator; tables in the
        body are skipped. This is the quick "give me the body text" helper
        requested in upstream#252 / upstream#72. For a breakdown that walks
        tables or non-body stories, iterate :attr:`paragraphs` /
        :attr:`tables` manually.

        .. versionadded:: 2026.05.0
        """
        return "\n".join(p.text for p in self.paragraphs)

    @property
    def recovery_warnings(self) -> list[str]:
        """List of parse-warning strings collected while opening this document.

        Populated only when the document was opened via :func:`docx.Document`
        with ``recover=True``. Empty for normally-opened documents and for
        well-formed documents opened in recovery mode.

        .. versionadded:: 2026.05.0
        """
        from docx.package import Package

        return cast("Package", self._part.package).recovery_warnings

    @property
    def paragraphs(self) -> list[Paragraph]:
        """The |Paragraph| instances in the document, in document order.

        Note that paragraphs within revision marks such as ``<w:ins>`` or ``<w:del>`` do
        not appear in this list.
        """
        return self._body.paragraphs

    @property
    def part(self) -> DocumentPart:
        """The |DocumentPart| object of this document."""
        return self._part

    def replace(
        self,
        old_text: str,
        new_text: str,
        case_sensitive: bool = True,
        whole_word: bool = False,
    ) -> int:
        """Replace occurrences of `old_text` with `new_text` in the document body paragraphs.

        Note: Only top-level body paragraphs are searched. Text inside table cells,
        headers, footers, footnotes, and endnotes is not affected.

        Preserves the run formatting of the first character's run for each replacement.
        Returns the number of replacements made.

        When `case_sensitive` is False, matching is case-insensitive. When `whole_word` is
        True, only whole-word matches are replaced.

        .. versionadded:: 2026.05.0
        """
        from docx.search import replace_in_paragraphs

        return replace_in_paragraphs(
            self.paragraphs, old_text, new_text, case_sensitive, whole_word
        )

    def replace_all(
        self,
        old_text: str,
        new_text: str,
        case_sensitive: bool = True,
        whole_word: bool = False,
    ) -> int:
        """Replace `old_text` with `new_text` in every story in this document.

        Unlike :meth:`replace`, which updates only top-level body paragraphs, this
        method walks every "story" in the package — the body (including top-level
        body tables), each section's non-inherited headers and footers, footnotes,
        endnotes, and comments — and applies the replacement to each.

        Paragraphs nested inside tables that live within a header, footer,
        footnote, endnote, or comment story are not descended into, and neither
        are tables nested inside body-level table cells; see
        :func:`docx.search._iter_all_paragraphs` for the full iteration contract.

        Returns the total number of replacements made across all stories.

        .. versionadded:: 2026.05.0
        """
        from docx.search import replace_in_all_paragraphs

        return replace_in_all_paragraphs(
            self, old_text, new_text, case_sensitive, whole_word
        )

    def replace_regex(
        self,
        pattern: str | re.Pattern[str],
        replacement: str,
        flags: int = 0,
    ) -> int:
        """Replace all regex matches of `pattern` with `replacement` in body paragraphs.

        `pattern` may be a string or a compiled `re.Pattern`. When `pattern` is a string,
        `flags` (e.g. `re.IGNORECASE`) is applied during compilation; when `pattern` is
        already compiled, `flags` is ignored. `replacement` follows `re.sub` semantics —
        backreferences such as ``\\1`` and ``\\g<name>`` are expanded per match.

        Note: Only top-level body paragraphs are processed. Text inside table cells,
        headers, footers, footnotes, and endnotes is not affected.

        Preserves the run formatting of the first character's run for each replacement.
        Returns the number of replacements made.

        .. versionadded:: 2026.05.0
        """
        from docx.search import replace_in_paragraphs_regex

        return replace_in_paragraphs_regex(
            self.paragraphs, pattern, replacement, flags
        )

    def replace_regex_all(
        self,
        pattern: str | re.Pattern[str],
        replacement: str,
        flags: int = 0,
    ) -> int:
        """Replace regex `pattern` with `replacement` across every story in this document.

        Like :meth:`replace_all` but using regex semantics. ``replacement`` follows
        :func:`re.sub` semantics — backreferences such as ``\\1`` and ``\\g<name>``
        are expanded per match.

        Returns the total number of replacements made across all stories.

        .. versionadded:: 2026.05.0
        """
        from docx.search import replace_in_all_paragraphs_regex

        return replace_in_all_paragraphs_regex(self, pattern, replacement, flags)

    def resolve_cross_references(self) -> int:
        """Resolve ``REF`` and ``PAGEREF`` fields in the document body.

        For each ``REF`` field whose bookmark exists, replaces the field's
        rendered :attr:`~docx.fields.Field.result_text` with the concatenated
        text of the referenced bookmark range. For each ``PAGEREF`` field
        whose cached result is empty, replaces it with ``"?"``; python-docx
        cannot compute real page numbers because it has no layout engine.

        Other field types (``PAGE``, ``DATE``, ``SEQ``, etc.) are ignored.
        Fields whose bookmark cannot be found are left unchanged. Returns the
        number of fields whose ``result_text`` was updated in place.

        Walks every field element in the body — including those inside
        tables, hyperlinks, and other containers — by descending all
        ``w:fldSimple`` elements and every run that opens a complex field.

        .. versionadded:: 2026.05.0
        """
        from docx.fields import Field

        body = self._element.body
        updated = 0
        for el in body.xpath(
            ".//w:fldSimple | .//w:r[w:fldChar[@w:fldCharType='begin']]"
        ):
            tag = el.tag.rsplit("}", 1)[-1]
            field = (
                Field.for_simple(el) if tag == "fldSimple" else Field.for_complex(el)
            )
            if field.type not in (
                "REF",
                "PAGEREF",
                "DOCPROPERTY",
                "AUTHOR",
                "TITLE",
                "SUBJECT",
                "KEYWORDS",
                "COMMENTS",
                "LASTSAVEDBY",
            ):
                continue
            resolved = field.resolve(self)
            if resolved == field.result_text:
                continue
            field.update_result_text(resolved)
            updated += 1
        return updated

    def revision_marks_text(
        self,
        open_ins: str = "[+",
        close_ins: str = "+]",
        open_del: str = "[-",
        close_del: str = "-]",
    ) -> str:
        """Return the document body's text with tracked-change markers applied.

        Each top-level paragraph is rendered via
        :meth:`Paragraph.revision_marks_text` and the results are joined with a
        blank line separator (``"\\n\\n"``). Tables in the body are skipped; this
        helper is intended as a quick CLI preview of inline insertions and
        deletions in the running prose.

        Pass ANSI escape sequences in place of the default square-bracket markers
        if you want styled terminal output.

        .. versionadded:: 2026.05.0
        """
        return "\n\n".join(
            p.revision_marks_text(open_ins, close_ins, open_del, close_del)
            for p in self.paragraphs
        )

    def append_document(self, other: Document) -> int:
        """Append the body of `other` to this document and return the number of
        block elements copied.

        Paragraphs, tables, and block-level SDT (structured document tag) elements
        from ``other`` are deep-copied into this document's body, inserted before
        the current section's sentinel ``w:sectPr`` so the destination's page
        setup is preserved. Relationships carried by the copied content —
        images, embedded objects, hyperlinks, charts, etc. — are imported into
        this document's package and rewritten to point at the new destination
        rIds. Paragraph / run styles referenced by the copied content are
        likewise copied (plus any ``basedOn`` / ``next`` / ``link`` dependencies).
        List numbering definitions referenced by the copied content are cloned
        under fresh numIds.

        Closes upstream#1457, upstream#558, upstream#543, upstream#437,
        upstream#460, upstream#44, upstream#709.

        .. versionadded:: 2026.05.0
        """
        from docx.append_document import append_document

        return append_document(self, other)

    def append_body(self, other: Document) -> int:
        """Alias for :meth:`append_document`.

        Provided as a second entry-point for users who think of the operation
        as "append the body" rather than "append the document".

        .. versionadded:: 2026.05.0
        """
        from docx.append_document import append_body

        return append_body(self, other)

    def append_paragraph(self, paragraph: Paragraph) -> Paragraph:
        """Copy `paragraph` from its owning document into this one and return the new paragraph.

        Any relationships referenced by the paragraph (images, hyperlinks,
        embedded objects) and any style / numbering references it carries are
        imported into this document the same way as for :meth:`append_document`.

        .. versionadded:: 2026.05.0
        """
        from docx.append_document import append_paragraph

        return append_paragraph(self, paragraph)

    def save(
        self,
        path_or_stream: str | IO[bytes],
        flat_opc: bool = False,
        reproducible: bool = False,
    ):
        """Save this document to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.

        When `path_or_stream` is a string, the filename component (the last path
        segment) is validated against the set of characters Windows disallows in
        file names (``< > : " | ? *``). If one of those characters is present,
        an :class:`OSError` is raised rather than writing a silently-empty or
        mis-named file (closes upstream#1111). The rest of the path — including
        drive-letter colons and forward/backward directory separators — is left
        to the underlying file system.

        When `flat_opc` is True, the document is serialised as Flat-OPC — the
        ``<pkg:package>`` single-XML-file representation defined in ECMA-376
        Part 2 — rather than a zip package. Closes upstream#892.

        When `reproducible` is True, the emitted zip archive uses a fixed
        timestamp for every member and writes members in sorted order, so
        repeated saves of the same content produce byte-identical output.
        This is the single bit of plumbing that closes upstream#1042 and
        upstream-PR#810.

        .. versionadded:: 2026.05.0
           The `flat_opc` and `reproducible` parameters.
        """
        if flat_opc:
            import io as _io

            from docx.opc.flat_opc import write_flat_opc

            buf = _io.BytesIO()
            self._part.save(buf, reproducible=reproducible)
            write_flat_opc(path_or_stream, buf.getvalue())
            return
        self._part.save(path_or_stream, reproducible=reproducible)

    def search(
        self,
        text: str,
        case_sensitive: bool = True,
        whole_word: bool = False,
    ) -> list[SearchMatch]:
        """Find all occurrences of `text` in the document body paragraphs.

        Note: Only top-level body paragraphs are searched. Text inside table cells,
        headers, footers, footnotes, and endnotes is not included.

        Returns a list of |SearchMatch| objects, one for each occurrence found. Each match
        provides access to the paragraph, run indices, and character offsets.

        When `case_sensitive` is False, matching is case-insensitive. When `whole_word` is
        True, only whole-word matches are returned.

        .. versionadded:: 2026.05.0
        """
        from docx.search import search_paragraphs

        return search_paragraphs(self.paragraphs, text, case_sensitive, whole_word)

    def search_all(
        self,
        text: str,
        case_sensitive: bool = True,
        whole_word: bool = False,
    ) -> list[SearchMatch]:
        """Find `text` in every story across this document.

        Unlike :meth:`search`, which only looks at top-level body paragraphs, this
        walks all document "stories" — the body and its top-level tables, each
        section's non-inherited headers and footers, footnotes, endnotes, and
        comments — and returns a |SearchMatch| for every hit. Each match's
        :attr:`SearchMatch.location` identifies which story produced it (e.g.
        ``"body"``, ``"table:0:row:1:col:2"``, ``"header:section0:primary"``,
        ``"footnote:2"``, ``"endnote:3"``, or ``"comment:5"``).

        Tables nested inside other stories (headers, footers, footnotes, etc.) and
        tables nested inside body-table cells are not descended into; see
        :func:`docx.search._iter_all_paragraphs` for details.

        .. versionadded:: 2026.05.0
        """
        from docx.search import search_all_paragraphs

        return search_all_paragraphs(self, text, case_sensitive, whole_word)

    def search_regex(
        self,
        pattern: str | re.Pattern[str],
        flags: int = 0,
    ) -> list[SearchMatch]:
        """Find all regex matches of `pattern` in the document body paragraphs.

        `pattern` may be a string or a compiled `re.Pattern`. When `pattern` is a string,
        `flags` (e.g. `re.IGNORECASE`) is applied during compilation; when `pattern` is
        already compiled, `flags` is ignored.

        Note: Only top-level body paragraphs are searched. Text inside table cells,
        headers, footers, footnotes, and endnotes is not included.

        Returns a list of |SearchMatch| objects, one for each match found. Each match
        provides access to the paragraph, run indices, and character offsets.

        .. versionadded:: 2026.05.0
        """
        from docx.search import search_paragraphs_regex

        return search_paragraphs_regex(self.paragraphs, pattern, flags)

    def search_regex_all(
        self,
        pattern: str | re.Pattern[str],
        flags: int = 0,
    ) -> list[SearchMatch]:
        """Find regex matches of `pattern` in every story in this document.

        Like :meth:`search_all` but using regex semantics. ``pattern`` may be a
        string or a compiled :class:`re.Pattern`. Each returned |SearchMatch|
        carries a :attr:`SearchMatch.location` identifying the story that
        produced it.

        .. versionadded:: 2026.05.0
        """
        from docx.search import search_all_paragraphs_regex

        return search_all_paragraphs_regex(self, pattern, flags)

    @property
    def sections(self) -> Sections:
        """|Sections| object providing access to each section in this document."""
        return Sections(self._element, self._part)

    def validate_heading_structure(self) -> list[HeadingIssue]:
        """Return a list of |HeadingIssue| objects describing heading problems.

        Scans the body paragraphs for common accessibility issues in the heading
        outline: skipped levels (e.g. a Heading 3 directly following a Heading 1),
        multiple Heading 1 paragraphs, empty heading paragraphs, and documents that
        start below Heading 1. Returns an empty list when no issues are found.

        .. versionadded:: 2026.05.0
        """
        from docx.accessibility import validate_heading_structure

        return validate_heading_structure(self.paragraphs)

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the document-level settings."""
        return self._part.settings

    @property
    def statistics(self) -> DocumentStatistics:
        """A |DocumentStatistics| summarizing the document body's text.

        Provides counts of non-empty paragraphs, words, characters (including
        spaces), and characters excluding spaces for the main document story
        (the ``w:body``). Text in headers, footers, footnotes, endnotes, and
        comments is not included, mirroring Word's default "Word Count"
        behavior.

        A "word" is a whitespace-delimited token, matching ``str.split()``
        semantics. Paragraphs nested inside tables or block-level content
        controls are included because they are part of the body story.

        :attr:`DocumentStatistics.pages` is populated from the cached
        ``<Pages>`` value in the extended-properties part (``docProps/app.xml``)
        when present, otherwise |None| -- python-docx does not lay the document
        out so it cannot compute a true page count.

        .. versionadded:: 2026.05.0
        """
        from docx.statistics import compute_statistics

        pages: int | None = None
        try:
            raw_pages = self.extended_properties.pages
        except Exception:
            # -- if the extended-properties part is missing or malformed,
            # -- fall back silently to a |None| page count rather than
            # -- propagating the error into the caller's statistics call.
            raw_pages = None
        # -- only accept genuine integer / None values; anything else (e.g.
        # -- a Mock surfaced by test doubles) is treated as "unknown" --
        if isinstance(raw_pages, int) and not isinstance(raw_pages, bool):
            pages = raw_pages
        else:
            pages = None

        return compute_statistics(self._element.body, pages=pages)

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in this document."""
        return self._part.styles

    @property
    def glossary(self) -> Glossary | None:
        """A |Glossary| proxy, or |None| when no ``glossaryDocument`` part is related.

        The glossary-document part carries the AutoText / Quick Parts /
        cover-page building blocks. It is owned by Word, so python-docx
        exposes it read-only. Returns |None| when the document has no
        ``glossaryDocument`` relationship — which is the case for
        documents created via :func:`docx.Document` with the default
        template.

        .. versionadded:: 2026.05.0
        """
        return self._part.glossary

    @property
    def theme(self) -> Theme | None:
        """A |Theme| proxy, or |None| when no ``theme`` part is related.

        The theme part is owned by Word, so python-docx exposes it read-only.
        Returns |None| when the document has no ``theme`` relationship — which
        is uncommon for documents created by Word but possible for minimal
        documents synthesized by other tools.

        .. versionadded:: 2026.05.0
        """
        return self._part.theme

    @property
    def web_settings(self) -> WebSettings | None:
        """A |WebSettings| proxy, or |None| when no ``webSettings`` part is related.

        The web-settings part is owned by Word, so python-docx exposes it
        read-oriented. Returns |None| when the document has no ``webSettings``
        relationship — for example, documents created via :func:`docx.Document`
        with no template.

        .. versionadded:: 2026.05.0
        """
        return self._part.web_settings

    @property
    def tables(self) -> list[Table]:
        """All |Table| instances in the document, in document order.

        Note that only tables appearing at the top level of the document appear in this
        list; a table nested inside a table cell does not appear. A table within
        revision marks such as ``<w:ins>`` or ``<w:del>`` will also not appear in the
        list.
        """
        return self._body.tables

    @property
    def _block_width(self) -> Length:
        """A |Length| object specifying the space between margins in last section.

        Falls back to the US-Letter default (8.5" page width, 1" margins — a
        6.5" usable block) when the document body contains no ``w:sectPr``,
        which some third-party generators emit. See upstream#514.
        """
        sections = self.sections
        if len(sections) == 0:
            # -- no sectPr present: use standard US-Letter defaults --
            return Emu(Inches(8.5) - Inches(1) - Inches(1))
        section = sections[-1]
        page_width = section.page_width or Inches(8.5)
        left_margin = section.left_margin or Inches(1)
        right_margin = section.right_margin or Inches(1)
        return Emu(page_width - left_margin - right_margin)

    @property
    def _body(self) -> _Body:
        """The |_Body| instance containing the content for this document."""
        if self.__body is None:
            self.__body = _Body(self._element.body, self)
        return self.__body


class _Body(BlockItemContainer):
    """Proxy for `<w:body>` element in this document.

    It's primary role is a container for document content.
    """

    def __init__(self, body_elm: CT_Body, parent: t.ProvidesStoryPart):
        super().__init__(body_elm, parent)
        self._body = body_elm

    def _iter_inner_content_flat_sdt(self) -> Iterator[Paragraph | Table]:
        """Yield `Paragraph`/`Table` walking into `w:sdt/w:sdtContent` wrappers.

        Block-level ``w:sdt`` elements have a ``w:sdtContent`` child holding
        the wrapped paragraphs and tables; descend into it so those items
        surface in the iteration. Non-block-level `w:sdt` elements (inline
        inside a paragraph) are not reached here because they are not direct
        children of the body.
        """
        from docx.oxml.ns import qn
        from docx.oxml.text.paragraph import CT_P as _CT_P
        from docx.oxml.table import CT_Tbl as _CT_Tbl
        from docx.table import Table as _Table
        from docx.text.paragraph import Paragraph as _Paragraph

        for child in self._body:
            tag = child.tag
            if isinstance(child, _CT_P):
                yield _Paragraph(child, self)
            elif isinstance(child, _CT_Tbl):
                yield _Table(child, self)
            elif tag == qn("w:sdt"):
                for sub in child:
                    if sub.tag != qn("w:sdtContent"):
                        continue
                    for inner in sub:
                        if isinstance(inner, _CT_P):
                            yield _Paragraph(inner, self)
                        elif isinstance(inner, _CT_Tbl):
                            yield _Table(inner, self)

    def add_content_control(
        self,
        type: ContentControlType,
        tag: str | None = None,
        title: str | None = None,
    ) -> ContentControl:
        """Add a block-level content control at the end of the body.

        The new `w:sdt` is inserted before any trailing `w:sectPr` element, mirroring
        how paragraphs and tables are appended.

        .. versionadded:: 2026.05.0
        """
        from docx.content_controls import ContentControl, new_sdt

        sdt = new_sdt(type, tag=tag, title=title, inline=False)
        # -- insert before trailing sectPr, if present --
        self._body._insert_sdt(sdt)  # pyright: ignore[reportPrivateUsage]
        return ContentControl(sdt)

    def clear_content(self) -> _Body:
        """Return this |_Body| instance after clearing it of all content.

        Section properties for the main document story, if present, are preserved.
        """
        self._body.clear_content()
        return self

    @property
    def content_controls(self) -> list[ContentControl]:
        """List of block-level |ContentControl| objects in this body, in document order.

        .. versionadded:: 2026.05.0
        """
        from docx.content_controls import ContentControl

        return [
            ContentControl(cast("CT_Sdt", sdt)) for sdt in self._body.xpath("./w:sdt")
        ]
