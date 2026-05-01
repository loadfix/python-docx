"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, cast
from collections.abc import Iterator

from docx.drawing import Drawing
from docx.enum.section import WD_SECTION_START
from docx.enum.shape import WD_ANCHOR_H, WD_ANCHOR_V, WD_WRAP_TYPE
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.fields import Field
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.drawing import CT_Drawing
from docx.oxml.shape import CT_Anchor
from docx.oxml.table import CT_Tbl
from docx.oxml.text.run import CT_R
from docx.shape import FloatingImage
from docx.shared import Inches, StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.tracked_changes import TrackedChange
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.bookmarks import Bookmark
    from docx.content_controls import ContentControl, ContentControlType
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.content_controls import CT_Sdt
    from docx.oxml.document import CT_Body
    from docx.oxml.text.paragraph import CT_P
    from docx.section import Section
    from docx.shared import Length
    from docx.styles.style import CharacterStyle
    from docx.table import Table as _Table
    from docx.styles.style import _TableStyle  # pyright: ignore[reportPrivateUsage]


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._p = self._element = p

    def add_bookmark(
        self,
        name: str,
        start_run: Run | None = None,
        end_run: Run | None = None,
    ) -> Bookmark:
        """Add a bookmark to this paragraph and return it.

        `name` is the bookmark name, which must be unique within the document.

        When `start_run` and `end_run` are both |None|, the bookmark wraps the entire
        paragraph content. When `start_run` is provided, the bookmark starts before that
        run. When `end_run` is provided, the bookmark ends after that run. When only
        `start_run` is provided, `end_run` defaults to `start_run`.
        """
        from docx.bookmarks import Bookmark

        body = self._get_body()
        bookmark_id = self._next_bookmark_id(body)

        if start_run is None and end_run is None:
            self._p.add_bookmark(bookmark_id, name)
        else:
            if start_run is None:
                start_run = end_run
            if end_run is None:
                end_run = start_run
            assert start_run is not None
            assert end_run is not None
            start_run._r.insert_bookmark_start_before(bookmark_id, name)
            end_run._r.insert_bookmark_end_after(bookmark_id)

        bookmarkStart = self._p.xpath(f".//w:bookmarkStart[@w:id='{bookmark_id}']")
        return Bookmark(bookmarkStart[0], body)

    def _get_body(self) -> CT_Body:
        """Return the w:body ancestor element."""
        from docx.oxml.document import CT_Body

        ancestor = self._p.getparent()
        while ancestor is not None and not isinstance(ancestor, CT_Body):
            ancestor = ancestor.getparent()
        if ancestor is None:
            raise ValueError("paragraph is not contained in a document body")
        return ancestor

    @staticmethod
    def _next_bookmark_id(body) -> int:
        """Return the next available bookmark ID in the document body."""
        used_ids = [int(x) for x in body.xpath(".//w:bookmarkStart/@w:id")]
        return max(used_ids, default=-1) + 1

    def add_hyperlink(
        self,
        url: str | None = None,
        text: str | None = None,
        style: str | CharacterStyle | None = "Hyperlink",
        anchor: str | None = None,
    ) -> Hyperlink:
        """Append a hyperlink to this paragraph and return a |Hyperlink| object.

        `url` is the target URL for an external hyperlink (e.g. "https://example.com").
        `text` is the visible link text; defaults to `url` or `anchor` when not provided.
        `style` is the character style for the hyperlink run, defaulting to "Hyperlink".
        `anchor` is a bookmark name for an internal document link.

        Either `url` or `anchor` must be provided, but not both.
        """
        if url is None and anchor is None:
            raise ValueError("Either url or anchor must be provided")
        if url is not None and anchor is not None:
            raise ValueError("Only one of url or anchor may be provided, not both")

        display_text = text if text is not None else (url or anchor or "")

        rId = None
        if url is not None:
            rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)

        rPr = None
        if style is not None:
            from docx.oxml.ns import qn
            from docx.oxml.parser import OxmlElement

            style_id = self.part.get_style_id(style, WD_STYLE_TYPE.CHARACTER)
            if style_id is not None:
                rPr = OxmlElement("w:rPr")
                rStyle = OxmlElement("w:rStyle")
                rStyle.set(qn("w:val"), style_id)
                rPr.append(rStyle)

        hyperlink_elm = self._p.add_hyperlink(rId, anchor, display_text, rPr)
        return Hyperlink(hyperlink_elm, self)

    def add_run(self, text: str | None = None, style: str | CharacterStyle | None = None) -> Run:
        """Append run containing `text` and having character-style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break. When `text` is `None`, the new run is empty.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    def add_content_control(
        self,
        type: ContentControlType,
        tag: str | None = None,
        title: str | None = None,
    ) -> ContentControl:
        """Append an inline content control (structured document tag) to this paragraph.

        `type` is a :class:`ContentControlType` member. `tag` becomes the programmatic
        `w:sdtPr/w:tag/@w:val` value, and `title` becomes `w:sdtPr/w:alias/@w:val`.
        Returns the newly appended |ContentControl|.
        """
        from docx.content_controls import ContentControl, new_sdt

        sdt = new_sdt(type, tag=tag, title=title, inline=True)
        self._p.append(sdt)
        return ContentControl(sdt)

    def add_page_break(self) -> Paragraph:
        """Append a page-break run to this paragraph and return self."""
        run = self.add_run()
        run.add_break(WD_BREAK.PAGE)
        return self

    def add_simple_field(self, instr: str, text: str | None = None) -> Field:
        """Append a ``<w:fldSimple>`` field to this paragraph and return a |Field|.

        `instr` is the field instruction (e.g. ``"PAGE"`` or ``"REF bookmark1 \\h"``).
        `text` is the optional current rendered result, added as a single run
        inside the fldSimple element.
        """
        fldSimple = self._p.add_fldSimple(instr, text)
        return Field.for_simple(fldSimple)

    def add_complex_field(self, instr: str, result_text: str | None = None) -> Field:
        """Append a complex field (begin/separate/end) to this paragraph.

        Returns a |Field| wrapping the run that contains the ``begin``
        ``<w:fldChar>`` marker. `instr` is the field instruction (e.g.
        ``"PAGE"``) and `result_text`, if provided, is added as a plain
        ``<w:r><w:t>`` run between the ``separate`` and ``end`` markers.
        """
        begin_run = self._p.add_complex_field(instr, result_text)
        return Field.for_complex(begin_run)

    def add_floating_image(
        self,
        image_path_or_stream: str | IO[bytes],
        width: int | Length | None = None,
        height: int | Length | None = None,
        position: dict | None = None,
    ) -> FloatingImage:
        """Add a floating (anchored) image to this paragraph and return it.

        `image_path_or_stream` is a path (str) or binary file-like object for the image.
        `width` and `height` work the same way as for `add_picture`.

        `position` is an optional dict that may contain any of these keys:
        - `horizontal`: horizontal offset (int EMU or |Length|)
        - `vertical`: vertical offset (int EMU or |Length|)
        - `h_anchor`: |WD_ANCHOR_H| member (defaults to `COLUMN`)
        - `v_anchor`: |WD_ANCHOR_V| member (defaults to `PARAGRAPH`)
        - `wrap`: |WD_WRAP_TYPE| member (defaults to `SQUARE`)
        """
        anchor = self.part.new_pic_anchor(image_path_or_stream, width, height)

        # -- apply optional positioning overrides --
        if position is not None:
            h_anchor = position.get("h_anchor", WD_ANCHOR_H.COLUMN)
            v_anchor = position.get("v_anchor", WD_ANCHOR_V.PARAGRAPH)
            horizontal = position.get("horizontal", 0)
            vertical = position.get("vertical", 0)
            wrap = position.get("wrap", WD_WRAP_TYPE.SQUARE)

            if isinstance(h_anchor, WD_ANCHOR_H):
                h_anchor_value = h_anchor.value
            else:
                h_anchor_value = str(h_anchor)
            if isinstance(v_anchor, WD_ANCHOR_V):
                v_anchor_value = v_anchor.value
            else:
                v_anchor_value = str(v_anchor)
            if isinstance(wrap, WD_WRAP_TYPE):
                wrap_value = wrap.value
            else:
                wrap_value = str(wrap)

            anchor.set_horizontal_position(h_anchor_value, int(horizontal))
            anchor.set_vertical_position(v_anchor_value, int(vertical))
            anchor.set_wrap(wrap_value)

        # -- append the anchor inside a new run's `w:drawing` --
        run = self.add_run()
        run._r.add_drawing(anchor)
        return FloatingImage(anchor)

    @property
    def fields(self) -> list[Field]:
        """List of |Field| objects for each field in this paragraph.

        Includes both simple (``w:fldSimple``) and complex (``w:fldChar``)
        fields, in document order.
        """
        result: list[Field] = []
        for kind, el in self._p.iter_field_elements():
            if kind == "simple":
                result.append(Field.for_simple(el))
            else:
                result.append(Field.for_complex(el))
        return result

    @property
    def floating_images(self) -> list[FloatingImage]:
        """A |FloatingImage| instance for each `wp:anchor` in this paragraph."""
        return [
            FloatingImage(cast(CT_Anchor, a))
            for a in self._p.xpath(".//w:r/w:drawing/wp:anchor")
        ]

    @property
    def content_controls(self) -> list[ContentControl]:
        """List of inline |ContentControl| objects in this paragraph, in document order."""
        from docx.content_controls import ContentControl

        return [
            ContentControl(cast("CT_Sdt", sdt)) for sdt in self._p.xpath("./w:sdt")
        ]

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """A member of the :ref:`WdParagraphAlignment` enumeration specifying the
        justification setting for this paragraph.

        A value of |None| indicates the paragraph has no directly-applied alignment
        value and will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        self._p.alignment = value

    def clear(self):
        """Return this same paragraph after removing all its content.

        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    def delete(self) -> None:
        """Remove this paragraph from the document.

        The paragraph element is removed from its parent. After calling this method,
        this |Paragraph| object is "defunct" and should not be used further.
        """
        p = self._p
        parent = p.getparent()
        if parent is None:
            return
        parent.remove(p)

    def clear_page_breaks(self) -> None:
        """Remove all ``<w:br w:type="page"/>`` elements from this paragraph.

        If a run contains only a page break and no other content, the entire run is
        removed. If a run contains other content alongside the page break, only the
        ``<w:br>`` element is removed. Does nothing when no page breaks are present.
        """
        for br in self._p.xpath('.//w:br[@w:type="page"]'):
            r = br.getparent()
            r.remove(br)
            # --- remove the run if it's now empty (no child elements and no text) ---
            if len(r) == 0 and not r.text:
                r.getparent().remove(r)

    @property
    def has_section_break(self) -> bool:
        """``True`` if this paragraph contains a section break (``<w:sectPr>`` in its
        ``<w:pPr>``)."""
        pPr = self._p.pPr
        if pPr is None:
            return False
        return pPr.sectPr is not None

    @property
    def contains_page_break(self) -> bool:
        """`True` when one or more rendered page-breaks occur in this paragraph."""
        return bool(self._p.lastRenderedPageBreaks)

    @property
    def has_page_break(self) -> bool:
        """`True` if this paragraph contains at least one ``<w:br w:type="page"/>``."""
        return bool(self._p.xpath('.//w:br[@w:type="page"]'))

    @property
    def drawings(self) -> list[Drawing]:
        """A |Drawing| instance for each `<w:drawing>` element in this paragraph."""
        return [
            Drawing(cast(CT_Drawing, d), self)
            for d in self._p.xpath(".//w:drawing")
        ]

    @property
    def hyperlinks(self) -> list[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    def insert_section_break(
        self, start_type: WD_SECTION_START = WD_SECTION_START.NEW_PAGE
    ) -> Section:
        """Insert a section break in this paragraph and return the new |Section|.

        `start_type` is a member of :ref:`WdSectionStart` and defaults to
        ``WD_SECTION.NEW_PAGE``. If this paragraph already contains a section break,
        its type is replaced rather than a new one being added.
        """
        from docx.section import Section as SectionCls

        pPr = self._p.get_or_add_pPr()
        sectPr = pPr.get_or_add_sectPr()
        sectPr.start_type = start_type
        return SectionCls(sectPr, self.part)

    def remove_section_break(self) -> None:
        """Remove the section break from this paragraph, if one is present.

        Calling this on a paragraph that has no section break is a no-op.
        """
        pPr = self._p.pPr
        if pPr is None:
            return
        if pPr.sectPr is not None:
            pPr._remove_sectPr()

    def insert_paragraph_before(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def insert_paragraph_after(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly after this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph. The new
        paragraph is inserted into the same parent element as this paragraph (which
        may be a body, cell, header/footer, or other block-level container).
        """
        from docx.oxml.parser import OxmlElement

        new_p = cast("CT_P", OxmlElement("w:p"))
        self._p.addnext(new_p)
        paragraph = Paragraph(new_p, self._parent)
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def add_caption_before(
        self,
        text: str,
        label: str = "Figure",
        style: str = "Caption",
    ) -> Paragraph:
        """Insert a caption paragraph directly before this paragraph and return it.

        This is the common shape for a caption that sits *above* a figure
        or table. The inserted paragraph has the standard caption structure:
        ``"{label} N: {text}"`` where ``N`` is produced by a
        ``SEQ {label} \\* ARABIC`` field. See
        :meth:`docx.document.Document.add_caption` for details.
        """
        from docx.captions import new_caption_paragraph

        paragraph = self.insert_paragraph_before()
        return new_caption_paragraph(paragraph, text, label=label, style=style)

    def add_caption_after(
        self,
        text: str,
        label: str = "Figure",
        style: str = "Caption",
    ) -> Paragraph:
        """Insert a caption paragraph directly after this paragraph and return it.

        This is the common shape for a caption that sits *below* a figure
        or table. The inserted paragraph has the standard caption structure:
        ``"{label} N: {text}"`` where ``N`` is produced by a
        ``SEQ {label} \\* ARABIC`` field. See
        :meth:`docx.document.Document.add_caption` for details.
        """
        from docx.captions import new_caption_paragraph

        paragraph = self.insert_paragraph_after()
        return new_caption_paragraph(paragraph, text, label=label, style=style)

    def insert_table_before(
        self,
        rows: int,
        cols: int,
        style: str | _TableStyle | None = None,
        width: Length | None = None,
    ) -> _Table:
        """Return a new table with `rows` rows and `cols` cols, inserted directly
        before this paragraph.

        If `style` is supplied, that style is assigned to the new table. The new
        table is inserted as a sibling of this paragraph in its parent element.
        `width` is an optional total table width; if not provided it defaults to 6
        inches (a reasonable default for a US-Letter page with 1" margins).
        """
        from docx.table import Table

        table_width = width if width is not None else Inches(6)
        tbl = CT_Tbl.new_tbl(rows, cols, table_width)
        self._p.addprevious(tbl)
        table = Table(tbl, self._parent)
        if style is not None:
            table.style = style
        return table

    def insert_table_after(
        self,
        rows: int,
        cols: int,
        style: str | _TableStyle | None = None,
        width: Length | None = None,
    ) -> _Table:
        """Return a new table with `rows` rows and `cols` cols, inserted directly
        after this paragraph.

        If `style` is supplied, that style is assigned to the new table. The new
        table is inserted as a sibling of this paragraph in its parent element.
        `width` is an optional total table width; if not provided it defaults to 6
        inches.
        """
        from docx.table import Table

        table_width = width if width is not None else Inches(6)
        tbl = CT_Tbl.new_tbl(rows, cols, table_width)
        self._p.addnext(tbl)
        table = Table(tbl, self._parent)
        if style is not None:
            table.style = style
        return table

    def iter_inner_content(self) -> Iterator[Run | Hyperlink]:
        """Generate the runs and hyperlinks in this paragraph, in the order they appear.

        The content in a paragraph consists of both runs and hyperlinks. This method
        allows accessing each of those separately, in document order, for when the
        precise position of the hyperlink within the paragraph text is important. Note
        that a hyperlink itself contains runs.
        """
        for r_or_hlink in self._p.inner_content_elements:
            yield (
                Run(r_or_hlink, self)
                if isinstance(r_or_hlink, CT_R)
                else Hyperlink(r_or_hlink, self)
            )

    @property
    def paragraph_format(self):
        """The |ParagraphFormat| object providing access to the formatting properties
        for this paragraph, such as line spacing and indentation."""
        return ParagraphFormat(self._element)

    @property
    def list_level(self) -> int | None:
        """The integer list-level of this paragraph (``w:numPr/w:ilvl/@w:val``).

        Returns |None| when the paragraph has no ``w:numPr`` or ``w:ilvl``
        child. Valid values are ``0`` through ``8``.

        Assigning |None| removes the ``w:ilvl`` child. Assigning an integer
        outside the range 0..8 raises ``ValueError``.
        """
        pPr = self._p.pPr
        if pPr is None or pPr.numPr is None:
            return None
        return pPr.numPr.ilvl_val

    @list_level.setter
    def list_level(self, value: int | None) -> None:
        if value is not None:
            if not isinstance(value, int) or not 0 <= value <= 8:
                raise ValueError(
                    "list_level must be an int in 0..8 or None, got %r" % (value,)
                )
        pPr = self._p.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.ilvl_val = value

    @property
    def list_format(self):
        """Named tuple ``(numbering_definition, level)`` describing this paragraph's
        list settings.

        Both fields are |None| when the paragraph is not part of a list. The
        ``numbering_definition`` is resolved by looking up the paragraph's
        ``numId`` in the document's numbering part.

        To set a paragraph's list format, use
        :meth:`NumberingDefinition.apply_to`.
        """
        from docx.numbering import ListFormat, Numbering, NumberingDefinition

        pPr = self._p.pPr
        if pPr is None or pPr.numPr is None:
            return ListFormat(None, None)
        numPr = pPr.numPr
        num_id = numPr.numId_val
        level = numPr.ilvl_val
        if num_id is None:
            return ListFormat(None, level)

        numbering_part = getattr(self.part, "numbering_part", None)
        if numbering_part is None:
            return ListFormat(None, level)

        numbering_elm = numbering_part.numbering_element
        try:
            num = numbering_elm.num_having_numId(num_id)
        except KeyError:
            return ListFormat(None, level)

        abstractNumId_elm = num.abstractNumId
        abstract_num_id = abstractNumId_elm.val
        try:
            abstractNum = numbering_elm.abstractNum_having_abstractNumId(
                abstract_num_id
            )
        except KeyError:
            return ListFormat(None, level)

        numbering_proxy = Numbering(numbering_elm, numbering_part)
        return ListFormat(
            NumberingDefinition(abstractNum, numbering_proxy), level
        )

    @property
    def numbering_format(self):
        """Read-only |Level| describing this paragraph's current level in its list.

        Returns |None| if the paragraph is not part of a numbered list, or if the
        list-level entry cannot be found in the document's numbering part.
        """
        list_format = self.list_format
        if list_format.numbering_definition is None:
            return None
        level = list_format.level if list_format.level is not None else 0
        return list_format.numbering_definition.level(level)

    def restart_numbering(self, start: int = 1) -> None:
        """Create a new numbering instance that restarts the current list at `start`.

        The new ``w:num`` reuses the existing abstract definition but adds a
        ``w:lvlOverride/w:startOverride`` for this paragraph's level. The
        paragraph's ``w:numPr/w:numId`` is rewritten to point at the new
        instance, so subsequent siblings at the same level continue the fresh
        count.

        Raises ``ValueError`` when the paragraph is not currently part of a
        numbered list.
        """
        pPr = self._p.pPr
        if pPr is None or pPr.numPr is None or pPr.numPr.numId_val is None:
            raise ValueError(
                "paragraph is not part of a numbered list; apply a numbering "
                "definition before calling restart_numbering()"
            )
        numPr = pPr.numPr
        num_id = numPr.numId_val
        ilvl = numPr.ilvl_val or 0

        try:
            numbering_part = self.part.numbering_part  # type: ignore[attr-defined]
        except AttributeError as err:
            raise ValueError(
                "cannot locate numbering part for this paragraph"
            ) from err

        numbering_elm = numbering_part.numbering_element
        try:
            existing_num = numbering_elm.num_having_numId(num_id)
        except KeyError as err:
            raise ValueError(
                "paragraph's numId %d does not match any w:num" % num_id
            ) from err

        abstract_num_id = existing_num.abstractNumId.val
        new_num = numbering_elm.add_num(abstract_num_id)
        override = new_num.add_lvlOverride(ilvl=ilvl)
        override.add_startOverride(val=start)

        numPr.numId_val = new_num.numId

    @property
    def rsid(self) -> str | None:
        """The paragraph's revision-save ID (``w:p/@w:rsidR``) or |None|.

        Read-only. Returns the 8-character hex string Word assigns to mark the
        editing session in which this paragraph was last modified, or |None|
        when the ``@w:rsidR`` attribute is not present.
        """
        return self._p.rsidR

    @property
    def rendered_page_breaks(self) -> list[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    @property
    def runs(self) -> list[Run]:
        """Sequence of |Run| instances corresponding to the <w:r> elements in this
        paragraph."""
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self) -> ParagraphStyle | None:
        """Read/Write.

        |_ParagraphStyle| object representing the style assigned to this paragraph. If
        no explicit style is assigned to this paragraph, its value is the default
        paragraph style for the document. A paragraph style name can be assigned in lieu
        of a paragraph style object. Assigning |None| removes any applied style, making
        its effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        style = self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        return cast(ParagraphStyle, style)

    @style.setter
    def style(self, style_or_name: str | ParagraphStyle | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.PARAGRAPH)
        self._p.style = style_id

    @property
    def formatting_change(self):
        """A |FormattingChange| for this paragraph's `w:pPrChange`, or |None|.

        Present when the paragraph's formatting (its `w:pPr`) has been edited while
        track-changes is enabled. The returned object exposes the author, date, and
        the prior `w:pPr` via ``old_properties``.
        """
        from docx.tracked_changes import FormattingChange

        pPr = self._p.pPr
        if pPr is None:
            return None
        pPrChange = pPr.pPrChange  # pyright: ignore[reportAttributeAccessIssue]
        if pPrChange is None:
            return None
        return FormattingChange(pPrChange)

    @property
    def tracked_changes(self) -> list[TrackedChange]:
        """A list of |TrackedChange| objects for each insertion or deletion in this
        paragraph."""
        return [TrackedChange(tc) for tc in self._p.tracked_change_elements]

    def revision_marks_text(
        self,
        open_ins: str = "[+",
        close_ins: str = "+]",
        open_del: str = "[-",
        close_del: str = "-]",
    ) -> str:
        """Return this paragraph's text with tracked-change markers applied.

        Inserted runs (inside ``<w:ins>``) are wrapped with `open_ins`/`close_ins`
        and deleted runs (inside ``<w:del>``) with `open_del`/`close_del`. Runs
        outside of any track-change wrapper are rendered as plain text.

        When the paragraph contains no tracked changes, the return value matches
        :attr:`text`. The defaults are CLI-friendly square-bracket markers; callers
        can pass ANSI escape sequences (e.g. ``"\\033[4m"`` / ``"\\033[0m"``) to
        style terminal output instead.
        """
        from docx.tracked_changes import _render_paragraph_marks

        return _render_paragraph_marks(
            self._p, open_ins, close_ins, open_del, close_del
        )

    @property
    def text(self) -> str:
        """The textual content of this paragraph.

        The text includes the visible-text portion of any hyperlinks in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n`` characters
        respectively.

        Assigning text to this property causes all existing paragraph content to be
        replaced with a single run containing the assigned text. A ``\\t`` character in
        the text is mapped to a ``<w:tab/>`` element and each ``\\n`` or ``\\r``
        character is mapped to a line break. Paragraph-level formatting, such as style,
        is preserved. All run-level formatting, such as bold or italic, is removed.
        """
        return self._p.text

    @text.setter
    def text(self, text: str | None):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)
