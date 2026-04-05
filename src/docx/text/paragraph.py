"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterator, List, cast

from docx.drawing import Drawing
from docx.enum.section import WD_SECTION_START
from docx.enum.shape import WD_RELATIVE_HORZ_POS, WD_RELATIVE_VERT_POS, WD_WRAP_TYPE
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.oxml.drawing import CT_Drawing
from docx.oxml.text.run import CT_R
from docx.shape import FloatingImage
from docx.shared import StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.tracked_changes import TrackedChange
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.bookmarks import Bookmark
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.document import CT_Body
    from docx.oxml.text.paragraph import CT_P
    from docx.section import Section
    from docx.shared import Length
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
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

    def add_floating_image(
        self,
        image_path_or_stream: str | IO[bytes],
        width: int | Length | None = None,
        height: int | Length | None = None,
        horz_offset: int = 0,
        vert_offset: int = 0,
        horz_pos_relative: WD_RELATIVE_HORZ_POS = WD_RELATIVE_HORZ_POS.COLUMN,
        vert_pos_relative: WD_RELATIVE_VERT_POS = WD_RELATIVE_VERT_POS.PARAGRAPH,
        wrap_type: WD_WRAP_TYPE = WD_WRAP_TYPE.IN_FRONT,
    ) -> FloatingImage:
        """Return |FloatingImage| containing image identified by `image_path_or_stream`.

        The floating image is added to this paragraph using a `wp:anchor` element
        instead of `wp:inline`, allowing it to be positioned relative to the page,
        margin, column, or paragraph.

        `horz_offset` and `vert_offset` are in EMUs (English Metric Units).
        `horz_pos_relative` and `vert_pos_relative` specify the reference frame.
        `wrap_type` specifies how text wraps around the image.
        """
        # -- map enum wrap_type to the string used in the XML layer --
        wrap_map = {
            WD_WRAP_TYPE.SQUARE: ("square", False),
            WD_WRAP_TYPE.TIGHT: ("tight", False),
            WD_WRAP_TYPE.THROUGH: ("through", False),
            WD_WRAP_TYPE.TOP_AND_BOTTOM: ("topAndBottom", False),
            WD_WRAP_TYPE.IN_FRONT: ("none", False),
            WD_WRAP_TYPE.BEHIND: ("none", True),
        }
        wrap_str, behind_doc = wrap_map[wrap_type]

        anchor = self.part.new_pic_anchor(
            image_path_or_stream,
            width,
            height,
            horz_offset=horz_offset,
            vert_offset=vert_offset,
            horz_relative_from=horz_pos_relative.value,
            vert_relative_from=vert_pos_relative.value,
            wrap_type=wrap_str,
            behind_doc=behind_doc,
        )
        run = self.add_run()
        run._r.add_drawing(anchor)
        return FloatingImage(anchor)

    @property
    def floating_images(self) -> List[FloatingImage]:
        """A |FloatingImage| for each `<wp:anchor>` element in this paragraph."""
        from docx.oxml.shape import CT_Anchor

        return [
            FloatingImage(cast(CT_Anchor, a))
            for a in self._p.xpath(".//w:drawing/wp:anchor")
        ]

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

    def add_page_break(self) -> Paragraph:
        """Append a page-break run to this paragraph and return self."""
        run = self.add_run()
        run.add_break(WD_BREAK.PAGE)
        return self

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
    def drawings(self) -> List[Drawing]:
        """A |Drawing| instance for each `<w:drawing>` element in this paragraph."""
        return [
            Drawing(cast(CT_Drawing, d), self)
            for d in self._p.xpath(".//w:drawing")
        ]

    @property
    def hyperlinks(self) -> List[Hyperlink]:
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
    def rendered_page_breaks(self) -> List[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    @property
    def runs(self) -> List[Run]:
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
    def tracked_changes(self) -> List[TrackedChange]:
        """A list of |TrackedChange| objects for each insertion or deletion in this
        paragraph."""
        return [TrackedChange(tc) for tc in self._p.tracked_change_elements]

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
