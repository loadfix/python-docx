"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, cast

from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.oxml.text.run import CT_R
from docx.shared import StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.listformat import ListFormat
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.text.paragraph import CT_P
    from docx.section import Section
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

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
    def hyperlinks(self) -> List[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    @property
    def list_format(self) -> ListFormat:
        """A |ListFormat| object providing access to the list formatting properties
        for this paragraph, such as the numbering definition and indent level."""
        return ListFormat(self._p, self.part)

    @property
    def list_level(self) -> int | None:
        """The list indentation level (0-8) for this paragraph.

        Returns None if this paragraph is not part of a list. Assigning an int
        sets the level; assigning None removes it.
        """
        return self.list_format.level

    @list_level.setter
    def list_level(self, value: int | None) -> None:
        self.list_format.level = value

    @property
    def numbering_format(self) -> str | None:
        """The current numbering format string (e.g. "decimal", "bullet") for this
        paragraph, or None if the paragraph is not part of a list.

        This is read-only. To change numbering, use ``list_format``.
        """
        list_fmt = self.list_format
        num_id = list_fmt.num_id
        if num_id is None or num_id == 0:
            return None
        level = list_fmt.level or 0
        try:
            numbering_part = self.part.numbering_part
            numbering_elm = numbering_part.numbering_element
            num = numbering_elm.num_having_numId(num_id)
            abstract_num_id = num.abstractNumId_val
            abstract_num = numbering_elm.abstractNum_having_abstractNumId(abstract_num_id)
            lvl = abstract_num.lvl_for_ilvl(level)
            if lvl is not None:
                return lvl.numFmt_val
        except (KeyError, AttributeError):
            pass
        return None

    def restart_numbering(self) -> None:
        """Restart the numbered list counter at 1 for this paragraph.

        Creates a new ``<w:num>`` element referencing the same abstract numbering
        definition but with a ``<w:lvlOverride>/<w:startOverride>`` set to 1.
        """
        list_fmt = self.list_format
        num_id = list_fmt.num_id
        if num_id is None or num_id == 0:
            return

        numbering_part = self.part.numbering_part
        numbering_elm = numbering_part.numbering_element
        num = numbering_elm.num_having_numId(num_id)
        abstract_num_id = num.abstractNumId_val

        new_num = numbering_elm.add_num(abstract_num_id)
        current_level = list_fmt.level or 0
        lvl_override = new_num.add_lvlOverride(ilvl=current_level)
        lvl_override.add_startOverride(val=1)

        list_fmt.num_id = new_num.numId

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
