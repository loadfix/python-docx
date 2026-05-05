"""The |Section| object and related proxy classes."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, overload
from collections.abc import Iterator, Sequence

from docx.blkcntnr import BlockItemContainer
from docx.enum.section import WD_HEADER_FOOTER
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.text.paragraph import CT_P
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.watermark import Watermark

if TYPE_CHECKING:
    from docx.endnotes import EndnoteProperties
    from docx.enum.section import (
        WD_BORDER_DISPLAY,
        WD_BORDER_OFFSET_FROM,
        WD_DOC_GRID_TYPE,
        WD_LINE_NUMBERING_RESTART,
        WD_ORIENTATION,
        WD_SECTION_START,
        WD_VERTICAL_ALIGNMENT,
    )
    from docx.enum.table import WD_TEXT_DIRECTION
    from docx.enum.text import WD_BORDER_STYLE
    from docx.footnotes import FootnoteProperties
    from docx.oxml.document import CT_Document
    from docx.oxml.section import (
        CT_Col,
        CT_Cols,
        CT_DocGrid,
        CT_LineNumber,
        CT_PgBorders,
        CT_SectPr,
    )
    from docx.oxml.text.parfmt import CT_Border
    from docx.oxml.watermark import CT_VmlShape
    from docx.parts.document import DocumentPart
    from docx.parts.story import StoryPart
    from docx.shared import Length


__all__ = [
    "Column",
    "DocumentGrid",
    "LineNumbering",
    "PageBorder",
    "PageBorders",
    "Section",
    "SectionColumns",
    "Sections",
    # -- underscored-but-public header/footer proxies --
    "_BaseHeaderFooter",
    "_Footer",
    "_Header",
]


class Section:
    """Document section, providing access to section and page setup settings.

    Also provides access to headers and footers.
    """

    def __init__(self, sectPr: CT_SectPr, document_part: DocumentPart):
        super().__init__()
        self._sectPr = sectPr
        self._document_part = document_part

    @property
    def columns(self) -> SectionColumns:
        """|SectionColumns| object providing access to column layout settings.

        .. versionadded:: 2026.05.0
        """
        return SectionColumns(self._sectPr)

    def set_columns(
        self,
        count: int,
        space: "Length | None" = None,
        equal_width: bool | None = None,
    ) -> SectionColumns:
        """Set column layout for this section in one call.

        `count` is written to ``w:cols/@w:num``. When `space` is supplied it
        is written to ``w:cols/@w:space``; when |None| the attribute is left
        unchanged. `equal_width` maps to ``w:cols/@w:equalWidth`` the same
        way. Returns the |SectionColumns| proxy for the resulting element.

        Mirrors the ``set_page_border`` / ``set_line_numbering`` /
        ``set_document_grid`` convenience pattern — equivalent to::

            section.columns.count = count
            section.columns.space = space
            section.columns.equal_width = equal_width

        .. versionadded:: 2026.05.0
        """
        columns = self.columns
        columns.count = count
        if space is not None:
            columns.space = space
        if equal_width is not None:
            columns.equal_width = equal_width
        return columns

    @property
    def bottom_margin(self) -> Length | None:
        """Read/write. Bottom margin for pages in this section, in EMU.

        `None` when no bottom margin has been specified. Assigning |None| removes any
        bottom-margin setting.
        """
        return self._sectPr.bottom_margin

    @bottom_margin.setter
    def bottom_margin(self, value: int | Length | None):
        self._sectPr.bottom_margin = value

    @property
    def different_first_page_header_footer(self) -> bool:
        """True if this section displays a distinct first-page header and footer.

        Read/write. The definition of the first-page header and footer are accessed
        using :attr:`.first_page_header` and :attr:`.first_page_footer` respectively.
        """
        return self._sectPr.titlePg_val

    @different_first_page_header_footer.setter
    def different_first_page_header_footer(self, value: bool):
        self._sectPr.titlePg_val = value

    @property
    def formatting_change(self):
        """A |FormattingChange| for this section's `w:sectPrChange`, or |None|.

        Present when the section's formatting has been edited while track-changes is
        enabled. The returned object exposes the author, date, and the prior
        `w:sectPr` via ``old_properties``.

        .. versionadded:: 2026.05.0
        """
        from docx.tracked_changes import FormattingChange

        sectPrChange = self._sectPr.sectPrChange  # pyright: ignore[reportAttributeAccessIssue]
        if sectPrChange is None:
            return None
        return FormattingChange(sectPrChange)

    @property
    def different_odd_and_even_pages_header_footer(self) -> bool:
        """Read/write. |True| when this document displays distinct odd/even headers.

        This is a **document-level** setting, not a per-section setting. It maps to
        ``w:settings/w:evenAndOddHeaders`` in the settings part and applies to every
        section in the document. It is surfaced on |Section| for discoverability --
        authors intuitively look for it beside :attr:`different_first_page_header_footer`.

        When |False|, the content of :attr:`even_page_header` and
        :attr:`even_page_footer` is ignored by Word and the primary (odd) header/footer
        is used for every page. Setting this to |True| is required for
        :attr:`even_page_header` and :attr:`even_page_footer` to take effect.

        This property is a thin alias for
        :attr:`docx.settings.Settings.even_and_odd_headers` on the parent document.

        .. versionadded:: 2026.05.0
        """
        return self._document_part.settings.even_and_odd_headers

    @different_odd_and_even_pages_header_footer.setter
    def different_odd_and_even_pages_header_footer(self, value: bool):
        self._document_part.settings.even_and_odd_headers = value

    @property
    def even_page_footer(self) -> _Footer:
        """|_Footer| object defining footer content for even pages.

        The content of this footer definition is ignored unless the document setting
        :attr:`different_odd_and_even_pages_header_footer` is set |True| (equivalent to
        setting :attr:`docx.settings.Settings.even_and_odd_headers`). That setting is
        document-level, not per-section, so toggling it affects every section.
        """
        return _Footer(self._sectPr, self._document_part, WD_HEADER_FOOTER.EVEN_PAGE)

    @property
    def even_page_header(self) -> _Header:
        """|_Header| object defining header content for even pages.

        The content of this header definition is ignored unless the document setting
        :attr:`different_odd_and_even_pages_header_footer` is set |True| (equivalent to
        setting :attr:`docx.settings.Settings.even_and_odd_headers`). That setting is
        document-level, not per-section, so toggling it affects every section.
        """
        return _Header(self._sectPr, self._document_part, WD_HEADER_FOOTER.EVEN_PAGE)

    @property
    def first_page_footer(self) -> _Footer:
        """|_Footer| object defining footer content for the first page of this section.

        The content of this footer definition is ignored unless the property
        :attr:`.different_first_page_header_footer` is set True.
        """
        return _Footer(self._sectPr, self._document_part, WD_HEADER_FOOTER.FIRST_PAGE)

    @property
    def first_page_header(self) -> _Header:
        """|_Header| object defining header content for the first page of this section.

        The content of this header definition is ignored unless the property
        :attr:`.different_first_page_header_footer` is set True.
        """
        return _Header(self._sectPr, self._document_part, WD_HEADER_FOOTER.FIRST_PAGE)

    @property
    def footer(self) -> _Footer:
        """|_Footer| object representing default page footer for this section.

        The default footer is used for odd-numbered pages when separate odd/even footers
        are enabled. It is used for both odd and even-numbered pages otherwise.
        """
        return _Footer(self._sectPr, self._document_part, WD_HEADER_FOOTER.PRIMARY)

    @property
    def footer_distance(self) -> Length | None:
        """Distance from bottom edge of page to bottom edge of the footer.

        Read/write. |None| if no setting is present in the XML.
        """
        return self._sectPr.footer

    @footer_distance.setter
    def footer_distance(self, value: int | Length | None):
        self._sectPr.footer = value

    @property
    def gutter(self) -> Length | None:
        """|Length| object representing page gutter size in English Metric Units.

        Read/write. The page gutter is extra spacing added to the `inner` margin to
        ensure even margins after page binding. Generally only used in book-bound
        documents with double-sided and facing pages.

        This setting applies to all pages in this section.

        """
        return self._sectPr.gutter

    @gutter.setter
    def gutter(self, value: int | Length | None):
        self._sectPr.gutter = value

    @property
    def header(self) -> _Header:
        """|_Header| object representing default page header for this section.

        The default header is used for odd-numbered pages when separate odd/even headers
        are enabled. It is used for both odd and even-numbered pages otherwise.
        """
        return _Header(self._sectPr, self._document_part, WD_HEADER_FOOTER.PRIMARY)

    @property
    def header_distance(self) -> Length | None:
        """Distance from top edge of page to top edge of header.

        Read/write. |None| if no setting is present in the XML. Assigning |None| causes
        default value to be used.
        """
        return self._sectPr.header

    @header_distance.setter
    def header_distance(self, value: int | Length | None):
        self._sectPr.header = value

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        """Generate each Paragraph or Table object in this `section`.

        Items appear in document order.
        """
        for element in self._sectPr.iter_inner_content():
            yield (Paragraph(element, self) if isinstance(element, CT_P) else Table(element, self))

    @property
    def paragraphs(self) -> list[Paragraph]:
        """List of |Paragraph| objects bounded by this section's ``w:sectPr``.

        A section's paragraphs are the ``w:p`` elements located after the prior
        section's terminating ``w:sectPr`` (exclusive) and up to-and-including
        the paragraph that hosts this section's ``w:sectPr`` (for
        paragraph-hosted ``w:sectPr``) or every remaining paragraph (for the
        document-terminal ``w:body/w:sectPr``).

        Tables and other block-level items are skipped; use
        :meth:`iter_inner_content` for the full sequence. Items appear in
        document order.

        .. versionadded:: 2026.05.0
        """
        return [
            Paragraph(element, self)
            for element in self._sectPr.iter_inner_content()
            if isinstance(element, CT_P)
        ]

    @property
    def tables(self) -> list[Table]:
        """List of |Table| objects bounded by this section's ``w:sectPr``.

        Companion to :attr:`paragraphs`; returns the tables whose ``w:tbl``
        elements fall in this section's range. Items appear in document order.

        .. versionadded:: 2026.05.0
        """
        return [
            Table(element, self)
            for element in self._sectPr.iter_inner_content()
            if not isinstance(element, CT_P)
        ]

    def delete(self) -> None:
        """Remove this section, merging its content into the following section.

        The paragraphs and tables that were part of this section become part
        of the next section (or, if this section is the last, the preceding
        section absorbs them by losing its terminating ``w:sectPr``).

        For a paragraph-hosted ``w:sectPr`` (i.e. any section except the last
        one), the ``w:sectPr`` is removed from its ``w:pPr`` parent; the
        paragraph itself is preserved so its content is not lost. The
        following section's ``w:sectPr`` now controls the merged range.

        For the document-terminal ``w:body/w:sectPr`` (the last section), the
        preceding section's ``w:sectPr`` is promoted in its place. The
        previously-promoted ``w:sectPr`` and its owning paragraph are removed
        (the paragraph is merged away because its only purpose was to host
        the now-redundant section break).

        Calling :meth:`delete` on the only section in a document is a no-op --
        every document must have at least one ``w:sectPr``.

        After calling this method, this |Section| object is "defunct" and
        should not be used further.

        .. versionadded:: 2026.05.0
        """
        sectPr = self._sectPr
        parent = sectPr.getparent()
        if parent is None:
            return
        # -- identify whether this is a p-hosted sectPr or body-hosted --
        if parent.tag == qn("w:pPr"):
            # -- p-hosted: just drop the sectPr from its pPr.
            # -- The paragraph that hosted it survives and joins the next section.
            parent.remove(sectPr)
            # -- if the pPr is now effectively empty, leave it; an empty pPr is
            # -- harmless and Word tolerates it.
            return
        # -- body-hosted (last section): there must be a preceding p-hosted sectPr
        # -- to promote, otherwise this is the only section and we no-op.
        preceding = sectPr.preceding_sectPr
        if preceding is None:
            return
        # -- move the preceding sectPr out of its paragraph and into body --
        preceding_pPr = preceding.getparent()
        assert preceding_pPr is not None
        preceding_p = preceding_pPr.getparent()
        assert preceding_p is not None
        body = parent
        # -- remove the old body-sectPr --
        body.remove(sectPr)
        # -- remove the hosting paragraph (which only existed to carry a break) --
        preceding_pPr.remove(preceding)
        body_parent = preceding_p.getparent()
        assert body_parent is not None
        body_parent.remove(preceding_p)
        # -- append the promoted sectPr at the end of body --
        body.append(preceding)

    def copy_header_from(self, other_section: Section) -> None:
        """Copy the default page header definition from `other_section` into this one.

        The header part is duplicated (a new ``/word/headerN.xml`` part is
        added to the package), the header XML tree is deep-copied from
        `other_section`'s header, and ``w:headerReference`` on this section is
        rewired to the new part.

        Image relationships in the source header are *not* transplanted; the
        copied header's image ``r:embed`` values will still point at the
        source part's image relationships. For pure text headers (by far the
        common case) this produces a correct standalone copy.

        Does nothing when `other_section` has no default-header definition.

        .. versionadded:: 2026.05.0
        """
        self._copy_hdrftr(other_section, is_header=True)

    def copy_footer_from(self, other_section: Section) -> None:
        """Copy the default page footer definition from `other_section` into this one.

        Companion to :meth:`copy_header_from`; see that method for the full
        contract. Does nothing when `other_section` has no default-footer
        definition.

        .. versionadded:: 2026.05.0
        """
        self._copy_hdrftr(other_section, is_header=False)

    def _copy_hdrftr(self, other_section: Section, is_header: bool) -> None:
        """Shared worker for :meth:`copy_header_from` / :meth:`copy_footer_from`."""
        from copy import deepcopy

        # -- source proxy, via `_has_definition` to avoid triggering inheritance --
        if is_header:
            src = other_section.header
            dst = self.header
        else:
            src = other_section.footer
            dst = self.footer
        if not src._has_definition:
            return
        src_part = src.part  # triggers resolution, but we already checked it exists
        src_elm = src_part.element

        # -- ensure destination currently owns a part; drop any existing definition
        # -- so the new one is cleanly wired in (matches the "last write wins"
        # -- semantics of set-property style APIs in the rest of this module).
        if dst._has_definition:
            dst._drop_definition()
        new_part = dst._add_definition()
        # -- replace the new (empty) part's root element with a deep copy of src --
        new_elm = deepcopy(src_elm)
        # -- swap: clear new_part's current element content and copy src children --
        current = new_part.element
        # -- remove all existing children of the fresh hdr/ftr element --
        for child in list(current):
            current.remove(child)
        # -- copy attributes and children from the source root --
        for key, value in new_elm.attrib.items():
            current.set(key, value)
        for child in list(new_elm):
            current.append(child)

    @property
    def left_margin(self) -> Length | None:
        """|Length| object representing the left margin for all pages in this section in
        English Metric Units."""
        return self._sectPr.left_margin

    @left_margin.setter
    def left_margin(self, value: int | Length | None):
        self._sectPr.left_margin = value

    @property
    def orientation(self) -> WD_ORIENTATION:
        """:ref:`WdOrientation` member specifying page orientation for this section.

        One of ``WD_ORIENT.PORTRAIT`` or ``WD_ORIENT.LANDSCAPE``.
        """
        return self._sectPr.orientation

    @orientation.setter
    def orientation(self, value: WD_ORIENTATION | None):
        self._sectPr.orientation = value

    @property
    def vertical_alignment(self) -> WD_VERTICAL_ALIGNMENT | None:
        """Read/write. Vertical alignment of text for this section, or |None|.

        Maps to the ``w:val`` attribute of the ``w:vAlign`` child of ``w:sectPr``
        (ECMA-376 17.6.22, simple type ``ST_VerticalJc``). Assigning |None|
        removes the ``w:vAlign`` child, restoring the default top alignment.

        One of ``WD_VERTICAL_ALIGNMENT.TOP``, ``.CENTER``, ``.BOTH``, ``.BOTTOM``.
        """
        return self._sectPr.vertical_alignment

    @vertical_alignment.setter
    def vertical_alignment(self, value: WD_VERTICAL_ALIGNMENT | None):
        self._sectPr.vertical_alignment = value

    @property
    def page_height(self) -> Length | None:
        """Total page height used for this section.

        This value is inclusive of all edge spacing values such as margins.

        Page orientation is taken into account, so for example, its expected value
        would be ``Inches(8.5)`` for letter-sized paper when orientation is landscape.
        """
        return self._sectPr.page_height

    @page_height.setter
    def page_height(self, value: Length | None):
        self._sectPr.page_height = value

    @property
    def page_width(self) -> Length | None:
        """Total page width used for this section.

        This value is like "paper size" and includes all edge spacing values such as
        margins.

        Page orientation is taken into account, so for example, its expected value
        would be ``Inches(11)`` for letter-sized paper when orientation is landscape.
        """
        return self._sectPr.page_width

    @page_width.setter
    def page_width(self, value: Length | None):
        self._sectPr.page_width = value

    @property
    def part(self) -> StoryPart:
        return self._document_part

    @property
    def page_borders(self) -> PageBorders:
        """|PageBorders| proxy providing access to this section's page borders.

        The returned object lazily creates the underlying ``w:pgBorders`` element only
        when a border edge or attribute is actually assigned. When no
        ``w:pgBorders`` child is present, reads of ``.top``/``.bottom``/``.left``/
        ``.right`` return |PageBorder| proxies whose ``.style``, ``.width``,
        ``.color`` and ``.space`` are ``None``.

        .. versionadded:: 2026.05.0
        """
        return PageBorders(self._sectPr)

    def set_page_border(
        self,
        side: str,
        style: "WD_BORDER_STYLE | None" = None,
        width: "Length | None" = None,
        color: "RGBColor | None" = None,
        space: "Length | None" = None,
    ) -> PageBorder:
        """Set properties of a single page-border edge in one call.

        `side` must be one of ``"top"``, ``"bottom"``, ``"left"``, or ``"right"``.
        Any of `style`, `width`, `color`, `space` may be |None| (their existing value
        is left unchanged when already present; any argument explicitly set is
        applied to the edge). Returns the updated |PageBorder| proxy.

        .. versionadded:: 2026.05.0
        """
        if side not in ("top", "bottom", "left", "right"):
            raise ValueError(
                "side must be one of 'top', 'bottom', 'left', 'right'; got %r" % side
            )
        border = getattr(self.page_borders, side)
        if style is not None:
            border.style = style
        if width is not None:
            border.width = width
        if color is not None:
            border.color = color
        if space is not None:
            border.space = space
        return border

    def remove_page_borders(self) -> None:
        """Remove any ``w:pgBorders`` element from this section's ``w:sectPr``.

        Does nothing when no ``w:pgBorders`` child is present.

        .. versionadded:: 2026.05.0
        """
        self._sectPr._remove_pgBorders()  # pyright: ignore[reportPrivateUsage]

    @property
    def line_numbering(self) -> LineNumbering | None:
        """|LineNumbering| proxy or |None| when no ``w:lnNumType`` child is present.

        Line numbering is displayed in the margin alongside each numbered line.
        Use :meth:`set_line_numbering` to create or update the ``w:lnNumType``
        element and :meth:`remove_line_numbering` to remove it.

        .. versionadded:: 2026.05.0
        """
        lnNumType = self._sectPr.lnNumType
        if lnNumType is None:
            return None
        return LineNumbering(lnNumType)

    def set_line_numbering(
        self,
        count_by: int | None = None,
        start: int | None = None,
        distance: "Length | None" = None,
        restart: "WD_LINE_NUMBERING_RESTART | None" = None,
    ) -> LineNumbering:
        """Create or update this section's ``w:lnNumType`` with provided values.

        Any argument left as |None| leaves the corresponding attribute on an
        existing ``w:lnNumType`` element unchanged. Returns the |LineNumbering|
        proxy for the resulting element.

        .. versionadded:: 2026.05.0
        """
        lnNumType = self._sectPr.get_or_add_lnNumType()
        if count_by is not None:
            lnNumType.countBy = count_by
        if start is not None:
            lnNumType.start = start
        if distance is not None:
            lnNumType.distance = distance
        if restart is not None:
            lnNumType.restart = restart
        return LineNumbering(lnNumType)

    def remove_line_numbering(self) -> None:
        """Remove any ``w:lnNumType`` element from this section's ``w:sectPr``.

        Does nothing when no ``w:lnNumType`` child is present.

        .. versionadded:: 2026.05.0
        """
        self._sectPr._remove_lnNumType()  # pyright: ignore[reportPrivateUsage]

    @property
    def first_page_paper_source(self) -> int | None:
        """Read/write. Printer tray bin number to use for the first page of this section.

        Returns the ``w:first`` attribute of the ``w:paperSrc`` child of this
        section's ``w:sectPr``. |None| when no ``w:paperSrc`` child is present or
        the attribute is unset.

        Setting this to |None| clears the attribute; if ``other_pages_paper_source``
        is also unset, the ``w:paperSrc`` element is removed entirely.

        .. versionadded:: 2026.05.0
        """
        paperSrc = self._sectPr.paperSrc
        if paperSrc is None:
            return None
        return paperSrc.first

    @first_page_paper_source.setter
    def first_page_paper_source(self, value: int | None) -> None:
        if value is None:
            paperSrc = self._sectPr.paperSrc
            if paperSrc is None:
                return
            paperSrc.first = None
            if paperSrc.other is None:
                self._sectPr._remove_paperSrc()  # pyright: ignore[reportPrivateUsage]
            return
        self._sectPr.get_or_add_paperSrc().first = value

    @property
    def other_pages_paper_source(self) -> int | None:
        """Read/write. Printer tray bin number for non-first pages of this section.

        Returns the ``w:other`` attribute of the ``w:paperSrc`` child of this
        section's ``w:sectPr``. |None| when no ``w:paperSrc`` child is present or
        the attribute is unset.

        Setting this to |None| clears the attribute; if ``first_page_paper_source``
        is also unset, the ``w:paperSrc`` element is removed entirely.

        .. versionadded:: 2026.05.0
        """
        paperSrc = self._sectPr.paperSrc
        if paperSrc is None:
            return None
        return paperSrc.other

    @other_pages_paper_source.setter
    def other_pages_paper_source(self, value: int | None) -> None:
        if value is None:
            paperSrc = self._sectPr.paperSrc
            if paperSrc is None:
                return
            paperSrc.other = None
            if paperSrc.first is None:
                self._sectPr._remove_paperSrc()  # pyright: ignore[reportPrivateUsage]
            return
        self._sectPr.get_or_add_paperSrc().other = value

    @property
    def document_grid(self) -> DocumentGrid | None:
        """|DocumentGrid| proxy or |None| when no ``w:docGrid`` child is present.

        The document grid controls the East Asian character grid for this section.
        Use :meth:`set_document_grid` to create or update the ``w:docGrid`` element
        and :meth:`remove_document_grid` to remove it.

        .. versionadded:: 2026.05.0
        """
        docGrid = self._sectPr.docGrid
        if docGrid is None:
            return None
        return DocumentGrid(docGrid)

    def set_document_grid(
        self,
        type: "WD_DOC_GRID_TYPE | None" = None,
        line_pitch: int | None = None,
        char_space: int | None = None,
    ) -> DocumentGrid:
        """Create or update this section's ``w:docGrid`` with provided values.

        Any argument left as |None| leaves the corresponding attribute on an
        existing ``w:docGrid`` element unchanged. Returns the |DocumentGrid|
        proxy for the resulting element.

        .. versionadded:: 2026.05.0
        """
        docGrid = self._sectPr.get_or_add_docGrid()
        if type is not None:
            docGrid.type = type
        if line_pitch is not None:
            docGrid.linePitch = line_pitch
        if char_space is not None:
            docGrid.charSpace = char_space
        return DocumentGrid(docGrid)

    def remove_document_grid(self) -> None:
        """Remove any ``w:docGrid`` element from this section's ``w:sectPr``.

        Does nothing when no ``w:docGrid`` child is present.

        .. versionadded:: 2026.05.0
        """
        self._sectPr._remove_docGrid()  # pyright: ignore[reportPrivateUsage]

    @property
    def right_margin(self) -> Length | None:
        """|Length| object representing the right margin for all pages in this section
        in English Metric Units."""
        return self._sectPr.right_margin

    @right_margin.setter
    def right_margin(self, value: Length | None):
        self._sectPr.right_margin = value

    @property
    def right_to_left(self) -> bool:
        """Read/write. ``True`` when this section uses right-to-left text flow.

        Reflects the presence of the ``w:bidi`` child of this section's
        ``w:sectPr``. Returns ``False`` when no ``w:bidi`` element is present or
        its ``w:val`` attribute evaluates to false.

        Assigning ``True`` inserts a ``w:bidi`` element; assigning ``False`` (or
        |None|) removes any existing ``w:bidi`` child.

        .. versionadded:: 2026.05.0
        """
        return self._sectPr.bidi_val

    @right_to_left.setter
    def right_to_left(self, value: bool | None):
        self._sectPr.bidi_val = value

    @property
    def text_direction(self) -> "WD_TEXT_DIRECTION | None":
        """Read/write. Text-flow direction for this section, as a :ref:`WdTextDirection`.

        Maps to the ``w:val`` attribute of the ``w:textDirection`` child of this
        section's ``w:sectPr``. |None| when no ``w:textDirection`` child is
        present. Assigning |None| removes the ``w:textDirection`` element.

        .. versionadded:: 2026.05.0
        """
        return self._sectPr.text_direction

    @text_direction.setter
    def text_direction(self, value: "WD_TEXT_DIRECTION | None"):
        self._sectPr.text_direction = value

    @property
    def start_type(self) -> WD_SECTION_START:
        """Type of page-break (if any) inserted at the start of this section.

        For exmple, ``WD_SECTION_START.ODD_PAGE`` if the section should begin on the
        next odd page, possibly inserting two page-breaks instead of one.
        """
        return self._sectPr.start_type

    @start_type.setter
    def start_type(self, value: WD_SECTION_START | None):
        self._sectPr.start_type = value

    @property
    def top_margin(self) -> Length | None:
        """|Length| object representing the top margin for all pages in this section in
        English Metric Units."""
        return self._sectPr.top_margin

    @top_margin.setter
    def top_margin(self, value: Length | None):
        self._sectPr.top_margin = value

    @property
    def footnote_properties(self) -> FootnoteProperties | None:
        """A |FootnoteProperties| object or |None| when no ``w:footnotePr`` child is
        present on this section's ``w:sectPr``.

        Section-level footnote properties override document-level defaults. Use
        :meth:`add_footnote_properties` to add a ``w:footnotePr`` child if not present.

        .. versionadded:: 2026.05.0
        """
        from docx.footnotes import FootnoteProperties

        footnotePr = self._sectPr.footnotePr
        if footnotePr is None:
            return None
        return FootnoteProperties(footnotePr)

    def add_footnote_properties(self) -> FootnoteProperties:
        """Return a |FootnoteProperties| proxy, adding ``w:footnotePr`` if needed.

        If a ``w:footnotePr`` element is already present on this section, the existing
        element is used.

        .. versionadded:: 2026.05.0
        """
        from docx.footnotes import FootnoteProperties

        footnotePr = self._sectPr.get_or_add_footnotePr()
        return FootnoteProperties(footnotePr)

    def remove_footnote_properties(self) -> None:
        """Remove the ``w:footnotePr`` child element if present.

        .. versionadded:: 2026.05.0
        """
        self._sectPr._remove_footnotePr()  # pyright: ignore[reportPrivateUsage]

    @property
    def endnote_properties(self) -> EndnoteProperties | None:
        """An |EndnoteProperties| object or |None| when no ``w:endnotePr`` child is
        present on this section's ``w:sectPr``.

        Section-level endnote properties override document-level defaults. Use
        :meth:`add_endnote_properties` to add a ``w:endnotePr`` child if not present.

        .. versionadded:: 2026.05.0
        """
        from docx.endnotes import EndnoteProperties

        endnotePr = self._sectPr.endnotePr
        if endnotePr is None:
            return None
        return EndnoteProperties(endnotePr)

    def add_endnote_properties(self) -> EndnoteProperties:
        """Return an |EndnoteProperties| proxy, adding ``w:endnotePr`` if needed.

        If a ``w:endnotePr`` element is already present on this section, the existing
        element is used.

        .. versionadded:: 2026.05.0
        """
        from docx.endnotes import EndnoteProperties

        endnotePr = self._sectPr.get_or_add_endnotePr()
        return EndnoteProperties(endnotePr)

    def remove_endnote_properties(self) -> None:
        """Remove the ``w:endnotePr`` child element if present.

        .. versionadded:: 2026.05.0
        """
        self._sectPr._remove_endnotePr()  # pyright: ignore[reportPrivateUsage]

    # -- watermark API ---------------------------------------------------------------

    def add_text_watermark(
        self,
        text: str,
        font: str = "Calibri",
        size: Length | None = None,
        color: RGBColor | None = None,
        layout: str = "diagonal",
    ) -> Watermark:
        """Add a text watermark to this section's default page header.

        Replaces any existing watermark. Returns the |Watermark| proxy for the
        newly-added shape.

        `size` defaults to 72pt, `color` to silver (``#C0C0C0``), `layout` to
        ``"diagonal"``. `layout` accepts ``"diagonal"`` or ``"horizontal"``.

        .. versionadded:: 2026.05.0
        """
        if size is None:
            size = Pt(72)
        if color is None:
            color = RGBColor(0xC0, 0xC0, 0xC0)
        if layout not in ("diagonal", "horizontal"):
            raise ValueError(
                "layout must be 'diagonal' or 'horizontal', got %r" % layout
            )

        # -- ensure the section has a non-linked header ---
        if self.header.is_linked_to_previous:
            self.header.is_linked_to_previous = False
        hdr = self.header._element

        # -- remove any existing watermark first --
        self._remove_watermark_paragraphs(hdr)

        # -- build the watermark paragraph --
        font_pt = float(size) / 12700.0  # EMU -> pt (12700 EMU/pt)
        rotation = "-45" if layout == "diagonal" else "0"
        style = (
            "position:absolute;margin-left:0;margin-top:0;"
            "width:%.2fpt;height:%.2fpt;z-index:-251654144;"
            "mso-position-horizontal:center;mso-position-horizontal-relative:margin;"
            "mso-position-vertical:center;mso-position-vertical-relative:margin;"
            "rotation:%s" % (font_pt * len(text) * 0.6 + 100, font_pt * 1.5, rotation)
        )
        text_escaped = _xml_escape_attr(text)
        font_escaped = _xml_escape_attr(font)
        p_xml = (
            "<w:p %s>"
            "<w:pPr><w:pStyle w:val=\"Header\"/></w:pPr>"
            "<w:r>"
            "<w:pict>"
            "<v:shape id=\"PowerPlusWaterMarkObject\" type=\"#_x0000_t136\""
            " style=\"%s\">"
            "<v:fill color=\"#%s\"/>"
            "<v:textpath style=\"font:%.2fpt &quot;%s&quot;\" string=\"%s\"/>"
            "</v:shape>"
            "</w:pict>"
            "</w:r>"
            "</w:p>"
            % (
                nsdecls("w", "v", "o", "w10"),
                style,
                str(color),
                font_pt,
                font_escaped,
                text_escaped,
            )
        )
        p = parse_xml(p_xml)
        hdr.append(p)
        shape = p.find(".//" + qn("v:shape"))
        assert shape is not None
        return Watermark(shape)

    def add_image_watermark(
        self,
        image_path: str | IO[bytes],
        width: Length | None = None,
        height: Length | None = None,
    ) -> Watermark:
        """Add an image watermark to this section's default page header.

        `image_path` can be a filesystem path or a file-like object. `width`
        and `height` are |Length| values; when omitted the image's native
        dimensions are used.

        Replaces any existing watermark.

        .. versionadded:: 2026.05.0
        """
        if self.header.is_linked_to_previous:
            self.header.is_linked_to_previous = False
        header_part = self.header.part
        assert isinstance(header_part, HeaderPart)
        hdr = self.header._element

        # -- remove any existing watermark first --
        self._remove_watermark_paragraphs(hdr)

        rId, image = header_part.get_or_add_image(image_path)
        cx, cy = image.scaled_dimensions(width, height)
        # -- VML uses points, not EMU; 12700 EMU per pt --
        w_pt = float(cx) / 12700.0
        h_pt = float(cy) / 12700.0
        style = (
            "position:absolute;margin-left:0;margin-top:0;"
            "width:%.2fpt;height:%.2fpt;z-index:-251654144;"
            "mso-position-horizontal:center;mso-position-horizontal-relative:margin;"
            "mso-position-vertical:center;mso-position-vertical-relative:margin"
            % (w_pt, h_pt)
        )
        p_xml = (
            "<w:p %s>"
            "<w:pPr><w:pStyle w:val=\"Header\"/></w:pPr>"
            "<w:r>"
            "<w:pict>"
            "<v:shape id=\"PowerPlusWaterMarkObject\" type=\"#_x0000_t75\""
            " style=\"%s\">"
            "<v:imagedata r:id=\"%s\" o:title=\"watermark\"/>"
            "</v:shape>"
            "</w:pict>"
            "</w:r>"
            "</w:p>"
            % (
                nsdecls("w", "v", "o", "w10", "r"),
                style,
                rId,
            )
        )
        p = parse_xml(p_xml)
        hdr.append(p)
        shape = p.find(".//" + qn("v:shape"))
        assert shape is not None
        return Watermark(shape)

    def remove_watermark(self) -> None:
        """Remove the watermark from this section's default page header.

        Does nothing when the section has no header definition or when the
        header contains no watermark.

        .. versionadded:: 2026.05.0
        """
        if self.header.is_linked_to_previous:
            return
        hdr = self.header._element
        self._remove_watermark_paragraphs(hdr)

    @property
    def watermark(self) -> Watermark | None:
        """|Watermark| object if this section's header contains one, else ``None``.

        .. versionadded:: 2026.05.0
        """
        if self.header.is_linked_to_previous:
            return None
        hdr = self.header._element
        shape = self._find_watermark_shape(hdr)
        if shape is None:
            return None
        return Watermark(shape)

    @staticmethod
    def _find_watermark_shape(hdr) -> "CT_VmlShape | None":
        """Return the first ``v:shape`` element inside a ``w:pict`` in `hdr`."""
        picts = hdr.findall(".//" + qn("w:pict"))
        for pict in picts:
            shape = pict.find(qn("v:shape"))
            if shape is not None:
                return shape
        return None

    @staticmethod
    def _remove_watermark_paragraphs(hdr) -> None:
        """Remove paragraphs from `hdr` that contain a watermark ``v:shape``."""
        for p in list(hdr.findall(qn("w:p"))):
            if p.find(".//" + qn("v:shape")) is not None:
                hdr.remove(p)


def _xml_escape_attr(value: str) -> str:
    """Escape characters that would break an XML attribute value."""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


class Sections(Sequence[Section]):
    """Sequence of |Section| objects corresponding to the sections in the document.

    Supports ``len()``, iteration, and indexed access.
    """

    def __init__(self, document_elm: CT_Document, document_part: DocumentPart):
        super().__init__()
        self._document_elm = document_elm
        self._document_part = document_part

    @overload
    def __getitem__(self, key: int) -> Section: ...

    @overload
    def __getitem__(self, key: slice) -> list[Section]: ...

    def __getitem__(self, key: int | slice) -> Section | list[Section]:
        if isinstance(key, slice):
            return [
                Section(sectPr, self._document_part)
                for sectPr in self._document_elm.sectPr_lst[key]
            ]
        return Section(self._document_elm.sectPr_lst[key], self._document_part)

    def __iter__(self) -> Iterator[Section]:
        for sectPr in self._document_elm.sectPr_lst:
            yield Section(sectPr, self._document_part)

    def __len__(self) -> int:
        return len(self._document_elm.sectPr_lst)

    def pop(self, index: int = -1) -> Section:
        """Remove and return the section at `index`, merging it into its neighbour.

        Delegates to :meth:`Section.delete`. Returns the (now-defunct) |Section|
        proxy that was removed, for symmetry with :meth:`list.pop` -- callers
        sometimes want to inspect section properties before the break is
        merged away, but should not attempt to further mutate the returned
        object after this call.

        Raises :class:`IndexError` when `index` is out of range.

        .. versionadded:: 2026.05.0
        """
        sectPrs = self._document_elm.sectPr_lst
        # -- normalize negative index and bounds-check --
        n = len(sectPrs)
        if index < 0:
            index += n
        if not 0 <= index < n:
            raise IndexError("section index out of range")
        section = Section(sectPrs[index], self._document_part)
        section.delete()
        return section


class Column:
    """Proxy for a ``<w:col>`` element, representing an individual column definition.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, col: CT_Col):
        self._col = col

    @property
    def space(self) -> Length | None:
        """Read/write. Space after this column, in EMU.

        |None| when no space value has been specified.

        .. versionadded:: 2026.05.0
        """
        return self._col.space

    @space.setter
    def space(self, value: Length | None):
        self._col.space = value

    @property
    def width(self) -> Length | None:
        """Read/write. Width of this column, in EMU.

        |None| when no width has been specified.

        .. versionadded:: 2026.05.0
        """
        return self._col.w

    @width.setter
    def width(self, value: Length | None):
        self._col.w = value


class SectionColumns(Sequence[Column]):
    """Proxy for a ``<w:cols>`` element, providing access to column layout settings.

    Supports indexed access to individual |Column| objects when ``equal_width`` is False.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, sectPr: CT_SectPr):
        self._sectPr = sectPr

    @overload
    def __getitem__(self, key: int) -> Column: ...

    @overload
    def __getitem__(self, key: slice) -> list[Column]: ...

    def __getitem__(self, key: int | slice) -> Column | list[Column]:
        cols = self._sectPr.cols
        col_lst = cols.col_lst if cols is not None else []
        if isinstance(key, slice):
            return [Column(col) for col in col_lst[key]]
        return Column(col_lst[key])

    def __iter__(self) -> Iterator[Column]:
        cols = self._sectPr.cols
        if cols is not None:
            for col in cols.col_lst:
                yield Column(col)

    def __len__(self) -> int:
        cols = self._sectPr.cols
        if cols is None:
            return 0
        return len(cols.col_lst)

    @property
    def count(self) -> int:
        """Read/write. Number of columns in this section.

        Defaults to 1 when no ``w:cols`` element is present or when ``w:num`` attribute
        is not specified.

        .. versionadded:: 2026.05.0
        """
        cols = self._sectPr.cols
        if cols is None:
            return 1
        return cols.num if cols.num is not None else 1

    @count.setter
    def count(self, value: int):
        cols = self._sectPr.get_or_add_cols()
        cols.num = value

    @property
    def equal_width(self) -> bool:
        """Read/write. True when all columns have equal width.

        Defaults to True when no ``w:cols`` element is present or when ``w:equalWidth``
        attribute is not specified.

        .. versionadded:: 2026.05.0
        """
        cols = self._sectPr.cols
        if cols is None:
            return True
        return cols.equalWidth if cols.equalWidth is not None else True

    @equal_width.setter
    def equal_width(self, value: bool):
        cols = self._sectPr.get_or_add_cols()
        cols.equalWidth = value

    @property
    def space(self) -> Length | None:
        """Read/write. Default space between columns, in EMU.

        |None| when no ``w:cols`` element is present or no ``w:space`` attribute is set.

        .. versionadded:: 2026.05.0
        """
        cols = self._sectPr.cols
        if cols is None:
            return None
        return cols.space

    @space.setter
    def space(self, value: Length | None):
        cols = self._sectPr.get_or_add_cols()
        cols.space = value


class PageBorder:
    """Proxy for a single page-border edge on a ``w:pgBorders`` element.

    Accessed via |PageBorders| side properties, e.g. ``section.page_borders.top``.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, sectPr: CT_SectPr, side: str):
        self._sectPr = sectPr
        self._side = side

    @property
    def _border_elm(self) -> "CT_Border | None":
        pgBorders = self._sectPr.pgBorders
        if pgBorders is None:
            return None
        return getattr(pgBorders, self._side)

    def _get_or_add_border_elm(self) -> "CT_Border":
        pgBorders = self._sectPr.get_or_add_pgBorders()
        return getattr(pgBorders, f"get_or_add_{self._side}")()

    @property
    def style(self) -> "WD_BORDER_STYLE | None":
        """Read/write. Border style as a :ref:`WdBorderStyle` member.

        |None| when the edge element is not present or has no ``w:val`` attribute.

        .. versionadded:: 2026.05.0
        """
        border = self._border_elm
        if border is None:
            return None
        return border.val

    @style.setter
    def style(self, value: "WD_BORDER_STYLE | None") -> None:
        if value is None:
            border = self._border_elm
            if border is not None:
                border.val = None
            return
        self._get_or_add_border_elm().val = value

    @property
    def width(self) -> "Length | None":
        """Read/write. Border width as a |Length|, stored in eighths of a point.

        |None| when not present.

        .. versionadded:: 2026.05.0
        """
        border = self._border_elm
        if border is None:
            return None
        return border.sz

    @width.setter
    def width(self, value: "Length | None") -> None:
        if value is None:
            border = self._border_elm
            if border is not None:
                border.sz = None
            return
        self._get_or_add_border_elm().sz = value

    @property
    def color(self) -> RGBColor | None:
        """Read/write. Border color as an |RGBColor|.

        An ``"auto"`` value in the XML is returned as |None|. |None| when no color is
        specified on the edge element.

        .. versionadded:: 2026.05.0
        """
        border = self._border_elm
        if border is None:
            return None
        color = border.color
        if isinstance(color, str):
            return None
        return color

    @color.setter
    def color(self, value: RGBColor | None) -> None:
        if value is None:
            border = self._border_elm
            if border is not None:
                border.color = None
            return
        self._get_or_add_border_elm().color = value

    @property
    def space(self) -> "Length | None":
        """Read/write. Distance from page/text edge to border, as |Length| (points).

        |None| when not specified on the edge element.

        .. versionadded:: 2026.05.0
        """
        border = self._border_elm
        if border is None:
            return None
        return border.space

    @space.setter
    def space(self, value: "Length | None") -> None:
        if value is None:
            border = self._border_elm
            if border is not None:
                border.space = None
            return
        self._get_or_add_border_elm().space = value


class PageBorders:
    """Proxy for the ``<w:pgBorders>`` element of a section.

    Accessed via :attr:`Section.page_borders`. Provides read/write access to each
    of the four edge borders plus the ``display`` and ``offset_from`` attributes.
    The underlying ``w:pgBorders`` element is created lazily on first write.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, sectPr: CT_SectPr):
        self._sectPr = sectPr

    @property
    def _pgBorders(self) -> "CT_PgBorders | None":
        return self._sectPr.pgBorders

    @property
    def top(self) -> PageBorder:
        """The |PageBorder| for the top edge of the page.

        .. versionadded:: 2026.05.0
        """
        return PageBorder(self._sectPr, "top")

    @property
    def bottom(self) -> PageBorder:
        """The |PageBorder| for the bottom edge of the page.

        .. versionadded:: 2026.05.0
        """
        return PageBorder(self._sectPr, "bottom")

    @property
    def left(self) -> PageBorder:
        """The |PageBorder| for the left edge of the page.

        .. versionadded:: 2026.05.0
        """
        return PageBorder(self._sectPr, "left")

    @property
    def right(self) -> PageBorder:
        """The |PageBorder| for the right edge of the page.

        .. versionadded:: 2026.05.0
        """
        return PageBorder(self._sectPr, "right")

    @property
    def display(self) -> "WD_BORDER_DISPLAY | None":
        """Read/write. Member of :class:`WD_BORDER_DISPLAY` or |None|.

        Reads the ``w:display`` attribute of the ``w:pgBorders`` element. |None|
        when no ``w:pgBorders`` element is present or the attribute is unset.

        .. versionadded:: 2026.05.0
        """
        pgBorders = self._pgBorders
        if pgBorders is None:
            return None
        return pgBorders.display

    @display.setter
    def display(self, value: "WD_BORDER_DISPLAY | None") -> None:
        if value is None:
            pgBorders = self._pgBorders
            if pgBorders is not None:
                pgBorders.display = None
            return
        self._sectPr.get_or_add_pgBorders().display = value

    @property
    def offset_from(self) -> "WD_BORDER_OFFSET_FROM | None":
        """Read/write. Member of :class:`WD_BORDER_OFFSET_FROM` or |None|.

        Reads the ``w:offsetFrom`` attribute of the ``w:pgBorders`` element.
        |None| when no ``w:pgBorders`` element is present or the attribute is
        unset.

        .. versionadded:: 2026.05.0
        """
        pgBorders = self._pgBorders
        if pgBorders is None:
            return None
        return pgBorders.offsetFrom

    @offset_from.setter
    def offset_from(self, value: "WD_BORDER_OFFSET_FROM | None") -> None:
        if value is None:
            pgBorders = self._pgBorders
            if pgBorders is not None:
                pgBorders.offsetFrom = None
            return
        self._sectPr.get_or_add_pgBorders().offsetFrom = value


class LineNumbering:
    """Proxy for a ``<w:lnNumType>`` element on a section's ``w:sectPr``.

    Accessed via :attr:`Section.line_numbering`. Provides read/write access to
    the ``countBy``, ``start``, ``distance`` and ``restart`` attributes.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, lnNumType: "CT_LineNumber"):
        self._lnNumType = lnNumType

    @property
    def count_by(self) -> int | None:
        """Read/write. Interval between displayed line numbers.

        A value of ``N`` means only every ``Nth`` line is numbered. |None| when
        the ``w:countBy`` attribute is not specified on the element.

        .. versionadded:: 2026.05.0
        """
        return self._lnNumType.countBy

    @count_by.setter
    def count_by(self, value: int | None) -> None:
        self._lnNumType.countBy = value

    @property
    def start(self) -> int | None:
        """Read/write. Starting line number for this section.

        |None| when the ``w:start`` attribute is not specified.

        .. versionadded:: 2026.05.0
        """
        return self._lnNumType.start

    @start.setter
    def start(self, value: int | None) -> None:
        self._lnNumType.start = value

    @property
    def distance(self) -> "Length | None":
        """Read/write. Distance from the text to the line numbers as |Length|.

        |None| when the ``w:distance`` attribute is not specified.

        .. versionadded:: 2026.05.0
        """
        return self._lnNumType.distance

    @distance.setter
    def distance(self, value: "Length | None") -> None:
        self._lnNumType.distance = value

    @property
    def restart(self) -> "WD_LINE_NUMBERING_RESTART | None":
        """Read/write. |WD_LINE_NUMBERING_RESTART| member or |None|.

        Controls when the line-number counter restarts: ``CONTINUOUS``,
        ``NEW_SECTION``, or ``NEW_PAGE``. |None| when the ``w:restart`` attribute
        is not specified.

        .. versionadded:: 2026.05.0
        """
        return self._lnNumType.restart

    @restart.setter
    def restart(self, value: "WD_LINE_NUMBERING_RESTART | None") -> None:
        self._lnNumType.restart = value


class DocumentGrid:
    """Proxy for a ``<w:docGrid>`` element on a section's ``w:sectPr``.

    Accessed via :attr:`Section.document_grid`. Provides read/write access to the
    ``type``, ``linePitch`` and ``charSpace`` attributes, which control the East
    Asian character grid for the section.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, docGrid: "CT_DocGrid"):
        self._docGrid = docGrid

    @property
    def type(self) -> "WD_DOC_GRID_TYPE | None":
        """Read/write. |WD_DOC_GRID_TYPE| member or |None|.

        Controls the document grid type: ``DEFAULT``, ``LINES``, ``LINES_AND_CHARS``,
        or ``SNAP_TO_CHARS``. |None| when the ``w:type`` attribute is not specified.

        .. versionadded:: 2026.05.0
        """
        return self._docGrid.type

    @type.setter
    def type(self, value: "WD_DOC_GRID_TYPE | None") -> None:
        self._docGrid.type = value

    @property
    def line_pitch(self) -> int | None:
        """Read/write. Line pitch (lines per page height unit) as an integer.

        Maps to the ``w:linePitch`` attribute. |None| when the attribute is not
        specified.

        .. versionadded:: 2026.05.0
        """
        return self._docGrid.linePitch

    @line_pitch.setter
    def line_pitch(self, value: int | None) -> None:
        self._docGrid.linePitch = value

    @property
    def char_space(self) -> int | None:
        """Read/write. Additional character spacing in 1/1024pt units, as an integer.

        Maps to the ``w:charSpace`` attribute. |None| when the attribute is not
        specified.

        .. versionadded:: 2026.05.0
        """
        return self._docGrid.charSpace

    @char_space.setter
    def char_space(self, value: int | None) -> None:
        self._docGrid.charSpace = value


class _BaseHeaderFooter(BlockItemContainer):
    """Base class for header and footer classes."""

    def __init__(
        self,
        sectPr: CT_SectPr,
        document_part: DocumentPart,
        header_footer_index: WD_HEADER_FOOTER,
    ):
        self._sectPr = sectPr
        self._document_part = document_part
        self._hdrftr_index = header_footer_index

    @property
    def is_linked_to_previous(self) -> bool:
        """``True`` if this header/footer uses the definition from the prior section.

        ``False`` if this header/footer has an explicit definition.

        Assigning ``True`` to this property removes the header/footer definition for
        this section, causing it to "inherit" the corresponding definition of the prior
        section. Assigning ``False`` causes a new, empty definition to be added for this
        section, but only if no definition is already present.
        """
        # ---absence of a header/footer part indicates "linked" behavior---
        return not self._has_definition

    @is_linked_to_previous.setter
    def is_linked_to_previous(self, value: bool) -> None:
        new_state = bool(value)
        # ---do nothing when value is not being changed---
        if new_state == self.is_linked_to_previous:
            return
        if new_state is True:
            self._drop_definition()
        else:
            self._add_definition()

    @property
    def part(self) -> HeaderPart | FooterPart:
        """The |HeaderPart| or |FooterPart| for this header/footer.

        This overrides `BlockItemContainer.part` and is required to support image
        insertion and perhaps other content like hyperlinks.
        """
        # ---should not appear in documentation;
        # ---not an interface property, even though public
        return self._get_or_add_definition()

    def _add_definition(self) -> HeaderPart | FooterPart:
        """Return newly-added header/footer part."""
        raise NotImplementedError("must be implemented by each subclass")

    @property
    def _definition(self) -> HeaderPart | FooterPart:
        """|HeaderPart| or |FooterPart| object containing header/footer content."""
        raise NotImplementedError("must be implemented by each subclass")

    def _drop_definition(self) -> None:
        """Remove header/footer part containing the definition of this header/footer."""
        raise NotImplementedError("must be implemented by each subclass")

    @property
    def _element(self):
        """`w:hdr` or `w:ftr` element, root of header/footer part."""
        return self._get_or_add_definition().element

    def _get_or_add_definition(self) -> HeaderPart | FooterPart:
        """Return HeaderPart or FooterPart object for this section.

        If this header/footer inherits its content, the part for the prior header/footer
        is returned; this process continue recursively until a definition is found. If
        the definition cannot be inherited (because the header/footer belongs to the
        first section), a new definition is added for that first section and then
        returned.
        """
        # ---note this method is called recursively to access inherited definitions---
        # ---case-1: definition is not inherited---
        if self._has_definition:
            return self._definition
        # ---case-2: definition is inherited and belongs to second-or-later section---
        prior_headerfooter = self._prior_headerfooter
        if prior_headerfooter:
            return prior_headerfooter._get_or_add_definition()
        # ---case-3: definition is inherited, but belongs to first section---
        return self._add_definition()

    @property
    def _has_definition(self) -> bool:
        """True if this header/footer has a related part containing its definition."""
        raise NotImplementedError("must be implemented by each subclass")

    @property
    def _prior_headerfooter(self) -> _Header | _Footer | None:
        """|_Header| or |_Footer| proxy on prior sectPr element.

        Returns None if this is first section.
        """
        raise NotImplementedError("must be implemented by each subclass")


class _Footer(_BaseHeaderFooter):
    """Page footer, used for all three types (default, even-page, and first-page).

    Note that, like a document or table cell, a footer must contain a minimum of one
    paragraph and a new or otherwise "empty" footer contains a single empty paragraph.
    This first paragraph can be accessed as `footer.paragraphs[0]` for purposes of
    adding content to it. Using :meth:`add_paragraph()` by itself to add content will
    leave an empty paragraph above the newly added one.
    """

    def _add_definition(self) -> FooterPart:
        """Return newly-added footer part."""
        footer_part, rId = self._document_part.add_footer_part()
        self._sectPr.add_footerReference(self._hdrftr_index, rId)
        return footer_part

    @property
    def _definition(self):
        """|FooterPart| object containing content of this footer."""
        footerReference = self._sectPr.get_footerReference(self._hdrftr_index)
        # -- currently this is never called when `._has_definition` evaluates False --
        assert footerReference is not None
        return self._document_part.footer_part(footerReference.rId)

    def _drop_definition(self):
        """Remove footer definition (footer part) associated with this section."""
        rId = self._sectPr.remove_footerReference(self._hdrftr_index)
        self._document_part.drop_rel(rId)

    @property
    def _has_definition(self) -> bool:
        """True if a footer is defined for this section."""
        footerReference = self._sectPr.get_footerReference(self._hdrftr_index)
        return footerReference is not None

    @property
    def _prior_headerfooter(self):
        """|_Footer| proxy on prior sectPr element or None if this is first section."""
        preceding_sectPr = self._sectPr.preceding_sectPr
        return (
            None
            if preceding_sectPr is None
            else _Footer(preceding_sectPr, self._document_part, self._hdrftr_index)
        )


class _Header(_BaseHeaderFooter):
    """Page header, used for all three types (default, even-page, and first-page).

    Note that, like a document or table cell, a header must contain a minimum of one
    paragraph and a new or otherwise "empty" header contains a single empty paragraph.
    This first paragraph can be accessed as `header.paragraphs[0]` for purposes of
    adding content to it. Using :meth:`add_paragraph()` by itself to add content will
    leave an empty paragraph above the newly added one.
    """

    def _add_definition(self):
        """Return newly-added header part."""
        header_part, rId = self._document_part.add_header_part()
        self._sectPr.add_headerReference(self._hdrftr_index, rId)
        return header_part

    @property
    def _definition(self):
        """|HeaderPart| object containing content of this header."""
        headerReference = self._sectPr.get_headerReference(self._hdrftr_index)
        # -- currently this is never called when `._has_definition` evaluates False --
        assert headerReference is not None
        return self._document_part.header_part(headerReference.rId)

    def _drop_definition(self):
        """Remove header definition associated with this section."""
        rId = self._sectPr.remove_headerReference(self._hdrftr_index)
        self._document_part.drop_header_part(rId)

    @property
    def _has_definition(self) -> bool:
        """True if a header is explicitly defined for this section."""
        headerReference = self._sectPr.get_headerReference(self._hdrftr_index)
        return headerReference is not None

    @property
    def _prior_headerfooter(self):
        """|_Header| proxy on prior sectPr element or None if this is first section."""
        preceding_sectPr = self._sectPr.preceding_sectPr
        return (
            None
            if preceding_sectPr is None
            else _Header(preceding_sectPr, self._document_part, self._hdrftr_index)
        )
