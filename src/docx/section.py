"""The |Section| object and related proxy classes."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterator, List, Sequence, Union, overload

from docx.blkcntnr import BlockItemContainer
from docx.enum.section import WD_HEADER_FOOTER
from docx.oxml.text.paragraph import CT_P
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.shared import Length, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    from docx.enum.section import WD_ORIENTATION, WD_SECTION_START
    from docx.oxml.document import CT_Document
    from docx.oxml.section import CT_Col, CT_Cols, CT_SectPr
    from docx.parts.document import DocumentPart
    from docx.parts.story import StoryPart


class Section:
    """Document section, providing access to section and page setup settings.

    Also provides access to headers and footers.
    """

    def __init__(self, sectPr: CT_SectPr, document_part: DocumentPart):
        super(Section, self).__init__()
        self._sectPr = sectPr
        self._document_part = document_part

    @property
    def columns(self) -> SectionColumns:
        """|SectionColumns| object providing access to column layout settings."""
        return SectionColumns(self._sectPr)

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
    def even_page_footer(self) -> _Footer:
        """|_Footer| object defining footer content for even pages.

        The content of this footer definition is ignored unless the document setting
        :attr:`~.Settings.odd_and_even_pages_header_footer` is set True.
        """
        return _Footer(self._sectPr, self._document_part, WD_HEADER_FOOTER.EVEN_PAGE)

    @property
    def even_page_header(self) -> _Header:
        """|_Header| object defining header content for even pages.

        The content of this header definition is ignored unless the document setting
        :attr:`~.Settings.odd_and_even_pages_header_footer` is set True.
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

    def add_text_watermark(
        self,
        text: str,
        font: str = "Calibri",
        size: Length | int = Pt(72),
        color: RGBColor = RGBColor(0xC0, 0xC0, 0xC0),
        layout: str = "diagonal",
    ) -> None:
        """Add a text watermark to this section.

        The watermark is rendered as a VML shape in the default header. Any existing
        watermark is removed before the new one is added.

        Args:
            text: The watermark text to display.
            font: Font family name (default ``"Calibri"``).
            size: Font size as a |Length| value (default ``Pt(72)``).
            color: Text color as an |RGBColor| value (default silver).
            layout: Either ``"diagonal"`` or ``"horizontal"`` (default ``"diagonal"``).
        """
        self.remove_watermark()
        header_part = self.header.part
        hdr_element = header_part.element
        size_pt = Length(int(size)).pt
        rotation = "315" if layout == "diagonal" else "0"
        color_hex = str(color)
        pict_xml = (
            '<w:pict xmlns:v="urn:schemas-microsoft-com:vml"'
            ' xmlns:o="urn:schemas-microsoft-com:office:office"'
            ' xmlns:w10="urn:schemas-microsoft-com:office:word"'
            ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<v:shapetype id="_x0000_t136" coordsize="21600,21600"'
            ' o:spt="136" adj="10800"'
            ' path="m@7,l@8,m@5,21600l@6,21600e">'
            '<v:formulas>'
            '<v:f eqn="sum #0 0 10800"/>'
            '<v:f eqn="prod #0 2 1"/>'
            '<v:f eqn="sum 21600 0 @1"/>'
            '<v:f eqn="sum 0 0 @2"/>'
            '<v:f eqn="sum 21600 0 @3"/>'
            '<v:f eqn="if @0 @3 0"/>'
            '<v:f eqn="if @0 21600 @1"/>'
            '<v:f eqn="if @0 0 @2"/>'
            '<v:f eqn="if @0 @4 21600"/>'
            '<v:f eqn="mid @5 @6"/>'
            '<v:f eqn="mid @8 @5"/>'
            '<v:f eqn="mid @7 @8"/>'
            '<v:f eqn="mid @6 @7"/>'
            '<v:f eqn="sum @6 0 @5"/>'
            '</v:formulas>'
            '<v:path textpathok="t" o:connecttype="custom"'
            ' o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"'
            ' o:connectangles="270,180,90,0"/>'
            '<v:textpath on="t" fitshape="t"/>'
            '<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/>'
            '</v:handles>'
            '<o:lock v:ext="edit" text="t" shapetype="t"/>'
            '</v:shapetype>'
            '<v:shape id="PowerPlusWaterMarkObject"'
            ' o:spid="_x0000_s2049" type="#_x0000_t136"'
            f' style="position:absolute;margin-left:0;margin-top:0;'
            f'width:468pt;height:{size_pt * 1.5:.0f}pt;'
            f'rotation:{rotation};z-index:-251657216;'
            f'mso-position-horizontal:center;'
            f'mso-position-horizontal-relative:margin;'
            f'mso-position-vertical:center;'
            f'mso-position-vertical-relative:margin"'
            f' o:allowincell="f"'
            f' fillcolor="#{color_hex}"'
            f' stroked="f">'
            f'<v:fill opacity=".5"/>'
            f'<v:textpath style="font-family:&quot;{font}&quot;;'
            f'font-size:{size_pt:.0f}pt" string="{text}"/>'
            '<w10:wrap anchorx="margin" anchory="margin"/>'
            '</v:shape>'
            '</w:pict>'
        )
        from docx.oxml.parser import parse_xml

        pict_element = parse_xml(pict_xml)
        # Insert a new paragraph with the watermark shape at the beginning of the header
        from docx.oxml.ns import qn

        p = hdr_element.makeelement(qn("w:p"), {})
        r = hdr_element.makeelement(qn("w:r"), {})
        r.append(pict_element)
        pPr = hdr_element.makeelement(qn("w:pPr"), {})
        pStyle = hdr_element.makeelement(qn("w:pStyle"), {qn("w:val"): "Header"})
        pPr.append(pStyle)
        p.append(pPr)
        p.append(r)
        hdr_element.insert(0, p)

    def add_image_watermark(
        self,
        image_path: str | IO[bytes],
        width: Length | int,
        height: Length | int,
    ) -> None:
        """Add an image watermark to this section.

        The watermark is rendered as a VML shape with image data in the default header.
        Any existing watermark is removed before the new one is added.

        Args:
            image_path: Path to the image file, or a file-like object.
            width: Width of the watermark image.
            height: Height of the watermark image.
        """
        self.remove_watermark()
        header_part = self.header.part
        hdr_element = header_part.element
        rId, _image = header_part.get_or_add_image(image_path)

        width_pt = Length(int(width)).pt
        height_pt = Length(int(height)).pt

        pict_xml = (
            '<w:pict xmlns:v="urn:schemas-microsoft-com:vml"'
            ' xmlns:o="urn:schemas-microsoft-com:office:office"'
            ' xmlns:w10="urn:schemas-microsoft-com:office:word"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<v:shapetype id="_x0000_t75" coordsize="21600,21600"'
            ' o:spt="75" o:preferrelative="t"'
            ' path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">'
            '<v:stroke joinstyle="miter"/>'
            '<v:formulas>'
            '<v:f eqn="if lineDrawn pixelLineWidth 0"/>'
            '<v:f eqn="sum @0 1 0"/>'
            '<v:f eqn="sum 0 0 @1"/>'
            '<v:f eqn="prod @2 1 2"/>'
            '<v:f eqn="prod @3 21600 pixelWidth"/>'
            '<v:f eqn="prod @3 21600 pixelHeight"/>'
            '<v:f eqn="sum @0 0 1"/>'
            '<v:f eqn="prod @6 1 2"/>'
            '<v:f eqn="prod @7 21600 pixelWidth"/>'
            '<v:f eqn="sum @8 21600 0"/>'
            '<v:f eqn="prod @7 21600 pixelHeight"/>'
            '<v:f eqn="sum @10 21600 0"/>'
            '</v:formulas>'
            '<v:path o:extrusionok="f" gradientshapeok="t"'
            ' o:connecttype="rect"/>'
            '<o:lock v:ext="edit" aspectratio="t"/>'
            '</v:shapetype>'
            '<v:shape id="PowerPlusWaterMarkObject"'
            ' o:spid="_x0000_s2049" type="#_x0000_t75"'
            f' style="position:absolute;margin-left:0;margin-top:0;'
            f'width:{width_pt:.0f}pt;height:{height_pt:.0f}pt;'
            f'z-index:-251657216;'
            f'mso-position-horizontal:center;'
            f'mso-position-horizontal-relative:margin;'
            f'mso-position-vertical:center;'
            f'mso-position-vertical-relative:margin"'
            f' o:allowincell="f">'
            f'<v:imagedata r:id="{rId}" o:title="watermark"/>'
            '<w10:wrap anchorx="margin" anchory="margin"/>'
            '</v:shape>'
            '</w:pict>'
        )
        from docx.oxml.parser import parse_xml
        from docx.oxml.ns import qn

        pict_element = parse_xml(pict_xml)
        p = hdr_element.makeelement(qn("w:p"), {})
        r = hdr_element.makeelement(qn("w:r"), {})
        r.append(pict_element)
        pPr = hdr_element.makeelement(qn("w:pPr"), {})
        pStyle = hdr_element.makeelement(qn("w:pStyle"), {qn("w:val"): "Header"})
        pPr.append(pStyle)
        p.append(pPr)
        p.append(r)
        hdr_element.insert(0, p)

    def remove_watermark(self) -> None:
        """Remove any watermark from this section's default header.

        Does nothing if this section has no header definition or no watermark.
        """
        from docx.oxml.ns import qn

        if not self.header._has_definition:
            return
        hdr_element = self.header._definition.element
        # Find and remove paragraphs containing a watermark VML shape
        for p in hdr_element.findall(qn("w:p")):
            for r in p.findall(qn("w:r")):
                for pict in r.findall(qn("w:pict")):
                    shapes = pict.xpath(
                        ".//v:shape[contains(@id, 'WaterMark')]",
                        namespaces={"v": "urn:schemas-microsoft-com:vml"},
                    )
                    if shapes:
                        hdr_element.remove(p)
                        break

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        """Generate each Paragraph or Table object in this `section`.

        Items appear in document order.
        """
        for element in self._sectPr.iter_inner_content():
            yield (Paragraph(element, self) if isinstance(element, CT_P) else Table(element, self))

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
    def right_margin(self) -> Length | None:
        """|Length| object representing the right margin for all pages in this section
        in English Metric Units."""
        return self._sectPr.right_margin

    @right_margin.setter
    def right_margin(self, value: Length | None):
        self._sectPr.right_margin = value

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


class Sections(Sequence[Section]):
    """Sequence of |Section| objects corresponding to the sections in the document.

    Supports ``len()``, iteration, and indexed access.
    """

    def __init__(self, document_elm: CT_Document, document_part: DocumentPart):
        super(Sections, self).__init__()
        self._document_elm = document_elm
        self._document_part = document_part

    @overload
    def __getitem__(self, key: int) -> Section: ...

    @overload
    def __getitem__(self, key: slice) -> List[Section]: ...

    def __getitem__(self, key: int | slice) -> Section | List[Section]:
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


class Column:
    """Proxy for a ``<w:col>`` element, representing an individual column definition."""

    def __init__(self, col: CT_Col):
        self._col = col

    @property
    def space(self) -> Length | None:
        """Read/write. Space after this column, in EMU.

        |None| when no space value has been specified.
        """
        return self._col.space

    @space.setter
    def space(self, value: Length | None):
        self._col.space = value

    @property
    def width(self) -> Length | None:
        """Read/write. Width of this column, in EMU.

        |None| when no width has been specified.
        """
        return self._col.w

    @width.setter
    def width(self, value: Length | None):
        self._col.w = value


class SectionColumns(Sequence[Column]):
    """Proxy for a ``<w:cols>`` element, providing access to column layout settings.

    Supports indexed access to individual |Column| objects when ``equal_width`` is False.
    """

    def __init__(self, sectPr: CT_SectPr):
        self._sectPr = sectPr

    @overload
    def __getitem__(self, key: int) -> Column: ...

    @overload
    def __getitem__(self, key: slice) -> List[Column]: ...

    def __getitem__(self, key: Union[int, slice]) -> Union[Column, List[Column]]:
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
        """
        cols = self._sectPr.cols
        if cols is None:
            return None
        return cols.space

    @space.setter
    def space(self, value: Length | None):
        cols = self._sectPr.get_or_add_cols()
        cols.space = value


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
