"""Custom element classes for shape-related elements like `<w:inline>`."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.exceptions import InvalidXmlError
from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import (
    ST_Coordinate,
    ST_DrawingElementId,
    ST_PositiveCoordinate,
    ST_RelationshipId,
    XsdBoolean,
    XsdString,
    XsdToken,
)
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.enum.drawing import WD_RELATIVE_HORZ_POS, WD_RELATIVE_VERT_POS, WD_WRAP_TYPE
    from docx.shared import Length


class CT_Anchor(BaseOxmlElement):
    """`<wp:anchor>` element, container for a "floating" shape.

    Child elements (in schema order)::

        wp:simplePos
        wp:positionH
        wp:positionV
        wp:extent
        wp:effectExtent
        wp:wrapNone / wp:wrapSquare / wp:wrapTight / wp:wrapThrough / wp:wrapTopAndBottom
        wp:docPr
        wp:cNvGraphicFramePr
        a:graphic
    """

    positionH: CT_PosH | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "wp:positionH",
        successors=(
            "wp:positionV",
            "wp:extent",
            "wp:effectExtent",
            "wp:wrapNone",
            "wp:wrapSquare",
            "wp:wrapTight",
            "wp:wrapThrough",
            "wp:wrapTopAndBottom",
            "wp:docPr",
            "wp:cNvGraphicFramePr",
            "a:graphic",
        ),
    )
    positionV: CT_PosV | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "wp:positionV",
        successors=(
            "wp:extent",
            "wp:effectExtent",
            "wp:wrapNone",
            "wp:wrapSquare",
            "wp:wrapTight",
            "wp:wrapThrough",
            "wp:wrapTopAndBottom",
            "wp:docPr",
            "wp:cNvGraphicFramePr",
            "a:graphic",
        ),
    )
    extent: CT_PositiveSize2D = OneAndOnlyOne("wp:extent")  # pyright: ignore[reportAssignmentType]
    wrapNone: CT_WrapNone | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "wp:wrapNone",
        successors=(
            "wp:wrapSquare",
            "wp:wrapTight",
            "wp:wrapThrough",
            "wp:wrapTopAndBottom",
            "wp:docPr",
            "wp:cNvGraphicFramePr",
            "a:graphic",
        ),
    )
    wrapSquare: CT_WrapSquare | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "wp:wrapSquare",
        successors=(
            "wp:wrapTight",
            "wp:wrapThrough",
            "wp:wrapTopAndBottom",
            "wp:docPr",
            "wp:cNvGraphicFramePr",
            "a:graphic",
        ),
    )
    wrapTight: CT_WrapTight | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "wp:wrapTight",
        successors=(
            "wp:wrapThrough",
            "wp:wrapTopAndBottom",
            "wp:docPr",
            "wp:cNvGraphicFramePr",
            "a:graphic",
        ),
    )
    wrapThrough: CT_WrapThrough | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "wp:wrapThrough",
        successors=(
            "wp:wrapTopAndBottom",
            "wp:docPr",
            "wp:cNvGraphicFramePr",
            "a:graphic",
        ),
    )
    wrapTopAndBottom: CT_WrapTopAndBottom | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "wp:wrapTopAndBottom",
        successors=(
            "wp:docPr",
            "wp:cNvGraphicFramePr",
            "a:graphic",
        ),
    )
    docPr: CT_NonVisualDrawingProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "wp:docPr"
    )
    graphic: CT_GraphicalObject = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphic"
    )

    behindDoc: bool = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "behindDoc", XsdBoolean
    )
    distT: int | None = OptionalAttribute("distT", ST_Coordinate)  # pyright: ignore
    distB: int | None = OptionalAttribute("distB", ST_Coordinate)  # pyright: ignore
    distL: int | None = OptionalAttribute("distL", ST_Coordinate)  # pyright: ignore
    distR: int | None = OptionalAttribute("distR", ST_Coordinate)  # pyright: ignore
    layoutInCell: bool = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "layoutInCell", XsdBoolean
    )
    allowOverlap: bool = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "allowOverlap", XsdBoolean
    )

    @classmethod
    def new(
        cls,
        cx: Length,
        cy: Length,
        shape_id: int,
        pic: CT_Picture,
        pos_h: int,
        pos_v: int,
        relative_from_h: WD_RELATIVE_HORZ_POS,
        relative_from_v: WD_RELATIVE_VERT_POS,
        wrap_type: WD_WRAP_TYPE,
        behind_doc: bool = False,
    ) -> CT_Anchor:
        """Return a new `<wp:anchor>` element populated with the given values."""
        from docx.enum.drawing import WD_WRAP_TYPE as WrapType

        anchor = cast(CT_Anchor, parse_xml(cls._anchor_xml()))
        anchor.extent.cx = cx
        anchor.extent.cy = cy
        anchor.docPr.id = shape_id
        anchor.docPr.name = "Picture %d" % shape_id
        anchor.graphic.graphicData.uri = (
            "http://schemas.openxmlformats.org/drawingml/2006/picture"
        )
        anchor.graphic.graphicData._insert_pic(pic)

        # -- set positioning --
        posH = anchor.positionH
        assert posH is not None
        posH.relativeFrom = relative_from_h.value
        posH.posOffset = pos_h

        posV = anchor.positionV
        assert posV is not None
        posV.relativeFrom = relative_from_v.value
        posV.posOffset = pos_v

        # -- set wrap type --
        if wrap_type == WrapType.NONE:
            anchor._add_wrapNone()
        elif wrap_type == WrapType.SQUARE:
            wrap_sq = anchor._add_wrapSquare()
            wrap_sq.wrapText = "bothSides"
        elif wrap_type == WrapType.TIGHT:
            anchor._add_wrapTight()
        elif wrap_type == WrapType.THROUGH:
            anchor._add_wrapThrough()
        elif wrap_type == WrapType.TOP_AND_BOTTOM:
            anchor._add_wrapTopAndBottom()

        # -- set behindDoc attribute --
        anchor.behindDoc = behind_doc

        return anchor

    @classmethod
    def new_pic_anchor(
        cls,
        shape_id: int,
        rId: str,
        filename: str,
        cx: Length,
        cy: Length,
        pos_h: int,
        pos_v: int,
        relative_from_h: WD_RELATIVE_HORZ_POS,
        relative_from_v: WD_RELATIVE_VERT_POS,
        wrap_type: WD_WRAP_TYPE,
        behind_doc: bool = False,
    ) -> CT_Anchor:
        """Create a `wp:anchor` element containing a `pic:pic` element."""
        pic_id = 0
        pic = CT_Picture.new(pic_id, filename, rId, cx, cy)
        return cls.new(
            cx,
            cy,
            shape_id,
            pic,
            pos_h,
            pos_v,
            relative_from_h,
            relative_from_v,
            wrap_type,
            behind_doc,
        )

    @property
    def wrap_type(self) -> WD_WRAP_TYPE:
        """The wrap type for this anchor."""
        from docx.enum.drawing import WD_WRAP_TYPE as WrapType

        if self.wrapNone is not None:
            return WrapType.NONE
        if self.wrapSquare is not None:
            return WrapType.SQUARE
        if self.wrapTight is not None:
            return WrapType.TIGHT
        if self.wrapThrough is not None:
            return WrapType.THROUGH
        if self.wrapTopAndBottom is not None:
            return WrapType.TOP_AND_BOTTOM
        return WrapType.NONE

    @classmethod
    def _anchor_xml(cls) -> str:
        return (
            "<wp:anchor %s"
            '  behindDoc="0" distT="0" distB="0" distL="114300" distR="114300"'
            '  simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">\n'
            '  <wp:simplePos x="0" y="0"/>\n'
            '  <wp:positionH relativeFrom="column">\n'
            "    <wp:posOffset>0</wp:posOffset>\n"
            "  </wp:positionH>\n"
            '  <wp:positionV relativeFrom="paragraph">\n'
            "    <wp:posOffset>0</wp:posOffset>\n"
            "  </wp:positionV>\n"
            '  <wp:extent cx="914400" cy="914400"/>\n'
            '  <wp:effectExtent l="0" t="0" r="0" b="0"/>\n'
            '  <wp:docPr id="666" name="unnamed"/>\n'
            "  <wp:cNvGraphicFramePr>\n"
            '    <a:graphicFrameLocks noChangeAspect="1"/>\n'
            "  </wp:cNvGraphicFramePr>\n"
            "  <a:graphic>\n"
            '    <a:graphicData uri="URI not set"/>\n'
            "  </a:graphic>\n"
            "</wp:anchor>" % nsdecls("wp", "a", "pic", "r")
        )


class CT_EffectExtent(BaseOxmlElement):
    """`<wp:effectExtent>` element, specifies additional extent for effects."""


class _CT_PosBase(BaseOxmlElement):
    """Common base for `CT_PosH` and `CT_PosV` — shared positioning logic."""

    relativeFrom: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "relativeFrom", XsdString
    )

    @property
    def posOffset(self) -> int | None:
        """Value of the `<wp:posOffset>` child element, or None."""
        children = self.xpath("wp:posOffset")
        if not children:
            return None
        text = children[0].text
        if text is None:
            return None
        return int(text)

    @posOffset.setter
    def posOffset(self, value: int) -> None:
        children = self.xpath("wp:posOffset")
        if children:
            children[0].text = str(value)
        else:
            raise InvalidXmlError(
                "<wp:posOffset> child element not present; element may use wp:align instead"
            )


class CT_PosH(_CT_PosBase):
    """`<wp:positionH>` element, specifies horizontal positioning."""


class CT_PosV(_CT_PosBase):
    """`<wp:positionV>` element, specifies vertical positioning."""


class CT_WrapNone(BaseOxmlElement):
    """`<wp:wrapNone>` element — no text wrapping."""


class CT_WrapSquare(BaseOxmlElement):
    """`<wp:wrapSquare>` element — text wraps in a square around the object."""

    wrapText: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "wrapText", XsdString
    )


class CT_WrapTight(BaseOxmlElement):
    """`<wp:wrapTight>` element — text wraps tightly around the object contour."""

    wrapPolygon = ZeroOrOne("wp:wrapPolygon")


class CT_WrapThrough(BaseOxmlElement):
    """`<wp:wrapThrough>` element — text wraps through the object."""

    wrapPolygon = ZeroOrOne("wp:wrapPolygon")


class CT_WrapTopAndBottom(BaseOxmlElement):
    """`<wp:wrapTopAndBottom>` element — text appears above and below only."""


class CT_Blip(BaseOxmlElement):
    """``<a:blip>`` element, specifies image source and adjustments such as alpha and
    tint."""

    embed: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:embed", ST_RelationshipId
    )
    link: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:link", ST_RelationshipId
    )


class CT_BlipFillProperties(BaseOxmlElement):
    """``<pic:blipFill>`` element, specifies picture properties."""

    blip: CT_Blip = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:blip", successors=("a:srcRect", "a:tile", "a:stretch")
    )


class CT_GraphicalObject(BaseOxmlElement):
    """``<a:graphic>`` element, container for a DrawingML object."""

    graphicData: CT_GraphicalObjectData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphicData"
    )


class CT_GraphicalObjectData(BaseOxmlElement):
    """``<a:graphicData>`` element, container for the XML of a DrawingML object."""

    pic: CT_Picture = ZeroOrOne("pic:pic")  # pyright: ignore[reportAssignmentType]
    uri: str = RequiredAttribute("uri", XsdToken)  # pyright: ignore[reportAssignmentType]


class CT_Inline(BaseOxmlElement):
    """`<wp:inline>` element, container for an inline shape."""

    extent: CT_PositiveSize2D = OneAndOnlyOne("wp:extent")  # pyright: ignore[reportAssignmentType]
    docPr: CT_NonVisualDrawingProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "wp:docPr"
    )
    graphic: CT_GraphicalObject = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphic"
    )

    @classmethod
    def new(cls, cx: Length, cy: Length, shape_id: int, pic: CT_Picture) -> CT_Inline:
        """Return a new ``<wp:inline>`` element populated with the values passed as
        parameters."""
        inline = cast(CT_Inline, parse_xml(cls._inline_xml()))
        inline.extent.cx = cx
        inline.extent.cy = cy
        inline.docPr.id = shape_id
        inline.docPr.name = "Picture %d" % shape_id
        inline.graphic.graphicData.uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
        inline.graphic.graphicData._insert_pic(pic)
        return inline

    @classmethod
    def new_pic_inline(
        cls, shape_id: int, rId: str, filename: str, cx: Length, cy: Length
    ) -> CT_Inline:
        """Create `wp:inline` element containing a `pic:pic` element.

        The contents of the `pic:pic` element is taken from the argument values.
        """
        pic_id = 0  # Word doesn't seem to use this, but does not omit it
        pic = CT_Picture.new(pic_id, filename, rId, cx, cy)
        inline = cls.new(cx, cy, shape_id, pic)
        return inline

    @classmethod
    def _inline_xml(cls):
        return (
            "<wp:inline %s>\n"
            '  <wp:extent cx="914400" cy="914400"/>\n'
            '  <wp:docPr id="666" name="unnamed"/>\n'
            "  <wp:cNvGraphicFramePr>\n"
            '    <a:graphicFrameLocks noChangeAspect="1"/>\n'
            "  </wp:cNvGraphicFramePr>\n"
            "  <a:graphic>\n"
            '    <a:graphicData uri="URI not set"/>\n'
            "  </a:graphic>\n"
            "</wp:inline>" % nsdecls("wp", "a", "pic", "r")
        )


class CT_NonVisualDrawingProps(BaseOxmlElement):
    """Used for ``<wp:docPr>`` element, and perhaps others.

    Specifies the id and name of a DrawingML drawing.
    """

    id = RequiredAttribute("id", ST_DrawingElementId)
    name = RequiredAttribute("name", XsdString)


class CT_NonVisualPictureProperties(BaseOxmlElement):
    """``<pic:cNvPicPr>`` element, specifies picture locking and resize behaviors."""


class CT_Picture(BaseOxmlElement):
    """``<pic:pic>`` element, a DrawingML picture."""

    nvPicPr: CT_PictureNonVisual = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "pic:nvPicPr"
    )
    blipFill: CT_BlipFillProperties = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "pic:blipFill"
    )
    spPr: CT_ShapeProperties = OneAndOnlyOne("pic:spPr")  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(cls, pic_id: int, filename: str, rId: str, cx: Length, cy: Length) -> CT_Picture:
        """A new minimum viable `<pic:pic>` (picture) element."""
        pic = parse_xml(cls._pic_xml())
        pic.nvPicPr.cNvPr.id = pic_id
        pic.nvPicPr.cNvPr.name = filename
        pic.blipFill.blip.embed = rId
        pic.spPr.cx = cx
        pic.spPr.cy = cy
        return pic

    @classmethod
    def _pic_xml(cls):
        return (
            "<pic:pic %s>\n"
            "  <pic:nvPicPr>\n"
            '    <pic:cNvPr id="666" name="unnamed"/>\n'
            "    <pic:cNvPicPr/>\n"
            "  </pic:nvPicPr>\n"
            "  <pic:blipFill>\n"
            "    <a:blip/>\n"
            "    <a:stretch>\n"
            "      <a:fillRect/>\n"
            "    </a:stretch>\n"
            "  </pic:blipFill>\n"
            "  <pic:spPr>\n"
            "    <a:xfrm>\n"
            '      <a:off x="0" y="0"/>\n'
            '      <a:ext cx="914400" cy="914400"/>\n'
            "    </a:xfrm>\n"
            '    <a:prstGeom prst="rect"/>\n'
            "  </pic:spPr>\n"
            "</pic:pic>" % nsdecls("pic", "a", "r")
        )


class CT_PictureNonVisual(BaseOxmlElement):
    """``<pic:nvPicPr>`` element, non-visual picture properties."""

    cNvPr = OneAndOnlyOne("pic:cNvPr")


class CT_Point2D(BaseOxmlElement):
    """Used for ``<a:off>`` element, and perhaps others.

    Specifies an x, y coordinate (point).
    """

    x = RequiredAttribute("x", ST_Coordinate)
    y = RequiredAttribute("y", ST_Coordinate)


class CT_PositiveSize2D(BaseOxmlElement):
    """Used for ``<wp:extent>`` element, and perhaps others later.

    Specifies the size of a DrawingML drawing.
    """

    cx: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cx", ST_PositiveCoordinate
    )
    cy: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cy", ST_PositiveCoordinate
    )


class CT_PresetGeometry2D(BaseOxmlElement):
    """``<a:prstGeom>`` element, specifies an preset autoshape geometry, such as
    ``rect``."""


class CT_RelativeRect(BaseOxmlElement):
    """``<a:fillRect>`` element, specifying picture should fill containing rectangle
    shape."""


class CT_ShapeProperties(BaseOxmlElement):
    """``<pic:spPr>`` element, specifies size and shape of picture container."""

    xfrm = ZeroOrOne(
        "a:xfrm",
        successors=(
            "a:custGeom",
            "a:prstGeom",
            "a:ln",
            "a:effectLst",
            "a:effectDag",
            "a:scene3d",
            "a:sp3d",
            "a:extLst",
        ),
    )

    @property
    def cx(self):
        """Shape width as an instance of Emu, or None if not present."""
        xfrm = self.xfrm
        if xfrm is None:
            return None
        return xfrm.cx

    @cx.setter
    def cx(self, value):
        xfrm = self.get_or_add_xfrm()
        xfrm.cx = value

    @property
    def cy(self):
        """Shape height as an instance of Emu, or None if not present."""
        xfrm = self.xfrm
        if xfrm is None:
            return None
        return xfrm.cy

    @cy.setter
    def cy(self, value):
        xfrm = self.get_or_add_xfrm()
        xfrm.cy = value


class CT_StretchInfoProperties(BaseOxmlElement):
    """``<a:stretch>`` element, specifies how picture should fill its containing
    shape."""


class CT_Transform2D(BaseOxmlElement):
    """``<a:xfrm>`` element, specifies size and shape of picture container."""

    off = ZeroOrOne("a:off", successors=("a:ext",))
    ext = ZeroOrOne("a:ext", successors=())

    @property
    def cx(self):
        ext = self.ext
        if ext is None:
            return None
        return ext.cx

    @cx.setter
    def cx(self, value):
        ext = self.get_or_add_ext()
        ext.cx = value

    @property
    def cy(self):
        ext = self.ext
        if ext is None:
            return None
        return ext.cy

    @cy.setter
    def cy(self, value):
        ext = self.get_or_add_ext()
        ext.cy = value
