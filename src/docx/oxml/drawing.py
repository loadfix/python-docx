"""Custom element-classes for DrawingML-related elements like `<w:drawing>`.

For legacy reasons, many DrawingML-related elements are in `docx.oxml.shape`. Expect
those to move over here as we have reason to touch them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.shape import CT_Picture
    from docx.oxml.text.paragraph import CT_P
    from docx.shared import Length


class CT_Drawing(BaseOxmlElement):
    """`<w:drawing>` element, containing a DrawingML object like a picture or chart."""

    @property
    def txbxContent_lst(self) -> list[CT_TxbxContent]:
        """All `<w:txbxContent>` descendants (text frames in shapes)."""
        return self.xpath(".//wps:txbx/w:txbxContent")

    @property
    def grpSp_lst(self) -> list[CT_GroupShape]:
        """All top-level `<wpg:grpSp>` or `<wpg:wgp>` group descendants."""
        return cast(
            "list[CT_GroupShape]",
            self.xpath(
                "./wp:inline/a:graphic/a:graphicData/wpg:grpSp"
                " | ./wp:anchor/a:graphic/a:graphicData/wpg:grpSp"
                " | ./wp:inline/a:graphic/a:graphicData/wpg:wgp"
                " | ./wp:anchor/a:graphic/a:graphicData/wpg:wgp"
            ),
        )


class CT_WordprocessingShape(BaseOxmlElement):
    """`<wps:wsp>` element, a WordprocessingML shape."""

    txbx: CT_TextBox | None = ZeroOrOne("wps:txbx")  # pyright: ignore[reportAssignmentType]

    @property
    def name(self) -> str | None:
        """Value of `wps:cNvPr/@name`, or None when not present.

        Mirrors how group shapes expose `wpg:cNvPr/@name`. The shape's own
        non-visual name is stored here, distinct from any `wp:docPr/@name`
        applied on the enclosing `wp:inline`/`wp:anchor`.
        """
        cNvPr = self.find(qn("wps:cNvPr"))
        if cNvPr is None:
            return None
        return cNvPr.get("name")

    @property
    def prst(self) -> str | None:
        """Preset-geometry token from `wps:spPr/a:prstGeom/@prst`, or None.

        This is the raw DrawingML preset name (e.g. ``"rect"`` or
        ``"roundRect"``). Returns |None| when no preset geometry element is
        present (for example on a custom-geometry shape).
        """
        prstGeom = self.find(f"{qn('wps:spPr')}/{qn('a:prstGeom')}")
        if prstGeom is None:
            return None
        return prstGeom.get("prst")

    def set_text(self, text: str) -> None:
        """Replace the shape's text-frame content with a single-run paragraph.

        When the shape has no existing `wps:txbx`, one is created with a
        `w:txbxContent/w:p/w:r/w:t` structure holding `text`. Any pre-existing
        `wps:txbx` content is fully replaced.
        """
        # -- remove any existing wps:txbx --
        existing = self.find(qn("wps:txbx"))
        if existing is not None:
            self.remove(existing)
        txbx_xml = (
            "<wps:txbx %s>"
            "<w:txbxContent><w:p><w:r><w:t/></w:r></w:p></w:txbxContent>"
            "</wps:txbx>" % nsdecls("wps", "w")
        )
        txbx = parse_xml(txbx_xml)
        # -- insert txbx with schema-correct ordering (before wps:bodyPr) --
        bodyPr = self.find(qn("wps:bodyPr"))
        if bodyPr is not None:
            bodyPr.addprevious(txbx)
        else:
            self.append(txbx)
        t_elm = txbx.find(
            f"{qn('w:txbxContent')}/{qn('w:p')}/{qn('w:r')}/{qn('w:t')}"
        )
        assert t_elm is not None  # just constructed above
        t_elm.text = text
        if text and text.strip() != text:
            t_elm.set(qn("xml:space"), "preserve")


class CT_TextBox(BaseOxmlElement):
    """`<wps:txbx>` element, containing a text box with `<w:txbxContent>`."""

    txbxContent: CT_TxbxContent | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:txbxContent"
    )


class CT_TxbxContent(BaseOxmlElement):
    """`<w:txbxContent>` element, containing paragraphs inside a text box."""

    p_lst: list[CT_P]

    p = ZeroOrMore("w:p")

    @property
    def text(self) -> str:
        """Concatenated text of all paragraphs, separated by newlines."""
        return "\n".join(p.text for p in self.p_lst)


class CT_GroupShape(BaseOxmlElement):
    """`<wpg:grpSp>` or `<wpg:wgp>` element, a WordprocessingML group shape.

    Contains a non-visual group-shape-properties element (`wpg:nvGrpSpPr`) followed
    by a collection of nested child shapes which may be `wps:wsp` (shapes/text boxes),
    `pic:pic` (pictures), or `wpg:grpSp` (nested groups).
    """

    nvGrpSpPr: CT_NonVisualGroupShapeProperties | None = (
        ZeroOrOne(  # pyright: ignore[reportAssignmentType]
            "wpg:nvGrpSpPr",
            successors=("wpg:grpSpPr", "wps:wsp", "wpg:grpSp", "pic:pic", "wpg:graphicFrame"),
        )
    )

    @property
    def name(self) -> str | None:
        """Value of `wpg:nvGrpSpPr/wpg:cNvPr/@name`, or None when not present."""
        nvGrpSpPr = self.nvGrpSpPr
        if nvGrpSpPr is None:
            return None
        cNvPr = nvGrpSpPr.find(qn("wpg:cNvPr"))
        if cNvPr is None:
            return None
        return cNvPr.get("name")

    @property
    def wsp_lst(self) -> list[CT_WordprocessingShape]:
        """Direct-child `<wps:wsp>` shape elements."""
        return cast("list[CT_WordprocessingShape]", self.findall(qn("wps:wsp")))

    @property
    def grpSp_lst(self) -> list[CT_GroupShape]:
        """Direct-child nested `<wpg:grpSp>` group elements."""
        return cast("list[CT_GroupShape]", self.findall(qn("wpg:grpSp")))

    @property
    def pic_lst(self) -> list[CT_Picture]:
        """Direct-child `<pic:pic>` picture elements."""
        return cast("list[CT_Picture]", self.findall(qn("pic:pic")))

    @property
    def shape_children(self) -> list[BaseOxmlElement]:
        """Direct-child shape elements in document order.

        Includes `wps:wsp`, `wpg:grpSp`, and `pic:pic` children. Other child types
        (e.g. `wpg:graphicFrame`) are ignored for now.
        """
        wanted = {qn("wps:wsp"), qn("wpg:grpSp"), qn("pic:pic")}
        return [cast(BaseOxmlElement, child) for child in self if child.tag in wanted]


class CT_NonVisualGroupShapeProperties(BaseOxmlElement):
    """`<wpg:nvGrpSpPr>` element, non-visual group-shape properties.

    Contains a `wpg:cNvPr` child carrying the group's id and name.
    """


class CT_WordprocessingCanvas(BaseOxmlElement):
    """`<wpc:wpc>` element, a DrawingML wordprocessing canvas.

    Groups a collection of child shapes (``wps:wsp``), pictures (``pic:pic``),
    or nested groups under a single ``w:drawing``. Each child is appended in
    the order the user adds them; child ordering in the schema is
    ``wpc:bg`` (canvas background) ``wpc:whole`` (frame) then the shape
    children.
    """

    @property
    def wsp_lst(self) -> list[CT_WordprocessingShape]:
        """Direct-child ``wps:wsp`` shape elements."""
        return cast("list[CT_WordprocessingShape]", self.findall(qn("wps:wsp")))


# -- DrawingML graphic-data URI for a `wps:wsp` wordprocessing shape --
_WPS_GRAPHIC_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
# -- DrawingML graphic-data URI for a `wpc:wpc` wordprocessing canvas --
_WPC_GRAPHIC_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"


def new_inline_shape_drawing(
    prst: str,
    cx: Length | int,
    cy: Length | int,
    shape_id: int,
    name: str,
    text: str | None = None,
) -> CT_Drawing:
    """Return a newly-constructed `w:drawing` element wrapping a `wps:wsp`.

    `prst` is the DrawingML preset-geometry token (e.g. ``"rect"``). `cx`/`cy`
    are the extent in EMU. `shape_id` populates both `wp:docPr/@id` and
    `wps:cNvPr/@id`. `name` populates both `@name` attributes. When `text` is
    provided a minimal text frame is attached.

    The element is populated with a generic blue fill (``4472C4``); callers
    that need finer control can post-modify the returned tree.
    """
    xml = (
        '<w:drawing %s>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="0" cy="0"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="0" name=""/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic>'
        '<a:graphicData uri="">'
        '<wps:wsp>'
        '<wps:cNvPr id="0" name=""/>'
        '<wps:cNvSpPr/>'
        '<wps:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>'
        '</wps:spPr>'
        '<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow"'
        ' horzOverflow="overflow" vert="horz" wrap="square"'
        ' lIns="91440" tIns="45720" rIns="91440" bIns="45720"'
        ' anchor="ctr" anchorCtr="0"/>'
        '</wps:wsp>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:inline>'
        '</w:drawing>' % nsdecls("w", "wp", "a", "wps")
    )
    drawing = cast("CT_Drawing", parse_xml(xml))

    # -- populate extent (wp:extent) --
    extent = drawing.find(f"{qn('wp:inline')}/{qn('wp:extent')}")
    assert extent is not None
    extent.set("cx", str(int(cx)))
    extent.set("cy", str(int(cy)))

    # -- populate docPr --
    docPr = drawing.find(f"{qn('wp:inline')}/{qn('wp:docPr')}")
    assert docPr is not None
    docPr.set("id", str(shape_id))
    docPr.set("name", name)

    # -- populate graphicData uri --
    graphicData = drawing.find(
        f"{qn('wp:inline')}/{qn('a:graphic')}/{qn('a:graphicData')}"
    )
    assert graphicData is not None
    graphicData.set("uri", _WPS_GRAPHIC_URI)

    wsp = cast(
        "CT_WordprocessingShape",
        drawing.find(
            f"{qn('wp:inline')}/{qn('a:graphic')}/{qn('a:graphicData')}/{qn('wps:wsp')}"
        ),
    )
    assert wsp is not None

    # -- populate wps:cNvPr --
    cNvPr = wsp.find(qn("wps:cNvPr"))
    assert cNvPr is not None
    cNvPr.set("id", str(shape_id))
    cNvPr.set("name", name)

    # -- populate shape extent (a:xfrm/a:ext) --
    ext = wsp.find(f"{qn('wps:spPr')}/{qn('a:xfrm')}/{qn('a:ext')}")
    assert ext is not None
    ext.set("cx", str(int(cx)))
    ext.set("cy", str(int(cy)))

    # -- set preset geometry --
    prstGeom = wsp.find(f"{qn('wps:spPr')}/{qn('a:prstGeom')}")
    assert prstGeom is not None
    prstGeom.set("prst", prst)

    # -- optional text --
    if text is not None:
        wsp.set_text(text)

    return drawing


def new_inline_canvas_drawing(
    cx: Length | int,
    cy: Length | int,
    shape_id: int,
    name: str,
) -> CT_Drawing:
    """Return a newly-constructed `w:drawing` wrapping an empty `wpc:wpc` canvas.

    The canvas is inserted inside a `wp:inline` container with the provided
    extent (EMU). `shape_id` populates `wp:docPr/@id`. The returned canvas has
    a minimal `wpc:bg` / `wpc:whole` chrome and no child shapes; callers add
    shapes by appending them to the `wpc:wpc` element via
    :meth:`docx.drawing.Canvas.add_shape`.

    Closes upstream#411.
    """
    xml = (
        '<w:drawing %s>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="0" cy="0"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="0" name=""/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic>'
        '<a:graphicData uri="">'
        '<wpc:wpc>'
        '<wpc:bg/>'
        '<wpc:whole/>'
        '</wpc:wpc>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:inline>'
        '</w:drawing>' % nsdecls("w", "wp", "a", "wpc", "wps", "pic")
    )
    drawing = cast("CT_Drawing", parse_xml(xml))

    extent = drawing.find(f"{qn('wp:inline')}/{qn('wp:extent')}")
    assert extent is not None
    extent.set("cx", str(int(cx)))
    extent.set("cy", str(int(cy)))

    docPr = drawing.find(f"{qn('wp:inline')}/{qn('wp:docPr')}")
    assert docPr is not None
    docPr.set("id", str(shape_id))
    docPr.set("name", name)

    graphicData = drawing.find(
        f"{qn('wp:inline')}/{qn('a:graphic')}/{qn('a:graphicData')}"
    )
    assert graphicData is not None
    graphicData.set("uri", _WPC_GRAPHIC_URI)

    return drawing
