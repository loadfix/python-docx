"""Custom element-classes for DrawingML-related elements like `<w:drawing>`.

For legacy reasons, many DrawingML-related elements are in `docx.oxml.shape`. Expect
those to move over here as we have reason to touch them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.shape import CT_Picture
    from docx.oxml.text.paragraph import CT_P


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
