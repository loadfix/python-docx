"""Custom element-classes for DrawingML-related elements like `<w:drawing>`.

For legacy reasons, many DrawingML-related elements are in `docx.oxml.shape`. Expect
those to move over here as we have reason to touch them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List

from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.text.paragraph import CT_P


class CT_Drawing(BaseOxmlElement):
    """`<w:drawing>` element, containing a DrawingML object like a picture or chart."""

    @property
    def inline_lst(self) -> list[CT_Inline_WP]:
        """All `<wp:inline>` children."""
        return self.xpath("./wp:inline")

    @property
    def anchor_lst(self) -> list[CT_Anchor_WP]:
        """All `<wp:anchor>` children."""
        return self.xpath("./wp:anchor")

    @property
    def txbxContent_lst(self) -> List[CT_TxbxContent]:
        """All `<w:txbxContent>` descendants (text frames in shapes)."""
        return self.xpath(".//wps:txbx/w:txbxContent")


class CT_Inline_WP(BaseOxmlElement):
    """Stub for type-checking `<wp:inline>` in drawing context."""


class CT_Anchor_WP(BaseOxmlElement):
    """Stub for type-checking `<wp:anchor>` in drawing context."""


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

    p_lst: List[CT_P]

    p = ZeroOrMore("w:p")

    @property
    def text(self) -> str:
        """Concatenated text of all paragraphs, separated by newlines."""
        return "\n".join(p.text for p in self.p_lst)
