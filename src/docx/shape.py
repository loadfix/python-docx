"""Objects related to shapes.

A shape is a visual object that appears on the drawing layer of a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.shape import WD_ANCHOR_H, WD_ANCHOR_V, WD_INLINE_SHAPE, WD_WRAP_TYPE
from docx.oxml.ns import nsmap, qn
from docx.shared import Emu, Parented

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body
    from docx.oxml.shape import CT_Anchor, CT_Inline
    from docx.parts.story import StoryPart
    from docx.shared import Length


class InlineShapes(Parented):
    """Sequence of |InlineShape| instances, supporting len(), iteration, and indexed access."""

    def __init__(self, body_elm: CT_Body, parent: StoryPart):
        super(InlineShapes, self).__init__(parent)
        self._body = body_elm

    def __getitem__(self, idx: int):
        """Provide indexed access, e.g. 'inline_shapes[idx]'."""
        try:
            inline = self._inline_lst[idx]
        except IndexError:
            msg = "inline shape index [%d] out of range" % idx
            raise IndexError(msg)

        return InlineShape(inline)

    def __iter__(self):
        return (InlineShape(inline) for inline in self._inline_lst)

    def __len__(self):
        return len(self._inline_lst)

    @property
    def _inline_lst(self):
        body = self._body
        xpath = ".//w:p/w:r/w:drawing/wp:inline"
        return body.xpath(xpath)


class InlineShape:
    """Proxy for an ``<wp:inline>`` element, representing the container for an inline
    graphical object."""

    def __init__(self, inline: CT_Inline):
        super(InlineShape, self).__init__()
        self._inline = inline

    @property
    def height(self) -> Length:
        """Read/write.

        The display height of this inline shape as an |Emu| instance.
        """
        return self._inline.extent.cy

    @height.setter
    def height(self, cy: Length):
        self._inline.extent.cy = cy
        self._inline.graphic.graphicData.pic.spPr.cy = cy

    @property
    def type(self):
        """The type of this inline shape as a member of
        ``docx.enum.shape.WD_INLINE_SHAPE``, e.g. ``LINKED_PICTURE``.

        Read-only.
        """
        graphicData = self._inline.graphic.graphicData
        uri = graphicData.uri
        if uri == nsmap["pic"]:
            blip = graphicData.pic.blipFill.blip
            if blip.link is not None:
                return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == nsmap["c"]:
            return WD_INLINE_SHAPE.CHART
        if uri == nsmap["dgm"]:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED

    @property
    def width(self):
        """Read/write.

        The display width of this inline shape as an |Emu| instance.
        """
        return self._inline.extent.cx

    @width.setter
    def width(self, cx: Length):
        self._inline.extent.cx = cx
        self._inline.graphic.graphicData.pic.spPr.cx = cx


class FloatingImage:
    """Proxy for a `<wp:anchor>` element, representing a floating (non-inline) image.

    Provides read-access to the anchor's positioning, wrap type, and offset, and is
    returned by :func:`docx.text.paragraph.Paragraph.add_floating_image`.
    """

    def __init__(self, anchor: CT_Anchor):
        super().__init__()
        self._anchor = anchor

    @property
    def width(self) -> Length:
        """Display width of this floating image as an |Emu| instance."""
        return self._anchor.extent.cx

    @property
    def height(self) -> Length:
        """Display height of this floating image as an |Emu| instance."""
        return self._anchor.extent.cy

    @property
    def horizontal_anchor(self) -> WD_ANCHOR_H:
        """The horizontal frame-of-reference for the image's position."""
        positionH = self._anchor.positionH
        value = positionH.relativeFrom if positionH is not None else "column"
        return WD_ANCHOR_H(value)

    @property
    def vertical_anchor(self) -> WD_ANCHOR_V:
        """The vertical frame-of-reference for the image's position."""
        positionV = self._anchor.positionV
        value = positionV.relativeFrom if positionV is not None else "paragraph"
        return WD_ANCHOR_V(value)

    @property
    def horizontal_offset(self) -> Length:
        """Horizontal offset (EMU) from the horizontal anchor.

        Zero when not specified in the XML.
        """
        positionH = self._anchor.positionH
        if positionH is None or positionH.posOffset is None:
            return Emu(0)
        try:
            return Emu(int(positionH.posOffset.text or "0"))
        except (TypeError, ValueError):
            return Emu(0)

    @property
    def vertical_offset(self) -> Length:
        """Vertical offset (EMU) from the vertical anchor.

        Zero when not specified in the XML.
        """
        positionV = self._anchor.positionV
        if positionV is None or positionV.posOffset is None:
            return Emu(0)
        try:
            return Emu(int(positionV.posOffset.text or "0"))
        except (TypeError, ValueError):
            return Emu(0)

    @property
    def offset(self) -> tuple[Length, Length]:
        """Tuple ``(horizontal_offset, vertical_offset)`` in EMU."""
        return self.horizontal_offset, self.vertical_offset

    @property
    def position(self) -> dict:
        """A dict describing the position of this floating image.

        Keys: ``h_anchor`` (WD_ANCHOR_H), ``v_anchor`` (WD_ANCHOR_V),
        ``horizontal`` (EMU offset), ``vertical`` (EMU offset).
        """
        return {
            "h_anchor": self.horizontal_anchor,
            "v_anchor": self.vertical_anchor,
            "horizontal": self.horizontal_offset,
            "vertical": self.vertical_offset,
        }

    @property
    def wrap_type(self) -> WD_WRAP_TYPE:
        """The text-wrap style of this floating image, a |WD_WRAP_TYPE| member."""
        return WD_WRAP_TYPE(self._anchor.wrap_type)

    @property
    def type(self):
        """The type of this floating shape, a member of `WD_INLINE_SHAPE`."""
        graphic = self._anchor.graphic
        if graphic is None:
            return WD_INLINE_SHAPE.NOT_IMPLEMENTED
        graphicData = graphic.graphicData
        uri = graphicData.uri
        if uri == nsmap["pic"]:
            pic = graphicData.find(qn("pic:pic"))
            if pic is not None:
                blip = pic.find(".//" + qn("a:blip"))
                if blip is not None and blip.get(qn("r:link")) is not None:
                    return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == nsmap["c"]:
            return WD_INLINE_SHAPE.CHART
        if uri == nsmap["dgm"]:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED
