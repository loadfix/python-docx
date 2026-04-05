"""Objects related to shapes.

A shape is a visual object that appears on the drawing layer of a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.drawing import WD_RELATIVE_HORZ_POS, WD_RELATIVE_VERT_POS, WD_WRAP_TYPE
from docx.enum.shape import WD_INLINE_SHAPE
from docx.oxml.ns import nsmap
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
        xpath = "//w:p/w:r/w:drawing/wp:inline"
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
    """Proxy for a `<wp:anchor>` element, representing a floating (non-inline) image."""

    def __init__(self, anchor: CT_Anchor):
        super().__init__()
        self._anchor = anchor

    @property
    def height(self) -> Length:
        """The display height of this floating image as an |Emu| instance."""
        return self._anchor.extent.cy

    @height.setter
    def height(self, cy: Length) -> None:
        self._anchor.extent.cy = cy
        self._anchor.graphic.graphicData.pic.spPr.cy = cy

    @property
    def width(self) -> Length:
        """The display width of this floating image as an |Emu| instance."""
        return self._anchor.extent.cx

    @width.setter
    def width(self, cx: Length) -> None:
        self._anchor.extent.cx = cx
        self._anchor.graphic.graphicData.pic.spPr.cx = cx

    @property
    def wrap_type(self) -> WD_WRAP_TYPE:
        """The text wrapping mode of this floating image."""
        return self._anchor.wrap_type

    @property
    def pos_h(self) -> Length:
        """Horizontal offset of this floating image from its reference frame, as EMU."""
        posH = self._anchor.positionH
        if posH is not None and posH.posOffset is not None:
            return Emu(posH.posOffset)
        return Emu(0)

    @pos_h.setter
    def pos_h(self, value: int | Length) -> None:
        posH = self._anchor.positionH
        if posH is not None:
            posH.posOffset = int(value)

    @property
    def pos_v(self) -> Length:
        """Vertical offset of this floating image from its reference frame, as EMU."""
        posV = self._anchor.positionV
        if posV is not None and posV.posOffset is not None:
            return Emu(posV.posOffset)
        return Emu(0)

    @pos_v.setter
    def pos_v(self, value: int | Length) -> None:
        posV = self._anchor.positionV
        if posV is not None:
            posV.posOffset = int(value)

    @property
    def relative_from_h(self) -> WD_RELATIVE_HORZ_POS:
        """The horizontal reference frame for positioning."""
        posH = self._anchor.positionH
        if posH is not None:
            return WD_RELATIVE_HORZ_POS(posH.relativeFrom)
        return WD_RELATIVE_HORZ_POS.COLUMN

    @relative_from_h.setter
    def relative_from_h(self, value: WD_RELATIVE_HORZ_POS) -> None:
        posH = self._anchor.positionH
        if posH is not None:
            posH.relativeFrom = value.value

    @property
    def relative_from_v(self) -> WD_RELATIVE_VERT_POS:
        """The vertical reference frame for positioning."""
        posV = self._anchor.positionV
        if posV is not None:
            return WD_RELATIVE_VERT_POS(posV.relativeFrom)
        return WD_RELATIVE_VERT_POS.PARAGRAPH

    @relative_from_v.setter
    def relative_from_v(self, value: WD_RELATIVE_VERT_POS) -> None:
        posV = self._anchor.positionV
        if posV is not None:
            posV.relativeFrom = value.value

    @property
    def behind_doc(self) -> bool:
        """True when image is positioned behind document text."""
        return self._anchor.behindDoc

    @behind_doc.setter
    def behind_doc(self, value: bool) -> None:
        self._anchor.behindDoc = value
