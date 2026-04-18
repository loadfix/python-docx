"""Objects related to shapes.

A shape is a visual object that appears on the drawing layer of a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.shape import WD_INLINE_SHAPE, WD_RELATIVE_HORZ_POS, WD_RELATIVE_VERT_POS, WD_WRAP_TYPE
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
    """Proxy for a ``<wp:anchor>`` element, representing a floating (anchored) image."""

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
        """The text wrapping mode for this floating image."""
        anchor = self._anchor
        wrap_str = anchor.wrap_type_str
        if wrap_str == "none":
            if anchor.behind_doc:
                return WD_WRAP_TYPE.BEHIND
            return WD_WRAP_TYPE.IN_FRONT
        wrap_map = {
            "square": WD_WRAP_TYPE.SQUARE,
            "tight": WD_WRAP_TYPE.TIGHT,
            "through": WD_WRAP_TYPE.THROUGH,
            "topAndBottom": WD_WRAP_TYPE.TOP_AND_BOTTOM,
        }
        return wrap_map.get(wrap_str, WD_WRAP_TYPE.IN_FRONT)

    @wrap_type.setter
    def wrap_type(self, value: WD_WRAP_TYPE) -> None:
        from docx.oxml.shape import _set_wrap_type

        wrap_map = {
            WD_WRAP_TYPE.SQUARE: ("square", False),
            WD_WRAP_TYPE.TIGHT: ("tight", False),
            WD_WRAP_TYPE.THROUGH: ("through", False),
            WD_WRAP_TYPE.TOP_AND_BOTTOM: ("topAndBottom", False),
            WD_WRAP_TYPE.IN_FRONT: ("none", False),
            WD_WRAP_TYPE.BEHIND: ("none", True),
        }
        wrap_str, behind_doc = wrap_map[value]
        _set_wrap_type(self._anchor, wrap_str)
        self._anchor.set("behindDoc", "1" if behind_doc else "0")

    @property
    def horz_pos_relative(self) -> WD_RELATIVE_HORZ_POS:
        """The horizontal reference frame for positioning."""
        value = self._anchor.horz_relative_from
        return WD_RELATIVE_HORZ_POS(value)

    @horz_pos_relative.setter
    def horz_pos_relative(self, value: WD_RELATIVE_HORZ_POS) -> None:
        self._anchor.horz_relative_from = value.value

    @property
    def vert_pos_relative(self) -> WD_RELATIVE_VERT_POS:
        """The vertical reference frame for positioning."""
        value = self._anchor.vert_relative_from
        return WD_RELATIVE_VERT_POS(value)

    @vert_pos_relative.setter
    def vert_pos_relative(self, value: WD_RELATIVE_VERT_POS) -> None:
        self._anchor.vert_relative_from = value.value

    @property
    def horz_offset(self) -> Length:
        """Horizontal offset from the reference frame, in EMUs."""
        return Emu(self._anchor.horz_offset)

    @horz_offset.setter
    def horz_offset(self, value: int | Length) -> None:
        self._anchor.horz_offset = int(value)

    @property
    def vert_offset(self) -> Length:
        """Vertical offset from the reference frame, in EMUs."""
        return Emu(self._anchor.vert_offset)

    @vert_offset.setter
    def vert_offset(self, value: int | Length) -> None:
        self._anchor.vert_offset = int(value)
