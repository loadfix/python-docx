"""DrawingML-related objects are in this subpackage."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.shape import WD_DRAWING_TYPE
from docx.oxml.drawing import CT_Drawing, CT_GroupShape, CT_WordprocessingShape
from docx.oxml.shape import CT_Picture
from docx.shared import ElementProxy, Parented

if TYPE_CHECKING:
    import docx.types as t
    from docx.image.image import Image
    from docx.text.paragraph import Paragraph


class Drawing(Parented):
    """Container for a DrawingML object."""

    def __init__(self, drawing: CT_Drawing, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._parent = parent
        self._drawing = self._element = drawing

    @property
    def has_picture(self) -> bool:
        """True when `drawing` contains an embedded picture.

        A drawing can contain a picture, but it can also contain a chart, SmartArt, or a
        drawing canvas. Methods related to a picture, like `.image`, will raise when the drawing
        does not contain a picture. Use this value to determine whether image methods will succeed.

        This value is `False` when a linked picture is present. This should be relatively rare and
        the image would only be retrievable from the filesystem.

        Note this does not distinguish between inline and floating images. The presence of either
        one will cause this value to be `True`.
        """
        xpath_expr = (
            # -- an inline picture --
            "./wp:inline/a:graphic/a:graphicData/pic:pic"
            # -- a floating picture --
            " | ./wp:anchor/a:graphic/a:graphicData/pic:pic"
        )
        # -- xpath() will return a list, empty if there are no matches --
        return bool(self._drawing.xpath(xpath_expr))

    @property
    def image(self) -> Image:
        """An `Image` proxy object for the image in this (picture) drawing.

        Raises `ValueError` when this drawing does contains something other than a picture. Use
        `.has_picture` to qualify drawing objects before using this property.
        """
        picture_rIds = self._drawing.xpath(".//pic:blipFill/a:blip/@r:embed")
        if not picture_rIds:
            raise ValueError("drawing does not contain a picture")
        rId = picture_rIds[0]
        doc_part = self.part
        image_part = doc_part.related_parts[rId]
        return image_part.image

    @property
    def text(self) -> str:
        """Concatenated text from all text frames in this drawing.

        Text from multiple text boxes is separated by newlines. Returns an empty
        string when the drawing contains no text content (e.g. a picture).
        """
        txbxContent_elements = self._drawing.txbxContent_lst
        if not txbxContent_elements:
            return ""
        return "\n".join(txbx.text for txbx in txbxContent_elements)

    @property
    def paragraphs(self) -> list[Paragraph]:
        """All paragraphs inside this drawing's text frames.

        Returns an empty list when the drawing contains no text content.
        """
        from docx.text.paragraph import Paragraph as ParagraphCls

        paragraphs: list[Paragraph] = []
        for txbxContent in self._drawing.txbxContent_lst:
            for p in txbxContent.p_lst:
                paragraphs.append(ParagraphCls(p, self._parent))
        return paragraphs

    @property
    def is_group(self) -> bool:
        """True when this drawing's root DrawingML object is a `wpg:grpSp` group."""
        return bool(self._drawing.grpSp_lst)

    @property
    def group_shape(self) -> GroupShape | None:
        """The top-level `GroupShape` for this drawing, or None if not a group.

        Returns the first top-level group shape when the drawing wraps a `wpg:grpSp`
        (or its legacy `wpg:wgp` alias). Returns None for drawings that contain a
        picture, chart, or other non-group content at the root.
        """
        grpSp_lst = self._drawing.grpSp_lst
        if not grpSp_lst:
            return None
        return GroupShape(grpSp_lst[0], self._parent)

    @property
    def group_shapes(self) -> list[GroupShape]:
        """All top-level group shapes in this drawing.

        Returns an empty list when the drawing doesn't wrap a `wpg:grpSp` at its
        root. A single drawing normally contains at most one top-level group shape,
        but the return type is a list to mirror existing `*_lst` accessors.
        """
        return [GroupShape(grpSp, self._parent) for grpSp in self._drawing.grpSp_lst]

    @property
    def type(self) -> WD_DRAWING_TYPE:
        """The type of content in this drawing.

        Returns a member of :ref:`WD_DRAWING_TYPE` indicating whether this drawing
        contains a shape, text_box, group, chart, diagram, or picture.
        """
        drawing = self._drawing

        # -- check for picture first (most common) --
        if drawing.xpath(
            "./wp:inline/a:graphic/a:graphicData/pic:pic"
            " | ./wp:anchor/a:graphic/a:graphicData/pic:pic"
        ):
            return WD_DRAWING_TYPE.PICTURE

        # -- check for chart --
        if drawing.xpath(
            "./wp:inline/a:graphic/a:graphicData/c:chart"
            " | ./wp:anchor/a:graphic/a:graphicData/c:chart"
        ):
            return WD_DRAWING_TYPE.CHART

        # -- check for diagram --
        if drawing.xpath(
            "./wp:inline/a:graphic/a:graphicData/dgm:*"
            " | ./wp:anchor/a:graphic/a:graphicData/dgm:*"
        ):
            return WD_DRAWING_TYPE.DIAGRAM

        # -- check for group shape --
        if drawing.xpath(
            "./wp:inline/a:graphic/a:graphicData/wpg:*"
            " | ./wp:anchor/a:graphic/a:graphicData/wpg:*"
        ):
            return WD_DRAWING_TYPE.GROUP

        # -- check for text box (shape with txbx content) --
        if drawing.xpath(".//wps:wsp/wps:txbx/w:txbxContent"):
            return WD_DRAWING_TYPE.TEXT_BOX

        return WD_DRAWING_TYPE.SHAPE


class GroupShape(ElementProxy):
    """Proxy for a `<wpg:grpSp>` grouped-shapes element inside a `<w:drawing>`.

    A group shape contains a collection of nested child shapes which may themselves
    be shapes (`WordprocessingShape`), pictures (`Picture`), or nested groups
    (`GroupShape`).
    """

    def __init__(self, grpSp: CT_GroupShape, parent: t.ProvidesXmlPart):
        super().__init__(grpSp, parent)
        self._grpSp = grpSp
        self._parent_: t.ProvidesXmlPart = parent

    @property
    def name(self) -> str | None:
        """Value of the group's `wpg:cNvPr/@name`, or None when not set.

        This is the name Word assigns to the group in the document outline
        (e.g. "Group 1").
        """
        return self._grpSp.name

    @property
    def shapes(self) -> list[GroupShape | WordprocessingShape | Picture]:
        """Flat list of nested child shapes in document order.

        Each entry is a proxy matching the child element type:

        * `wps:wsp` -> `WordprocessingShape`
        * `wpg:grpSp` -> `GroupShape` (recursive)
        * `pic:pic` -> `Picture`

        Other child element types (e.g. `wpg:graphicFrame`) are omitted.
        """
        result: list[GroupShape | WordprocessingShape | Picture] = []
        for child in self._grpSp.shape_children:
            if isinstance(child, CT_WordprocessingShape):
                result.append(WordprocessingShape(child, self._parent_))
            elif isinstance(child, CT_GroupShape):
                result.append(GroupShape(child, self._parent_))
            elif isinstance(child, CT_Picture):
                result.append(Picture(child, self._parent_))
        return result


class WordprocessingShape(ElementProxy):
    """Proxy for a `<wps:wsp>` shape element inside a group shape or drawing.

    Provides read-only access to the shape's text content when it wraps a text box.
    """

    def __init__(self, wsp: CT_WordprocessingShape, parent: t.ProvidesXmlPart):
        super().__init__(wsp, parent)
        self._wsp = wsp

    @property
    def text(self) -> str:
        """Concatenated text from this shape's text frame, or '' when it has none."""
        txbx = self._wsp.txbx
        if txbx is None or txbx.txbxContent is None:
            return ""
        return txbx.txbxContent.text


class Picture(ElementProxy):
    """Proxy for a `<pic:pic>` picture element inside a group shape."""

    def __init__(self, pic: CT_Picture, parent: t.ProvidesXmlPart):
        super().__init__(pic, parent)
        self._pic = pic
