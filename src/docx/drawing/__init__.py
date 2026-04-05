"""DrawingML-related objects are in this subpackage."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.shape import WD_DRAWING_TYPE
from docx.oxml.drawing import CT_Drawing
from docx.shared import Parented

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
