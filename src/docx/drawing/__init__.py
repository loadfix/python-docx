"""DrawingML-related objects are in this subpackage."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.enum.shape import WD_DRAWING_TYPE, WD_SHAPE
from docx.oxml.drawing import (
    CT_Drawing,
    CT_GroupShape,
    CT_WordprocessingCanvas,
    CT_WordprocessingShape,
)
from docx.oxml.shape import CT_Picture
from docx.shared import ElementProxy, Length, Parented

if TYPE_CHECKING:
    import docx.types as t
    from docx.chart import Chart
    from docx.image.image import Image
    from docx.smart_art import SmartArt
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
    def has_chart(self) -> bool:
        """True when this drawing contains a `c:chart` reference element.

        Note this checks only the *presence* of the chart placeholder; the
        related chart part must exist in the package for :attr:`chart` to
        return a non-None value.
        """
        xpath_expr = (
            "./wp:inline/a:graphic/a:graphicData/c:chart"
            " | ./wp:anchor/a:graphic/a:graphicData/c:chart"
        )
        return bool(self._drawing.xpath(xpath_expr))

    @property
    def chart(self) -> Chart | None:
        """A |Chart| proxy for the chart in this drawing, or |None|.

        Returns |None| when the drawing does not contain a chart reference or
        when the referenced chart part cannot be resolved via the document's
        relationship graph.
        """
        chart_refs = self._drawing.xpath(
            "./wp:inline/a:graphic/a:graphicData/c:chart/@r:id"
            " | ./wp:anchor/a:graphic/a:graphicData/c:chart/@r:id"
        )
        if not chart_refs:
            return None
        rId = chart_refs[0]
        try:
            chart_part = self.part.related_parts[rId]
        except KeyError:
            return None
        # -- local import to avoid a circular-import at module load time --
        from docx.chart import Chart
        from docx.parts.chart import ChartPart

        if not isinstance(chart_part, ChartPart):
            return None
        return Chart(chart_part)

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
    def is_smart_art(self) -> bool:
        """True when this drawing contains a SmartArt (DrawingML diagram) reference.

        A SmartArt reference is a ``dgm:relIds`` element nested inside the
        drawing's ``a:graphicData``. This check is independent of whether
        the companion data part can be resolved — it reports structural
        presence only.
        """
        from docx.oxml.smart_art import dgm_relIds_from_drawing

        return dgm_relIds_from_drawing(self._drawing) is not None

    @property
    def smart_art(self) -> SmartArt | None:
        """A :class:`~docx.smart_art.SmartArt` proxy, or ``None`` when not SmartArt.

        Returns ``None`` when the drawing does not contain a ``dgm:relIds``
        reference. When the reference is present but the companion data part
        cannot be resolved from the document's relationships, the returned
        :class:`SmartArt` still carries the detection but reports an empty
        node list.
        """
        from docx.smart_art import smart_art_for_drawing

        return smart_art_for_drawing(self._drawing, self._parent.part)

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

    Provides read access to the shape's preset type, name, and text, and allows
    replacing the text-frame contents.
    """

    def __init__(self, wsp: CT_WordprocessingShape, parent: t.ProvidesXmlPart):
        super().__init__(wsp, parent)
        self._wsp = wsp

    @property
    def name(self) -> str | None:
        """Value of this shape's `wps:cNvPr/@name`, or |None| when not set."""
        return self._wsp.name

    @property
    def shape_type(self) -> WD_SHAPE | None:
        """The :class:`WD_SHAPE` member for this shape's preset geometry.

        Returns |None| when no preset-geometry element is present, or when the
        preset value doesn't correspond to a known :class:`WD_SHAPE` member.
        """
        prst = self._wsp.prst
        if prst is None:
            return None
        try:
            return WD_SHAPE(prst)
        except ValueError:
            return None

    @property
    def text(self) -> str:
        """Concatenated text from this shape's text frame, or '' when it has none."""
        txbx = self._wsp.txbx
        if txbx is None or txbx.txbxContent is None:
            return ""
        return txbx.txbxContent.text

    @text.setter
    def text(self, value: str) -> None:
        self._wsp.set_text(value)

    def add_paragraph(self, text: str = "") -> Paragraph:
        """Append a paragraph to this shape's text frame and return it.

        A ``wps:txbx/w:txbxContent`` is created lazily if the shape has no
        existing text frame. `text`, when non-empty, is placed in a single run
        inside the new paragraph.

        Raises |ValueError| when the shape's preset geometry cannot carry text
        (e.g. a line), which python-docx does not currently surface distinctly
        from any other geometry.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.ns import nsdecls, qn
        from docx.oxml.parser import parse_xml
        from docx.text.paragraph import Paragraph

        txbx = self._wsp.txbx
        if txbx is None:
            txbx_xml = (
                "<wps:txbx %s><w:txbxContent/></wps:txbx>"
                % nsdecls("wps", "w")
            )
            new_txbx = parse_xml(txbx_xml)
            bodyPr = self._wsp.find(qn("wps:bodyPr"))
            if bodyPr is not None:
                bodyPr.addprevious(new_txbx)
            else:
                self._wsp.append(new_txbx)
            txbx = self._wsp.txbx
        assert txbx is not None
        content = txbx.txbxContent
        if content is None:
            # -- create empty w:txbxContent inside the existing txbx --
            txbx.get_or_add_txbxContent()
            content = txbx.txbxContent
        assert content is not None
        p = content.add_p()  # pyright: ignore[reportAttributeAccessIssue]
        paragraph = Paragraph(p, self)
        if text:
            paragraph.add_run(text)
        return paragraph


class Canvas(ElementProxy):
    """Proxy for a ``<wpc:wpc>`` DrawingML wordprocessing canvas.

    A canvas is a container for one or more child shapes (or pictures) inside
    a single ``w:drawing``. V1 supports appending preset DrawingML shapes via
    :meth:`add_shape`; richer child types (pictures, grouped sub-canvases)
    round-trip but are not yet writable here.

    .. todo:: Flesh out writable picture/grouped-shape children and positional
       placement for canvas shapes (upstream#411).

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, wpc: CT_WordprocessingCanvas, parent: t.ProvidesStoryPart):
        super().__init__(wpc, parent)
        self._wpc = wpc
        self._parent_: t.ProvidesStoryPart = parent

    @property
    def shapes(self) -> list[WordprocessingShape]:
        """List of ``WordprocessingShape`` proxies for this canvas's shapes."""
        return [WordprocessingShape(wsp, self._parent_) for wsp in self._wpc.wsp_lst]

    def add_shape(
        self,
        shape_type: WD_SHAPE,
        width: Length | None = None,
        height: Length | None = None,
        text: str | None = None,
    ) -> WordprocessingShape:
        """Append a ``wps:wsp`` shape to this canvas and return its proxy.

        `shape_type` is a :class:`WD_SHAPE` member (e.g.
        ``WD_SHAPE.ROUNDED_RECTANGLE``). `width` and `height` are |Length|
        values for the shape's extent; they default to 2" x 1".

        .. versionadded:: 1.3.0.dev0
        """
        from docx.oxml.drawing import new_inline_shape_drawing
        from docx.oxml.ns import qn
        from docx.shared import Inches
        from docx.text.paragraph import _shape_name_for

        if not isinstance(shape_type, WD_SHAPE):
            raise TypeError(
                "shape_type must be a WD_SHAPE member, got %r" % (shape_type,)
            )

        cx = int(width) if width is not None else int(Inches(2))
        cy = int(height) if height is not None else int(Inches(1))

        shape_id = self._parent_.part.next_id
        name = "%s %d" % (_shape_name_for(shape_type), shape_id)

        # -- reuse the shared `new_inline_shape_drawing` helper to build a full
        # -- wps:wsp subtree; we then detach the wps:wsp and append it onto the
        # -- canvas. --
        throwaway = new_inline_shape_drawing(
            shape_type.value, cx, cy, shape_id, name, text=text
        )
        wsp = throwaway.find(
            f"{qn('wp:inline')}/{qn('a:graphic')}/{qn('a:graphicData')}/{qn('wps:wsp')}"
        )
        assert wsp is not None
        wsp.getparent().remove(wsp)  # pyright: ignore[reportOptionalMemberAccess]
        self._wpc.append(wsp)

        return WordprocessingShape(
            cast("CT_WordprocessingShape", wsp), self._parent_
        )


class Picture(ElementProxy):
    """Proxy for a `<pic:pic>` picture element inside a group shape."""

    def __init__(self, pic: CT_Picture, parent: t.ProvidesXmlPart):
        super().__init__(pic, parent)
        self._pic = pic
