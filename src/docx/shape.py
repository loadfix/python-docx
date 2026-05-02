"""Objects related to shapes.

A shape is a visual object that appears on the drawing layer of a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.shape import WD_ANCHOR_H, WD_ANCHOR_V, WD_INLINE_SHAPE, WD_WRAP_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import nsmap, qn
from docx.shared import Emu, Parented, RGBColor

if TYPE_CHECKING:
    from typing import IO

    from docx.image.image import Image
    from docx.opc.part import Part
    from docx.oxml.document import CT_Body
    from docx.oxml.shape import (
        CT_Anchor,
        CT_Inline,
        CT_Picture,
        CT_ShapeProperties,
    )
    from docx.parts.image import ImagePart
    from docx.parts.story import StoryPart
    from docx.shared import Length


def _pic_from_graphicData(graphicData) -> CT_Picture | None:
    """Return the `pic:pic` element under `graphicData`, or None when absent."""
    # -- `graphicData.pic` is available when uri points at picture namespace --
    try:
        pic = graphicData.pic
    except AttributeError:
        pic = None
    if pic is not None:
        return pic
    return graphicData.find(qn("pic:pic"))


def _percent_to_thousandths(value: float | int) -> int:
    """Convert a fractional or percent crop/transparency value to 1000ths of a percent.

    Accepts a float in the range 0.0-1.0 (interpreted as a fraction) or an
    integer 0-100 (interpreted as a whole-number percent). Values outside
    those ranges are clamped so that the XML attribute stays within the
    ``ST_PositiveFixedPercentage`` domain (0..100000).
    """
    if isinstance(value, bool):  # -- guard: bool is a subclass of int --
        raise TypeError("crop/transparency must be a number, not bool")
    f = float(value)
    if not isinstance(value, float) and f > 1.0:
        # -- whole-number percent --
        f = f / 100.0
    if f < 0.0:
        f = 0.0
    if f > 1.0:
        f = 1.0
    return int(round(f * 100000))


def _thousandths_to_fraction(value: int | None) -> float:
    """Inverse of :func:`_percent_to_thousandths`, mapping back to 0.0-1.0."""
    if value is None:
        return 0.0
    return value / 100000.0


class PictureOutline:
    """Proxy for the ``pic:spPr/a:ln`` outline/border of a picture.

    Setting any of the writable properties materialises the ``a:ln`` element
    (and its ``a:solidFill/a:srgbClr`` colour subtree) if absent.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, spPr: CT_ShapeProperties):
        self._spPr = spPr

    @property
    def width(self) -> Length | None:
        """Outline line width as an |Emu| instance, or |None| when unset.

        .. versionadded:: 2026.05.0
        """
        ln = self._spPr.ln
        if ln is None or ln.w is None:
            return None
        return Emu(int(ln.w))

    @width.setter
    def width(self, value: int | None):
        if value is None:
            ln = self._spPr.ln
            if ln is not None and "w" in ln.attrib:
                del ln.attrib["w"]
            return
        ln = self._spPr.get_or_add_ln()
        ln.w = int(value)

    @property
    def color(self) -> RGBColor | None:
        """Outline colour as an |RGBColor|, or |None| when unset.

        .. versionadded:: 2026.05.0
        """
        ln = self._spPr.ln
        if ln is None:
            return None
        solid = ln.solidFill
        if solid is None:
            return None
        srgb = solid.srgbClr
        if srgb is None:
            return None
        return RGBColor.from_string(srgb.val)

    @color.setter
    def color(self, value: RGBColor | str | None):
        if value is None:
            ln = self._spPr.ln
            if ln is None:
                return
            if ln.solidFill is not None:
                ln._remove_solidFill()
            return
        if isinstance(value, str):
            value = RGBColor.from_string(value)
        ln = self._spPr.get_or_add_ln()
        solid = ln.get_or_add_solidFill()
        srgb = solid.get_or_add_srgbClr()
        srgb.val = "%02X%02X%02X" % value

    @property
    def transparency(self) -> float:
        """Outline transparency as a float in the range 0.0-1.0 (0.0 = opaque).

        .. versionadded:: 2026.05.0
        """
        ln = self._spPr.ln
        if ln is None:
            return 0.0
        solid = ln.solidFill
        if solid is None:
            return 0.0
        srgb = solid.srgbClr
        if srgb is None:
            return 0.0
        alpha = srgb.find(qn("a:alpha"))
        if alpha is None:
            return 0.0
        try:
            alpha_val = int(alpha.get("val") or 100000)
        except ValueError:
            return 0.0
        return 1.0 - (alpha_val / 100000.0)

    @transparency.setter
    def transparency(self, value: float | int):
        # -- transparency 0.0 (opaque) => alpha 100000; 1.0 (transparent) => alpha 0.
        fractional = _thousandths_to_fraction(_percent_to_thousandths(value))
        alpha_val = int(round((1.0 - fractional) * 100000))
        ln = self._spPr.get_or_add_ln()
        solid = ln.get_or_add_solidFill()
        srgb = solid.get_or_add_srgbClr()
        # -- reuse existing a:alpha or create one --
        existing = srgb.find(qn("a:alpha"))
        if existing is None:
            from docx.oxml.parser import OxmlElement

            alpha = OxmlElement("a:alpha")
            srgb.append(alpha)
        else:
            alpha = existing
        alpha.set("val", str(alpha_val))


class PictureCrop:
    """Proxy for the ``pic:blipFill/a:srcRect`` crop rectangle.

    Each side is expressed as a fraction (0.0-1.0) of the source image
    that is hidden by the crop. ``left=0.25`` crops a quarter of the image
    off the left edge. Assigning |None| (or ``0``) removes the corresponding
    attribute.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, pic: CT_Picture):
        self._pic = pic

    def _srcRect_or_none(self):
        return self._pic.blipFill.srcRect

    def _get_or_add_srcRect(self):
        return self._pic.blipFill.get_or_add_srcRect()

    def _read(self, attr: str) -> float:
        src = self._srcRect_or_none()
        if src is None:
            return 0.0
        value = getattr(src, attr)
        return _thousandths_to_fraction(value)

    def _write(self, attr: str, value: float | int | None):
        if value is None or value == 0:
            src = self._srcRect_or_none()
            if src is None:
                return
            if attr in src.attrib:
                del src.attrib[attr]
            return
        src = self._get_or_add_srcRect()
        setattr(src, attr, _percent_to_thousandths(value))

    @property
    def left(self) -> float:
        """Left crop as a fraction of the source width (0.0-1.0).

        .. versionadded:: 2026.05.0
        """
        return self._read("l")

    @left.setter
    def left(self, value: float | int | None):
        self._write("l", value)

    @property
    def top(self) -> float:
        """Top crop as a fraction of the source height (0.0-1.0).

        .. versionadded:: 2026.05.0
        """
        return self._read("t")

    @top.setter
    def top(self, value: float | int | None):
        self._write("t", value)

    @property
    def right(self) -> float:
        """Right crop as a fraction of the source width (0.0-1.0).

        .. versionadded:: 2026.05.0
        """
        return self._read("r")

    @right.setter
    def right(self, value: float | int | None):
        self._write("r", value)

    @property
    def bottom(self) -> float:
        """Bottom crop as a fraction of the source height (0.0-1.0).

        .. versionadded:: 2026.05.0
        """
        return self._read("b")

    @bottom.setter
    def bottom(self, value: float | int | None):
        self._write("b", value)

    def set(
        self,
        left: float | int | None = None,
        top: float | int | None = None,
        right: float | int | None = None,
        bottom: float | int | None = None,
    ) -> None:
        """Set all four crop values in a single call.

        Any argument left as |None| leaves the current value unchanged.

        .. versionadded:: 2026.05.0
        """
        if left is not None:
            self.left = left
        if top is not None:
            self.top = top
        if right is not None:
            self.right = right
        if bottom is not None:
            self.bottom = bottom


class ShadowFormat:
    """Proxy for the ``a:effectLst/a:outerShdw`` drop-shadow effect.

    Writable attributes: ``blur_radius`` (EMU), ``distance`` (EMU),
    ``angle`` (degrees, float; stored as ``ST_PositiveFixedAngle`` in
    60000ths of a degree), ``color`` (|RGBColor|).

    Calling :meth:`apply` (or setting any writable property) materialises
    the ``a:outerShdw`` element under ``a:effectLst``; :meth:`clear`
    removes it.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, spPr: CT_ShapeProperties):
        self._spPr = spPr

    # -- helpers --

    def _outerShdw_or_none(self):
        effectLst = self._spPr.effectLst
        if effectLst is None:
            return None
        return effectLst.outerShdw

    def _get_or_add_outerShdw(self):
        effectLst = self._spPr.get_or_add_effectLst()
        return effectLst.get_or_add_outerShdw()

    # -- public API --

    @property
    def exists(self) -> bool:
        """Whether an ``a:outerShdw`` element is currently present.

        .. versionadded:: 2026.05.0
        """
        return self._outerShdw_or_none() is not None

    @property
    def blur_radius(self) -> Length | None:
        """Shadow blur radius in |Emu|, or |None| when unset.

        .. versionadded:: 2026.05.0
        """
        shdw = self._outerShdw_or_none()
        if shdw is None or shdw.blurRad is None:
            return None
        return Emu(int(shdw.blurRad))

    @blur_radius.setter
    def blur_radius(self, value: int | None):
        if value is None:
            shdw = self._outerShdw_or_none()
            if shdw is not None and "blurRad" in shdw.attrib:
                del shdw.attrib["blurRad"]
            return
        self._get_or_add_outerShdw().blurRad = int(value)

    @property
    def distance(self) -> Length | None:
        """Shadow offset distance in |Emu|, or |None| when unset.

        .. versionadded:: 2026.05.0
        """
        shdw = self._outerShdw_or_none()
        if shdw is None or shdw.dist is None:
            return None
        return Emu(int(shdw.dist))

    @distance.setter
    def distance(self, value: int | None):
        if value is None:
            shdw = self._outerShdw_or_none()
            if shdw is not None and "dist" in shdw.attrib:
                del shdw.attrib["dist"]
            return
        self._get_or_add_outerShdw().dist = int(value)

    @property
    def angle(self) -> float | None:
        """Shadow direction in degrees (0-360), or |None| when unset.

        .. versionadded:: 2026.05.0
        """
        shdw = self._outerShdw_or_none()
        if shdw is None or shdw.dir is None:
            return None
        return shdw.dir / 60000.0

    @angle.setter
    def angle(self, value: float | int | None):
        if value is None:
            shdw = self._outerShdw_or_none()
            if shdw is not None and "dir" in shdw.attrib:
                del shdw.attrib["dir"]
            return
        deg = float(value) % 360.0
        self._get_or_add_outerShdw().dir = int(round(deg * 60000))

    @property
    def color(self) -> RGBColor | None:
        """Shadow colour, or |None| when unset.

        .. versionadded:: 2026.05.0
        """
        shdw = self._outerShdw_or_none()
        if shdw is None:
            return None
        srgb = shdw.srgbClr
        if srgb is None:
            return None
        return RGBColor.from_string(srgb.val)

    @color.setter
    def color(self, value: RGBColor | str | None):
        if value is None:
            shdw = self._outerShdw_or_none()
            if shdw is None:
                return
            if shdw.srgbClr is not None:
                shdw._remove_srgbClr()
            return
        if isinstance(value, str):
            value = RGBColor.from_string(value)
        shdw = self._get_or_add_outerShdw()
        srgb = shdw.get_or_add_srgbClr()
        srgb.val = "%02X%02X%02X" % value

    def apply(
        self,
        blur_radius: int | None = None,
        distance: int | None = None,
        angle: float | int | None = None,
        color: RGBColor | str | None = None,
    ) -> ShadowFormat:
        """Configure an outer drop-shadow in a single call.

        Any argument left as |None| leaves the corresponding attribute
        unchanged. Returns ``self`` for chaining.

        .. versionadded:: 2026.05.0
        """
        # -- ensure the a:outerShdw element exists even when no attribute
        #    is supplied, so `outline.shadow.apply()` yields a default
        #    outer-shadow with the schema-specified defaults. --
        self._get_or_add_outerShdw()
        if blur_radius is not None:
            self.blur_radius = blur_radius
        if distance is not None:
            self.distance = distance
        if angle is not None:
            self.angle = angle
        if color is not None:
            self.color = color
        return self

    def clear(self) -> None:
        """Remove the ``a:outerShdw`` element, if present.

        .. versionadded:: 2026.05.0
        """
        effectLst = self._spPr.effectLst
        if effectLst is None:
            return
        if effectLst.outerShdw is not None:
            effectLst._remove_outerShdw()


class EffectsFormat:
    """Container for picture-level visual effects.

    Currently exposes :attr:`shadow` for an outer drop-shadow. Additional
    effects can be added here without changing the public entry points on
    :class:`InlineShape` / :class:`FloatingImage`.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, spPr: CT_ShapeProperties):
        self._spPr = spPr

    @property
    def shadow(self) -> ShadowFormat:
        """The outer drop-shadow proxy for this picture.

        .. versionadded:: 2026.05.0
        """
        return ShadowFormat(self._spPr)


class InlineShapes(Parented):
    """Sequence of |InlineShape| instances, supporting len(), iteration, and indexed access."""

    def __init__(self, body_elm: CT_Body, parent: StoryPart):
        super().__init__(parent)
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
        """List of ``wp:inline`` elements reachable from this body.

        Inline shapes may appear either directly under ``w:drawing`` or nested
        inside an ``mc:AlternateContent`` compatibility block. Word wraps
        newer features (e.g. SVG-with-PNG-fallback, certain DrawingML
        effects) in ``mc:AlternateContent/mc:Choice`` with an
        ``mc:Fallback`` holding a down-level alternative. We prefer each
        ``mc:Choice`` where present and fall back to ``mc:Fallback`` for
        compatibility blocks that have no surviving ``Choice``, so that each
        alternate-content block contributes at most one inline shape.
        """
        body = self._body
        direct = body.xpath(".//w:p/w:r/w:drawing/wp:inline")
        # -- enumerate mc:AlternateContent blocks by positional index so we
        #    can run per-block xpath expressions via `body.xpath()` (the
        #    enhanced method carries the docx namespace map; generic lxml
        #    elements returned from descendant queries may not). --
        alt_block_count = int(body.xpath("count(.//mc:AlternateContent)"))
        alt_inlines: list[CT_Inline] = []
        for idx in range(1, alt_block_count + 1):
            # -- XPath positions are 1-indexed --
            choice_xpath = (
                "(.//mc:AlternateContent)[%d]/mc:Choice//wp:inline" % idx
            )
            fallback_xpath = (
                "(.//mc:AlternateContent)[%d]/mc:Fallback//wp:inline" % idx
            )
            chosen = body.xpath(choice_xpath)
            if not chosen:
                chosen = body.xpath(fallback_xpath)
            alt_inlines.extend(chosen)
        return direct + alt_inlines


def _collect_blip_rIds(drawing_or_element) -> list[str]:
    """Return every ``r:embed`` / ``r:link`` rId referenced by blips under `element`."""
    rIds: list[str] = []
    for attr in ("r:embed", "r:link"):
        rIds.extend(drawing_or_element.xpath(".//a:blip/@%s" % attr))
    # -- SVG blips and other extension elements also carry r:embed --
    rIds.extend(drawing_or_element.xpath(".//asvg:svgBlip/@r:embed"))
    return [rId for rId in rIds if rId]


def _blip_ref_count(part, rId: str) -> int:
    """Return the number of blip references to `rId` in `part`'s XML.

    `Part._rel_ref_count` counts only ``@r:id`` (hyperlinks, chart refs,
    etc.); image blips carry the rId in ``@r:embed`` / ``@r:link``, so we
    count those explicitly. The caller provides the reference-count *after*
    any XML the caller intends to keep is still in place.
    """
    count = 0
    element = getattr(part, "_element", None)
    if element is None:
        return 0
    for attr in ("r:embed", "r:link", "r:id"):
        count += sum(
            1 for val in element.xpath("//@%s" % attr) if val == rId
        )
    return count


def _prune_unused_image_rIds(part, rIds: list[str]) -> None:
    """Drop each rId in `rIds` from `part` when no other XML reference remains.

    When the referenced part is an |ImagePart| that becomes unreachable from
    the package after the drop, it is also removed from the package's
    ``image_parts`` cache so the XML-free image is not accidentally written
    back to the package on save.

    Must be called *after* the referring drawing has been removed from the
    XML so the reference count reflects the final document state.
    """
    from docx.parts.image import ImagePart as _ImagePart

    package = part.package
    for rId in rIds:
        rel = part.rels.get(rId)
        if rel is None:
            continue
        if _blip_ref_count(part, rId) > 0:
            continue

        target_part: _ImagePart | None = None
        if not rel.is_external:
            try:
                candidate = rel.target_part
            except ValueError:
                candidate = None
            if isinstance(candidate, _ImagePart):
                target_part = candidate

        # -- remove the rel directly; Part.drop_rel uses an XmlPart-level
        #    reference count that looks only at `@r:id`, so it won't notice
        #    image blips. --
        del part.rels[rId]
        target_parts_by_rId = getattr(part.rels, "_target_parts_by_rId", None)
        if target_parts_by_rId is not None:
            target_parts_by_rId.pop(rId, None)

        if target_part is None or package is None:
            continue
        if _image_part_still_referenced(package, target_part):
            continue
        image_parts_coll = getattr(package, "image_parts", None)
        if image_parts_coll is None:
            continue
        backing_list = getattr(image_parts_coll, "_image_parts", image_parts_coll)
        if target_part in backing_list:
            backing_list.remove(target_part)


def _image_part_still_referenced(package, image_part) -> bool:
    """Return True if any relationship in `package` still targets `image_part`."""
    iter_rels = getattr(package, "iter_rels", None)
    if iter_rels is None:  # pragma: no cover - defensive
        return False
    for rel in iter_rels():
        if rel.is_external:
            continue
        try:
            if rel.target_part is image_part:
                return True
        except ValueError:  # pragma: no cover - defensive
            continue
    return False


class InlineShape:
    """Proxy for an ``<wp:inline>`` element, representing the container for an inline
    graphical object."""

    def __init__(self, inline: CT_Inline, part: Part | None = None):
        super().__init__()
        self._inline = inline
        self._part = part

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

    @property
    def alt_text(self) -> str | None:
        """Alternative text (accessibility description) for this inline shape.

        Maps to ``wp:inline/wp:docPr/@descr``. Returns |None| when the attribute
        is not present. Assigning |None| removes the attribute.

        .. versionadded:: 2026.05.0
        """
        return self._inline.docPr.descr

    @alt_text.setter
    def alt_text(self, value: str | None):
        self._inline.docPr.descr = value

    @property
    def _pic(self) -> CT_Picture | None:
        """Underlying ``pic:pic`` element, or |None| for non-picture shapes."""
        graphicData = self._inline.graphic.graphicData
        if graphicData.uri != nsmap["pic"]:
            return None
        return _pic_from_graphicData(graphicData)

    @property
    def outline(self) -> PictureOutline:
        """Picture outline/border proxy.

        Writing to ``outline.width``/``outline.color``/``outline.transparency``
        adds a ``pic:spPr/a:ln`` element with a single-colour solid fill
        (upstream#1510).

        Raises |ValueError| when this shape is not a picture.

        .. versionadded:: 2026.05.0
        """
        pic = self._pic
        if pic is None:
            raise ValueError("outline is only available on picture shapes")
        return PictureOutline(pic.spPr)

    @property
    def crop(self) -> PictureCrop:
        """Picture crop (``pic:blipFill/a:srcRect``) proxy.

        Closes upstream#1463 and upstream#1331.

        Raises |ValueError| when this shape is not a picture.

        .. versionadded:: 2026.05.0
        """
        pic = self._pic
        if pic is None:
            raise ValueError("crop is only available on picture shapes")
        return PictureCrop(pic)

    @property
    def effects(self) -> EffectsFormat:
        """Picture visual-effects container, currently exposing :attr:`shadow`.

        Closes upstream#419.

        Raises |ValueError| when this shape is not a picture.

        .. versionadded:: 2026.05.0
        """
        pic = self._pic
        if pic is None:
            raise ValueError("effects are only available on picture shapes")
        return EffectsFormat(pic.spPr)

    @property
    def title(self) -> str | None:
        """Title (accessibility title) for this inline shape.

        Maps to ``wp:inline/wp:docPr/@title``. Returns |None| when the attribute
        is not present. Assigning |None| removes the attribute.

        .. versionadded:: 2026.05.0
        """
        return self._inline.docPr.title

    @title.setter
    def title(self, value: str | None):
        self._inline.docPr.title = value

    @property
    def opacity(self) -> float | None:
        """Image opacity as a float in ``[0.0, 1.0]``, or |None| when unset.

        Reads ``a:blip/a:alphaModFix/@amt`` (a positive percentage expressed
        in 1/1000ths of a percent, per DrawingML ``ST_PositivePercentage``)
        and returns the equivalent fraction — e.g. ``amt=50000`` yields
        ``0.5``. Returns |None| when no ``a:alphaModFix`` element is
        present; since the spec default is 100%, callers can treat |None|
        as fully opaque.

        Assigning a float clamps the value to ``[0.0, 1.0]`` and writes the
        corresponding ``@amt``. Assigning |None| removes the
        ``a:alphaModFix`` element. Closes upstream#1316.

        .. versionadded:: 2026.05.0
        """
        return _get_opacity(self._blip())

    @opacity.setter
    def opacity(self, value: float | None):
        _set_opacity(self._blip(), value)

    @property
    def lock_aspect_ratio(self) -> bool:
        """Whether this picture's aspect ratio is locked in Word.

        Reads ``pic:nvPicPr/pic:cNvPicPr/a:picLocks/@noChangeAspect``. When
        the element or attribute is absent the spec default is ``False``
        (aspect *not* locked).

        Assigning a bool creates the ``a:picLocks`` child element on demand
        and writes the attribute. Assigning ``False`` clears the lock that
        python-docx writes by default on newly-added pictures, matching the
        upstream#1314 request. Closes upstream#1314.

        .. versionadded:: 2026.05.0
        """
        return _get_lock_aspect(self._cNvPicPr())

    @lock_aspect_ratio.setter
    def lock_aspect_ratio(self, value: bool):
        _set_lock_aspect(self._cNvPicPr(), value)

    @property
    def image(self) -> Image:
        """The :class:`~docx.image.image.Image` referenced by this inline shape.

        Resolves the related |ImagePart| via ``a:blip/@r:embed`` on the
        containing part's relationships and returns its underlying |Image|.
        Read-only. Raises :class:`ValueError` when this shape carries no
        part reference (for example an |InlineShape| constructed directly
        from XML) or when no ``r:embed`` rId is present. Closes upstream#249.

        .. versionadded:: 2026.05.0
        """
        if self._part is None:
            raise ValueError(
                "InlineShape.image requires a part reference; "
                "construct the shape via add_picture() or InlineShapes[..]"
            )
        blip = self._blip()
        if blip is None or blip.embed is None:
            raise ValueError("inline shape has no embedded image reference")
        image_part = self._part.related_parts[blip.embed]
        return image_part.image

    # -- helpers ---------------------------------------------------------------

    def _blip(self):
        """Return the `a:blip` child of this inline, or |None| when not a picture."""
        graphicData = self._inline.graphic.graphicData
        if graphicData.uri != nsmap["pic"]:
            return None
        pic = graphicData.pic
        if pic is None:
            return None
        return pic.blipFill.blip

    def _cNvPicPr(self):
        """Return the `pic:cNvPicPr` child, or |None| when not a picture."""
        graphicData = self._inline.graphic.graphicData
        if graphicData.uri != nsmap["pic"]:
            return None
        pic = graphicData.pic
        if pic is None:
            return None
        return pic.nvPicPr.cNvPicPr

    def delete(self, part: StoryPart | None = None) -> None:
        """Remove this inline shape's ``w:drawing`` from its run and prune the rId.

        When the inline shape's image relationship is no longer referenced
        from anywhere in the owning part, the relationship is removed; when
        the underlying |ImagePart| is no longer referenced anywhere in the
        package, it is also removed from the package's image-parts cache.

        `part` is the owning |StoryPart|; when |None|, the ancestor chain of
        the underlying XML element must provide a way to recover it, which is
        not currently available on bare ``wp:inline`` elements — tests can
        pass the part explicitly.

        .. versionadded:: 2026.05.0
        """
        _delete_drawing_carrying(self._inline, part)

    def replace_image(
        self,
        image_path_or_stream: "str | IO[bytes]",
        part: StoryPart | None = None,
    ) -> None:
        """Swap the underlying image binary while preserving this shape's rId/position.

        When this shape's blip rId has *no other* references in the part's
        XML, the |ImagePart| blob is rewritten in-place so the existing
        relationship keeps working. When the rId is shared with another
        drawing, a new |ImagePart| is created and the shape's ``a:blip``
        ``@r:embed`` is rebound to a freshly-allocated rId so the other
        drawing is not affected.

        `part` is the owning |StoryPart|; it is required to reach the package
        and must be supplied when the inline shape is not wired to a
        story part on its own.

        .. versionadded:: 2026.05.0
        """
        _replace_drawing_image(self._inline, image_path_or_stream, part)


class FloatingImage:
    """Proxy for a `<wp:anchor>` element, representing a floating (non-inline) image.

    Provides read-access to the anchor's positioning, wrap type, and offset, and is
    returned by :func:`docx.text.paragraph.Paragraph.add_floating_image`.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, anchor: CT_Anchor):
        super().__init__()
        self._anchor = anchor

    @property
    def width(self) -> Length:
        """Display width of this floating image as an |Emu| instance.

        .. versionadded:: 2026.05.0
        """
        return self._anchor.extent.cx

    @property
    def height(self) -> Length:
        """Display height of this floating image as an |Emu| instance.

        .. versionadded:: 2026.05.0
        """
        return self._anchor.extent.cy

    @property
    def horizontal_anchor(self) -> WD_ANCHOR_H:
        """The horizontal frame-of-reference for the image's position.

        .. versionadded:: 2026.05.0
        """
        positionH = self._anchor.positionH
        value = positionH.relativeFrom if positionH is not None else "column"
        return WD_ANCHOR_H(value)

    @property
    def vertical_anchor(self) -> WD_ANCHOR_V:
        """The vertical frame-of-reference for the image's position.

        .. versionadded:: 2026.05.0
        """
        positionV = self._anchor.positionV
        value = positionV.relativeFrom if positionV is not None else "paragraph"
        return WD_ANCHOR_V(value)

    @property
    def horizontal_offset(self) -> Length:
        """Horizontal offset (EMU) from the horizontal anchor.

        Zero when not specified in the XML.

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
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
        """Tuple ``(horizontal_offset, vertical_offset)`` in EMU.


.. versionadded:: 2026.05.0

"""
        return self.horizontal_offset, self.vertical_offset

    @property
    def position(self) -> dict:
        """A dict describing the position of this floating image.

        Keys: ``h_anchor`` (WD_ANCHOR_H), ``v_anchor`` (WD_ANCHOR_V),
        ``horizontal`` (EMU offset), ``vertical`` (EMU offset).

        .. versionadded:: 2026.05.0
        """
        return {
            "h_anchor": self.horizontal_anchor,
            "v_anchor": self.vertical_anchor,
            "horizontal": self.horizontal_offset,
            "vertical": self.vertical_offset,
        }

    @property
    def wrap_type(self) -> WD_WRAP_TYPE:
        """The text-wrap style of this floating image, a |WD_WRAP_TYPE| member.

    .. versionadded:: 2026.05.0
    """
        return WD_WRAP_TYPE(self._anchor.wrap_type)

    @property
    def alt_text(self) -> str | None:
        """Alternative text (accessibility description) for this floating image.

        Maps to ``wp:anchor/wp:docPr/@descr``. Returns |None| when the attribute
        (or the ``wp:docPr`` element) is not present. Assigning |None| removes
        the attribute.

        .. versionadded:: 2026.05.0
        """
        docPr = self._anchor.docPr
        if docPr is None:
            return None
        return docPr.descr

    @alt_text.setter
    def alt_text(self, value: str | None):
        docPr = self._anchor.get_or_add_docPr()
        docPr.descr = value

    @property
    def title(self) -> str | None:
        """Title (accessibility title) for this floating image.

        Maps to ``wp:anchor/wp:docPr/@title``. Returns |None| when the attribute
        (or the ``wp:docPr`` element) is not present. Assigning |None| removes
        the attribute.

        .. versionadded:: 2026.05.0
        """
        docPr = self._anchor.docPr
        if docPr is None:
            return None
        return docPr.title

    @title.setter
    def title(self, value: str | None):
        docPr = self._anchor.get_or_add_docPr()
        docPr.title = value

    @property
    def type(self):
        """The type of this floating shape, a member of `WD_INLINE_SHAPE`.

        .. versionadded:: 2026.05.0
        """
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

    @property
    def opacity(self) -> float | None:
        """Image opacity as a fraction in ``[0.0, 1.0]``, or |None| when unset.

        See :attr:`InlineShape.opacity` for semantics. Closes upstream#1316.

        .. versionadded:: 2026.05.0
        """
        return _get_opacity(self._blip())

    @opacity.setter
    def opacity(self, value: float | None):
        _set_opacity(self._blip(), value)

    @property
    def lock_aspect_ratio(self) -> bool:
        """Whether this floating picture's aspect ratio is locked in Word.

        See :attr:`InlineShape.lock_aspect_ratio` for semantics. Closes
        upstream#1314.

        .. versionadded:: 2026.05.0
        """
        return _get_lock_aspect(self._cNvPicPr())

    @lock_aspect_ratio.setter
    def lock_aspect_ratio(self, value: bool):
        _set_lock_aspect(self._cNvPicPr(), value)

    @property
    def _pic(self) -> CT_Picture | None:
        """Underlying ``pic:pic`` element, or |None| for non-picture anchors."""
        graphic = self._anchor.graphic
        if graphic is None:
            return None
        graphicData = graphic.graphicData
        if graphicData.uri != nsmap["pic"]:
            return None
        return _pic_from_graphicData(graphicData)

    @property
    def outline(self) -> PictureOutline:
        """Picture outline/border proxy for this floating image.

        Closes upstream#1510.

        .. versionadded:: 2026.05.0
        """
        pic = self._pic
        if pic is None:
            raise ValueError("outline is only available on picture anchors")
        return PictureOutline(pic.spPr)

    @property
    def crop(self) -> PictureCrop:
        """Picture crop (``pic:blipFill/a:srcRect``) proxy.

        .. versionadded:: 2026.05.0
        """
        pic = self._pic
        if pic is None:
            raise ValueError("crop is only available on picture anchors")
        return PictureCrop(pic)

    @property
    def effects(self) -> EffectsFormat:
        """Picture visual-effects container for this floating image.

        .. versionadded:: 2026.05.0
        """
        pic = self._pic
        if pic is None:
            raise ValueError("effects are only available on picture anchors")
        return EffectsFormat(pic.spPr)

    # -- helpers ---------------------------------------------------------------

    def _blip(self):
        """Return the `a:blip` element for this anchor's picture, or |None|."""
        pic = self._pic
        if pic is None:
            return None
        return pic.find(".//" + qn("a:blip"))

    def _cNvPicPr(self):
        """Return the `pic:cNvPicPr` element for this anchor's picture, or |None|."""
        pic = self._pic
        if pic is None:
            return None
        return pic.find(qn("pic:nvPicPr") + "/" + qn("pic:cNvPicPr"))

    def delete(self, part: StoryPart | None = None) -> None:
        """Remove this floating image's ``w:drawing`` and prune its unused rId.

        Follows the same contract as :meth:`InlineShape.delete`: the rId is
        dropped when no other reference remains in the owning part's XML,
        and the backing |ImagePart| is evicted from the package's image-parts
        cache when no other relationship keeps it alive.

        .. versionadded:: 2026.05.0
        """
        _delete_drawing_carrying(self._anchor, part)


# ---------------------------------------------------------------------------
# Shared opacity / lock_aspect_ratio helpers for InlineShape and FloatingImage
# ---------------------------------------------------------------------------


def _get_opacity(blip):
    """Return the opacity encoded on `blip` as a ``[0.0, 1.0]`` float, or |None|."""
    if blip is None:
        return None
    amf = getattr(blip, "alphaModFix", None)
    if amf is None:
        return None
    amt = amf.amt
    if amt is None:
        # -- default per spec is 100% when the element is present without @amt --
        return 1.0
    return max(0.0, min(1.0, amt / 100000.0))


def _set_opacity(blip, value):
    """Write `value` (``[0.0, 1.0]`` float or |None|) to `blip/a:alphaModFix/@amt`."""
    if blip is None:
        raise ValueError("cannot set opacity on a shape that is not a picture")
    if value is None:
        blip._remove_alphaModFix()  # pyright: ignore[reportPrivateUsage]
        return
    clamped = max(0.0, min(1.0, float(value)))
    amf = blip.get_or_add_alphaModFix()
    amf.amt = int(round(clamped * 100000))


def _get_lock_aspect(cNvPicPr) -> bool:
    if cNvPicPr is None:
        return False
    picLocks = getattr(cNvPicPr, "picLocks", None)
    if picLocks is None:
        return False
    return bool(picLocks.noChangeAspect)


def _set_lock_aspect(cNvPicPr, value: bool) -> None:
    if cNvPicPr is None:
        raise ValueError(
            "cannot set lock_aspect_ratio on a shape that is not a picture"
        )
    picLocks = cNvPicPr.get_or_add_picLocks()
    picLocks.noChangeAspect = bool(value)


def _ascending_drawing(element) -> "object | None":
    """Return the ``w:drawing`` ancestor of `element`, or |None|."""
    current = element
    drawing_tag = qn("w:drawing")
    while current is not None:
        if current.tag == drawing_tag:
            return current
        current = current.getparent()
    return None


def _delete_drawing_carrying(element, part: StoryPart | None) -> None:
    """Remove the ``w:drawing`` ancestor of `element` and prune any freed rIds."""
    drawing = _ascending_drawing(element)
    if drawing is None:
        return
    rIds = _collect_blip_rIds(drawing)
    drawing_parent = drawing.getparent()
    if drawing_parent is not None:
        drawing_parent.remove(drawing)
        # -- drop the enclosing run if the removal left it empty --
        if drawing_parent.tag == qn("w:r"):
            has_children = len(drawing_parent) > 0
            has_text = bool(drawing_parent.text)
            if not has_children and not has_text:
                run_parent = drawing_parent.getparent()
                if run_parent is not None:
                    run_parent.remove(drawing_parent)
    if part is not None and rIds:
        _prune_unused_image_rIds(part, rIds)


def _replace_drawing_image(
    element,
    image_path_or_stream: "str | IO[bytes]",
    part: StoryPart | None,
) -> None:
    """Swap the image binary referenced by `element`'s first blip.

    When other XML references still point at the same rId, a new image
    part and rId are allocated and this element's blip is rebound to the
    new rId so the other references are not disturbed.
    """
    if part is None:
        raise ValueError(
            "replace_image() requires a StoryPart; pass `part=paragraph.part` "
            "from the call site"
        )

    blip = element.find(".//" + qn("a:blip"))
    if blip is None:
        raise ValueError("no a:blip element found; nothing to replace")

    embed_attr = qn("r:embed")
    rId = blip.get(embed_attr)
    if rId is None:
        # -- linked picture (r:link); rebind to a freshly embedded image --
        rId = None
    package = part.package
    if package is None:  # pragma: no cover - defensive
        raise ValueError("part is not attached to a package")

    # -- count blip references to the existing rId in the owning part's XML --
    if rId is not None and _blip_ref_count(part, rId) == 1:
        # -- exclusive owner: swap the ImagePart behind this rId --
        rel = part.rels.get(rId)
        if rel is not None and not rel.is_external:
            image_part = rel.target_part
            new_image_part = package.get_or_add_image_part(image_path_or_stream)
            if new_image_part is image_part:
                return
            # -- rebind this rId to point at the new image part --
            part.rels[rId]._target = new_image_part  # type: ignore[attr-defined]
            part.rels._target_parts_by_rId[rId] = new_image_part  # type: ignore[attr-defined]
            return

    # -- shared rId or external/link: create a fresh embedded image part --
    new_image_part = package.get_or_add_image_part(image_path_or_stream)
    new_rId = part.relate_to(new_image_part, RT.IMAGE)
    blip.set(embed_attr, new_rId)
    link_attr = qn("r:link")
    if blip.get(link_attr) is not None:
        del blip.attrib[link_attr]
