"""Watermark-related oxml objects and XML builders."""

from __future__ import annotations

from lxml import etree

# -- VML namespace URIs used in watermark shapes --
_VML_NS = "urn:schemas-microsoft-com:vml"
_OFFICE_NS = "urn:schemas-microsoft-com:office:office"
_WORD_NS = "urn:schemas-microsoft-com:office:word"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_WATERMARK_SHAPE_ID = "PowerPlusWaterMarkObject"
_WATERMARK_SHAPE_NAME = "PowerPlusWaterMarkObject"

# -- VML style strings for diagonal and horizontal text watermarks --
_DIAGONAL_STYLE = (
    "position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;"
    "rotation:315;z-index:-251656192;mso-position-horizontal:center;"
    "mso-position-horizontal-relative:margin;mso-position-vertical:center;"
    "mso-position-vertical-relative:margin"
)

_HORIZONTAL_STYLE = (
    "position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;"
    "z-index:-251656192;mso-position-horizontal:center;"
    "mso-position-horizontal-relative:margin;mso-position-vertical:center;"
    "mso-position-vertical-relative:margin"
)

# -- VML style string for image watermarks --
_IMAGE_STYLE_TMPL = (
    "position:absolute;margin-left:0;margin-top:0;width:{width}pt;height:{height}pt;"
    "z-index:-251656192;mso-position-horizontal:center;"
    "mso-position-horizontal-relative:margin;mso-position-vertical:center;"
    "mso-position-vertical-relative:margin"
)


def _build_nsmap():
    """Return namespace map for watermark VML elements."""
    return {
        "v": _VML_NS,
        "o": _OFFICE_NS,
        "w10": _WORD_NS,
        "r": _R_NS,
        "w": _W_NS,
    }


NSMAP = _build_nsmap()


def text_watermark_xml(
    text: str,
    font: str,
    size_pt: float,
    color_hex: str,
    layout: str,
) -> bytes:
    """Return XML bytes for a `w:p` element containing a text watermark VML shape.

    The paragraph is suitable for insertion into a header part (`w:hdr`).
    """
    style = _DIAGONAL_STYLE if layout == "diagonal" else _HORIZONTAL_STYLE

    # -- build the w:p/w:r/w:pict/v:shape structure --
    p = etree.SubElement(etree.Element("_dummy"), f"{{{_W_NS}}}p")
    r = etree.SubElement(p, f"{{{_W_NS}}}r")
    pict = etree.SubElement(r, f"{{{_W_NS}}}pict")

    shape = etree.SubElement(pict, f"{{{_VML_NS}}}shape", nsmap={"v": _VML_NS, "o": _OFFICE_NS})
    shape.set("id", _WATERMARK_SHAPE_ID)
    shape.set(f"{{{_OFFICE_NS}}}spid", "_x0000_s2049")
    shape.set(f"{{{_OFFICE_NS}}}spt", "136")
    shape.set("type", "#_x0000_t136")
    shape.set("style", style)
    shape.set("fillcolor", f"#{color_hex}")
    shape.set("stroked", "f")

    fill = etree.SubElement(shape, f"{{{_VML_NS}}}fill", nsmap={"v": _VML_NS})
    fill.set("opacity", ".5")

    textpath = etree.SubElement(shape, f"{{{_VML_NS}}}textpath", nsmap={"v": _VML_NS})
    textpath.set("style", f"font-family:&quot;{font}&quot;;font-size:{int(size_pt)}pt")
    textpath.set("string", text)

    wrap = etree.SubElement(shape, f"{{{_WORD_NS}}}wrap", nsmap={"w10": _WORD_NS})
    wrap.set("anchorx", "margin")
    wrap.set("anchory", "margin")

    # -- return just the w:p element --
    return etree.tostring(p, xml_declaration=False)


def image_watermark_xml(rId: str, width_pt: float, height_pt: float) -> bytes:
    """Return XML bytes for a `w:p` element containing an image watermark VML shape.

    The paragraph is suitable for insertion into a header part (`w:hdr`).
    `rId` is the relationship id linking to the image part.
    """
    style = _IMAGE_STYLE_TMPL.format(width=width_pt, height=height_pt)

    p = etree.SubElement(etree.Element("_dummy"), f"{{{_W_NS}}}p")
    r = etree.SubElement(p, f"{{{_W_NS}}}r")
    pict = etree.SubElement(r, f"{{{_W_NS}}}pict")

    shape = etree.SubElement(pict, f"{{{_VML_NS}}}shape", nsmap={"v": _VML_NS, "o": _OFFICE_NS, "r": _R_NS})
    shape.set("id", _WATERMARK_SHAPE_ID)
    shape.set(f"{{{_OFFICE_NS}}}spid", "_x0000_s2049")
    shape.set("type", "#_x0000_t75")
    shape.set("style", style)

    imagedata = etree.SubElement(shape, f"{{{_VML_NS}}}imagedata", nsmap={"v": _VML_NS, "r": _R_NS})
    imagedata.set(f"{{{_R_NS}}}id", rId)
    imagedata.set(f"{{{_OFFICE_NS}}}title", "Watermark")
    imagedata.set("gain", "19661f")
    imagedata.set("blacklevel", "22938f")

    wrap = etree.SubElement(shape, f"{{{_WORD_NS}}}wrap", nsmap={"w10": _WORD_NS})
    wrap.set("anchorx", "margin")
    wrap.set("anchory", "margin")

    return etree.tostring(p, xml_declaration=False)


def has_watermark(hdr_element: etree._Element) -> bool:
    """Return True if the header element contains a watermark shape."""
    nsmap = {"v": _VML_NS, "w": _W_NS}
    shapes = hdr_element.xpath(
        f".//v:shape[@id='{_WATERMARK_SHAPE_ID}']",
        namespaces=nsmap,
    )
    return len(shapes) > 0


def remove_watermark_from_header(hdr_element: etree._Element) -> None:
    """Remove watermark paragraph(s) from a header element.

    Finds and removes any `w:p` elements containing a VML shape with the
    watermark shape id.
    """
    nsmap = {"v": _VML_NS, "w": _W_NS}
    # -- find all v:shape elements with the watermark id --
    shapes = hdr_element.xpath(
        f".//v:shape[@id='{_WATERMARK_SHAPE_ID}']",
        namespaces=nsmap,
    )
    for shape in shapes:
        # -- walk up to find the w:p ancestor and remove it --
        p = shape.getparent()
        while p is not None:
            if p.tag == f"{{{_W_NS}}}p":
                p.getparent().remove(p)
                break
            p = p.getparent()
