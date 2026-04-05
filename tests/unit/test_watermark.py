# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.oxml.watermark module."""

from __future__ import annotations

from lxml import etree

from docx.oxml.watermark import (
    has_watermark,
    image_watermark_xml,
    remove_watermark_from_header,
    text_watermark_xml,
)


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_VML_NS = "urn:schemas-microsoft-com:vml"
_OFFICE_NS = "urn:schemas-microsoft-com:office:office"
_WORD_NS = "urn:schemas-microsoft-com:office:word"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _make_hdr(*children_xml: str) -> etree._Element:
    """Return a w:hdr element, optionally with child XML inserted."""
    children = "".join(children_xml)
    xml = (
        f'<w:hdr xmlns:w="{_W_NS}" xmlns:v="{_VML_NS}"'
        f' xmlns:o="{_OFFICE_NS}" xmlns:w10="{_WORD_NS}"'
        f' xmlns:r="{_R_NS}">'
        f"<w:p><w:pPr><w:pStyle w:val=\"Header\"/></w:pPr></w:p>"
        f"{children}"
        f"</w:hdr>"
    )
    return etree.fromstring(xml.encode("utf-8"))


class DescribeTextWatermarkXml:
    def it_produces_a_w_p_element(self):
        xml_bytes = text_watermark_xml("DRAFT", "Calibri", 72.0, "C0C0C0", "diagonal")
        p = etree.fromstring(xml_bytes)
        assert p.tag == f"{{{_W_NS}}}p"

    def it_contains_a_vml_shape(self):
        xml_bytes = text_watermark_xml("DRAFT", "Calibri", 72.0, "C0C0C0", "diagonal")
        p = etree.fromstring(xml_bytes)
        shapes = p.xpath(".//v:shape", namespaces={"v": _VML_NS})
        assert len(shapes) == 1

    def it_sets_the_watermark_shape_id(self):
        xml_bytes = text_watermark_xml("DRAFT", "Calibri", 72.0, "C0C0C0", "diagonal")
        p = etree.fromstring(xml_bytes)
        shapes = p.xpath(".//v:shape", namespaces={"v": _VML_NS})
        assert shapes[0].get("id") == "PowerPlusWaterMarkObject"

    def it_contains_textpath_with_correct_text(self):
        xml_bytes = text_watermark_xml("CONFIDENTIAL", "Arial", 48.0, "FF0000", "horizontal")
        p = etree.fromstring(xml_bytes)
        textpaths = p.xpath(".//v:textpath", namespaces={"v": _VML_NS})
        assert len(textpaths) == 1
        assert textpaths[0].get("string") == "CONFIDENTIAL"

    def it_applies_diagonal_layout(self):
        xml_bytes = text_watermark_xml("DRAFT", "Calibri", 72.0, "C0C0C0", "diagonal")
        p = etree.fromstring(xml_bytes)
        shapes = p.xpath(".//v:shape", namespaces={"v": _VML_NS})
        style = shapes[0].get("style")
        assert "rotation:315" in style

    def it_applies_horizontal_layout(self):
        xml_bytes = text_watermark_xml("DRAFT", "Calibri", 72.0, "C0C0C0", "horizontal")
        p = etree.fromstring(xml_bytes)
        shapes = p.xpath(".//v:shape", namespaces={"v": _VML_NS})
        style = shapes[0].get("style")
        assert "rotation" not in style

    def it_sets_fill_color(self):
        xml_bytes = text_watermark_xml("DRAFT", "Calibri", 72.0, "FF0000", "diagonal")
        p = etree.fromstring(xml_bytes)
        shapes = p.xpath(".//v:shape", namespaces={"v": _VML_NS})
        assert shapes[0].get("fillcolor") == "#FF0000"

    def it_includes_word_wrap_element(self):
        xml_bytes = text_watermark_xml("DRAFT", "Calibri", 72.0, "C0C0C0", "diagonal")
        p = etree.fromstring(xml_bytes)
        wraps = p.xpath(".//w10:wrap", namespaces={"w10": _WORD_NS})
        assert len(wraps) == 1
        assert wraps[0].get("anchorx") == "margin"


class DescribeImageWatermarkXml:
    def it_produces_a_w_p_element(self):
        xml_bytes = image_watermark_xml("rId1", 400.0, 300.0)
        p = etree.fromstring(xml_bytes)
        assert p.tag == f"{{{_W_NS}}}p"

    def it_contains_a_vml_shape(self):
        xml_bytes = image_watermark_xml("rId1", 400.0, 300.0)
        p = etree.fromstring(xml_bytes)
        shapes = p.xpath(".//v:shape", namespaces={"v": _VML_NS})
        assert len(shapes) == 1
        assert shapes[0].get("id") == "PowerPlusWaterMarkObject"

    def it_contains_imagedata_with_rId(self):
        xml_bytes = image_watermark_xml("rId7", 400.0, 300.0)
        p = etree.fromstring(xml_bytes)
        imagedata = p.xpath(".//v:imagedata", namespaces={"v": _VML_NS})
        assert len(imagedata) == 1
        assert imagedata[0].get(f"{{{_R_NS}}}id") == "rId7"

    def it_sets_dimensions_in_style(self):
        xml_bytes = image_watermark_xml("rId1", 400.0, 300.0)
        p = etree.fromstring(xml_bytes)
        shapes = p.xpath(".//v:shape", namespaces={"v": _VML_NS})
        style = shapes[0].get("style")
        assert "width:400.0pt" in style
        assert "height:300.0pt" in style


class DescribeHasWatermark:
    def it_returns_False_for_empty_header(self):
        hdr = _make_hdr()
        assert has_watermark(hdr) is False

    def it_returns_True_when_watermark_present(self):
        wm_xml = (
            '<w:p xmlns:w="{w}" xmlns:v="{v}">'
            "<w:r><w:pict>"
            '<v:shape id="PowerPlusWaterMarkObject">'
            '<v:textpath string="DRAFT"/>'
            "</v:shape>"
            "</w:pict></w:r></w:p>"
        ).format(w=_W_NS, v=_VML_NS)
        hdr = _make_hdr(wm_xml)
        assert has_watermark(hdr) is True


class DescribeRemoveWatermarkFromHeader:
    def it_removes_watermark_paragraph(self):
        wm_xml = (
            '<w:p xmlns:w="{w}" xmlns:v="{v}">'
            "<w:r><w:pict>"
            '<v:shape id="PowerPlusWaterMarkObject">'
            '<v:textpath string="DRAFT"/>'
            "</v:shape>"
            "</w:pict></w:r></w:p>"
        ).format(w=_W_NS, v=_VML_NS)
        hdr = _make_hdr(wm_xml)
        assert has_watermark(hdr) is True

        remove_watermark_from_header(hdr)
        assert has_watermark(hdr) is False

    def it_preserves_non_watermark_paragraphs(self):
        wm_xml = (
            '<w:p xmlns:w="{w}" xmlns:v="{v}">'
            "<w:r><w:pict>"
            '<v:shape id="PowerPlusWaterMarkObject">'
            '<v:textpath string="DRAFT"/>'
            "</v:shape>"
            "</w:pict></w:r></w:p>"
        ).format(w=_W_NS, v=_VML_NS)
        hdr = _make_hdr(wm_xml)
        # -- header has the style paragraph + watermark paragraph = 2 --
        p_elements = hdr.xpath(".//w:p", namespaces={"w": _W_NS})
        assert len(p_elements) == 2

        remove_watermark_from_header(hdr)

        # -- only the style paragraph should remain --
        p_elements = hdr.xpath(".//w:p", namespaces={"w": _W_NS})
        assert len(p_elements) == 1

    def it_does_nothing_when_no_watermark(self):
        hdr = _make_hdr()
        p_count_before = len(hdr.xpath(".//w:p", namespaces={"w": _W_NS}))
        remove_watermark_from_header(hdr)
        p_count_after = len(hdr.xpath(".//w:p", namespaces={"w": _W_NS}))
        assert p_count_before == p_count_after
