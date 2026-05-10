# pyright: reportPrivateUsage=false

"""Round 14 SmartArt adoption — InlineShape / Document / SmartArt accessors.

These tests cover the typed accessors added in 2026.05.10 that route
SmartArt discovery through :class:`~docx.shape.InlineShape`:

* :attr:`InlineShape.is_smart_art` — structural SmartArt detection.
* :attr:`InlineShape.smart_art` — SmartArt proxy fully wired with
  colour / style transforms.
* :attr:`~docx.document.Document.smart_arts` — plural alias.
* :meth:`~docx.document.Document.iter_smart_arts` — generator walk.
* :attr:`~docx.smart_art.SmartArt.color_transform` /
  :attr:`~docx.smart_art.SmartArt.style_transform` — typed proxies
  routed through ``python-ooxml-smartart`` 0.3.
* :attr:`~docx.smart_art.SmartArt.graphic_frame_xml` — raw bytes of
  the wrapping ``w:drawing`` for pptx migration workflows.
"""

from __future__ import annotations

from typing import cast

import pytest

from docx.document import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.document import CT_Document
from docx.oxml.parser import parse_xml
from docx.oxml.shape import CT_Inline
from docx.oxml.smart_art import CT_DataModel
from docx.parts.document import DocumentPart
from docx.parts.smart_art import (
    DiagramColorsPart,
    DiagramDataPart,
    DiagramStylePart,
)
from docx.shape import InlineShape
from docx.smart_art import SmartArt

from .unitutil.cxml import element
from .unitutil.mock import FixtureRequest, instance_mock


# -- minimal SmartArt data-model blob (three flat content nodes) --

DATA_XML = (
    b'<?xml version="1.0"?>\n'
    b'<dgm:dataModel'
    b' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
    b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
    b"  <dgm:ptLst>\n"
    b'    <dgm:pt modelId="n1" type="node">\n'
    b'      <dgm:t><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></dgm:t>\n'
    b"    </dgm:pt>\n"
    b'    <dgm:pt modelId="n2" type="node">\n'
    b'      <dgm:t><a:p><a:r><a:t>Beta</a:t></a:r></a:p></dgm:t>\n'
    b"    </dgm:pt>\n"
    b"  </dgm:ptLst>\n"
    b"</dgm:dataModel>\n"
)

# -- minimal SmartArt colours part (dgm:colorsDef) --
COLORS_XML = (
    b'<?xml version="1.0"?>\n'
    b'<dgm:colorsDef'
    b' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
    b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    b' uniqueId="urn:test/colors1" minVer="12.0.0.0000">\n'
    b"  <dgm:title lang=\"en-US\" val=\"Test Colours\"/>\n"
    b"  <dgm:desc lang=\"en-US\" val=\"\"/>\n"
    b"  <dgm:catLst>\n"
    b'    <dgm:cat type="accent1" pri="10100"/>\n'
    b"  </dgm:catLst>\n"
    b'  <dgm:styleLbl name="node0">\n'
    b"    <dgm:fillClrLst meth=\"repeat\"><a:schemeClr val=\"accent1\"/></dgm:fillClrLst>\n"
    b"    <dgm:linClrLst meth=\"repeat\"><a:schemeClr val=\"accent1\"/></dgm:linClrLst>\n"
    b"    <dgm:effectClrLst meth=\"repeat\"><a:schemeClr val=\"accent1\"/></dgm:effectClrLst>\n"
    b"    <dgm:txFillClrLst meth=\"repeat\"><a:schemeClr val=\"lt1\"/></dgm:txFillClrLst>\n"
    b"    <dgm:txLinClrLst meth=\"repeat\"><a:schemeClr val=\"lt1\"/></dgm:txLinClrLst>\n"
    b"    <dgm:txEffectClrLst meth=\"repeat\"><a:schemeClr val=\"lt1\"/></dgm:txEffectClrLst>\n"
    b"  </dgm:styleLbl>\n"
    b"</dgm:colorsDef>\n"
)

# -- minimal SmartArt quickStyle part (dgm:styleDef) --
STYLE_XML = (
    b'<?xml version="1.0"?>\n'
    b'<dgm:styleDef'
    b' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
    b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    b' uniqueId="urn:test/style1" minVer="12.0.0.0000">\n'
    b"  <dgm:title lang=\"en-US\" val=\"Test Style\"/>\n"
    b"  <dgm:desc lang=\"en-US\" val=\"\"/>\n"
    b"  <dgm:catLst>\n"
    b'    <dgm:cat type="simple" pri="10100"/>\n'
    b"  </dgm:catLst>\n"
    b'  <dgm:scene3d>\n'
    b'    <a:camera prst="orthographicFront"/>\n'
    b'    <a:lightRig rig="flat" dir="t"/>\n'
    b"  </dgm:scene3d>\n"
    b'  <dgm:styleLbl name="node0">\n'
    b"    <dgm:scene3d>\n"
    b'      <a:camera prst="orthographicFront"/>\n'
    b'      <a:lightRig rig="flat" dir="t"/>\n'
    b"    </dgm:scene3d>\n"
    b"    <dgm:sp3d/>\n"
    b"    <dgm:txPr/>\n"
    b"    <dgm:style><a:lnRef idx=\"0\"/><a:fillRef idx=\"0\"/><a:effectRef idx=\"0\"/><a:fontRef idx=\"minor\"/></dgm:style>\n"
    b"  </dgm:styleLbl>\n"
    b"</dgm:styleDef>\n"
)


# -- helpers -----------------------------------------------------------------


def _data_part(idx: int = 1) -> DiagramDataPart:
    element_ = cast(CT_DataModel, parse_xml(DATA_XML))
    return DiagramDataPart(
        PackURI("/word/diagrams/data%d.xml" % idx),
        CT.DML_DIAGRAM_DATA,
        element_,
        cast(object, None),  # pyright: ignore[reportArgumentType]
    )


def _colors_part(idx: int = 1) -> DiagramColorsPart:
    return DiagramColorsPart(
        PackURI("/word/diagrams/colors%d.xml" % idx),
        CT.DML_DIAGRAM_COLORS,
        parse_xml(COLORS_XML),
        cast(object, None),  # pyright: ignore[reportArgumentType]
    )


def _style_part(idx: int = 1) -> DiagramStylePart:
    return DiagramStylePart(
        PackURI("/word/diagrams/quickStyle%d.xml" % idx),
        CT.DML_DIAGRAM_STYLE,
        parse_xml(STYLE_XML),
        cast(object, None),  # pyright: ignore[reportArgumentType]
    )


# -- a realistic <w:drawing> with the four SmartArt rIds --
def _smart_art_drawing_xml(
    dm_rId: str = "rId4",
    lo_rId: str = "rId5",
    qs_rId: str = "rId6",
    cs_rId: str = "rId7",
) -> bytes:
    return (
        b'<w:drawing'
        b' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        b' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        b' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        b'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        b'<wp:extent cx="5486400" cy="3200400"/>'
        b'<wp:docPr id="1" name="Diagram 1"/>'
        b'<a:graphic>'
        b'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        b'<dgm:relIds r:dm="%s" r:lo="%s" r:qs="%s" r:cs="%s"/>'
        b'</a:graphicData>'
        b'</a:graphic>'
        b'</wp:inline>'
        b'</w:drawing>' % (
            dm_rId.encode(), lo_rId.encode(), qs_rId.encode(), cs_rId.encode(),
        )
    )


def _inline_from_smart_art_drawing(part: DocumentPart) -> InlineShape:
    """Build a real w:drawing subtree and return an InlineShape over its wp:inline."""
    drawing = parse_xml(_smart_art_drawing_xml())
    inline = drawing.xpath("./wp:inline")[0]
    return InlineShape(cast(CT_Inline, inline), part)


# -- tests -------------------------------------------------------------------


class DescribeInlineShape_is_smart_art:
    def it_is_True_for_an_inline_that_carries_dgm_relIds(
        self, request: FixtureRequest
    ):
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {}
        shape = _inline_from_smart_art_drawing(document_part_)

        assert shape.is_smart_art is True

    def it_is_False_for_a_plain_picture_inline(
        self, request: FixtureRequest
    ):
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {}
        inline = cast(
            CT_Inline,
            element(
                "wp:inline/a:graphic/a:graphicData{uri="
                "http://schemas.openxmlformats.org/drawingml/2006/picture}"
                "/pic:pic"
            ),
        )
        shape = InlineShape(inline, document_part_)

        assert shape.is_smart_art is False


class DescribeInlineShape_smart_art:
    def it_returns_None_when_not_smart_art(self, request: FixtureRequest):
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {}
        inline = cast(
            CT_Inline,
            element(
                "wp:inline/a:graphic/a:graphicData{uri="
                "http://schemas.openxmlformats.org/drawingml/2006/picture}"
                "/pic:pic"
            ),
        )

        assert InlineShape(inline, document_part_).smart_art is None

    def it_returns_a_SmartArt_proxy_when_is_smart_art(
        self, request: FixtureRequest
    ):
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {
            "rId4": _data_part(),
            "rId6": _style_part(),
            "rId7": _colors_part(),
        }
        shape = _inline_from_smart_art_drawing(document_part_)

        sa = shape.smart_art

        assert isinstance(sa, SmartArt)
        assert sa.dm_rId == "rId4"
        assert [n.text for n in sa.nodes] == ["Alpha", "Beta"]

    def it_raises_when_inline_has_no_part_reference(self):
        inline = parse_xml(
            _smart_art_drawing_xml()
        ).xpath("./wp:inline")[0]
        shape = InlineShape(cast(CT_Inline, inline), part=None)

        with pytest.raises(ValueError, match="requires a part reference"):
            _ = shape.smart_art


class DescribeSmartArt_color_and_style_transforms:
    def it_exposes_color_transform_when_colors_part_resolves(
        self, request: FixtureRequest
    ):
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {
            "rId4": _data_part(),
            "rId6": _style_part(),
            "rId7": _colors_part(),
        }
        shape = _inline_from_smart_art_drawing(document_part_)

        sa = shape.smart_art

        assert sa is not None
        ct = sa.color_transform
        assert ct is not None
        # -- the colour transform proxy exposes at least one colour style
        # -- label round-tripped from the colorsDef source --
        assert list(ct.color_styles)  # truthy, non-empty

    def it_exposes_style_transform_when_qs_part_resolves(
        self, request: FixtureRequest
    ):
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {
            "rId4": _data_part(),
            "rId6": _style_part(),
            "rId7": _colors_part(),
        }
        shape = _inline_from_smart_art_drawing(document_part_)

        sa = shape.smart_art

        assert sa is not None
        st = sa.style_transform
        assert st is not None
        assert list(st.style_labels)

    def it_returns_None_for_transforms_when_related_parts_missing(
        self, request: FixtureRequest
    ):
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {}
        shape = _inline_from_smart_art_drawing(document_part_)

        sa = shape.smart_art
        assert sa is not None
        assert sa.color_transform is None
        assert sa.style_transform is None


class DescribeSmartArt_graphic_frame_xml:
    def it_re_serialises_the_host_w_drawing_element(
        self, request: FixtureRequest
    ):
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {
            "rId4": _data_part(),
        }
        shape = _inline_from_smart_art_drawing(document_part_)

        sa = shape.smart_art

        assert sa is not None
        xml_bytes = sa.graphic_frame_xml
        assert isinstance(xml_bytes, bytes)
        # -- contains the signature SmartArt bits from the drawing --
        assert b"dgm:relIds" in xml_bytes
        assert b'r:dm="rId4"' in xml_bytes
        # -- and is wrapped by the w:drawing envelope, not just wp:inline --
        assert xml_bytes.startswith(b"<w:drawing") or b"w:drawing" in xml_bytes[:20]

    def it_is_None_when_SmartArt_has_no_host_drawing(self):
        from ooxml_smartart.oxml.relIds import CT_RelIds

        relIds_xml = (
            b'<dgm:relIds'
            b' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            b' r:dm="rId4"/>'
        )
        relIds = cast(CT_RelIds, parse_xml(relIds_xml))
        sa = SmartArt(relIds, None)

        assert sa.graphic_frame_xml is None


class DescribeDocument_smart_arts_and_iter:
    def _document_with_two_smart_arts(
        self, request: FixtureRequest
    ) -> Document:
        document_part_ = instance_mock(request, DocumentPart)
        document_part_.related_parts = {
            "rId4": _data_part(idx=1),
            "rId8": _data_part(idx=2),
        }
        # -- craft a body with two w:drawing elements carrying dgm:relIds --
        body_xml = (
            b'<w:document'
            b' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            b' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
            b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            b' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
            b'<w:body>'
            b'<w:p><w:r>' + _smart_art_drawing_xml("rId4") + b'</w:r></w:p>'
            b'<w:p><w:r>' + _smart_art_drawing_xml("rId8") + b'</w:r></w:p>'
            b'</w:body>'
            b'</w:document>'
        )
        doc_elm = cast(CT_Document, parse_xml(body_xml))
        return Document(doc_elm, document_part_)

    def it_returns_a_list_from_smart_arts_plural(
        self, request: FixtureRequest
    ):
        document = self._document_with_two_smart_arts(request)

        smart_arts = document.smart_arts

        assert len(smart_arts) == 2
        assert smart_arts[0].data_partname == "/word/diagrams/data1.xml"
        assert smart_arts[1].data_partname == "/word/diagrams/data2.xml"

    def it_matches_the_historical_smart_art_singular(
        self, request: FixtureRequest
    ):
        document = self._document_with_two_smart_arts(request)

        assert [sa.dm_rId for sa in document.smart_arts] == [
            sa.dm_rId for sa in document.smart_art
        ]

    def it_iter_smart_arts_yields_the_same_sequence(
        self, request: FixtureRequest
    ):
        document = self._document_with_two_smart_arts(request)

        streamed = list(document.iter_smart_arts())

        assert len(streamed) == 2
        assert [sa.dm_rId for sa in streamed] == ["rId4", "rId8"]

    def it_iter_smart_arts_is_a_generator(self, request: FixtureRequest):
        import types

        document = self._document_with_two_smart_arts(request)

        assert isinstance(document.iter_smart_arts(), types.GeneratorType)
