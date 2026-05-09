"""Unit test suite for the ``docx.parts.vml`` module."""

from __future__ import annotations

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.parts.vml import LegacyDrawingPart, VmlDrawingPart

from ..unitutil.mock import FixtureRequest, Mock, instance_mock

# A minimal but well-formed VML watermark snippet modelled on the one
# Word writes for a header/footer "Draft" banner.  Any byte mutation
# during round-trip breaks the assertion.
VML_BLOB = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<xml xmlns:v="urn:schemas-microsoft-com:vml"'
    b' xmlns:o="urn:schemas-microsoft-com:office:office">\r\n'
    b'  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/>'
    b'</o:shapelayout>\r\n'
    b'  <v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s1025"'
    b' type="#_x0000_t136" style="width:415pt;height:103.5pt"'
    b' fillcolor="silver" stroked="f">\r\n'
    b'    <v:textpath style="font-family:&quot;Calibri&quot;;"'
    b' string="Draft"/>\r\n'
    b'  </v:shape>\r\n'
    b'</xml>\r\n'
)


class DescribeVmlDrawingPart:
    """Unit test suite for ``docx.parts.vml.VmlDrawingPart`` objects."""

    def it_is_used_by_the_part_loader_to_construct_a_vml_drawing_part(
        self, package_: Mock
    ):
        partname = PackURI("/word/vmlDrawing1.vml")
        content_type = CT.OFC_VML_DRAWING

        part = PartFactory(
            partname, content_type, "irrelevant-rel-type", VML_BLOB, package_
        )

        assert isinstance(part, VmlDrawingPart)
        assert part.partname == partname
        assert part.content_type == content_type

    def it_preserves_the_blob_byte_identical_on_round_trip(
        self, package_: Mock
    ):
        partname = PackURI("/word/vmlDrawing1.vml")

        part = VmlDrawingPart.load(
            partname, CT.OFC_VML_DRAWING, VML_BLOB, package_
        )

        assert part.blob == VML_BLOB
        # -- .blob is the re-emit path the package writer calls --
        assert part.blob is not VML_BLOB or part.blob == VML_BLOB

    def it_exposes_an_ooxml_vml_vml_drawing_part_facade(
        self, package_: Mock
    ):
        from ooxml_vml import VmlDrawingPart as SharedVmlDrawingPart

        partname = PackURI("/word/vmlDrawing1.vml")
        part = VmlDrawingPart.load(
            partname, CT.OFC_VML_DRAWING, VML_BLOB, package_
        )

        assert isinstance(part.vml_part, SharedVmlDrawingPart)
        assert part.vml_part.blob == VML_BLOB
        assert part.vml_part.partname == str(partname)

    def it_memoises_the_vml_part_facade_across_accesses(
        self, package_: Mock
    ):
        partname = PackURI("/word/vmlDrawing1.vml")
        part = VmlDrawingPart.load(
            partname, CT.OFC_VML_DRAWING, VML_BLOB, package_
        )

        assert part.vml_part is part.vml_part

    def it_is_registered_for_the_vml_drawing_content_type(self):
        # -- importing ``docx`` registers the class into the PartFactory
        # -- registry.  Idempotent re-import is fine.
        import docx  # noqa: F401  (side-effect import)

        assert PartFactory.part_type_for[CT.OFC_VML_DRAWING] is VmlDrawingPart

    # -- fixtures ---------------------------------------------------------------

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        from docx.package import Package

        return instance_mock(request, Package)


class DescribeLegacyDrawingPart:
    """Alias of :class:`VmlDrawingPart` for Word legacy-drawing callers."""

    def it_is_a_subclass_of_vml_drawing_part(self):
        assert issubclass(LegacyDrawingPart, VmlDrawingPart)

    def it_preserves_the_blob_byte_identical_on_round_trip(
        self, package_: Mock
    ):
        partname = PackURI("/word/vmlDrawing1.vml")

        part = LegacyDrawingPart.load(
            partname, CT.OFC_VML_DRAWING, VML_BLOB, package_
        )

        assert part.blob == VML_BLOB
        from ooxml_vml import VmlDrawingPart as SharedVmlDrawingPart

        assert isinstance(part.vml_part, SharedVmlDrawingPart)

    # -- fixtures ---------------------------------------------------------------

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        from docx.package import Package

        return instance_mock(request, Package)
