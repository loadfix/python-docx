# pyright: reportPrivateUsage=false

"""Unit-test suite for `Run.add_ole_object`."""

from __future__ import annotations

import io
from typing import cast

import pytest

from docx.embedded_objects import EmbeddedObject
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R
from docx.parts.embedded_object import EmbeddedObjectPart
from docx.text.paragraph import Paragraph
from docx.text.run import Run, _content_type_for_ole

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, instance_mock


FAKE_XLSX = b"PK\x03\x04" + b"\x00" * 16 + b"xlsx-payload"
FAKE_PDF = b"%PDF-1.4\n%abcd pdf-payload"


class _FakePackage:
    def __init__(self):
        self._counter = 0

    def next_partname(self, template: str) -> PackURI:
        self._counter += 1
        return PackURI(template % self._counter)


class _FakeStoryPart:
    """Minimal StoryPart stand-in sufficient for `add_ole_object`."""

    def __init__(self, package):
        self.package = package
        self.related_parts: dict[str, object] = {}
        self._rel_counter = 0

    def relate_to(self, target, reltype: str, is_external: bool = False) -> str:
        self._rel_counter += 1
        rId = f"rId{self._rel_counter}"
        self.related_parts[rId] = target
        return rId

    def get_or_add_image(self, descriptor):
        # -- minimal stub: map an rId to the descriptor and return it --
        self._rel_counter += 1
        rId = f"rId{self._rel_counter}"
        self.related_parts[rId] = descriptor
        return rId, None


class DescribeContentTypeForOle:
    def it_maps_excel_prog_id_to_xlsx_content_type(self):
        assert _content_type_for_ole("Excel.Sheet.12", "", b"") == CT.SML_SHEET

    def it_maps_acro_prog_id_to_pdf_content_type(self):
        assert _content_type_for_ole("AcroExch.Document.DC", "", b"") == CT.PDF

    def it_falls_back_to_extension_hint_for_xlsx(self):
        assert _content_type_for_ole("Unknown", "xlsx", b"") == CT.SML_SHEET

    def it_falls_back_to_extension_hint_for_pdf(self):
        assert _content_type_for_ole("Unknown", "pdf", b"") == CT.PDF

    def it_falls_back_to_extension_hint_for_zip(self):
        assert _content_type_for_ole("Unknown", "zip", b"") == CT.ZIP

    def it_sniffs_pk_bytes_as_zip(self):
        assert _content_type_for_ole("Unknown", "", b"PK\x03\x04") == CT.ZIP

    def it_sniffs_pdf_magic_bytes(self):
        assert _content_type_for_ole("Unknown", "", b"%PDF-1.4") == CT.PDF

    def it_falls_back_to_generic_ole(self):
        assert (
            _content_type_for_ole("Unknown", "bin", b"\x00\x01") == CT.OFC_OLE_OBJECT
        )


class DescribeRun_add_ole_object:
    """Unit-test suite for `Run.add_ole_object`."""

    def it_creates_an_embedding_part_and_object_element(
        self, request: FixtureRequest, tmp_path
    ):
        package = _FakePackage()
        story_part = _FakeStoryPart(package)

        class FakeParagraph:
            @property
            def part(self):
                return story_part

        xlsx_path = tmp_path / "data.xlsx"
        xlsx_path.write_bytes(FAKE_XLSX)

        p_elm = cast(CT_P, element("w:p/w:r"))
        r = Run(cast(CT_R, p_elm[0]), FakeParagraph())  # pyright: ignore[reportArgumentType]

        eo = r.add_ole_object(str(xlsx_path), prog_id="Excel.Sheet.12")

        assert isinstance(eo, EmbeddedObject)
        # -- w:object/o:OLEObject appended to run --
        objs = r._r.xpath(".//w:object/o:OLEObject")
        assert len(objs) == 1
        ole = objs[0]
        assert ole.get("ProgID") == "Excel.Sheet.12"
        assert ole.get("Type") == "Embed"
        assert ole.get(qn("r:id")) == "rId1"
        # -- embedded part created with the right content type --
        part = story_part.related_parts["rId1"]
        assert isinstance(part, EmbeddedObjectPart)
        assert part.content_type == CT.SML_SHEET
        assert part.blob == FAKE_XLSX
        assert str(part.partname) == "/word/embeddings/oleObject1.bin"

    def it_accepts_a_file_like_stream(self, request: FixtureRequest):
        package = _FakePackage()
        story_part = _FakeStoryPart(package)

        class FakeParagraph:
            @property
            def part(self):
                return story_part

        stream = io.BytesIO(FAKE_PDF)
        p_elm = cast(CT_P, element("w:p/w:r"))
        r = Run(cast(CT_R, p_elm[0]), FakeParagraph())  # pyright: ignore[reportArgumentType]

        eo = r.add_ole_object(stream, prog_id="AcroExch.Document.DC")

        part = story_part.related_parts["rId1"]
        assert isinstance(part, EmbeddedObjectPart)
        assert part.content_type == CT.PDF
        assert part.blob == FAKE_PDF
        assert eo.prog_id == "AcroExch.Document.DC"

    def it_selects_generic_content_type_for_unknown_prog_id(
        self, request: FixtureRequest, tmp_path
    ):
        package = _FakePackage()
        story_part = _FakeStoryPart(package)

        class FakeParagraph:
            @property
            def part(self):
                return story_part

        payload = b"arbitrary-ole-bytes"
        blob_path = tmp_path / "weird.bin"
        blob_path.write_bytes(payload)
        p_elm = cast(CT_P, element("w:p/w:r"))
        r = Run(cast(CT_R, p_elm[0]), FakeParagraph())  # pyright: ignore[reportArgumentType]

        r.add_ole_object(str(blob_path), prog_id="Unknown.Prog")

        part = story_part.related_parts["rId1"]
        assert isinstance(part, EmbeddedObjectPart)
        assert part.content_type == CT.OFC_OLE_OBJECT

    def it_emits_shape_fallback_when_icon_supplied(
        self, request: FixtureRequest, tmp_path
    ):
        package = _FakePackage()
        story_part = _FakeStoryPart(package)

        class FakeParagraph:
            @property
            def part(self):
                return story_part

        xlsx_path = tmp_path / "data.xlsx"
        xlsx_path.write_bytes(FAKE_XLSX)
        icon_stream = io.BytesIO(b"fake-png")

        p_elm = cast(CT_P, element("w:p/w:r"))
        r = Run(cast(CT_R, p_elm[0]), FakeParagraph())  # pyright: ignore[reportArgumentType]

        r.add_ole_object(
            str(xlsx_path), prog_id="Excel.Sheet.12", icon_path_or_stream=icon_stream
        )

        shapes = r._r.xpath(".//w:object/v:shape")
        assert len(shapes) == 1
        imagedata = shapes[0].xpath("./v:imagedata")[0]
        # -- icon rId is rId2 (OLE payload took rId1). --
        assert imagedata.get(qn("r:id")) == "rId2"
        # -- ensure OLE object references its own rId (not the icon) --
        ole = r._r.xpath(".//o:OLEObject")[0]
        assert ole.get(qn("r:id")) == "rId1"
