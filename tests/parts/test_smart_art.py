"""Unit test suite for the `docx.parts.smart_art` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.parser import parse_xml
from docx.oxml.smart_art import CT_DataModel
from docx.package import Package
from docx.parts.smart_art import DiagramDataPart

from ..unitutil.mock import FixtureRequest, Mock, instance_mock


DATA_MODEL_XML = (
    b'<?xml version="1.0"?>\n'
    b'<dgm:dataModel'
    b' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
    b' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
    b"  <dgm:ptLst>\n"
    b'    <dgm:pt modelId="n1" type="node">\n'
    b'      <dgm:t><a:p><a:r><a:t>First</a:t></a:r></a:p></dgm:t>\n'
    b"    </dgm:pt>\n"
    b'    <dgm:pt modelId="n2" type="node">\n'
    b'      <dgm:t><a:p><a:r><a:t>Second</a:t></a:r></a:p></dgm:t>\n'
    b"    </dgm:pt>\n"
    b"  </dgm:ptLst>\n"
    b"</dgm:dataModel>\n"
)


class DescribeDiagramDataPart:
    """Unit test suite for `docx.parts.smart_art.DiagramDataPart`."""

    def it_exposes_its_data_model(self, package_: Mock):
        element = cast(CT_DataModel, parse_xml(DATA_MODEL_XML))
        part = DiagramDataPart(
            PackURI("/word/diagrams/data1.xml"),
            CT.DML_DIAGRAM_DATA,
            element,
            package_,
        )

        data_model = part.data_model

        assert isinstance(data_model, CT_DataModel)
        assert len(data_model.pt_lst) == 2
        assert data_model.pt_lst[0].modelId == "n1"

    def it_is_loaded_by_the_part_factory(self, request: FixtureRequest):
        package_ = instance_mock(request, Package)
        partname = PackURI("/word/diagrams/data1.xml")

        part = PartFactory(
            partname,
            CT.DML_DIAGRAM_DATA,
            RT.DIAGRAM_DATA,
            DATA_MODEL_XML,
            package_,
        )

        assert isinstance(part, DiagramDataPart)
        assert part.partname == partname
        assert part.content_type == CT.DML_DIAGRAM_DATA
        assert isinstance(part.data_model, CT_DataModel)

    def it_round_trips_the_blob(self, package_: Mock):
        element = cast(CT_DataModel, parse_xml(DATA_MODEL_XML))
        part = DiagramDataPart(
            PackURI("/word/diagrams/data1.xml"),
            CT.DML_DIAGRAM_DATA,
            element,
            package_,
        )

        # -- the blob is the serialized XML; check it at least contains the text --
        assert b"First" in part.blob
        assert b"Second" in part.blob

    # -- fixtures -----------------------------------------------------------------

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
