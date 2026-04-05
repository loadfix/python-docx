"""Unit test suite for the docx.parts.endnotes module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.endnotes import CT_Endnotes
from docx.package import Package
from docx.parts.endnotes import EndnotesPart

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock, method_mock


class DescribeEndnotesPart:
    """Unit test suite for `docx.parts.endnotes.EndnotesPart` objects."""

    def it_is_used_by_the_part_loader_to_construct_an_endnotes_part(
        self, package_: Mock, EndnotesPart_load_: Mock, endnotes_part_: Mock
    ):
        partname = PackURI("/word/endnotes.xml")
        content_type = CT.WML_ENDNOTES
        reltype = RT.ENDNOTES
        blob = b"<w:endnotes/>"
        EndnotesPart_load_.return_value = endnotes_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        EndnotesPart_load_.assert_called_once_with(partname, content_type, blob, package_)
        assert part is endnotes_part_

    def it_provides_access_to_its_endnotes_element(self, package_: Mock):
        endnotes_elm = cast(CT_Endnotes, element("w:endnotes"))
        endnotes_part = EndnotesPart(
            PackURI("/word/endnotes.xml"), CT.WML_ENDNOTES, endnotes_elm, package_
        )

        assert endnotes_part.endnotes_element is endnotes_elm

    def it_constructs_a_default_endnotes_part_to_help(self):
        package = Package()

        endnotes_part = EndnotesPart.default(package)

        assert isinstance(endnotes_part, EndnotesPart)
        assert endnotes_part.partname == "/word/endnotes.xml"
        assert endnotes_part.content_type == CT.WML_ENDNOTES
        assert endnotes_part.package is package
        assert endnotes_part.element.tag == (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}endnotes"
        )
        # default template has separator (id=0) and continuation separator (id=1)
        endnote_elms = endnotes_part.element.xpath("./w:endnote")
        assert len(endnote_elms) == 2
        assert endnote_elms[0].id == 0
        assert endnote_elms[1].id == 1

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def endnotes_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, EndnotesPart)

    @pytest.fixture
    def EndnotesPart_load_(self, request: FixtureRequest) -> Mock:
        return method_mock(request, EndnotesPart, "load", autospec=False)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
