"""Unit test suite for the docx.parts.footnotes module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import PartFactory
from docx.oxml.footnotes import CT_Footnotes
from docx.package import Package
from docx.parts.footnotes import FootnotesPart

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock, method_mock


class DescribeFootnotesPart:
    """Unit test suite for `docx.parts.footnotes.FootnotesPart` objects."""

    def it_is_used_by_the_part_loader_to_construct_a_footnotes_part(
        self, package_: Mock, FootnotesPart_load_: Mock, footnotes_part_: Mock
    ):
        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        reltype = RT.FOOTNOTES
        blob = b"<w:footnotes/>"
        FootnotesPart_load_.return_value = footnotes_part_

        part = PartFactory(partname, content_type, reltype, blob, package_)

        FootnotesPart_load_.assert_called_once_with(partname, content_type, blob, package_)
        assert part is footnotes_part_

    def it_provides_access_to_its_footnotes_element(self, package_: Mock):
        footnotes_elm = cast(CT_Footnotes, element("w:footnotes"))
        footnotes_part = FootnotesPart(
            PackURI("/word/footnotes.xml"), CT.WML_FOOTNOTES, footnotes_elm, package_
        )

        assert footnotes_part.footnotes_element is footnotes_elm

    def it_constructs_a_default_footnotes_part_to_help(self):
        package = Package()

        footnotes_part = FootnotesPart.default(package)

        assert isinstance(footnotes_part, FootnotesPart)
        assert footnotes_part.partname == "/word/footnotes.xml"
        assert footnotes_part.content_type == CT.WML_FOOTNOTES
        assert footnotes_part.package is package
        assert footnotes_part.element.tag == (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnotes"
        )
        # default template has separator (id=0) and continuation separator (id=1)
        footnote_elms = footnotes_part.element.xpath("./w:footnote")
        assert len(footnote_elms) == 2
        assert footnote_elms[0].id == 0
        assert footnote_elms[1].id == 1

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def footnotes_part_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, FootnotesPart)

    @pytest.fixture
    def FootnotesPart_load_(self, request: FixtureRequest) -> Mock:
        return method_mock(request, FootnotesPart, "load", autospec=False)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
