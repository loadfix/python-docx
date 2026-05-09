"""Unit test suite for the `docx.parts.glossary` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.glossary import Glossary
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.glossary import CT_GlossaryDocument
from docx.package import Package
from docx.parts.glossary import GlossaryPart

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock


class DescribeGlossaryPart:
    """Unit test suite for `docx.parts.glossary.GlossaryPart`."""

    def it_provides_access_to_its_glossary_proxy(
        self, Glossary_: Mock, glossary_: Mock, package_: Mock
    ):
        Glossary_.return_value = glossary_
        glossary_elm = cast(CT_GlossaryDocument, element("w:glossaryDocument"))
        glossary_part = GlossaryPart(
            PackURI("/word/glossary/document.xml"),
            CT.WML_DOCUMENT_GLOSSARY,
            glossary_elm,
            package_,
        )

        glossary = glossary_part.glossary

        Glossary_.assert_called_once_with(glossary_elm, glossary_part)
        assert glossary is glossary_

    def it_exposes_its_glossary_element(self, package_: Mock):
        glossary_elm = cast(CT_GlossaryDocument, element("w:glossaryDocument"))
        glossary_part = GlossaryPart(
            PackURI("/word/glossary/document.xml"),
            CT.WML_DOCUMENT_GLOSSARY,
            glossary_elm,
            package_,
        )

        assert glossary_part.glossary_element is glossary_elm

    def it_provides_a_default_empty_glossary_part(self, package_: Mock):
        part = GlossaryPart.default(package_)
        assert isinstance(part, GlossaryPart)
        assert part.partname == PackURI("/word/glossary/document.xml")
        assert part.content_type == CT.WML_DOCUMENT_GLOSSARY
        # -- the root has an empty w:docParts container ready to append to --
        assert part.glossary_element.docParts is not None
        assert len(part.glossary_element.docPart_lst) == 0

    # -- fixtures ------------------------------------------------------------

    @pytest.fixture
    def Glossary_(self, request: FixtureRequest) -> Mock:
        return class_mock(request, "docx.parts.glossary.Glossary")

    @pytest.fixture
    def glossary_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Glossary)

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)
