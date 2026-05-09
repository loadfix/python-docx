"""Unit-test suite for `docx.parts.comments_extended`."""

from __future__ import annotations

import pytest

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.comments_extended import CT_CommentExtendedList
from docx.package import Package
from docx.parts.comments import CommentsPart
from docx.parts.comments_extended import CommentsExtendedPart

from ..unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeCommentsExtendedPart:
    """Unit-test suite for `docx.parts.comments_extended.CommentsExtendedPart`."""

    def it_constructs_a_default_commentsExtended_part_to_help(self):
        package = Package()

        part = CommentsExtendedPart.default(package)

        assert isinstance(part, CommentsExtendedPart)
        assert part.partname == "/word/commentsExtended.xml"
        assert part.content_type == CT.WML_COMMENTS_EXTENDED
        assert part.package is package
        assert isinstance(part.element, CT_CommentExtendedList)
        assert part.element.tag == (
            "{http://schemas.microsoft.com/office/word/2012/wordml}commentsEx"
        )
        assert len(part.element) == 0

    def it_exposes_the_root_element_via_the_element_property(self):
        package = Package()
        part = CommentsExtendedPart.default(package)

        assert part.element is part._comments_ex  # type: ignore[attr-defined]


class DescribeCommentsPart_extendedAccessors:
    """Integration suite for the new `comments_extended_part*` getters on CommentsPart."""

    def it_returns_None_when_no_extended_part_is_related(self):
        package = Package()
        comments_part = CommentsPart.default(package)

        assert comments_part.comments_extended_part is None

    def it_creates_and_relates_the_extended_part_on_demand(self):
        package = Package()
        comments_part = CommentsPart.default(package)

        ex_part = comments_part.comments_extended_part_or_add()

        assert isinstance(ex_part, CommentsExtendedPart)
        # -- the relationship is established from the comments part --
        related = comments_part.part_related_by(RT.COMMENTS_EXTENDED)
        assert related is ex_part

    def it_returns_the_same_part_on_subsequent_calls(self):
        package = Package()
        comments_part = CommentsPart.default(package)

        first = comments_part.comments_extended_part_or_add()
        second = comments_part.comments_extended_part_or_add()
        third = comments_part.comments_extended_part

        assert first is second
        assert first is third

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def package_(self, request: FixtureRequest) -> Mock:
        return instance_mock(request, Package)

    @pytest.fixture
    def partname_(self) -> PackURI:
        return PackURI("/word/commentsExtended.xml")
