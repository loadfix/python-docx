# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.comments` module."""

from __future__ import annotations

from typing import cast
from unittest.mock import patch

import pytest

from docx.oxml.comments import CT_Comment, CT_Comments
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml

from ..unitutil.cxml import element


class DescribeCT_Comments:
    """Unit-test suite for `docx.oxml.comments.CT_Comments`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:comments", 0),
            ("w:comments/(w:comment{w:id=1})", 2),
            ("w:comments/(w:comment{w:id=4},w:comment{w:id=2147483646})", 2147483647),
            ("w:comments/(w:comment{w:id=1},w:comment{w:id=2147483647})", 0),
            ("w:comments/(w:comment{w:id=1},w:comment{w:id=2},w:comment{w:id=3})", 4),
        ],
    )
    def it_finds_the_next_available_comment_id_to_help(self, cxml: str, expected_value: int):
        comments_elm = cast(CT_Comments, element(cxml))
        assert comments_elm._next_available_comment_id() == expected_value

    def it_can_add_a_comment_with_a_paraId(self):
        comments_elm = cast(CT_Comments, element("w:comments"))

        comment = comments_elm.add_comment()

        assert comment.paraId is not None
        assert len(comment.paraId) == 8
        # -- paraId should be a hex string --
        int(comment.paraId, 16)

    def it_generates_unique_paraIds(self):
        comments_elm = cast(CT_Comments, element("w:comments"))

        comment1 = comments_elm.add_comment()
        comment2 = comments_elm.add_comment()

        assert comment1.paraId != comment2.paraId

    def it_can_add_a_reply_comment(self):
        comments_elm = cast(CT_Comments, element("w:comments"))
        parent = comments_elm.add_comment()
        parent_para_id = parent.paraId
        assert parent_para_id is not None

        reply = comments_elm.add_reply(parent_para_id)

        assert reply.paraIdParent == parent_para_id
        assert reply.paraId is not None
        assert reply.paraId != parent_para_id
        assert reply.id != parent.id

    def it_can_find_replies_for_a_comment(self):
        comments_elm = cast(CT_Comments, element("w:comments"))
        parent = comments_elm.add_comment()
        parent_para_id = parent.paraId
        assert parent_para_id is not None
        reply1 = comments_elm.add_reply(parent_para_id)
        reply2 = comments_elm.add_reply(parent_para_id)
        # -- add an unrelated comment to make sure it's not included --
        comments_elm.add_comment()

        replies = comments_elm.get_replies_for(parent_para_id)

        assert len(replies) == 2
        assert replies[0] is reply1
        assert replies[1] is reply2

    def but_it_returns_empty_list_when_no_replies(self):
        comments_elm = cast(CT_Comments, element("w:comments"))
        parent = comments_elm.add_comment()
        parent_para_id = parent.paraId
        assert parent_para_id is not None

        replies = comments_elm.get_replies_for(parent_para_id)

        assert replies == []


class DescribeCT_Comment:
    """Unit-test suite for `docx.oxml.comments.CT_Comment`."""

    def it_can_get_and_set_paraId(self):
        comment_elm = cast(CT_Comment, element("w:comment{w:id=1}"))

        assert comment_elm.paraId is None

        comment_elm.paraId = "AABB0011"
        assert comment_elm.paraId == "AABB0011"

        comment_elm.paraId = None
        assert comment_elm.paraId is None

    def it_can_get_and_set_paraIdParent(self):
        comment_elm = cast(CT_Comment, element("w:comment{w:id=1}"))

        assert comment_elm.paraIdParent is None

        comment_elm.paraIdParent = "CCDD2233"
        assert comment_elm.paraIdParent == "CCDD2233"

        comment_elm.paraIdParent = None
        assert comment_elm.paraIdParent is None
