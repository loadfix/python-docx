# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.oxml.comments_extended` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.comments_extended import (
    CT_CommentExtended,
    CT_CommentExtendedList,
    CT_PresenceInfo,
)

from ..unitutil.cxml import element


class DescribeCT_CommentExtendedList:
    """Unit-test suite for `docx.oxml.comments_extended.CT_CommentExtendedList`."""

    def it_creates_an_empty_root_element(self):
        ex_list = CT_CommentExtendedList.new()

        assert ex_list.tag == (
            "{http://schemas.microsoft.com/office/word/2012/wordml}commentsEx"
        )
        assert len(ex_list) == 0

    def it_can_add_a_commentEx_entry(self):
        ex_list = CT_CommentExtendedList.new()

        commentEx = ex_list.add_commentEx(paraId="ABCD1234", done=True)

        assert isinstance(commentEx, CT_CommentExtended)
        assert commentEx.paraId == "ABCD1234"
        assert commentEx.done is True
        assert commentEx.paraIdParent is None

    def it_can_add_a_commentEx_with_parent(self):
        ex_list = CT_CommentExtendedList.new()

        commentEx = ex_list.add_commentEx(
            paraId="11112222", done=False, parentParaId="ABCD1234"
        )

        assert commentEx.paraId == "11112222"
        assert commentEx.done is False
        assert commentEx.paraIdParent == "ABCD1234"

    def it_can_lookup_a_commentEx_by_paraId(self):
        ex_list = cast(
            CT_CommentExtendedList,
            element(
                "w15:commentsEx/("
                "w15:commentEx{w15:paraId=AAAA0001,w15:done=0},"
                "w15:commentEx{w15:paraId=BBBB0002,w15:done=1})"
            ),
        )

        match = ex_list.get_commentEx_by_paraId("BBBB0002")

        assert match is not None
        assert match.paraId == "BBBB0002"
        assert match.done is True

    def but_it_returns_None_when_paraId_is_absent(self):
        ex_list = cast(
            CT_CommentExtendedList,
            element("w15:commentsEx/w15:commentEx{w15:paraId=AAAA,w15:done=0}"),
        )

        assert ex_list.get_commentEx_by_paraId("ZZZZ") is None

    def it_can_list_children_of_a_parent_paraId(self):
        ex_list = CT_CommentExtendedList.new()
        ex_list.add_commentEx(paraId="PARENT01", done=False)
        c1 = ex_list.add_commentEx(
            paraId="CHILD001", done=False, parentParaId="PARENT01"
        )
        c2 = ex_list.add_commentEx(
            paraId="CHILD002", done=True, parentParaId="PARENT01"
        )
        # -- an unrelated child of a different parent --
        ex_list.add_commentEx(paraId="OTHER001", done=False, parentParaId="XXXX1111")

        children = ex_list.get_children_for("PARENT01")

        assert len(children) == 2
        assert children[0] is c1
        assert children[1] is c2


class DescribeCT_CommentExtended:
    """Unit-test suite for `docx.oxml.comments_extended.CT_CommentExtended`."""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [("1", True), ("0", False), ("true", True), ("false", False)],
    )
    def it_reads_xsd_boolean_done_values(self, raw: str, expected: bool):
        commentEx = cast(
            CT_CommentExtended,
            element("w15:commentEx{w15:paraId=A,w15:done=%s}" % raw),
        )

        assert commentEx.done is expected

    def it_can_toggle_the_done_attribute(self):
        commentEx = cast(
            CT_CommentExtended,
            element("w15:commentEx{w15:paraId=A,w15:done=0}"),
        )

        commentEx.done = True

        assert commentEx.done is True


class DescribeCT_PresenceInfo:
    """Unit-test suite for `docx.oxml.comments_extended.CT_PresenceInfo`."""

    def it_exposes_the_provider_and_user_id_pair(self):
        presence = cast(
            CT_PresenceInfo,
            element(
                "w15:presenceInfo{w15:providerId=AD,w15:userId=S-1-5-21-abc}"
            ),
        )

        assert presence.providerId == "AD"
        assert presence.userId == "S-1-5-21-abc"
