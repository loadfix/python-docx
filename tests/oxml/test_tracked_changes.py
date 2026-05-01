# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.tracked_changes` module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.oxml.tracked_changes import (
    CT_CellDel,
    CT_CellIns,
    CT_Del,
    CT_DelText,
    CT_Ins,
    CT_MoveFrom,
    CT_MoveTo,
    CT_TblPrChange,
    CT_TcPrChange,
    CT_TrPrChange,
)

from ..unitutil.cxml import element


class DescribeCT_Ins:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_Ins`."""

    def it_knows_its_id(self):
        ins = cast(CT_Ins, element("w:ins{w:id=1,w:author=Alice}"))
        assert ins.id == 1

    def it_knows_its_author(self):
        ins = cast(CT_Ins, element("w:ins{w:id=1,w:author=Alice}"))
        assert ins.author == "Alice"

    def it_knows_its_date(self):
        ins = cast(CT_Ins, element("w:ins{w:id=1,w:author=Alice,w:date=2023-10-01T12:00:00Z}"))
        assert ins.date == dt.datetime(2023, 10, 1, 12, 0, 0, tzinfo=dt.timezone.utc)

    def it_returns_None_when_date_is_absent(self):
        ins = cast(CT_Ins, element("w:ins{w:id=1,w:author=Alice}"))
        assert ins.date is None

    @pytest.mark.parametrize(
        ("cxml", "expected_text"),
        [
            ("w:ins{w:id=1,w:author=A}", ""),
            ('w:ins{w:id=1,w:author=A}/w:r/w:t"hello"', "hello"),
            (
                'w:ins{w:id=1,w:author=A}/(w:r/w:t"hello ",w:r/w:t"world")',
                "hello world",
            ),
        ],
    )
    def it_can_produce_its_text(self, cxml: str, expected_text: str):
        ins = cast(CT_Ins, element(cxml))
        assert ins.text == expected_text

    def it_provides_access_to_its_runs(self):
        ins = cast(CT_Ins, element('w:ins{w:id=1,w:author=A}/(w:r/w:t"a",w:r/w:t"b")'))
        assert len(ins.r_lst) == 2


class DescribeCT_Del:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_Del`."""

    def it_knows_its_id(self):
        del_elm = cast(CT_Del, element("w:del{w:id=2,w:author=Bob}"))
        assert del_elm.id == 2

    def it_knows_its_author(self):
        del_elm = cast(CT_Del, element("w:del{w:id=2,w:author=Bob}"))
        assert del_elm.author == "Bob"

    def it_knows_its_date(self):
        del_elm = cast(
            CT_Del, element("w:del{w:id=2,w:author=Bob,w:date=2023-11-15T09:30:00Z}")
        )
        assert del_elm.date == dt.datetime(2023, 11, 15, 9, 30, 0, tzinfo=dt.timezone.utc)

    def it_returns_None_when_date_is_absent(self):
        del_elm = cast(CT_Del, element("w:del{w:id=2,w:author=Bob}"))
        assert del_elm.date is None

    @pytest.mark.parametrize(
        ("cxml", "expected_text"),
        [
            ("w:del{w:id=2,w:author=B}", ""),
            ('w:del{w:id=2,w:author=B}/w:r/w:delText"removed"', "removed"),
            (
                'w:del{w:id=2,w:author=B}/(w:r/w:delText"foo ",w:r/w:delText"bar")',
                "foo bar",
            ),
        ],
    )
    def it_can_produce_its_text(self, cxml: str, expected_text: str):
        del_elm = cast(CT_Del, element(cxml))
        assert del_elm.text == expected_text


class DescribeCT_DelText:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_DelText`."""

    def it_can_report_its_text(self):
        dt_elm = cast(CT_DelText, element('w:delText"some deleted text"'))
        assert str(dt_elm) == "some deleted text"

    def it_returns_empty_string_when_no_content(self):
        dt_elm = cast(CT_DelText, element("w:delText"))
        assert str(dt_elm) == ""


class DescribeCT_Ins_acceptReject:
    """Accept/reject behavior for `<w:ins>`."""

    def it_unwraps_itself_on_accept_keeping_inserted_runs(self):
        p = element(
            'w:p/(w:r/w:t"before",w:ins{w:id=1,w:author=A}/w:r/w:t"inserted",w:r/w:t"after")'
        )
        ins = p.xpath("./w:ins")[0]
        ins.accept()
        assert p.xpath("./w:ins") == []
        assert [r.text for r in p.xpath("./w:r/w:t")] == ["before", "inserted", "after"]

    def it_removes_itself_on_reject_discarding_inserted_runs(self):
        p = element(
            'w:p/(w:r/w:t"before",w:ins{w:id=1,w:author=A}/w:r/w:t"inserted",w:r/w:t"after")'
        )
        ins = p.xpath("./w:ins")[0]
        ins.reject()
        assert p.xpath("./w:ins") == []
        assert [r.text for r in p.xpath("./w:r/w:t")] == ["before", "after"]


class DescribeCT_Del_acceptReject:
    """Accept/reject behavior for `<w:del>`."""

    def it_removes_itself_on_accept_discarding_deleted_content(self):
        p = element(
            'w:p/(w:r/w:t"keep",w:del{w:id=2,w:author=B}/w:r/w:delText"gone")'
        )
        del_ = p.xpath("./w:del")[0]
        del_.accept()
        assert p.xpath("./w:del") == []
        assert [r.text for r in p.xpath("./w:r/w:t")] == ["keep"]

    def it_restores_content_on_reject_converting_delText_to_t(self):
        p = element(
            'w:p/(w:r/w:t"keep ",w:del{w:id=2,w:author=B}/w:r/w:delText"restore")'
        )
        del_ = p.xpath("./w:del")[0]
        del_.reject()
        assert p.xpath("./w:del") == []
        assert p.xpath("./w:r/w:delText") == []
        # Both runs survive; their text values are "keep " and "restore"
        texts = [t.text for t in p.xpath("./w:r/w:t")]
        assert texts == ["keep ", "restore"]


class DescribeCT_MoveFrom:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_MoveFrom`."""

    def it_knows_its_id(self):
        mf = cast(
            CT_MoveFrom,
            element("w:moveFrom{w:id=1,w:author=Alice,w:name=m1}"),
        )
        assert mf.id == 1

    def it_knows_its_author(self):
        mf = cast(
            CT_MoveFrom,
            element("w:moveFrom{w:id=1,w:author=Alice,w:name=m1}"),
        )
        assert mf.author == "Alice"

    def it_knows_its_name(self):
        mf = cast(
            CT_MoveFrom,
            element("w:moveFrom{w:id=1,w:author=Alice,w:name=m1}"),
        )
        assert mf.name == "m1"

    def it_returns_None_when_name_is_absent(self):
        mf = cast(CT_MoveFrom, element("w:moveFrom{w:id=1,w:author=A}"))
        assert mf.name is None

    def it_can_produce_its_text_from_delText_children(self):
        mf = cast(
            CT_MoveFrom,
            element(
                'w:moveFrom{w:id=1,w:author=A,w:name=m1}/w:r/w:delText"moved away"'
            ),
        )
        assert mf.text == "moved away"

    def it_is_recognized_as_CT_Del_for_polymorphism(self):
        # -- CT_MoveFrom inherits from CT_Del so _resolve_all_changes's type
        # -- dispatch treats them uniformly --
        mf = cast(CT_MoveFrom, element("w:moveFrom{w:id=1,w:author=A,w:name=m1}"))
        assert isinstance(mf, CT_Del)


class DescribeCT_MoveTo:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_MoveTo`."""

    def it_knows_its_id(self):
        mt = cast(CT_MoveTo, element("w:moveTo{w:id=2,w:author=Bob,w:name=m1}"))
        assert mt.id == 2

    def it_knows_its_author(self):
        mt = cast(CT_MoveTo, element("w:moveTo{w:id=2,w:author=Bob,w:name=m1}"))
        assert mt.author == "Bob"

    def it_knows_its_name(self):
        mt = cast(CT_MoveTo, element("w:moveTo{w:id=2,w:author=Bob,w:name=m1}"))
        assert mt.name == "m1"

    def it_returns_None_when_name_is_absent(self):
        mt = cast(CT_MoveTo, element("w:moveTo{w:id=2,w:author=B}"))
        assert mt.name is None

    def it_can_produce_its_text_from_t_children(self):
        mt = cast(
            CT_MoveTo,
            element('w:moveTo{w:id=2,w:author=B,w:name=m1}/w:r/w:t"moved here"'),
        )
        assert mt.text == "moved here"

    def it_is_recognized_as_CT_Ins_for_polymorphism(self):
        mt = cast(CT_MoveTo, element("w:moveTo{w:id=2,w:author=B,w:name=m1}"))
        assert isinstance(mt, CT_Ins)


class DescribeCT_MoveFrom_acceptReject:
    """Accept/reject behavior for `<w:moveFrom>`."""

    def it_removes_itself_on_accept_completing_the_move(self):
        p = element(
            'w:p/(w:r/w:t"keep",'
            'w:moveFrom{w:id=1,w:author=A,w:name=m1}/w:r/w:delText"gone")'
        )
        mf = p.xpath("./w:moveFrom")[0]
        mf.accept()
        assert p.xpath("./w:moveFrom") == []
        assert [r.text for r in p.xpath("./w:r/w:t")] == ["keep"]

    def it_restores_content_on_reject_converting_delText_to_t(self):
        p = element(
            'w:p/(w:r/w:t"keep ",'
            'w:moveFrom{w:id=1,w:author=A,w:name=m1}/w:r/w:delText"restored")'
        )
        mf = p.xpath("./w:moveFrom")[0]
        mf.reject()
        assert p.xpath("./w:moveFrom") == []
        assert p.xpath("./w:r/w:delText") == []
        assert [t.text for t in p.xpath("./w:r/w:t")] == ["keep ", "restored"]


class DescribeCT_MoveTo_acceptReject:
    """Accept/reject behavior for `<w:moveTo>`."""

    def it_unwraps_itself_on_accept_keeping_content(self):
        p = element(
            'w:p/(w:r/w:t"before ",'
            'w:moveTo{w:id=2,w:author=B,w:name=m1}/w:r/w:t"moved")'
        )
        mt = p.xpath("./w:moveTo")[0]
        mt.accept()
        assert p.xpath("./w:moveTo") == []
        assert [r.text for r in p.xpath("./w:r/w:t")] == ["before ", "moved"]

    def it_removes_itself_on_reject_cancelling_the_move(self):
        p = element(
            'w:p/(w:r/w:t"before ",'
            'w:moveTo{w:id=2,w:author=B,w:name=m1}/w:r/w:t"moved")'
        )
        mt = p.xpath("./w:moveTo")[0]
        mt.reject()
        assert p.xpath("./w:moveTo") == []
        assert [r.text for r in p.xpath("./w:r/w:t")] == ["before "]


class DescribeCT_CellIns:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_CellIns`."""

    def it_knows_its_id(self):
        ci = cast(CT_CellIns, element("w:cellIns{w:id=3,w:author=Alice}"))
        assert ci.id == 3

    def it_knows_its_author(self):
        ci = cast(CT_CellIns, element("w:cellIns{w:id=3,w:author=Alice}"))
        assert ci.author == "Alice"

    def it_knows_its_date(self):
        ci = cast(
            CT_CellIns,
            element("w:cellIns{w:id=3,w:author=Alice,w:date=2024-01-02T03:04:05Z}"),
        )
        assert ci.date == dt.datetime(2024, 1, 2, 3, 4, 5, tzinfo=dt.timezone.utc)

    def it_returns_None_when_date_is_absent(self):
        ci = cast(CT_CellIns, element("w:cellIns{w:id=3,w:author=Alice}"))
        assert ci.date is None


class DescribeCT_CellDel:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_CellDel`."""

    def it_knows_its_id(self):
        cd = cast(CT_CellDel, element("w:cellDel{w:id=4,w:author=Bob}"))
        assert cd.id == 4

    def it_knows_its_author(self):
        cd = cast(CT_CellDel, element("w:cellDel{w:id=4,w:author=Bob}"))
        assert cd.author == "Bob"

    def it_knows_its_date(self):
        cd = cast(
            CT_CellDel,
            element("w:cellDel{w:id=4,w:author=Bob,w:date=2024-06-07T08:09:10Z}"),
        )
        assert cd.date == dt.datetime(2024, 6, 7, 8, 9, 10, tzinfo=dt.timezone.utc)


class DescribeCT_TcPrChange:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_TcPrChange`."""

    def it_knows_its_id(self):
        tpc = cast(CT_TcPrChange, element("w:tcPrChange{w:id=5,w:author=Carol}"))
        assert tpc.id == 5

    def it_knows_its_author(self):
        tpc = cast(CT_TcPrChange, element("w:tcPrChange{w:id=5,w:author=Carol}"))
        assert tpc.author == "Carol"

    def it_exposes_its_inner_tcPr(self):
        tpc = cast(
            CT_TcPrChange,
            element(
                "w:tcPrChange{w:id=5,w:author=C}/w:tcPr/w:vAlign{w:val=center}"
            ),
        )
        assert tpc.tcPr is not None
        assert tpc.tcPr.xpath("./w:vAlign")

    def it_returns_None_for_tcPr_when_absent(self):
        tpc = cast(CT_TcPrChange, element("w:tcPrChange{w:id=5,w:author=C}"))
        assert tpc.tcPr is None


class DescribeCT_TrPrChange:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_TrPrChange`."""

    def it_knows_its_id(self):
        rpc = cast(CT_TrPrChange, element("w:trPrChange{w:id=6,w:author=Dave}"))
        assert rpc.id == 6

    def it_exposes_its_inner_trPr(self):
        rpc = cast(
            CT_TrPrChange,
            element("w:trPrChange{w:id=6,w:author=D}/w:trPr/w:cantSplit"),
        )
        assert rpc.trPr is not None
        assert rpc.trPr.xpath("./w:cantSplit")


class DescribeCT_TblPrChange:
    """Unit-test suite for `docx.oxml.tracked_changes.CT_TblPrChange`."""

    def it_knows_its_id(self):
        tpc = cast(CT_TblPrChange, element("w:tblPrChange{w:id=7,w:author=Eve}"))
        assert tpc.id == 7

    def it_exposes_its_inner_tblPr(self):
        tpc = cast(
            CT_TblPrChange,
            element(
                "w:tblPrChange{w:id=7,w:author=E}/w:tblPr/w:tblW{w:w=5000,w:type=dxa}"
            ),
        )
        assert tpc.tblPr is not None
        assert tpc.tblPr.xpath("./w:tblW")
