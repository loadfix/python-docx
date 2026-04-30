# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.tracked_changes` module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.oxml.tracked_changes import (
    CT_Del,
    CT_Ins,
    CT_PPrChange,
    CT_RPrChange,
    CT_SectPrChange,
)
from docx.tracked_changes import FormattingChange, TrackedChange

from .unitutil.cxml import element


class DescribeTrackedChange:
    """Unit-test suite for `docx.tracked_changes.TrackedChange`."""

    def it_reports_insertion_type_for_w_ins(self):
        ins = cast(CT_Ins, element("w:ins{w:id=1,w:author=Alice}"))
        tc = TrackedChange(ins)
        assert tc.type == "insertion"

    def it_reports_deletion_type_for_w_del(self):
        del_elm = cast(CT_Del, element("w:del{w:id=2,w:author=Bob}"))
        tc = TrackedChange(del_elm)
        assert tc.type == "deletion"

    def it_knows_its_author(self):
        ins = cast(CT_Ins, element("w:ins{w:id=1,w:author=Alice}"))
        tc = TrackedChange(ins)
        assert tc.author == "Alice"

    def it_knows_its_date(self):
        ins = cast(
            CT_Ins, element("w:ins{w:id=1,w:author=Alice,w:date=2023-10-01T12:00:00Z}")
        )
        tc = TrackedChange(ins)
        assert tc.date == dt.datetime(2023, 10, 1, 12, 0, 0, tzinfo=dt.timezone.utc)

    def it_returns_None_for_date_when_absent(self):
        ins = cast(CT_Ins, element("w:ins{w:id=1,w:author=Alice}"))
        tc = TrackedChange(ins)
        assert tc.date is None

    def it_knows_its_text_for_an_insertion(self):
        ins = cast(CT_Ins, element('w:ins{w:id=1,w:author=A}/w:r/w:t"inserted text"'))
        tc = TrackedChange(ins)
        assert tc.text == "inserted text"

    def it_knows_its_text_for_a_deletion(self):
        del_elm = cast(CT_Del, element('w:del{w:id=2,w:author=B}/w:r/w:delText"deleted text"'))
        tc = TrackedChange(del_elm)
        assert tc.text == "deleted text"

    def it_delegates_accept_to_the_underlying_element(self):
        p = element(
            'w:p/(w:ins{w:id=1,w:author=A}/w:r/w:t"i",w:del{w:id=2,w:author=B}/w:r/w:delText"d")'
        )
        ins = cast(CT_Ins, p.xpath("./w:ins")[0])
        del_ = cast(CT_Del, p.xpath("./w:del")[0])
        TrackedChange(ins).accept()
        TrackedChange(del_).accept()
        assert p.xpath("./w:ins") == []
        assert p.xpath("./w:del") == []
        # inserted run survived, deleted run discarded
        assert [t.text for t in p.xpath("./w:r/w:t")] == ["i"]

    def it_delegates_reject_to_the_underlying_element(self):
        p = element(
            'w:p/(w:ins{w:id=1,w:author=A}/w:r/w:t"i",w:del{w:id=2,w:author=B}/w:r/w:delText"d")'
        )
        ins = cast(CT_Ins, p.xpath("./w:ins")[0])
        del_ = cast(CT_Del, p.xpath("./w:del")[0])
        TrackedChange(ins).reject()
        TrackedChange(del_).reject()
        assert p.xpath("./w:ins") == []
        assert p.xpath("./w:del") == []
        # inserted run discarded, deleted run restored (now with w:t)
        assert [t.text for t in p.xpath("./w:r/w:t")] == ["d"]


class DescribeFormattingChange:
    """Unit-test suite for `docx.tracked_changes.FormattingChange`."""

    def it_knows_its_author(self):
        rPrChange = cast(CT_RPrChange, element("w:rPrChange{w:id=1,w:author=Alice}"))
        fc = FormattingChange(rPrChange)
        assert fc.author == "Alice"

    def it_knows_its_date(self):
        rPrChange = cast(
            CT_RPrChange,
            element("w:rPrChange{w:id=1,w:author=Alice,w:date=2024-05-20T14:30:00Z}"),
        )
        fc = FormattingChange(rPrChange)
        assert fc.date == dt.datetime(2024, 5, 20, 14, 30, 0, tzinfo=dt.timezone.utc)

    def it_returns_None_for_date_when_absent(self):
        rPrChange = cast(CT_RPrChange, element("w:rPrChange{w:id=1,w:author=A}"))
        fc = FormattingChange(rPrChange)
        assert fc.date is None

    def it_exposes_old_rPr_for_rPrChange(self):
        rPrChange = cast(
            CT_RPrChange,
            element("w:rPrChange{w:id=1,w:author=A}/w:rPr/w:b"),
        )
        fc = FormattingChange(rPrChange)
        assert fc.old_properties is not None
        assert fc.old_properties.xpath("./w:b")

    def it_exposes_old_pPr_for_pPrChange(self):
        pPrChange = cast(
            CT_PPrChange,
            element("w:pPrChange{w:id=2,w:author=B}/w:pPr/w:jc{w:val=center}"),
        )
        fc = FormattingChange(pPrChange)
        assert fc.old_properties is not None
        assert fc.old_properties.xpath("./w:jc")

    def it_exposes_old_sectPr_for_sectPrChange(self):
        sectPrChange = cast(
            CT_SectPrChange,
            element("w:sectPrChange{w:id=3,w:author=C}/w:sectPr/w:pgSz{w:w=12240,w:h=15840}"),
        )
        fc = FormattingChange(sectPrChange)
        assert fc.old_properties is not None
        assert fc.old_properties.xpath("./w:pgSz")

    def it_returns_None_for_old_properties_when_inner_element_missing(self):
        rPrChange = cast(CT_RPrChange, element("w:rPrChange{w:id=1,w:author=A}"))
        fc = FormattingChange(rPrChange)
        assert fc.old_properties is None
