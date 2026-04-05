# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.tracked_changes` module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.oxml.tracked_changes import CT_Del, CT_DelText, CT_Ins

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
