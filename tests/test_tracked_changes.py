# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.tracked_changes` module."""

from __future__ import annotations

import datetime as dt
from typing import cast

import pytest

from docx.oxml.tracked_changes import CT_Del, CT_Ins
from docx.tracked_changes import TrackedChange

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
