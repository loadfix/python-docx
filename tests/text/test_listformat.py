# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.text.listformat` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.numbering import CT_NumPr
from docx.oxml.text.paragraph import CT_P
from docx.text.listformat import ListFormat

from ..unitutil.cxml import element
from ..unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeListFormat:
    def it_can_get_the_level_when_no_numPr(self):
        p = cast(CT_P, element("w:p"))
        list_format = ListFormat(p, Mock())
        assert list_format.level is None

    def it_can_get_the_level_with_numPr(self):
        p = cast(CT_P, element("w:p/w:pPr/w:numPr/w:ilvl{w:val=2}"))
        list_format = ListFormat(p, Mock())
        assert list_format.level == 2

    def it_can_set_the_level(self):
        p = cast(CT_P, element("w:p"))
        list_format = ListFormat(p, Mock())
        list_format.level = 3
        assert list_format.level == 3

    def it_can_clear_the_level(self):
        p = cast(CT_P, element("w:p/w:pPr/w:numPr/w:ilvl{w:val=2}"))
        list_format = ListFormat(p, Mock())
        list_format.level = None
        assert list_format.level is None

    def it_can_get_the_num_id_when_no_numPr(self):
        p = cast(CT_P, element("w:p"))
        list_format = ListFormat(p, Mock())
        assert list_format.num_id is None

    def it_can_get_the_num_id_with_numPr(self):
        p = cast(CT_P, element("w:p/w:pPr/w:numPr/w:numId{w:val=5}"))
        list_format = ListFormat(p, Mock())
        assert list_format.num_id == 5

    def it_can_set_the_num_id(self):
        p = cast(CT_P, element("w:p"))
        list_format = ListFormat(p, Mock())
        list_format.num_id = 7
        assert list_format.num_id == 7

    def it_can_clear_the_num_id(self):
        p = cast(CT_P, element("w:p/w:pPr/w:numPr/w:numId{w:val=5}"))
        list_format = ListFormat(p, Mock())
        list_format.num_id = None
        assert list_format.num_id is None

    def it_can_apply_numbering(self):
        p = cast(CT_P, element("w:p"))
        list_format = ListFormat(p, Mock())
        list_format.apply(num_id=3, level=1)
        assert list_format.num_id == 3
        assert list_format.level == 1

    def it_can_apply_numbering_with_default_level(self):
        p = cast(CT_P, element("w:p"))
        list_format = ListFormat(p, Mock())
        list_format.apply(num_id=3)
        assert list_format.num_id == 3
        assert list_format.level == 0
