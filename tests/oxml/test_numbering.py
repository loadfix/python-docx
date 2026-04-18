# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.numbering` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.numbering import CT_Num, CT_NumLvl, CT_NumPr, CT_Numbering

from ..unitutil.cxml import element


class DescribeCT_Num:
    """Unit-test suite for `docx.oxml.numbering.CT_Num`."""

    def it_provides_access_to_its_numId(self):
        num = cast(CT_Num, element("w:num{w:numId=42}"))
        assert num.numId == 42

    def it_can_add_a_lvlOverride(self):
        num = cast(CT_Num, element("w:num{w:numId=1}/w:abstractNumId{w:val=0}"))
        lvl_override = num.add_lvlOverride(ilvl=3)
        assert lvl_override.ilvl == 3
        assert len(num.lvlOverride_lst) == 1

    def it_can_construct_a_new_num_element(self):
        num = CT_Num.new(num_id=7, abstractNum_id=3)
        assert num.numId == 7
        assert num.abstractNumId.val == 3


class DescribeCT_NumLvl:
    """Unit-test suite for `docx.oxml.numbering.CT_NumLvl`."""

    def it_provides_access_to_its_ilvl(self):
        lvl_override = cast(CT_NumLvl, element("w:lvlOverride{w:ilvl=2}"))
        assert lvl_override.ilvl == 2

    def it_can_add_a_startOverride(self):
        lvl_override = cast(CT_NumLvl, element("w:lvlOverride{w:ilvl=0}"))
        start_override = lvl_override.add_startOverride(val=5)
        assert start_override.val == 5


class DescribeCT_NumPr:
    """Unit-test suite for `docx.oxml.numbering.CT_NumPr`."""

    def it_provides_access_to_its_ilvl_child(self):
        numPr = cast(CT_NumPr, element("w:numPr/w:ilvl{w:val=2}"))
        assert numPr.ilvl is not None
        assert numPr.ilvl.val == 2

    def it_returns_None_when_ilvl_is_absent(self):
        numPr = cast(CT_NumPr, element("w:numPr"))
        assert numPr.ilvl is None

    def it_provides_access_to_its_numId_child(self):
        numPr = cast(CT_NumPr, element("w:numPr/w:numId{w:val=7}"))
        assert numPr.numId is not None
        assert numPr.numId.val == 7

    def it_returns_None_when_numId_is_absent(self):
        numPr = cast(CT_NumPr, element("w:numPr"))
        assert numPr.numId is None


class DescribeCT_Numbering:
    """Unit-test suite for `docx.oxml.numbering.CT_Numbering`."""

    def it_provides_access_to_its_num_children(self):
        numbering = cast(
            CT_Numbering,
            element("w:numbering/(w:num{w:numId=1},w:num{w:numId=2})"),
        )
        assert len(numbering.num_lst) == 2

    def it_can_add_a_num_element(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        num = numbering.add_num(abstractNum_id=0)
        assert num.numId == 1
        assert num.abstractNumId.val == 0

    def it_assigns_sequential_numIds(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        num1 = numbering.add_num(abstractNum_id=0)
        num2 = numbering.add_num(abstractNum_id=1)
        assert num1.numId == 1
        assert num2.numId == 2

    def it_fills_gaps_in_numId_sequence(self):
        numbering = cast(
            CT_Numbering,
            element(
                "w:numbering/(w:num{w:numId=1}/w:abstractNumId{w:val=0}"
                ",w:num{w:numId=3}/w:abstractNumId{w:val=1})"
            ),
        )
        num = numbering.add_num(abstractNum_id=2)
        # should fill the gap at numId=2
        assert num.numId == 2

    def it_can_find_a_num_by_numId(self):
        numbering = cast(
            CT_Numbering,
            element(
                "w:numbering/(w:num{w:numId=1}/w:abstractNumId{w:val=0}"
                ",w:num{w:numId=2}/w:abstractNumId{w:val=1})"
            ),
        )
        num = numbering.num_having_numId(2)
        assert num.numId == 2

    def it_raises_on_num_not_found(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        with pytest.raises(KeyError, match="no <w:num> element with numId 99"):
            numbering.num_having_numId(99)
