# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.oxml.numbering` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.numbering import (
    CT_AbstractNum,
    CT_Lvl,
    CT_Num,
    CT_Numbering,
    CT_NumLvl,
    CT_NumPr,
)

from ..unitutil.cxml import element


class DescribeCT_AbstractNum:
    def it_can_construct_a_new_element(self):
        abstract_num = CT_AbstractNum.new(42)
        assert abstract_num.abstractNumId == 42

    def it_can_add_a_level(self):
        abstract_num = CT_AbstractNum.new(0)
        lvl = abstract_num.add_lvl(0)
        assert lvl.ilvl == 0
        assert len(abstract_num.lvl_lst) == 1

    def it_can_find_a_lvl_for_a_given_ilvl(self):
        abstract_num = CT_AbstractNum.new(0)
        lvl0 = abstract_num.add_lvl(0)
        lvl1 = abstract_num.add_lvl(1)
        assert abstract_num.lvl_for_ilvl(0) is lvl0
        assert abstract_num.lvl_for_ilvl(1) is lvl1
        assert abstract_num.lvl_for_ilvl(2) is None


class DescribeCT_Lvl:
    def it_can_get_and_set_start_val(self):
        abstract_num = CT_AbstractNum.new(0)
        lvl = abstract_num.add_lvl(0)
        assert lvl.start_val == 1
        lvl.start_val = 5
        assert lvl.start_val == 5

    def it_can_get_and_set_numFmt_val(self):
        abstract_num = CT_AbstractNum.new(0)
        lvl = abstract_num.add_lvl(0)
        assert lvl.numFmt_val is None
        lvl.numFmt_val = "decimal"
        assert lvl.numFmt_val == "decimal"
        lvl.numFmt_val = None
        assert lvl.numFmt_val is None

    def it_can_get_and_set_lvlText_val(self):
        abstract_num = CT_AbstractNum.new(0)
        lvl = abstract_num.add_lvl(0)
        assert lvl.lvlText_val is None
        lvl.lvlText_val = "%1."
        assert lvl.lvlText_val == "%1."
        lvl.lvlText_val = None
        assert lvl.lvlText_val is None


class DescribeCT_NumPr:
    def it_can_get_and_set_ilvl_val(self):
        numPr = cast(CT_NumPr, element("w:numPr"))
        assert numPr.ilvl_val is None
        numPr.ilvl_val = 2
        assert numPr.ilvl_val == 2
        numPr.ilvl_val = None
        assert numPr.ilvl_val is None

    def it_can_get_and_set_numId_val(self):
        numPr = cast(CT_NumPr, element("w:numPr"))
        assert numPr.numId_val is None
        numPr.numId_val = 5
        assert numPr.numId_val == 5
        numPr.numId_val = None
        assert numPr.numId_val is None


class DescribeCT_Numbering:
    def it_can_add_an_abstractNum(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        abstract_num = CT_AbstractNum.new(0)
        numbering.add_abstractNum(abstract_num)
        assert len(numbering.abstractNum_lst) == 1

    def it_can_find_an_abstractNum_by_id(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        abstract_num = CT_AbstractNum.new(42)
        numbering.add_abstractNum(abstract_num)
        found = numbering.abstractNum_having_abstractNumId(42)
        assert found is abstract_num

    def it_raises_on_missing_abstractNum(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        with pytest.raises(KeyError):
            numbering.abstractNum_having_abstractNumId(99)

    def it_computes_the_next_abstractNumId(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        assert numbering._next_abstractNumId == 0
        numbering.add_abstractNum(CT_AbstractNum.new(0))
        assert numbering._next_abstractNumId == 1

    def it_can_add_a_num(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        num = numbering.add_num(0)
        assert num.numId == 1
        assert num.abstractNumId_val == 0

    def it_computes_the_next_numId(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        assert numbering._next_numId == 1
        numbering.add_num(0)
        assert numbering._next_numId == 2


class DescribeCT_Num:
    def it_can_construct_a_new_element(self):
        num = CT_Num.new(1, 0)
        assert num.numId == 1
        assert num.abstractNumId_val == 0

    def it_can_add_a_lvlOverride(self):
        num = CT_Num.new(1, 0)
        lvl_override = num.add_lvlOverride(ilvl=0)
        assert lvl_override.ilvl == 0


class DescribeCT_NumLvl:
    def it_can_add_a_startOverride(self):
        num = CT_Num.new(1, 0)
        lvl_override = num.add_lvlOverride(ilvl=0)
        start_override = lvl_override.add_startOverride(val=1)
        assert start_override.val == 1
