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


class DescribeCT_Numbering:
    """Unit-test suite for `docx.oxml.numbering.CT_Numbering` objects."""

    def it_can_add_an_abstractNum(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        abstractNum = numbering.add_abstractNum()

        assert isinstance(abstractNum, CT_AbstractNum)
        assert abstractNum.abstractNumId == 0
        assert len(numbering.abstractNum_lst) == 1

    def it_can_add_multiple_abstractNums(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        a0 = numbering.add_abstractNum()
        a1 = numbering.add_abstractNum()

        assert a0.abstractNumId == 0
        assert a1.abstractNumId == 1
        assert len(numbering.abstractNum_lst) == 2

    def it_can_add_a_num(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        num = numbering.add_num(0)

        assert isinstance(num, CT_Num)
        assert num.numId == 1
        assert num.abstractNumId_val == 0

    def it_can_find_a_num_by_numId(self):
        numbering = cast(
            CT_Numbering,
            element("w:numbering/(w:num{w:numId=1}/(w:abstractNumId{w:val=0}))"),
        )

        num = numbering.num_having_numId(1)

        assert num.numId == 1

    def it_raises_on_missing_numId(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        with pytest.raises(KeyError):
            numbering.num_having_numId(99)

    def it_computes_next_abstractNumId(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        assert numbering._next_abstractNumId == 0
        numbering.add_abstractNum()
        assert numbering._next_abstractNumId == 1

    def it_inserts_abstractNum_before_num(self):
        numbering = cast(
            CT_Numbering,
            element("w:numbering/(w:num{w:numId=1}/(w:abstractNumId{w:val=0}))"),
        )

        abstractNum = numbering.add_abstractNum()

        # abstractNum should be before num in document order
        children = list(numbering)
        abstractNum_idx = children.index(abstractNum)
        num_idx = children.index(numbering.num_lst[0])
        assert abstractNum_idx < num_idx


class DescribeCT_AbstractNum:
    """Unit-test suite for `docx.oxml.numbering.CT_AbstractNum` objects."""

    def it_can_be_constructed(self):
        abstractNum = CT_AbstractNum.new(0)

        assert isinstance(abstractNum, CT_AbstractNum)
        assert abstractNum.abstractNumId == 0

    def it_can_add_a_level(self):
        abstractNum = CT_AbstractNum.new(0)

        lvl = abstractNum.add_lvl(0)

        assert isinstance(lvl, CT_Lvl)
        assert lvl.ilvl == 0
        assert len(abstractNum.lvl_lst) == 1

    def it_can_add_multiple_levels(self):
        abstractNum = CT_AbstractNum.new(0)

        abstractNum.add_lvl(0)
        abstractNum.add_lvl(1)
        abstractNum.add_lvl(2)

        assert len(abstractNum.lvl_lst) == 3


class DescribeCT_Lvl:
    """Unit-test suite for `docx.oxml.numbering.CT_Lvl` objects."""

    def it_can_get_and_set_numFmt_val(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering.add_abstractNum()
        lvl = abstractNum.add_lvl(0)

        assert lvl.numFmt_val is None

        lvl.numFmt_val = "decimal"
        assert lvl.numFmt_val == "decimal"

        lvl.numFmt_val = None
        assert lvl.numFmt_val is None

    def it_can_get_and_set_lvlText_val(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering.add_abstractNum()
        lvl = abstractNum.add_lvl(0)

        assert lvl.lvlText_val is None

        lvl.lvlText_val = "%1."
        assert lvl.lvlText_val == "%1."

        lvl.lvlText_val = None
        assert lvl.lvlText_val is None

    def it_can_get_and_set_start_val(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering.add_abstractNum()
        lvl = abstractNum.add_lvl(0)

        assert lvl.start_val is None

        lvl.start_val = 1
        assert lvl.start_val == 1

        lvl.start_val = None
        assert lvl.start_val is None


class DescribeCT_NumPr:
    """Unit-test suite for `docx.oxml.numbering.CT_NumPr` objects."""

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


class DescribeCT_Num:
    """Unit-test suite for `docx.oxml.numbering.CT_Num` objects."""

    def it_provides_access_to_abstractNumId_val(self):
        numbering = cast(
            CT_Numbering,
            element("w:numbering/(w:num{w:numId=1}/(w:abstractNumId{w:val=42}))"),
        )
        num = numbering.num_lst[0]

        assert num.abstractNumId_val == 42

    def it_can_add_a_lvlOverride(self):
        numbering = cast(
            CT_Numbering,
            element("w:numbering/(w:num{w:numId=1}/(w:abstractNumId{w:val=0}))"),
        )
        num = numbering.num_lst[0]

        lvlOverride = num.add_lvlOverride(0)

        assert isinstance(lvlOverride, CT_NumLvl)
        assert lvlOverride.ilvl == 0
