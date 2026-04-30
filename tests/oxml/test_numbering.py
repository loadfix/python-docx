# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.oxml.numbering` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.numbering import (
    CT_AbstractNum,
    CT_Lvl,
    CT_Num,
    CT_Numbering,
    CT_NumPr,
)

from ..unitutil.cxml import element


class DescribeCT_Numbering:
    """Unit-test suite for `docx.oxml.numbering.CT_Numbering`."""

    def it_can_add_an_abstractNum_to_an_empty_numbering(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        abstractNum = numbering.add_abstractNum()

        assert isinstance(abstractNum, CT_AbstractNum)
        assert abstractNum.abstractNumId == 0

    def it_assigns_next_abstractNumId_for_consecutive_calls(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        a = numbering.add_abstractNum()
        b = numbering.add_abstractNum()
        c = numbering.add_abstractNum()

        assert [a.abstractNumId, b.abstractNumId, c.abstractNumId] == [0, 1, 2]

    def it_can_add_an_abstractNum_with_an_explicit_id(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        numbering.add_abstractNum(7)

        abstractNum = numbering.add_abstractNum()

        assert abstractNum.abstractNumId == 0

    def it_can_add_a_num_with_an_explicit_id(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        num = numbering.add_num(abstractNum_id=0, num_id=5)

        assert isinstance(num, CT_Num)
        assert num.numId == 5

    def it_finds_an_abstractNum_by_id(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        a = numbering.add_abstractNum()
        b = numbering.add_abstractNum()

        assert numbering.abstractNum_having_abstractNumId(a.abstractNumId) is a
        assert numbering.abstractNum_having_abstractNumId(b.abstractNumId) is b


class DescribeCT_AbstractNum:
    """Unit-test suite for `docx.oxml.numbering.CT_AbstractNum`."""

    def it_can_add_a_level(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering.add_abstractNum()

        lvl = abstractNum.add_lvl()
        lvl.ilvl = 2

        assert isinstance(lvl, CT_Lvl)
        assert lvl.ilvl == 2

    def it_can_retrieve_a_level_by_ilvl(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering.add_abstractNum()
        l0 = abstractNum.add_lvl()
        l0.ilvl = 0
        l1 = abstractNum.add_lvl()
        l1.ilvl = 1

        assert abstractNum.get_lvl(0) is l0
        assert abstractNum.get_lvl(1) is l1
        assert abstractNum.get_lvl(5) is None


class DescribeCT_Lvl:
    """Unit-test suite for `docx.oxml.numbering.CT_Lvl`."""

    def it_round_trips_start_numFmt_and_lvlText_values(self):
        from docx.enum.text import WD_NUMBER_FORMAT

        numbering = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering.add_abstractNum()
        lvl = abstractNum.add_lvl()
        lvl.ilvl = 0

        lvl.start_val = 3
        lvl.numFmt_val = WD_NUMBER_FORMAT.UPPER_ROMAN
        lvl.lvlText_val = "%1)"

        assert lvl.start_val == 3
        assert lvl.numFmt_val == WD_NUMBER_FORMAT.UPPER_ROMAN
        assert lvl.lvlText_val == "%1)"

    def it_defaults_start_to_1_when_no_start_child(self):
        numbering = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering.add_abstractNum()
        lvl = abstractNum.add_lvl()
        lvl.ilvl = 0

        assert lvl.start_val == 1


class DescribeCT_NumPr:
    """Unit-test suite for `docx.oxml.numbering.CT_NumPr`."""

    def it_exposes_ilvl_val_and_numId_val(self):
        numPr = cast(
            CT_NumPr,
            element(
                "w:numPr/(w:ilvl{w:val=2},w:numId{w:val=7})"
            ),
        )

        assert numPr.ilvl_val == 2
        assert numPr.numId_val == 7

    def it_accepts_writes_for_ilvl_and_numId(self):
        numPr = cast(CT_NumPr, element("w:numPr"))

        numPr.ilvl_val = 3
        numPr.numId_val = 4

        assert numPr.ilvl_val == 3
        assert numPr.numId_val == 4

    def it_can_clear_ilvl_and_numId_by_assigning_None(self):
        numPr = cast(
            CT_NumPr,
            element(
                "w:numPr/(w:ilvl{w:val=1},w:numId{w:val=1})"
            ),
        )

        numPr.ilvl_val = None
        numPr.numId_val = None

        assert numPr.ilvl_val is None
        assert numPr.numId_val is None
