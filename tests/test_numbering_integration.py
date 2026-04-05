# pyright: reportPrivateUsage=false

"""Unit test suite for numbering/list features on Paragraph."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.numbering import CT_Numbering, CT_NumPr
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.parfmt import CT_PPr
from docx.parts.numbering import NumberingPart
from docx.text.paragraph import Paragraph, _ListFormat

from .unitutil.cxml import element
from .unitutil.mock import Mock, PropertyMock


class DescribeParagraph_list_level:
    """Unit-test suite for `Paragraph.list_level` property."""

    def it_returns_None_when_no_numPr(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _fake_parent())
        assert paragraph.list_level is None

    def it_returns_None_when_pPr_but_no_numPr(self):
        p = cast(CT_P, element("w:p/w:pPr"))
        paragraph = Paragraph(p, _fake_parent())
        assert paragraph.list_level is None

    def it_returns_the_ilvl_value(self):
        p = cast(CT_P, element("w:p/w:pPr/w:numPr/w:ilvl{w:val=2}"))
        paragraph = Paragraph(p, _fake_parent())
        assert paragraph.list_level == 2

    def it_can_set_the_list_level(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _fake_parent())

        paragraph.list_level = 3

        assert p.pPr is not None
        assert p.pPr.numPr is not None
        assert p.pPr.numPr.ilvl_val == 3

    def it_can_set_level_to_None(self):
        p = cast(CT_P, element("w:p/w:pPr/w:numPr/w:ilvl{w:val=1}"))
        paragraph = Paragraph(p, _fake_parent())

        paragraph.list_level = None

        assert p.pPr.numPr.ilvl_val is None


class DescribeParagraph_list_format:
    """Unit-test suite for `Paragraph.list_format` property."""

    def it_returns_a_ListFormat_object(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _fake_parent())

        list_format = paragraph.list_format

        assert isinstance(list_format, _ListFormat)

    def it_can_get_and_set_level(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _fake_parent())
        lf = paragraph.list_format

        assert lf.level is None

        lf.level = 2
        assert lf.level == 2

    def it_can_get_and_set_num_id(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _fake_parent())
        lf = paragraph.list_format

        assert lf.num_id is None

        lf.num_id = 5
        assert lf.num_id == 5


class DescribeParagraph_numbering_format:
    """Unit-test suite for `Paragraph.numbering_format` property."""

    def it_returns_None_when_no_numPr(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _fake_parent())
        assert paragraph.numbering_format is None

    def it_returns_None_when_numId_is_zero(self):
        p = cast(CT_P, element("w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=0})"))
        paragraph = Paragraph(p, _fake_parent())
        assert paragraph.numbering_format is None

    def it_returns_the_format_when_numbering_exists(self):
        # -- set up numbering part with abstract num --
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering_elm.add_abstractNum()
        lvl = abstractNum.add_lvl(0)
        lvl.numFmt_val = "decimal"
        num = numbering_elm.add_num(0)

        numbering_part = Mock(spec=NumberingPart)
        numbering_part.numbering_element = numbering_elm

        # -- set up paragraph with numPr pointing to our num --
        p = cast(
            CT_P,
            element(
                "w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=%d})" % num.numId
            ),
        )

        part_mock = Mock()
        part_mock.numbering_part = numbering_part
        parent = Mock()
        type(parent).part = PropertyMock(return_value=part_mock)

        paragraph = Paragraph(p, parent)

        assert paragraph.numbering_format == "decimal"


class DescribeParagraph_restart_numbering:
    """Unit-test suite for `Paragraph.restart_numbering()` method."""

    def it_raises_when_paragraph_not_in_list(self):
        p = cast(CT_P, element("w:p"))
        paragraph = Paragraph(p, _fake_parent())

        with pytest.raises(ValueError, match="not in a numbered list"):
            paragraph.restart_numbering()

    def it_raises_when_numId_is_zero(self):
        p = cast(CT_P, element("w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=0})"))
        paragraph = Paragraph(p, _fake_parent())

        with pytest.raises(ValueError, match="not in a numbered list"):
            paragraph.restart_numbering()

    def it_creates_a_new_num_with_restart_override(self):
        # -- set up numbering part --
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering_elm.add_abstractNum()
        lvl = abstractNum.add_lvl(0)
        lvl.numFmt_val = "decimal"
        num = numbering_elm.add_num(0)
        orig_num_id = num.numId

        numbering_part = Mock(spec=NumberingPart)
        numbering_part.numbering_element = numbering_elm

        # -- set up paragraph --
        p = cast(
            CT_P,
            element(
                "w:p/w:pPr/w:numPr/(w:ilvl{w:val=0},w:numId{w:val=%d})" % orig_num_id
            ),
        )

        part_mock = Mock()
        part_mock.numbering_part = numbering_part
        parent = Mock()
        type(parent).part = PropertyMock(return_value=part_mock)

        paragraph = Paragraph(p, parent)

        paragraph.restart_numbering()

        # -- paragraph now points to a different numId --
        new_num_id = p.pPr.numPr.numId_val
        assert new_num_id != orig_num_id

        # -- the new num has a lvlOverride with startOverride --
        new_num = numbering_elm.num_having_numId(new_num_id)
        assert len(new_num.lvlOverride_lst) == 1
        lvl_override = new_num.lvlOverride_lst[0]
        assert lvl_override.ilvl == 0
        assert lvl_override.startOverride.val == 1


class DescribeCT_PPr_numPr_properties:
    """Unit-test suite for numPr convenience properties on CT_PPr."""

    def it_can_get_and_set_numPr_ilvl_val(self):
        pPr = cast(CT_PPr, element("w:pPr"))

        assert pPr.numPr_ilvl_val is None

        pPr.numPr_ilvl_val = 3
        assert pPr.numPr_ilvl_val == 3

        pPr.numPr_ilvl_val = None
        assert pPr.numPr_ilvl_val is None

    def it_can_get_and_set_numPr_numId_val(self):
        pPr = cast(CT_PPr, element("w:pPr"))

        assert pPr.numPr_numId_val is None

        pPr.numPr_numId_val = 7
        assert pPr.numPr_numId_val == 7

        # Setting to None should remove numPr when both children are gone
        pPr.numPr_numId_val = None
        assert pPr.numPr_numId_val is None
        assert pPr.numPr is None


def _fake_parent():
    """Return a mock that satisfies the ProvidesStoryPart interface."""
    parent = Mock()
    type(parent).part = PropertyMock(return_value=Mock())
    return parent
