# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.numbering` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.numbering import LevelFormat, Numbering, NumberingDefinition
from docx.oxml.numbering import CT_AbstractNum, CT_Lvl, CT_Numbering
from docx.parts.numbering import NumberingPart

from .unitutil.cxml import element
from .unitutil.mock import Mock


class DescribeNumbering:
    """Unit-test suite for `docx.numbering.Numbering` objects."""

    def it_can_list_definitions(self):
        numbering_elm = cast(
            CT_Numbering,
            element("w:numbering/(w:num{w:numId=1}/(w:abstractNumId{w:val=0}))"),
        )
        numbering_part = Mock(spec=NumberingPart)
        numbering = Numbering(numbering_elm, numbering_part)

        definitions = numbering.definitions

        assert len(definitions) == 1
        assert isinstance(definitions[0], NumberingDefinition)
        assert definitions[0].num_id == 1

    def it_can_list_multiple_definitions(self):
        numbering_elm = cast(
            CT_Numbering,
            element(
                "w:numbering/(w:num{w:numId=1}/(w:abstractNumId{w:val=0})"
                ",w:num{w:numId=2}/(w:abstractNumId{w:val=1}))"
            ),
        )
        numbering_part = Mock(spec=NumberingPart)
        numbering = Numbering(numbering_elm, numbering_part)

        definitions = numbering.definitions

        assert len(definitions) == 2
        assert definitions[0].num_id == 1
        assert definitions[1].num_id == 2

    def it_returns_empty_list_when_no_definitions(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering_part = Mock(spec=NumberingPart)
        numbering = Numbering(numbering_elm, numbering_part)

        assert numbering.definitions == []

    def it_can_add_a_simple_numbering_definition(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering_part = Mock(spec=NumberingPart)
        numbering = Numbering(numbering_elm, numbering_part)

        defn = numbering.add_numbering_definition()

        assert isinstance(defn, NumberingDefinition)
        assert defn.num_id == 1
        assert len(numbering_elm.abstractNum_lst) == 1
        assert len(numbering_elm.num_lst) == 1

    def it_can_add_a_multi_level_definition(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering_part = Mock(spec=NumberingPart)
        numbering = Numbering(numbering_elm, numbering_part)

        levels = [
            {"number_format": "decimal", "text": "%1.", "start": 1},
            {"number_format": "lowerLetter", "text": "%2)", "start": 1},
            {"number_format": "lowerRoman", "text": "%3.", "start": 1},
        ]
        defn = numbering.add_numbering_definition(levels)

        assert isinstance(defn, NumberingDefinition)
        level_fmts = defn.level_formats
        assert len(level_fmts) == 3
        assert level_fmts[0].number_format == "decimal"
        assert level_fmts[0].text == "%1."
        assert level_fmts[1].number_format == "lowerLetter"
        assert level_fmts[1].text == "%2)"
        assert level_fmts[2].number_format == "lowerRoman"

    def it_can_add_a_bullet_definition_with_font(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering_part = Mock(spec=NumberingPart)
        numbering = Numbering(numbering_elm, numbering_part)

        levels = [
            {"number_format": "bullet", "text": "\u2022", "font": "Symbol"},
        ]
        defn = numbering.add_numbering_definition(levels)

        level_fmts = defn.level_formats
        assert len(level_fmts) == 1
        assert level_fmts[0].number_format == "bullet"

    def it_can_add_a_definition_with_indent(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering_part = Mock(spec=NumberingPart)
        numbering = Numbering(numbering_elm, numbering_part)

        levels = [
            {"number_format": "decimal", "text": "%1.", "indent": 720},
        ]
        defn = numbering.add_numbering_definition(levels)

        assert defn.num_id == 1


class DescribeNumberingDefinition:
    """Unit-test suite for `docx.numbering.NumberingDefinition` objects."""

    def it_provides_access_to_its_num_id(self):
        numbering_elm = cast(
            CT_Numbering,
            element("w:numbering/(w:num{w:numId=42}/(w:abstractNumId{w:val=0}))"),
        )
        num = numbering_elm.num_lst[0]
        defn = NumberingDefinition(num, numbering_elm)

        assert defn.num_id == 42

    def it_provides_access_to_abstract_num_id(self):
        numbering_elm = cast(
            CT_Numbering,
            element("w:numbering/(w:num{w:numId=1}/(w:abstractNumId{w:val=7}))"),
        )
        num = numbering_elm.num_lst[0]
        defn = NumberingDefinition(num, numbering_elm)

        assert defn.abstract_num_id == 7


class DescribeLevelFormat:
    """Unit-test suite for `docx.numbering.LevelFormat` objects."""

    def it_provides_access_to_level_properties(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        abstractNum = numbering_elm.add_abstractNum()
        lvl = abstractNum.add_lvl(0)
        lvl.numFmt_val = "decimal"
        lvl.lvlText_val = "%1."
        lvl.start_val = 1

        level_fmt = LevelFormat(lvl)

        assert level_fmt.level_index == 0
        assert level_fmt.number_format == "decimal"
        assert level_fmt.text == "%1."
        assert level_fmt.start == 1
