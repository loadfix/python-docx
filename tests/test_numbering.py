# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.numbering` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.numbering import LevelFormat, Numbering, NumberingDefinition
from docx.oxml.numbering import CT_AbstractNum, CT_Numbering

from .unitutil.cxml import element


class DescribeNumbering:
    def it_can_list_definitions(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        abstract_num = CT_AbstractNum.new(0)
        numbering_elm.add_abstractNum(abstract_num)
        numbering_elm.add_num(0)
        numbering = Numbering(numbering_elm, None)

        defs = numbering.definitions
        assert len(defs) == 1
        assert isinstance(defs[0], NumberingDefinition)

    def it_can_add_a_numbering_definition_with_defaults(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, None)

        defn = numbering.add_numbering_definition()

        assert isinstance(defn, NumberingDefinition)
        assert defn.num_id == 1
        assert len(defn.levels) == 1
        assert defn.levels[0].number_format == "decimal"
        assert defn.levels[0].text_pattern == "%1."

    def it_can_add_a_multi_level_definition(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, None)

        levels = [
            {"format": "decimal", "text": "%1.", "start": 1},
            {"format": "lowerAlpha", "text": "%2)", "start": 1},
            {"format": "upperRoman", "text": "%3.", "start": 1},
        ]
        defn = numbering.add_numbering_definition(levels)

        assert defn.num_id == 1
        assert len(defn.levels) == 3
        assert defn.levels[0].number_format == "decimal"
        assert defn.levels[1].number_format == "lowerAlpha"
        assert defn.levels[2].number_format == "upperRoman"
        assert defn.levels[0].text_pattern == "%1."
        assert defn.levels[1].text_pattern == "%2)"

    def it_can_add_a_bullet_definition(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, None)

        levels = [{"format": "bullet", "text": "\u2022", "start": 1}]
        defn = numbering.add_numbering_definition(levels)

        assert defn.levels[0].number_format == "bullet"


class DescribeNumberingDefinition:
    def it_knows_its_num_id(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        abstract_num = CT_AbstractNum.new(0)
        numbering_elm.add_abstractNum(abstract_num)
        num = numbering_elm.add_num(0)
        defn = NumberingDefinition(num, numbering_elm)

        assert defn.num_id == 1

    def it_knows_its_abstract_num_id(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        abstract_num = CT_AbstractNum.new(0)
        numbering_elm.add_abstractNum(abstract_num)
        num = numbering_elm.add_num(0)
        defn = NumberingDefinition(num, numbering_elm)

        assert defn.abstract_num_id == 0

    def it_can_restart_numbering(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        abstract_num = CT_AbstractNum.new(0)
        numbering_elm.add_abstractNum(abstract_num)
        num = numbering_elm.add_num(0)
        defn = NumberingDefinition(num, numbering_elm)

        new_defn = defn.restart()

        assert new_defn.num_id != defn.num_id
        assert new_defn.abstract_num_id == defn.abstract_num_id
        # verify lvlOverride/startOverride was added
        new_num = numbering_elm.num_having_numId(new_defn.num_id)
        assert len(new_num.lvlOverride_lst) == 1
        assert new_num.lvlOverride_lst[0].startOverride.val == 1


class DescribeLevelFormat:
    def it_knows_its_properties(self):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        abstract_num = CT_AbstractNum.new(0)
        lvl = abstract_num.add_lvl(0)
        lvl.start_val = 1
        lvl.numFmt_val = "decimal"
        lvl.lvlText_val = "%1."
        numbering_elm.add_abstractNum(abstract_num)

        level_fmt = LevelFormat(lvl)

        assert level_fmt.level == 0
        assert level_fmt.number_format == "decimal"
        assert level_fmt.text_pattern == "%1."
        assert level_fmt.start == 1
