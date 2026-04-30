# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.numbering` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.text import WD_NUMBER_FORMAT
from docx.numbering import Level, Numbering, NumberingDefinition
from docx.oxml.numbering import CT_Numbering
from docx.shared import Inches

from .unitutil.cxml import element
from .unitutil.mock import instance_mock
from docx.parts.numbering import NumberingPart


class DescribeNumbering:
    """Unit-test suite for `docx.numbering.Numbering`."""

    def it_is_empty_for_a_fresh_numbering_element(self, request: pytest.FixtureRequest):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        part = instance_mock(request, NumberingPart)

        numbering = Numbering(numbering_elm, part)

        assert len(numbering) == 0
        assert list(numbering) == []
        assert numbering.definitions == []

    def it_can_add_a_numbering_definition_from_mapping_specs(
        self, request: pytest.FixtureRequest
    ):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))

        defn = numbering.add_numbering_definition(
            levels=[
                {
                    "format": WD_NUMBER_FORMAT.DECIMAL,
                    "text": "%1.",
                    "indent": Inches(0.5),
                },
                {
                    "format": WD_NUMBER_FORMAT.LOWER_LETTER,
                    "text": "%2)",
                    "indent": Inches(1.0),
                },
            ]
        )

        assert isinstance(defn, NumberingDefinition)
        assert defn.abstract_num_id == 0
        assert len(defn.levels) == 2
        # -- a matching w:num instance was created so the definition is usable --
        assert len(numbering_elm.num_lst) == 1
        assert numbering_elm.num_lst[0].abstractNumId.val == 0

        level_0 = defn.levels[0]
        assert level_0.ilvl == 0
        assert level_0.number_format == WD_NUMBER_FORMAT.DECIMAL
        assert level_0.text == "%1."

        level_1 = defn.levels[1]
        assert level_1.ilvl == 1
        assert level_1.number_format == WD_NUMBER_FORMAT.LOWER_LETTER
        assert level_1.text == "%2)"

    def it_accepts_positional_tuple_level_specs(self, request: pytest.FixtureRequest):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))

        defn = numbering.add_numbering_definition(
            levels=[(WD_NUMBER_FORMAT.UPPER_ROMAN, "%1.")]
        )

        assert defn.levels[0].number_format == WD_NUMBER_FORMAT.UPPER_ROMAN

    def it_accepts_raw_string_format_values(self, request: pytest.FixtureRequest):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))

        defn = numbering.add_numbering_definition(
            levels=[{"format": "bullet", "text": "•", "font": "Symbol"}]
        )

        lvl = defn.levels[0]
        assert lvl.number_format == WD_NUMBER_FORMAT.BULLET
        assert lvl.text == "•"

    def it_supports_multiple_definitions(self, request: pytest.FixtureRequest):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))

        defn_a = numbering.add_numbering_definition(
            levels=[{"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%1."}]
        )
        defn_b = numbering.add_numbering_definition(
            levels=[{"format": WD_NUMBER_FORMAT.UPPER_LETTER, "text": "%1."}]
        )

        assert defn_a.abstract_num_id == 0
        assert defn_b.abstract_num_id == 1
        assert len(numbering) == 2
        # -- each definition has its own w:num instance --
        assert len(numbering_elm.num_lst) == 2


class DescribeNumberingDefinition:
    """Unit-test suite for `docx.numbering.NumberingDefinition`."""

    def it_exposes_levels_in_order(self, request: pytest.FixtureRequest):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))
        defn = numbering.add_numbering_definition(
            levels=[
                {"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%1."},
                {"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%2."},
                {"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%3."},
            ]
        )

        levels = defn.levels

        assert [lvl.ilvl for lvl in levels] == [0, 1, 2]
        assert all(isinstance(lvl, Level) for lvl in levels)

    def it_can_return_a_level_by_ilvl(self, request: pytest.FixtureRequest):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))
        defn = numbering.add_numbering_definition(
            levels=[
                {"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%1."},
                {"format": WD_NUMBER_FORMAT.LOWER_LETTER, "text": "%2)"},
            ]
        )

        lvl = defn.level(1)

        assert lvl is not None
        assert lvl.ilvl == 1
        assert lvl.number_format == WD_NUMBER_FORMAT.LOWER_LETTER

    def it_returns_None_for_missing_level(self, request: pytest.FixtureRequest):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))
        defn = numbering.add_numbering_definition(
            levels=[{"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%1."}]
        )

        assert defn.level(5) is None


class DescribeLevel:
    """Unit-test suite for `docx.numbering.Level`."""

    def it_reports_level_properties_from_the_underlying_lvl(
        self, request: pytest.FixtureRequest
    ):
        numbering_elm = cast(CT_Numbering, element("w:numbering"))
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))
        defn = numbering.add_numbering_definition(
            levels=[
                {
                    "format": WD_NUMBER_FORMAT.DECIMAL,
                    "text": "%1.",
                    "indent": Inches(0.75),
                }
            ]
        )

        lvl = defn.levels[0]

        assert lvl.ilvl == 0
        assert lvl.number_format == WD_NUMBER_FORMAT.DECIMAL
        assert lvl.text == "%1."
        assert lvl.start == 1
        assert lvl.indent is not None
        assert lvl.indent.inches == pytest.approx(0.75, rel=1e-3)

    def it_returns_None_number_format_for_unknown_value(
        self, request: pytest.FixtureRequest
    ):
        numbering_elm = cast(
            CT_Numbering,
            element(
                "w:numbering/w:abstractNum{w:abstractNumId=0}/"
                "w:lvl{w:ilvl=0}/w:numFmt{w:val=chicagoManual}"
            ),
        )
        numbering = Numbering(numbering_elm, instance_mock(request, NumberingPart))
        defn = numbering.definitions[0]
        lvl = defn.level(0)

        assert lvl is not None
        assert lvl.number_format is None
