"""Unit-test suite for docx.oxml.simpletypes (tolerant numeric parsers)."""

from __future__ import annotations

import pytest

from docx.oxml.simpletypes import ST_HpsMeasure, ST_TwipsMeasure
from docx.shared import Emu


class DescribeST_TwipsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_TwipsMeasure`.

    Covers upstream issues #1475, #1539 and PR #1478 — fractional twips
    written by some third-party tools must not crash the loader.
    """

    def it_parses_an_integer_twips_value(self):
        length = ST_TwipsMeasure.convert_from_xml("283")
        assert int(length.twips) == 283

    def it_tolerates_a_decimal_twips_value(self):
        # -- "283.5" should round to 284 twips instead of raising ValueError --
        length = ST_TwipsMeasure.convert_from_xml("283.5")
        assert int(length.twips) == 284

    def it_rounds_half_down_twips_to_nearest_integer(self):
        length = ST_TwipsMeasure.convert_from_xml("283.49")
        assert int(length.twips) == 283

    def it_still_parses_a_universal_measure_with_units(self):
        length = ST_TwipsMeasure.convert_from_xml("1in")
        assert length.inches == pytest.approx(1.0)

    def it_serializes_an_emu_value_to_twips(self):
        assert ST_TwipsMeasure.convert_to_xml(Emu(914400)) == "1440"


class DescribeST_HpsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_HpsMeasure`.

    Covers upstream issues #1475, #1539 and PR #1478 — fractional half-points
    (e.g. ``"23.5"``) must not crash the loader.
    """

    def it_parses_an_integer_half_point_value(self):
        length = ST_HpsMeasure.convert_from_xml("24")
        assert length.pt == pytest.approx(12.0)

    def it_tolerates_a_decimal_half_point_value(self):
        # -- "23.5" means 11.75 points; prior behavior raised ValueError --
        length = ST_HpsMeasure.convert_from_xml("23.5")
        assert length.pt == pytest.approx(11.75)

    def it_still_parses_a_universal_measure_with_units(self):
        length = ST_HpsMeasure.convert_from_xml("12pt")
        assert length.pt == pytest.approx(12.0)

    def it_serializes_an_emu_value_to_half_points(self):
        # -- 12 pt == 24 half-points --
        from docx.shared import Pt

        assert ST_HpsMeasure.convert_to_xml(Pt(12)) == "24"
