# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.simpletypes` module."""

from __future__ import annotations

import datetime as dt

import pytest

from docx.exceptions import InvalidXmlError
from docx.oxml.simpletypes import (
    BaseIntType,
    BaseSimpleType,
    BaseStringEnumerationType,
    BaseStringType,
    ST_BrClear,
    ST_BrType,
    ST_Coordinate,
    ST_CoordinateUnqualified,
    ST_DateTime,
    ST_HexColor,
    ST_HpsMeasure,
    ST_OnOff,
    ST_PositiveCoordinate,
    ST_SignedTwipsMeasure,
    ST_TblLayoutType,
    ST_TblWidth,
    ST_TwipsMeasure,
    ST_UniversalMeasure,
    XsdBoolean,
    XsdInt,
    XsdLong,
    XsdUnsignedInt,
    XsdUnsignedLong,
)
from docx.shared import Emu, Pt, RGBColor, Twips


class DescribeBaseSimpleType:
    """Unit-test suite for `docx.oxml.simpletypes.BaseSimpleType`."""

    def it_can_convert_from_xml(self):
        assert BaseSimpleType.from_xml("42") == 42

    def it_can_convert_to_xml(self):
        assert BaseSimpleType.to_xml(42) == BaseSimpleType.convert_to_xml(42)

    def it_validates_int(self):
        BaseSimpleType.validate_int(42)
        with pytest.raises(TypeError, match="value must be <type 'int'>"):
            BaseSimpleType.validate_int("not an int")

    def it_validates_int_in_range(self):
        BaseSimpleType.validate_int_in_range(5, 0, 10)
        with pytest.raises(ValueError, match="value must be in range"):
            BaseSimpleType.validate_int_in_range(11, 0, 10)
        with pytest.raises(ValueError, match="value must be in range"):
            BaseSimpleType.validate_int_in_range(-1, 0, 10)

    def it_validates_string(self):
        assert BaseSimpleType.validate_string("hello") == "hello"
        with pytest.raises(TypeError, match="value must be a string"):
            BaseSimpleType.validate_string(42)


class DescribeBaseIntType:
    """Unit-test suite for `docx.oxml.simpletypes.BaseIntType`."""

    def it_converts_from_xml(self):
        assert BaseIntType.convert_from_xml("99") == 99

    def it_converts_to_xml(self):
        assert BaseIntType.convert_to_xml(99) == "99"

    def it_validates_int(self):
        BaseIntType.validate(42)
        with pytest.raises(TypeError):
            BaseIntType.validate("not int")


class DescribeBaseStringType:
    """Unit-test suite for `docx.oxml.simpletypes.BaseStringType`."""

    def it_converts_from_xml(self):
        assert BaseStringType.convert_from_xml("hello") == "hello"

    def it_converts_to_xml(self):
        assert BaseStringType.convert_to_xml("hello") == "hello"

    def it_validates(self):
        BaseStringType.validate("hello")
        with pytest.raises(TypeError):
            BaseStringType.validate(42)


class DescribeBaseStringEnumerationType:
    """Unit-test suite for `docx.oxml.simpletypes.BaseStringEnumerationType`."""

    def it_validates_membership(self):
        class TestEnum(BaseStringEnumerationType):
            _members = ("a", "b", "c")

        TestEnum.validate("a")
        with pytest.raises(ValueError, match="must be one of"):
            TestEnum.validate("d")
        with pytest.raises(TypeError):
            TestEnum.validate(42)


class DescribeXsdBoolean:
    """Unit-test suite for `docx.oxml.simpletypes.XsdBoolean`."""

    @pytest.mark.parametrize(
        ("str_value", "expected"),
        [("1", True), ("0", False), ("true", True), ("false", False)],
    )
    def it_converts_from_xml(self, str_value: str, expected: bool):
        assert XsdBoolean.convert_from_xml(str_value) == expected

    def it_raises_on_invalid_xml_value(self):
        with pytest.raises(InvalidXmlError):
            XsdBoolean.convert_from_xml("yes")

    @pytest.mark.parametrize(
        ("value", "expected"),
        [(True, "1"), (False, "0")],
    )
    def it_converts_to_xml(self, value: bool, expected: str):
        assert XsdBoolean.convert_to_xml(value) == expected

    def it_validates(self):
        XsdBoolean.validate(True)
        XsdBoolean.validate(False)
        with pytest.raises(TypeError, match="only True or False"):
            XsdBoolean.validate("yes")


class DescribeXsdInt:
    """Unit-test suite for `docx.oxml.simpletypes.XsdInt`."""

    def it_validates_in_range(self):
        XsdInt.validate(0)
        XsdInt.validate(-2147483648)
        XsdInt.validate(2147483647)
        with pytest.raises(ValueError):
            XsdInt.validate(2147483648)


class DescribeXsdLong:
    """Unit-test suite for `docx.oxml.simpletypes.XsdLong`."""

    def it_validates_in_range(self):
        XsdLong.validate(0)
        with pytest.raises(ValueError):
            XsdLong.validate(9223372036854775808)


class DescribeXsdUnsignedInt:
    """Unit-test suite for `docx.oxml.simpletypes.XsdUnsignedInt`."""

    def it_validates_in_range(self):
        XsdUnsignedInt.validate(0)
        XsdUnsignedInt.validate(4294967295)
        with pytest.raises(ValueError):
            XsdUnsignedInt.validate(-1)


class DescribeXsdUnsignedLong:
    """Unit-test suite for `docx.oxml.simpletypes.XsdUnsignedLong`."""

    def it_validates_in_range(self):
        XsdUnsignedLong.validate(0)
        with pytest.raises(ValueError):
            XsdUnsignedLong.validate(-1)


class DescribeST_BrClear:
    """Unit-test suite for `docx.oxml.simpletypes.ST_BrClear`."""

    def it_validates_valid_values(self):
        for val in ("none", "left", "right", "all"):
            ST_BrClear.validate(val)

    def it_raises_on_invalid_value(self):
        with pytest.raises(ValueError, match="must be one of"):
            ST_BrClear.validate("invalid")


class DescribeST_BrType:
    """Unit-test suite for `docx.oxml.simpletypes.ST_BrType`."""

    def it_validates_valid_values(self):
        for val in ("page", "column", "textWrapping"):
            ST_BrType.validate(val)

    def it_raises_on_invalid_value(self):
        with pytest.raises(ValueError, match="must be one of"):
            ST_BrType.validate("invalid")


class DescribeST_Coordinate:
    """Unit-test suite for `docx.oxml.simpletypes.ST_Coordinate`."""

    def it_converts_emu_from_xml(self):
        result = ST_Coordinate.convert_from_xml("914400")
        assert result == Emu(914400)

    def it_converts_universal_measure_from_xml(self):
        result = ST_Coordinate.convert_from_xml("1in")
        assert result == Emu(914400)

    def it_validates(self):
        ST_Coordinate.validate(0)


class DescribeST_CoordinateUnqualified:
    """Unit-test suite for `docx.oxml.simpletypes.ST_CoordinateUnqualified`."""

    def it_validates_in_range(self):
        ST_CoordinateUnqualified.validate(0)
        with pytest.raises(ValueError):
            ST_CoordinateUnqualified.validate(27273042316901)


class DescribeST_DateTime:
    """Unit-test suite for `docx.oxml.simpletypes.ST_DateTime`."""

    def it_converts_zulu_time_from_xml(self):
        result = ST_DateTime.convert_from_xml("2023-10-01T12:00:00Z")
        assert result == dt.datetime(2023, 10, 1, 12, 0, 0, tzinfo=dt.timezone.utc)

    def it_converts_zulu_time_with_fractional_seconds(self):
        result = ST_DateTime.convert_from_xml("2023-10-01T12:00:00.500Z")
        assert result.tzinfo == dt.timezone.utc
        assert result.microsecond == 500000

    def it_converts_iso_format_with_offset(self):
        result = ST_DateTime.convert_from_xml("2023-10-01T12:00:00+00:00")
        assert result.year == 2023

    def it_falls_back_on_unparseable_input(self):
        result = ST_DateTime.convert_from_xml("not-a-date")
        assert result == dt.datetime(1970, 1, 1, tzinfo=dt.timezone.utc)

    def it_converts_to_xml(self):
        value = dt.datetime(2023, 10, 1, 12, 0, 0, tzinfo=dt.timezone.utc)
        assert ST_DateTime.convert_to_xml(value) == "2023-10-01T12:00:00Z"

    def it_converts_naive_datetime_to_xml(self):
        value = dt.datetime(2023, 10, 1, 12, 0, 0)
        result = ST_DateTime.convert_to_xml(value)
        # result should end with Z (UTC)
        assert result.endswith("Z")

    def it_validates(self):
        ST_DateTime.validate(dt.datetime.now())
        with pytest.raises(TypeError, match="only a datetime.datetime"):
            ST_DateTime.validate("2023-01-01")


class DescribeST_HexColor:
    """Unit-test suite for `docx.oxml.simpletypes.ST_HexColor`."""

    def it_converts_auto_from_xml(self):
        result = ST_HexColor.convert_from_xml("auto")
        assert result == "auto"

    def it_converts_hex_color_from_xml(self):
        result = ST_HexColor.convert_from_xml("FF0000")
        assert isinstance(result, RGBColor)
        assert str(result) == "FF0000"

    def it_converts_to_xml(self):
        result = ST_HexColor.convert_to_xml(RGBColor(0xFF, 0x00, 0x00))
        assert result == "FF0000"

    def it_validates(self):
        ST_HexColor.validate(RGBColor(0, 0, 0))
        with pytest.raises(ValueError, match="rgb color value must be RGBColor"):
            ST_HexColor.validate("FF0000")


class DescribeST_HpsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_HpsMeasure`."""

    def it_converts_half_points_from_xml(self):
        result = ST_HpsMeasure.convert_from_xml("24")
        assert result == Pt(12.0)

    def it_converts_universal_measure_from_xml(self):
        result = ST_HpsMeasure.convert_from_xml("12pt")
        assert result == Emu(12 * 12700)

    def it_converts_to_xml(self):
        result = ST_HpsMeasure.convert_to_xml(Pt(12))
        assert result == "24"


class DescribeST_OnOff:
    """Unit-test suite for `docx.oxml.simpletypes.ST_OnOff`."""

    @pytest.mark.parametrize(
        ("str_value", "expected"),
        [
            ("1", True),
            ("0", False),
            ("true", True),
            ("false", False),
            ("on", True),
            ("off", False),
        ],
    )
    def it_converts_from_xml(self, str_value: str, expected: bool):
        assert ST_OnOff.convert_from_xml(str_value) == expected

    def it_raises_on_invalid_xml_value(self):
        with pytest.raises(InvalidXmlError):
            ST_OnOff.convert_from_xml("yes")


class DescribeST_PositiveCoordinate:
    """Unit-test suite for `docx.oxml.simpletypes.ST_PositiveCoordinate`."""

    def it_converts_from_xml(self):
        assert ST_PositiveCoordinate.convert_from_xml("914400") == Emu(914400)

    def it_validates(self):
        ST_PositiveCoordinate.validate(0)
        with pytest.raises(ValueError):
            ST_PositiveCoordinate.validate(-1)


class DescribeST_SignedTwipsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_SignedTwipsMeasure`."""

    def it_converts_twips_from_xml(self):
        result = ST_SignedTwipsMeasure.convert_from_xml("720")
        assert result == Twips(720)

    def it_converts_universal_measure_from_xml(self):
        result = ST_SignedTwipsMeasure.convert_from_xml("1in")
        assert result == Emu(914400)

    def it_converts_to_xml(self):
        result = ST_SignedTwipsMeasure.convert_to_xml(Twips(720))
        twips_val = Emu(Twips(720)).twips
        assert result == str(twips_val)


class DescribeST_TblLayoutType:
    """Unit-test suite for `docx.oxml.simpletypes.ST_TblLayoutType`."""

    def it_validates_valid_values(self):
        ST_TblLayoutType.validate("fixed")
        ST_TblLayoutType.validate("autofit")

    def it_raises_on_invalid_value(self):
        with pytest.raises(ValueError, match="must be one of"):
            ST_TblLayoutType.validate("invalid")


class DescribeST_TblWidth:
    """Unit-test suite for `docx.oxml.simpletypes.ST_TblWidth`."""

    def it_validates_valid_values(self):
        for val in ("auto", "dxa", "nil", "pct"):
            ST_TblWidth.validate(val)

    def it_raises_on_invalid_value(self):
        with pytest.raises(ValueError, match="must be one of"):
            ST_TblWidth.validate("invalid")


class DescribeST_TwipsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_TwipsMeasure`."""

    def it_converts_from_xml(self):
        result = ST_TwipsMeasure.convert_from_xml("1440")
        assert result == Twips(1440)

    def it_converts_universal_measure_from_xml(self):
        result = ST_TwipsMeasure.convert_from_xml("1in")
        assert result == Emu(914400)

    def it_converts_to_xml(self):
        result = ST_TwipsMeasure.convert_to_xml(Twips(1440))
        twips_val = Emu(Twips(1440)).twips
        assert result == str(twips_val)


class DescribeST_UniversalMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_UniversalMeasure`."""

    @pytest.mark.parametrize(
        ("str_value", "expected_emu"),
        [
            ("1in", 914400),
            ("1mm", 36000),
            ("1cm", 360000),
            ("1pt", 12700),
            ("1pc", 152400),
            ("1pi", 152400),
            ("2.5in", 2286000),
        ],
    )
    def it_converts_from_xml(self, str_value: str, expected_emu: int):
        result = ST_UniversalMeasure.convert_from_xml(str_value)
        assert result == Emu(expected_emu)
