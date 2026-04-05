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

    def it_validates_and_converts_to_xml(self):
        assert BaseIntType.to_xml(42) == "42"

    def it_validates_int(self):
        BaseSimpleType.validate_int(42)

    def it_raises_on_non_int(self):
        with pytest.raises(TypeError, match="value must be <type 'int'>"):
            BaseSimpleType.validate_int("not an int")

    def it_validates_int_in_range(self):
        BaseSimpleType.validate_int_in_range(5, 0, 10)

    def it_raises_on_int_out_of_range(self):
        with pytest.raises(ValueError, match="value must be in range"):
            BaseSimpleType.validate_int_in_range(11, 0, 10)

    def it_validates_string(self):
        assert BaseSimpleType.validate_string("hello") == "hello"

    def it_raises_on_non_string(self):
        with pytest.raises(TypeError, match="value must be a string"):
            BaseSimpleType.validate_string(42)


class DescribeBaseStringType:
    """Unit-test suite for `docx.oxml.simpletypes.BaseStringType`."""

    def it_can_convert_from_xml(self):
        assert BaseStringType.convert_from_xml("hello") == "hello"

    def it_can_convert_to_xml(self):
        assert BaseStringType.convert_to_xml("hello") == "hello"

    def it_validates_string_values(self):
        BaseStringType.validate("hello")

    def it_raises_on_non_string(self):
        with pytest.raises(TypeError, match="value must be a string"):
            BaseStringType.validate(42)


class DescribeBaseStringEnumerationType:
    """Unit-test suite for `docx.oxml.simpletypes.BaseStringEnumerationType`."""

    def it_validates_member_values(self):

        class TestEnum(BaseStringEnumerationType):
            _members = ("a", "b", "c")

        TestEnum.validate("a")

    def it_raises_on_non_member_values(self):

        class TestEnum(BaseStringEnumerationType):
            _members = ("a", "b")

        with pytest.raises(ValueError, match="must be one of"):
            TestEnum.validate("z")


class DescribeXsdBoolean:
    """Unit-test suite for `docx.oxml.simpletypes.XsdBoolean`."""

    @pytest.mark.parametrize(
        ("str_value", "expected"),
        [("1", True), ("0", False), ("true", True), ("false", False)],
    )
    def it_can_convert_from_xml(self, str_value: str, expected: bool):
        assert XsdBoolean.convert_from_xml(str_value) is expected

    def it_raises_on_invalid_xml_value(self):
        with pytest.raises(InvalidXmlError):
            XsdBoolean.convert_from_xml("yes")

    @pytest.mark.parametrize(
        ("value", "expected"), [(True, "1"), (False, "0")]
    )
    def it_can_convert_to_xml(self, value: bool, expected: str):
        assert XsdBoolean.convert_to_xml(value) == expected

    def it_validates_bool_values(self):
        XsdBoolean.validate(True)
        XsdBoolean.validate(False)

    def it_raises_on_non_bool(self):
        with pytest.raises(TypeError, match="only True or False"):
            XsdBoolean.validate("true")


class DescribeXsdInt:
    """Unit-test suite for `docx.oxml.simpletypes.XsdInt`."""

    def it_accepts_valid_int(self):
        XsdInt.validate(0)
        XsdInt.validate(-2147483648)
        XsdInt.validate(2147483647)

    def it_raises_on_out_of_range(self):
        with pytest.raises(ValueError, match="value must be in range"):
            XsdInt.validate(2147483648)


class DescribeXsdLong:
    """Unit-test suite for `docx.oxml.simpletypes.XsdLong`."""

    def it_raises_on_out_of_range(self):
        with pytest.raises(ValueError, match="value must be in range"):
            XsdLong.validate(9223372036854775808)


class DescribeXsdUnsignedInt:
    """Unit-test suite for `docx.oxml.simpletypes.XsdUnsignedInt`."""

    def it_raises_on_out_of_range(self):
        with pytest.raises(ValueError, match="value must be in range"):
            XsdUnsignedInt.validate(-1)


class DescribeXsdUnsignedLong:
    """Unit-test suite for `docx.oxml.simpletypes.XsdUnsignedLong`."""

    def it_raises_on_out_of_range(self):
        with pytest.raises(ValueError, match="value must be in range"):
            XsdUnsignedLong.validate(-1)


class DescribeST_BrClear:
    """Unit-test suite for `docx.oxml.simpletypes.ST_BrClear`."""

    def it_accepts_valid_values(self):
        for val in ("none", "left", "right", "all"):
            ST_BrClear.validate(val)

    def it_raises_on_invalid_value(self):
        with pytest.raises(ValueError, match="must be one of"):
            ST_BrClear.validate("invalid")


class DescribeST_BrType:
    """Unit-test suite for `docx.oxml.simpletypes.ST_BrType`."""

    def it_accepts_valid_values(self):
        for val in ("page", "column", "textWrapping"):
            ST_BrType.validate(val)

    def it_raises_on_invalid_value(self):
        with pytest.raises(ValueError, match="must be one of"):
            ST_BrType.validate("invalid")


class DescribeST_Coordinate:
    """Unit-test suite for `docx.oxml.simpletypes.ST_Coordinate`."""

    def it_can_convert_EMU_from_xml(self):
        result = ST_Coordinate.convert_from_xml("914400")
        assert result == Emu(914400)

    def it_can_convert_universal_measure_from_xml(self):
        result = ST_Coordinate.convert_from_xml("1in")
        assert result == Emu(914400)

    def it_validates_coordinate_values(self):
        ST_Coordinate.validate(0)


class DescribeST_CoordinateUnqualified:
    """Unit-test suite for `docx.oxml.simpletypes.ST_CoordinateUnqualified`."""

    def it_raises_on_out_of_range(self):
        with pytest.raises(ValueError, match="value must be in range"):
            ST_CoordinateUnqualified.validate(27273042316901)


class DescribeST_DateTime:
    """Unit-test suite for `docx.oxml.simpletypes.ST_DateTime`."""

    def it_can_convert_Z_suffix_from_xml(self):
        result = ST_DateTime.convert_from_xml("2023-10-01T12:00:00Z")
        assert result == dt.datetime(2023, 10, 1, 12, 0, 0, tzinfo=dt.timezone.utc)

    def it_can_convert_Z_suffix_with_fractional_seconds(self):
        result = ST_DateTime.convert_from_xml("2023-10-01T12:00:00.500Z")
        assert result.microsecond == 500000

    def it_can_convert_iso_format_with_offset(self):
        result = ST_DateTime.convert_from_xml("2023-10-01T12:00:00+00:00")
        assert result.tzinfo is not None

    def it_falls_back_to_epoch_on_garbage(self):
        result = ST_DateTime.convert_from_xml("not-a-date")
        assert result == dt.datetime(1970, 1, 1, tzinfo=dt.timezone.utc)

    def it_can_convert_to_xml(self):
        value = dt.datetime(2023, 10, 1, 12, 0, 0, tzinfo=dt.timezone.utc)
        assert ST_DateTime.convert_to_xml(value) == "2023-10-01T12:00:00Z"

    def it_validates_datetime_values(self):
        ST_DateTime.validate(dt.datetime.now())

    def it_raises_on_non_datetime(self):
        with pytest.raises(TypeError, match="only a datetime.datetime object"):
            ST_DateTime.validate("2023-01-01")


class DescribeST_HexColor:
    """Unit-test suite for `docx.oxml.simpletypes.ST_HexColor`."""

    def it_can_convert_auto_from_xml(self):
        result = ST_HexColor.convert_from_xml("auto")
        assert result == "auto"

    def it_can_convert_hex_from_xml(self):
        result = ST_HexColor.convert_from_xml("FF0000")
        assert result == RGBColor(0xFF, 0x00, 0x00)

    def it_can_convert_to_xml(self):
        assert ST_HexColor.convert_to_xml(RGBColor(0xFF, 0x00, 0x00)) == "FF0000"

    def it_validates_rgb_color(self):
        ST_HexColor.validate(RGBColor(0, 0, 0))

    def it_raises_on_non_rgb(self):
        with pytest.raises(ValueError, match="rgb color value must be RGBColor"):
            ST_HexColor.validate("FF0000")


class DescribeST_HpsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_HpsMeasure`."""

    def it_can_convert_half_points_from_xml(self):
        result = ST_HpsMeasure.convert_from_xml("24")
        assert result == Pt(12.0)

    def it_can_convert_universal_measure_from_xml(self):
        result = ST_HpsMeasure.convert_from_xml("1in")
        assert result == Emu(914400)

    def it_can_convert_to_xml(self):
        result = ST_HpsMeasure.convert_to_xml(Pt(12))
        assert result == "24"


class DescribeST_OnOff:
    """Unit-test suite for `docx.oxml.simpletypes.ST_OnOff`."""

    @pytest.mark.parametrize(
        ("str_value", "expected"),
        [
            ("1", True), ("0", False), ("true", True), ("false", False),
            ("on", True), ("off", False),
        ],
    )
    def it_can_convert_from_xml(self, str_value: str, expected: bool):
        assert ST_OnOff.convert_from_xml(str_value) is expected

    def it_raises_on_invalid_value(self):
        with pytest.raises(InvalidXmlError):
            ST_OnOff.convert_from_xml("yes")


class DescribeST_PositiveCoordinate:
    """Unit-test suite for `docx.oxml.simpletypes.ST_PositiveCoordinate`."""

    def it_can_convert_from_xml(self):
        result = ST_PositiveCoordinate.convert_from_xml("914400")
        assert result == Emu(914400)

    def it_raises_on_negative(self):
        with pytest.raises(ValueError, match="value must be in range"):
            ST_PositiveCoordinate.validate(-1)


class DescribeST_SignedTwipsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_SignedTwipsMeasure`."""

    def it_can_convert_twips_from_xml(self):
        result = ST_SignedTwipsMeasure.convert_from_xml("720")
        assert result == Twips(720)

    def it_can_convert_universal_measure_from_xml(self):
        result = ST_SignedTwipsMeasure.convert_from_xml("1in")
        assert result == Emu(914400)

    def it_can_convert_to_xml(self):
        result = ST_SignedTwipsMeasure.convert_to_xml(Twips(720))
        assert result == "720"


class DescribeST_TblLayoutType:
    """Unit-test suite for `docx.oxml.simpletypes.ST_TblLayoutType`."""

    def it_accepts_valid_values(self):
        ST_TblLayoutType.validate("fixed")
        ST_TblLayoutType.validate("autofit")

    def it_raises_on_invalid_value(self):
        with pytest.raises(ValueError, match="must be one of"):
            ST_TblLayoutType.validate("invalid")


class DescribeST_TblWidth:
    """Unit-test suite for `docx.oxml.simpletypes.ST_TblWidth`."""

    def it_accepts_valid_values(self):
        for val in ("auto", "dxa", "nil", "pct"):
            ST_TblWidth.validate(val)

    def it_raises_on_invalid_value(self):
        with pytest.raises(ValueError, match="must be one of"):
            ST_TblWidth.validate("invalid")


class DescribeST_TwipsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_TwipsMeasure`."""

    def it_can_convert_from_xml(self):
        assert ST_TwipsMeasure.convert_from_xml("720") == Twips(720)

    def it_can_convert_universal_measure_from_xml(self):
        assert ST_TwipsMeasure.convert_from_xml("1in") == Emu(914400)

    def it_can_convert_to_xml(self):
        assert ST_TwipsMeasure.convert_to_xml(Twips(720)) == "720"


class DescribeST_UniversalMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_UniversalMeasure`."""

    @pytest.mark.parametrize(
        ("str_value", "expected_emu"),
        [
            ("1in", 914400),
            ("1cm", 360000),
            ("1mm", 36000),
            ("1pt", 12700),
            ("1pc", 152400),
            ("1pi", 152400),
        ],
    )
    def it_can_convert_various_units_from_xml(self, str_value: str, expected_emu: int):
        assert ST_UniversalMeasure.convert_from_xml(str_value) == Emu(expected_emu)
