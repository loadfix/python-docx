# pyright: reportImportCycles=false

"""Simple-type classes, corresponding to ST_* schema items.

The generic XSD primitive base classes live in the shared
:mod:`ooxml_xmlchemy.simpletypes` package and are re-exported below so
existing ``docx.oxml.simpletypes.*`` import paths keep working.  The
WordprocessingML-specific concrete ``ST_*`` classes that depend on
docx's ``Emu`` / ``Pt`` / ``Twips`` / ``RGBColor`` value objects stay
local to docx.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Any

from ooxml_xmlchemy.simpletypes import (
    BaseFloatType,
    BaseIntType,
    BaseSimpleType,
    BaseStringEnumerationType,
    BaseStringType,
    XsdAnyUri,
    XsdBoolean,
    XsdDouble,
    XsdId,
    XsdInt,
    XsdLong,
    XsdString,
    XsdStringEnumeration,
    XsdToken,
    XsdTokenEnumeration,
    XsdUnsignedByte,
    XsdUnsignedInt,
    XsdUnsignedLong,
    XsdUnsignedShort,
)

from docx.exceptions import InvalidXmlError
from docx.shared import Emu, Pt, RGBColor, Twips

if TYPE_CHECKING:
    from docx.shared import Length


__all__ = [
    "BaseFloatType",
    "BaseIntType",
    "BaseSimpleType",
    "BaseStringEnumerationType",
    "BaseStringType",
    "ST_BrClear",
    "ST_BrType",
    "ST_Coordinate",
    "ST_CoordinateUnqualified",
    "ST_DateTime",
    "ST_DecimalNumber",
    "ST_DrawingElementId",
    "ST_EighthPointMeasure",
    "ST_HexColor",
    "ST_HexColorAuto",
    "ST_HpsMeasure",
    "ST_Merge",
    "ST_OnOff",
    "ST_PointMeasure",
    "ST_PositiveCoordinate",
    "ST_RelationshipId",
    "ST_SignedTwipsMeasure",
    "ST_String",
    "ST_TblLayoutType",
    "ST_TblWidth",
    "ST_TwipsMeasure",
    "ST_UniversalMeasure",
    "ST_VerticalAlignRun",
    "XsdAnyUri",
    "XsdBoolean",
    "XsdDouble",
    "XsdId",
    "XsdInt",
    "XsdLong",
    "XsdString",
    "XsdStringEnumeration",
    "XsdToken",
    "XsdTokenEnumeration",
    "XsdUnsignedByte",
    "XsdUnsignedInt",
    "XsdUnsignedLong",
    "XsdUnsignedShort",
]


class ST_BrClear(XsdString):
    @classmethod
    def validate(cls, value: str) -> None:
        cls.validate_string(value)
        valid_values = ("none", "left", "right", "all")
        if value not in valid_values:
            raise ValueError("must be one of %s, got '%s'" % (valid_values, value))


class ST_BrType(XsdString):
    @classmethod
    def validate(cls, value: Any) -> None:
        cls.validate_string(value)
        valid_values = ("page", "column", "textWrapping")
        if value not in valid_values:
            raise ValueError("must be one of %s, got '%s'" % (valid_values, value))


class ST_Coordinate(BaseIntType):
    @classmethod
    def convert_from_xml(cls, str_value: str) -> "Length":
        if "i" in str_value or "m" in str_value or "p" in str_value:
            return ST_UniversalMeasure.convert_from_xml(str_value)
        return Emu(int(str_value))

    @classmethod
    def validate(cls, value: Any) -> None:
        ST_CoordinateUnqualified.validate(value)


class ST_CoordinateUnqualified(XsdLong):
    @classmethod
    def validate(cls, value: Any) -> None:
        cls.validate_int_in_range(value, -27273042329600, 27273042316900)


class ST_EighthPointMeasure(BaseIntType):
    """Measurement in eighths of a point, e.g. sz="8" represents 1 point.

    Used for border widths (``w:sz`` attribute).  Prior to the
    ``python-ooxml-xmlchemy`` adoption, two ``ST_EighthPointMeasure``
    classes were defined in the same file — the second one (retained
    here) transparently interoperates with :class:`docx.shared.Length`
    values.  The earlier raw-integer definition has been removed.
    """

    @classmethod
    def convert_from_xml(cls, str_value: str) -> "Length":
        return Pt(int(str_value) / 8.0)

    @classmethod
    def convert_to_xml(cls, value: "int | Length") -> str:
        emu = Emu(value)
        eighth_points = int(round(emu.pt * 8))
        return str(eighth_points)


class ST_PointMeasure(BaseIntType):
    """Measurement in whole points, e.g. space="4" represents 4 points."""

    @classmethod
    def convert_from_xml(cls, str_value: str) -> "Length":
        return Pt(int(str_value))

    @classmethod
    def convert_to_xml(cls, value: "int | Length") -> str:
        emu = Emu(value)
        points = int(round(emu.pt))
        return str(points)


class ST_DateTime(BaseSimpleType):
    @classmethod
    def convert_from_xml(cls, str_value: str) -> dt.datetime:
        """Convert an xsd:dateTime string to a datetime object."""

        def parse_xsd_datetime(dt_str: str) -> dt.datetime:
            # -- handle trailing 'Z' (Zulu/UTC), common in Word files --
            if dt_str.endswith("Z"):
                try:
                    # -- optional fractional seconds case --
                    return dt.datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%S.%fZ").replace(
                        tzinfo=dt.timezone.utc
                    )
                except ValueError:
                    return dt.datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%SZ").replace(
                        tzinfo=dt.timezone.utc
                    )

            # -- handles explicit offsets like +00:00, -05:00, or naive datetimes --
            try:
                return dt.datetime.fromisoformat(dt_str)
            except ValueError:
                # -- fall-back to parsing as naive datetime (with or without fractional seconds) --
                try:
                    return dt.datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%S.%f")
                except ValueError:
                    return dt.datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%S")

        try:
            # -- parse anything reasonable, but never raise, just use default epoch time --
            return parse_xsd_datetime(str_value)
        except Exception:
            return dt.datetime(1970, 1, 1, tzinfo=dt.timezone.utc)

    @classmethod
    def convert_to_xml(cls, value: dt.datetime) -> str:
        # -- convert naive datetime to timezon-aware assuming local timezone --
        if value.tzinfo is None:
            value = value.astimezone()

        # -- convert to UTC if not already --
        value = value.astimezone(dt.timezone.utc)

        # -- format with 'Z' suffix for UTC --
        return value.strftime("%Y-%m-%dT%H:%M:%SZ")

    @classmethod
    def validate(cls, value: Any) -> None:
        if not isinstance(value, dt.datetime):
            raise TypeError("only a datetime.datetime object may be assigned, got '%s'" % value)


class ST_DecimalNumber(XsdInt):
    pass


class ST_DrawingElementId(XsdUnsignedInt):
    pass


class ST_HexColor(BaseStringType):
    @classmethod
    def convert_from_xml(  # pyright: ignore[reportIncompatibleMethodOverride]
        cls, str_value: str
    ) -> "RGBColor | str":
        if str_value == "auto":
            return ST_HexColorAuto.AUTO
        return RGBColor.from_string(str_value)

    @classmethod
    def convert_to_xml(  # pyright: ignore[reportIncompatibleMethodOverride]
        cls, value: RGBColor
    ) -> str:
        """Keep alpha hex numerals all uppercase just for consistency."""
        # expecting 3-tuple of ints in range 0-255
        return "%02X%02X%02X" % value

    @classmethod
    def validate(cls, value: Any) -> None:
        # must be an RGBColor object ---
        if not isinstance(value, RGBColor):
            raise ValueError(
                "rgb color value must be RGBColor object, got %s %s" % (type(value), value)
            )


class ST_HexColorAuto(XsdStringEnumeration):
    """Value for `w:color/[@val="auto"] attribute setting."""

    AUTO = "auto"

    _members = (AUTO,)


class ST_HpsMeasure(XsdUnsignedLong):
    """Half-point measure, e.g. 24.0 represents 12.0 points.

    Some .docx producers (including older Word versions on certain locales,
    and a handful of third-party tools) write non-integer half-point values
    such as ``"23.5"``. The schema reads ``xsd:unsignedLong``, but accepting
    a decimal string here lets us keep loading those documents instead of
    crashing. Fractional values are rounded to the nearest half-point when
    written back out. See upstream issues #1475, #1539 and PR #1478.
    """

    @classmethod
    def convert_from_xml(cls, str_value: str) -> "Length":
        if "m" in str_value or "n" in str_value or "p" in str_value:
            return ST_UniversalMeasure.convert_from_xml(str_value)
        # -- tolerate decimal half-points (e.g. "23.5") --
        return Pt(float(str_value) / 2.0)

    @classmethod
    def convert_to_xml(cls, value: "int | Length") -> str:
        emu = Emu(int(value))
        half_points = int(round(emu.pt * 2))
        return str(half_points)


class ST_Merge(XsdStringEnumeration):
    """Valid values for <w:xMerge val=""> attribute."""

    CONTINUE = "continue"
    RESTART = "restart"

    _members = (CONTINUE, RESTART)


class ST_OnOff(XsdBoolean):
    @classmethod
    def convert_from_xml(cls, str_value: str) -> bool:
        if str_value not in ("1", "0", "true", "false", "on", "off"):
            raise InvalidXmlError(
                "value must be one of '1', '0', 'true', 'false', 'on', or 'o"
                "ff', got '%s'" % str_value
            )
        return str_value in ("1", "true", "on")


class ST_PositiveCoordinate(XsdLong):
    @classmethod
    def convert_from_xml(cls, str_value: str) -> "Length":
        return Emu(int(str_value))

    @classmethod
    def validate(cls, value: Any) -> None:
        cls.validate_int_in_range(value, 0, 27273042316900)


class ST_RelationshipId(XsdString):
    pass


class ST_SignedTwipsMeasure(XsdInt):
    @classmethod
    def convert_from_xml(cls, str_value: str) -> "Length":
        if "i" in str_value or "m" in str_value or "p" in str_value:
            return ST_UniversalMeasure.convert_from_xml(str_value)
        return Twips(int(round(float(str_value))))

    @classmethod
    def convert_to_xml(cls, value: "int | Length") -> str:
        emu = Emu(value)
        twips = emu.twips
        return str(twips)


class ST_String(XsdString):
    pass


class ST_TblLayoutType(XsdString):
    @classmethod
    def validate(cls, value: Any) -> None:
        cls.validate_string(value)
        valid_values = ("fixed", "autofit")
        if value not in valid_values:
            raise ValueError("must be one of %s, got '%s'" % (valid_values, value))


class ST_TblWidth(XsdString):
    @classmethod
    def validate(cls, value: Any) -> None:
        cls.validate_string(value)
        valid_values = ("auto", "dxa", "nil", "pct")
        if value not in valid_values:
            raise ValueError("must be one of %s, got '%s'" % (valid_values, value))


class ST_TwipsMeasure(XsdUnsignedLong):
    """Twips measure (20ths of a point).

    Microsoft Word, when saving documents created by some third-party tools
    or older revisions, occasionally emits fractional twips like ``"283.5"``.
    The schema calls for ``xsd:unsignedLong``, but being tolerant here
    (rounding to the nearest whole twip) lets us load those documents
    instead of raising ``ValueError``. See upstream issues #1475, #1539
    and PR #1478.
    """

    @classmethod
    def convert_from_xml(cls, str_value: str) -> "Length":
        if "i" in str_value or "m" in str_value or "p" in str_value:
            return ST_UniversalMeasure.convert_from_xml(str_value)
        # -- tolerate decimal twips by rounding to the nearest whole twip --
        return Twips(int(round(float(str_value))))

    @classmethod
    def convert_to_xml(cls, value: "int | Length") -> str:
        emu = Emu(int(value))
        twips = emu.twips
        return str(twips)


class ST_UniversalMeasure(BaseSimpleType):
    @classmethod
    def convert_from_xml(cls, str_value: str) -> Emu:
        float_part, units_part = str_value[:-2], str_value[-2:]
        quantity = float(float_part)
        multiplier = {
            "mm": 36000,
            "cm": 360000,
            "in": 914400,
            "pt": 12700,
            "pc": 152400,
            "pi": 152400,
        }[units_part]
        return Emu(int(round(quantity * multiplier)))


class ST_VerticalAlignRun(XsdStringEnumeration):
    """Valid values for `w:vertAlign/@val`."""

    BASELINE = "baseline"
    SUPERSCRIPT = "superscript"
    SUBSCRIPT = "subscript"

    _members = (BASELINE, SUPERSCRIPT, SUBSCRIPT)
