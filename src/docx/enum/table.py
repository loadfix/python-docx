"""Enumerations related to tables in WordprocessingML files."""

from docx.enum.base import BaseEnum, BaseXmlEnum


class WD_CELL_VERTICAL_ALIGNMENT(BaseXmlEnum):
    """Alias: **WD_ALIGN_VERTICAL**

    Specifies the vertical alignment of text in one or more cells of a table.

    Example::

        from docx.enum.table import WD_ALIGN_VERTICAL

        table = document.add_table(3, 3)
        table.cell(0, 0).vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

    MS API name: `WdCellVerticalAlignment`

    https://msdn.microsoft.com/en-us/library/office/ff193345.aspx
    """

    TOP = (0, "top", "Text is aligned to the top border of the cell.")
    """Text is aligned to the top border of the cell."""

    CENTER = (1, "center", "Text is aligned to the center of the cell.")
    """Text is aligned to the center of the cell."""

    BOTTOM = (3, "bottom", "Text is aligned to the bottom border of the cell.")
    """Text is aligned to the bottom border of the cell."""

    BOTH = (
        101,
        "both",
        "This is an option in the OpenXml spec, but not in Word itself. It's not"
        " clear what Word behavior this setting produces. If you find out please"
        " let us know and we'll update this documentation. Otherwise, probably best"
        " to avoid this option.",
    )
    """This is an option in the OpenXml spec, but not in Word itself.

    It's not clear what Word behavior this setting produces. If you find out please let
    us know and we'll update this documentation. Otherwise, probably best to avoid this
    option.
    """


WD_ALIGN_VERTICAL = WD_CELL_VERTICAL_ALIGNMENT


class WD_ROW_HEIGHT_RULE(BaseXmlEnum):
    """Alias: **WD_ROW_HEIGHT**

    Specifies the rule for determining the height of a table row

    Example::

        from docx.enum.table import WD_ROW_HEIGHT_RULE

        table = document.add_table(3, 3)
        table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    MS API name: `WdRowHeightRule`

    https://msdn.microsoft.com/en-us/library/office/ff193620.aspx
    """

    AUTO = (
        0,
        "auto",
        "The row height is adjusted to accommodate the tallest value in the row.",
    )
    """The row height is adjusted to accommodate the tallest value in the row."""

    AT_LEAST = (1, "atLeast", "The row height is at least a minimum specified value.")
    """The row height is at least a minimum specified value."""

    EXACTLY = (2, "exact", "The row height is an exact value.")
    """The row height is an exact value."""


WD_ROW_HEIGHT = WD_ROW_HEIGHT_RULE


class WD_TABLE_ALIGNMENT(BaseXmlEnum):
    """Specifies table justification type.

    Example::

        from docx.enum.table import WD_TABLE_ALIGNMENT

        table = document.add_table(3, 3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

    MS API name: `WdRowAlignment`

    http://office.microsoft.com/en-us/word-help/HV080607259.aspx
    """

    LEFT = (0, "left", "Left-aligned")
    """Left-aligned"""

    CENTER = (1, "center", "Center-aligned.")
    """Center-aligned."""

    RIGHT = (2, "right", "Right-aligned.")
    """Right-aligned."""


class WD_SHADING_PATTERN(BaseXmlEnum):
    """Specifies the pattern style for cell shading.

    Example::

        from docx.enum.table import WD_SHADING_PATTERN

        table = document.add_table(3, 3)
        cell = table.cell(0, 0)
        cell.shading.pattern = WD_SHADING_PATTERN.CLEAR

    MS API name: `WdShadingPattern` (partial)
    """

    CLEAR = (0, "clear", "No pattern, just background fill color.")
    """No pattern, just background fill color."""

    SOLID = (1, "solid", "Solid pattern (foreground color fills entire area).")
    """Solid pattern (foreground color fills entire area)."""

    HORZ_STRIPE = (2, "horzStripe", "Horizontal stripe pattern.")
    """Horizontal stripe pattern."""

    VERT_STRIPE = (3, "vertStripe", "Vertical stripe pattern.")
    """Vertical stripe pattern."""

    REVERSE_DIAG_STRIPE = (4, "reverseDiagStripe", "Reverse diagonal stripe pattern.")
    """Reverse diagonal stripe pattern."""

    DIAG_STRIPE = (5, "diagStripe", "Diagonal stripe pattern.")
    """Diagonal stripe pattern."""

    HORZ_CROSS = (6, "horzCross", "Horizontal cross pattern.")
    """Horizontal cross pattern."""

    DIAG_CROSS = (7, "diagCross", "Diagonal cross pattern.")
    """Diagonal cross pattern."""

    THIN_HORZ_STRIPE = (8, "thinHorzStripe", "Thin horizontal stripe pattern.")
    """Thin horizontal stripe pattern."""

    THIN_VERT_STRIPE = (9, "thinVertStripe", "Thin vertical stripe pattern.")
    """Thin vertical stripe pattern."""

    THIN_REVERSE_DIAG_STRIPE = (
        10,
        "thinReverseDiagStripe",
        "Thin reverse diagonal stripe pattern.",
    )
    """Thin reverse diagonal stripe pattern."""

    THIN_DIAG_STRIPE = (11, "thinDiagStripe", "Thin diagonal stripe pattern.")
    """Thin diagonal stripe pattern."""

    THIN_HORZ_CROSS = (12, "thinHorzCross", "Thin horizontal cross pattern.")
    """Thin horizontal cross pattern."""

    THIN_DIAG_CROSS = (13, "thinDiagCross", "Thin diagonal cross pattern.")
    """Thin diagonal cross pattern."""

    PCT_5 = (14, "pct5", "5 percent fill pattern.")
    """5 percent fill pattern."""

    PCT_10 = (15, "pct10", "10 percent fill pattern.")
    """10 percent fill pattern."""

    PCT_12 = (16, "pct12", "12.5 percent fill pattern.")
    """12.5 percent fill pattern."""

    PCT_15 = (17, "pct15", "15 percent fill pattern.")
    """15 percent fill pattern."""

    PCT_20 = (18, "pct20", "20 percent fill pattern.")
    """20 percent fill pattern."""

    PCT_25 = (19, "pct25", "25 percent fill pattern.")
    """25 percent fill pattern."""

    PCT_30 = (20, "pct30", "30 percent fill pattern.")
    """30 percent fill pattern."""

    PCT_35 = (21, "pct35", "35 percent fill pattern.")
    """35 percent fill pattern."""

    PCT_37 = (22, "pct37", "37.5 percent fill pattern.")
    """37.5 percent fill pattern."""

    PCT_40 = (23, "pct40", "40 percent fill pattern.")
    """40 percent fill pattern."""

    PCT_45 = (24, "pct45", "45 percent fill pattern.")
    """45 percent fill pattern."""

    PCT_50 = (25, "pct50", "50 percent fill pattern.")
    """50 percent fill pattern."""

    PCT_55 = (26, "pct55", "55 percent fill pattern.")
    """55 percent fill pattern."""

    PCT_60 = (27, "pct60", "60 percent fill pattern.")
    """60 percent fill pattern."""

    PCT_62 = (28, "pct62", "62.5 percent fill pattern.")
    """62.5 percent fill pattern."""

    PCT_65 = (29, "pct65", "65 percent fill pattern.")
    """65 percent fill pattern."""

    PCT_70 = (30, "pct70", "70 percent fill pattern.")
    """70 percent fill pattern."""

    PCT_75 = (31, "pct75", "75 percent fill pattern.")
    """75 percent fill pattern."""

    PCT_80 = (32, "pct80", "80 percent fill pattern.")
    """80 percent fill pattern."""

    PCT_85 = (33, "pct85", "85 percent fill pattern.")
    """85 percent fill pattern."""

    PCT_87 = (34, "pct87", "87.5 percent fill pattern.")
    """87.5 percent fill pattern."""

    PCT_90 = (35, "pct90", "90 percent fill pattern.")
    """90 percent fill pattern."""

    PCT_95 = (36, "pct95", "95 percent fill pattern.")
    """95 percent fill pattern."""

    NIL = (37, "nil", "No shading.")
    """No shading."""


class WD_BORDER_STYLE(BaseXmlEnum):
    """Specifies the style of a table or cell border.

    Example::

        from docx.enum.table import WD_BORDER_STYLE

        table = document.add_table(3, 3)
        table.borders.top.style = WD_BORDER_STYLE.SINGLE

    Based on the ST_Border simple type in the Open XML spec.
    """

    NONE = (0, "none", "No border.")
    """No border."""

    SINGLE = (1, "single", "A single line.")
    """A single line."""

    DOUBLE = (2, "double", "A double line.")
    """A double line."""

    DOTTED = (3, "dotted", "A dotted line.")
    """A dotted line."""

    DASHED = (4, "dashed", "A dashed line.")
    """A dashed line."""

    DOT_DASH = (5, "dotDash", "A line with alternating dots and dashes.")
    """A line with alternating dots and dashes."""

    DOT_DOT_DASH = (6, "dotDotDash", "A line with a repeating dot-dot-dash pattern.")
    """A line with a repeating dot-dot-dash pattern."""

    TRIPLE = (7, "triple", "A triple line.")
    """A triple line."""

    THIN_THICK_SMALL_GAP = (8, "thinThickSmallGap", "A thin-thick line with a small gap.")
    """A thin-thick line with a small gap."""

    THICK_THIN_SMALL_GAP = (9, "thickThinSmallGap", "A thick-thin line with a small gap.")
    """A thick-thin line with a small gap."""

    THIN_THICK_THIN_SMALL_GAP = (
        10,
        "thinThickThinSmallGap",
        "A thin-thick-thin line with a small gap.",
    )
    """A thin-thick-thin line with a small gap."""

    THIN_THICK_MEDIUM_GAP = (11, "thinThickMediumGap", "A thin-thick line with a medium gap.")
    """A thin-thick line with a medium gap."""

    THICK_THIN_MEDIUM_GAP = (12, "thickThinMediumGap", "A thick-thin line with a medium gap.")
    """A thick-thin line with a medium gap."""

    THIN_THICK_THIN_MEDIUM_GAP = (
        13,
        "thinThickThinMediumGap",
        "A thin-thick-thin line with a medium gap.",
    )
    """A thin-thick-thin line with a medium gap."""

    THIN_THICK_LARGE_GAP = (14, "thinThickLargeGap", "A thin-thick line with a large gap.")
    """A thin-thick line with a large gap."""

    THICK_THIN_LARGE_GAP = (15, "thickThinLargeGap", "A thick-thin line with a large gap.")
    """A thick-thin line with a large gap."""

    THIN_THICK_THIN_LARGE_GAP = (
        16,
        "thinThickThinLargeGap",
        "A thin-thick-thin line with a large gap.",
    )
    """A thin-thick-thin line with a large gap."""

    WAVE = (17, "wave", "A wavy line.")
    """A wavy line."""

    DOUBLE_WAVE = (18, "doubleWave", "A double wavy line.")
    """A double wavy line."""

    DASH_SMALL_GAP = (19, "dashSmallGap", "A dashed line with small gaps.")
    """A dashed line with small gaps."""

    DASH_DOT_STROKED = (20, "dashDotStroked", "A dash-dot stroked line.")
    """A dash-dot stroked line."""

    THREE_D_EMBOSS = (21, "threeDEmboss", "A 3D embossed line.")
    """A 3D embossed line."""

    THREE_D_ENGRAVE = (22, "threeDEngrave", "A 3D engraved line.")
    """A 3D engraved line."""

    OUTSET = (23, "outset", "An outset line.")
    """An outset line."""

    INSET = (24, "inset", "An inset line.")
    """An inset line."""

    NIL = (25, "nil", "No border (used to override inherited border).")
    """No border (used to override inherited border)."""


class WD_TABLE_DIRECTION(BaseEnum):
    """Specifies the direction in which an application orders cells in the specified
    table or row.

    Example::

        from docx.enum.table import WD_TABLE_DIRECTION

        table = document.add_table(3, 3)
        table.direction = WD_TABLE_DIRECTION.RTL

    MS API name: `WdTableDirection`

    http://msdn.microsoft.com/en-us/library/ff835141.aspx
    """

    LTR = (
        0,
        "The table or row is arranged with the first column in the leftmost position.",
    )
    """The table or row is arranged with the first column in the leftmost position."""

    RTL = (
        1,
        "The table or row is arranged with the first column in the rightmost position.",
    )
    """The table or row is arranged with the first column in the rightmost position."""
