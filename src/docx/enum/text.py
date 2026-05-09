"""Enumerations related to text in WordprocessingML files."""

from __future__ import annotations

import enum

from docx.enum.base import BaseEnum, BaseXmlEnum


class WD_PARAGRAPH_ALIGNMENT(BaseXmlEnum):
    """Alias: **WD_ALIGN_PARAGRAPH**

    Specifies paragraph justification type.

    Example::

        from docx.enum.text import WD_ALIGN_PARAGRAPH

        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    """

    LEFT = (0, "left", "Left-aligned")
    """Left-aligned"""

    CENTER = (1, "center", "Center-aligned.")
    """Center-aligned."""

    RIGHT = (2, "right", "Right-aligned.")
    """Right-aligned."""

    JUSTIFY = (3, "both", "Fully justified.")
    """Fully justified."""

    # -- Bidi-aware aliases. ``start`` maps to left-alignment and ``end`` to
    # -- right-alignment for left-to-right paragraphs (which is what Word
    # -- resolves them to at render time). Sharing the same ``ms_api_value``
    # -- as LEFT/RIGHT makes ``member is WD_ALIGN_PARAGRAPH.LEFT`` and equality
    # -- comparisons against the legacy names continue to work for callers
    # -- that have branched on alignment before `start`/`end` existed.
    START = (0, "start", "Left-aligned (logical start; alias of LEFT).")
    """Left-aligned (logical start; alias of LEFT)."""

    END = (2, "end", "Right-aligned (logical end; alias of RIGHT).")
    """Right-aligned (logical end; alias of RIGHT)."""

    DISTRIBUTE = (
        4,
        "distribute",
        "Paragraph characters are distributed to fill entire width of paragraph.",
    )
    """Paragraph characters are distributed to fill entire width of paragraph."""

    JUSTIFY_MED = (
        5,
        "mediumKashida",
        "Justified with a medium character compression ratio.",
    )
    """Justified with a medium character compression ratio."""

    JUSTIFY_HI = (
        7,
        "highKashida",
        "Justified with a high character compression ratio.",
    )
    """Justified with a high character compression ratio."""

    JUSTIFY_LOW = (8, "lowKashida", "Justified with a low character compression ratio.")
    """Justified with a low character compression ratio."""

    THAI_JUSTIFY = (
        9,
        "thaiDistribute",
        "Justified according to Thai formatting layout.",
    )
    """Justified according to Thai formatting layout."""

    @classmethod
    def from_xml(cls, xml_value: str | None) -> "WD_PARAGRAPH_ALIGNMENT":
        """Return the enum member for ``xml_value``.

        Overridden to map the bidi-aware ``start``/``end`` XML values onto the
        :attr:`LEFT` / :attr:`RIGHT` members. Word writes ``start`` and
        ``end`` on RTL-aware paragraph alignment; mapping them to LEFT/RIGHT
        preserves compatibility with existing equality checks.
        """
        if xml_value == "start":
            return cls.LEFT
        if xml_value == "end":
            return cls.RIGHT
        return super().from_xml(xml_value)


WD_ALIGN_PARAGRAPH = WD_PARAGRAPH_ALIGNMENT


class WD_BREAK_TYPE(enum.Enum):
    """Corresponds to WdBreakType enumeration.

    http://msdn.microsoft.com/en-us/library/office/ff195905.aspx.
    """

    COLUMN = 8
    LINE = 6
    LINE_CLEAR_LEFT = 9
    LINE_CLEAR_RIGHT = 10
    LINE_CLEAR_ALL = 11  # -- added for consistency, not in MS version --
    PAGE = 7
    SECTION_CONTINUOUS = 3
    SECTION_EVEN_PAGE = 4
    SECTION_NEXT_PAGE = 2
    SECTION_ODD_PAGE = 5
    TEXT_WRAPPING = 11


WD_BREAK = WD_BREAK_TYPE


class WD_COLOR_INDEX(BaseXmlEnum):
    """Specifies a standard preset color to apply.

    Used for font highlighting and perhaps other applications.

    * MS API name: `WdColorIndex`
    * URL: https://msdn.microsoft.com/EN-US/library/office/ff195343.aspx
    """

    INHERITED = (-1, None, "Color is inherited from the style hierarchy.")
    """Color is inherited from the style hierarchy."""

    AUTO = (0, "default", "Automatic color. Default; usually black.")
    """Automatic color. Default; usually black."""

    BLACK = (1, "black", "Black color.")
    """Black color."""

    BLUE = (2, "blue", "Blue color")
    """Blue color"""

    BRIGHT_GREEN = (4, "green", "Bright green color.")
    """Bright green color."""

    DARK_BLUE = (9, "darkBlue", "Dark blue color.")
    """Dark blue color."""

    DARK_RED = (13, "darkRed", "Dark red color.")
    """Dark red color."""

    DARK_YELLOW = (14, "darkYellow", "Dark yellow color.")
    """Dark yellow color."""

    GRAY_25 = (16, "lightGray", "25% shade of gray color.")
    """25% shade of gray color."""

    GRAY_50 = (15, "darkGray", "50% shade of gray color.")
    """50% shade of gray color."""

    GREEN = (11, "darkGreen", "Green color.")
    """Green color."""

    PINK = (5, "magenta", "Pink color.")
    """Pink color."""

    RED = (6, "red", "Red color.")
    """Red color."""

    TEAL = (10, "darkCyan", "Teal color.")
    """Teal color."""

    TURQUOISE = (3, "cyan", "Turquoise color.")
    """Turquoise color."""

    VIOLET = (12, "darkMagenta", "Violet color.")
    """Violet color."""

    WHITE = (8, "white", "White color.")
    """White color."""

    YELLOW = (7, "yellow", "Yellow color.")
    """Yellow color."""


WD_COLOR = WD_COLOR_INDEX


class WD_LINE_SPACING(BaseXmlEnum):
    """Specifies a line spacing format to be applied to a paragraph.

    Example::

        from docx.enum.text import WD_LINE_SPACING

        paragraph = document.add_paragraph()
        paragraph.line_spacing_rule = WD_LINE_SPACING.EXACTLY


    MS API name: `WdLineSpacing`

    URL: http://msdn.microsoft.com/en-us/library/office/ff844910.aspx
    """

    SINGLE = (0, "UNMAPPED", "Single spaced (default).")
    """Single spaced (default)."""

    ONE_POINT_FIVE = (1, "UNMAPPED", "Space-and-a-half line spacing.")
    """Space-and-a-half line spacing."""

    DOUBLE = (2, "UNMAPPED", "Double spaced.")
    """Double spaced."""

    AT_LEAST = (
        3,
        "atLeast",
        "Minimum line spacing is specified amount. Amount is specified separately.",
    )
    """Minimum line spacing is specified amount. Amount is specified separately."""

    EXACTLY = (
        4,
        "exact",
        "Line spacing is exactly specified amount. Amount is specified separately.",
    )
    """Line spacing is exactly specified amount. Amount is specified separately."""

    MULTIPLE = (
        5,
        "auto",
        "Line spacing is specified as multiple of line heights. Changing font size"
        " will change line spacing proportionately.",
    )
    """Line spacing is specified as multiple of line heights. Changing font size will
       change the line spacing proportionately."""


class WD_TAB_ALIGNMENT(BaseXmlEnum):
    """Specifies the tab stop alignment to apply.

    MS API name: `WdTabAlignment`

    URL: https://msdn.microsoft.com/EN-US/library/office/ff195609.aspx
    """

    LEFT = (0, "left", "Left-aligned.")
    """Left-aligned."""

    CENTER = (1, "center", "Center-aligned.")
    """Center-aligned."""

    RIGHT = (2, "right", "Right-aligned.")
    """Right-aligned."""

    DECIMAL = (3, "decimal", "Decimal-aligned.")
    """Decimal-aligned."""

    BAR = (4, "bar", "Bar-aligned.")
    """Bar-aligned."""

    LIST = (6, "list", "List-aligned. (deprecated)")
    """List-aligned. (deprecated)"""

    CLEAR = (101, "clear", "Clear an inherited tab stop.")
    """Clear an inherited tab stop."""

    END = (102, "end", "Right-aligned.  (deprecated)")
    """Right-aligned.  (deprecated)"""

    NUM = (103, "num", "Left-aligned.  (deprecated)")
    """Left-aligned.  (deprecated)"""

    START = (104, "start", "Left-aligned.  (deprecated)")
    """Left-aligned.  (deprecated)"""


class WD_TAB_LEADER(BaseXmlEnum):
    """Specifies the character to use as the leader with formatted tabs.

    MS API name: `WdTabLeader`

    URL: https://msdn.microsoft.com/en-us/library/office/ff845050.aspx
    """

    SPACES = (0, "none", "Spaces. Default.")
    """Spaces. Default."""

    DOTS = (1, "dot", "Dots.")
    """Dots."""

    DASHES = (2, "hyphen", "Dashes.")
    """Dashes."""

    LINES = (3, "underscore", "Double lines.")
    """Double lines."""

    HEAVY = (4, "heavy", "A heavy line.")
    """A heavy line."""

    MIDDLE_DOT = (5, "middleDot", "A vertically-centered dot.")
    """A vertically-centered dot."""


class WD_BORDER_STYLE(BaseXmlEnum):
    """Specifies the style of a paragraph border.

    Example::

        from docx.enum.text import WD_BORDER_STYLE

        paragraph = document.add_paragraph()
        paragraph.paragraph_format.borders.bottom.style = WD_BORDER_STYLE.SINGLE

    .. versionadded:: 2026.05.0
    """

    NIL = (0, "nil", "No border.")
    """No border."""

    NONE = (1, "none", "No border.")
    """No border."""

    SINGLE = (2, "single", "A single line.")
    """A single line."""

    THICK = (3, "thick", "A single thick line.")
    """A single thick line."""

    DOUBLE = (4, "double", "A double line.")
    """A double line."""

    DOTTED = (5, "dotted", "A dotted line.")
    """A dotted line."""

    DASHED = (6, "dashed", "A dashed line.")
    """A dashed line."""

    DOT_DASH = (7, "dotDash", "An alternating dot-dash line.")
    """An alternating dot-dash line."""

    DOT_DOT_DASH = (8, "dotDotDash", "An alternating dot-dot-dash line.")
    """An alternating dot-dot-dash line."""

    TRIPLE = (9, "triple", "A triple line.")
    """A triple line."""

    THIN_THICK_SMALL_GAP = (10, "thinThickSmallGap", "A thin-thick line with a small gap.")
    """A thin-thick line with a small gap."""

    THICK_THIN_SMALL_GAP = (11, "thickThinSmallGap", "A thick-thin line with a small gap.")
    """A thick-thin line with a small gap."""

    THIN_THICK_THIN_SMALL_GAP = (
        12,
        "thinThickThinSmallGap",
        "A thin-thick-thin line with a small gap.",
    )
    """A thin-thick-thin line with a small gap."""

    THIN_THICK_MEDIUM_GAP = (13, "thinThickMediumGap", "A thin-thick line with a medium gap.")
    """A thin-thick line with a medium gap."""

    THICK_THIN_MEDIUM_GAP = (14, "thickThinMediumGap", "A thick-thin line with a medium gap.")
    """A thick-thin line with a medium gap."""

    THIN_THICK_THIN_MEDIUM_GAP = (
        15,
        "thinThickThinMediumGap",
        "A thin-thick-thin line with a medium gap.",
    )
    """A thin-thick-thin line with a medium gap."""

    THIN_THICK_LARGE_GAP = (16, "thinThickLargeGap", "A thin-thick line with a large gap.")
    """A thin-thick line with a large gap."""

    THICK_THIN_LARGE_GAP = (17, "thickThinLargeGap", "A thick-thin line with a large gap.")
    """A thick-thin line with a large gap."""

    THIN_THICK_THIN_LARGE_GAP = (
        18,
        "thinThickThinLargeGap",
        "A thin-thick-thin line with a large gap.",
    )
    """A thin-thick-thin line with a large gap."""

    WAVE = (19, "wave", "A wavy line.")
    """A wavy line."""

    DOUBLE_WAVE = (20, "doubleWave", "A double wavy line.")
    """A double wavy line."""

    DASH_SMALL_GAP = (21, "dashSmallGap", "A dashed line with a small gap.")
    """A dashed line with a small gap."""

    DASH_DOT_STROKED = (22, "dashDotStroked", "A dash-dot stroked line.")
    """A dash-dot stroked line."""

    THREE_D_EMBOSS = (23, "threeDEmboss", "A 3D embossed line.")
    """A 3D embossed line."""

    THREE_D_ENGRAVE = (24, "threeDEngrave", "A 3D engraved line.")
    """A 3D engraved line."""

    OUTSET = (25, "outset", "An outset line.")
    """An outset line."""

    INSET = (26, "inset", "An inset line.")
    """An inset line."""



class WD_UNDERLINE(BaseXmlEnum):
    """Specifies the style of underline applied to a run of characters.

    MS API name: `WdUnderline`

    URL: http://msdn.microsoft.com/en-us/library/office/ff822388.aspx
    """

    INHERITED = (-1, None, "Inherit underline setting from containing paragraph.")
    """Inherit underline setting from containing paragraph."""

    NONE = (
        0,
        "none",
        "No underline.\n\nThis setting overrides any inherited underline value, so can"
        " be used to remove underline from a run that inherits underlining from its"
        " containing paragraph. Note this is not the same as assigning |None| to"
        " Run.underline. |None| is a valid assignment value, but causes the run to"
        " inherit its underline value. Assigning `WD_UNDERLINE.NONE` causes"
        " underlining to be unconditionally turned off.",
    )
    """No underline.

    This setting overrides any inherited underline value, so can be used to remove
    underline from a run that inherits underlining from its containing paragraph. Note
    this is not the same as assigning |None| to Run.underline. |None| is a valid
    assignment value, but causes the run to inherit its underline value. Assigning
    ``WD_UNDERLINE.NONE`` causes underlining to be unconditionally turned off.
    """

    SINGLE = (
        1,
        "single",
        "A single line.\n\nNote that this setting is write-only in the sense that"
        " |True| (rather than `WD_UNDERLINE.SINGLE`) is returned for a run having"
        " this setting.",
    )
    """A single line.

    Note that this setting is write-only in the sense that |True|
    (rather than ``WD_UNDERLINE.SINGLE``) is returned for a run having this setting.
    """

    WORDS = (2, "words", "Underline individual words only.")
    """Underline individual words only."""

    DOUBLE = (3, "double", "A double line.")
    """A double line."""

    DOTTED = (4, "dotted", "Dots.")
    """Dots."""

    THICK = (6, "thick", "A single thick line.")
    """A single thick line."""

    DASH = (7, "dash", "Dashes.")
    """Dashes."""

    DOT_DASH = (9, "dotDash", "Alternating dots and dashes.")
    """Alternating dots and dashes."""

    DOT_DOT_DASH = (10, "dotDotDash", "An alternating dot-dot-dash pattern.")
    """An alternating dot-dot-dash pattern."""

    WAVY = (11, "wave", "A single wavy line.")
    """A single wavy line."""

    DOTTED_HEAVY = (20, "dottedHeavy", "Heavy dots.")
    """Heavy dots."""

    DASH_HEAVY = (23, "dashedHeavy", "Heavy dashes.")
    """Heavy dashes."""

    DOT_DASH_HEAVY = (25, "dashDotHeavy", "Alternating heavy dots and heavy dashes.")
    """Alternating heavy dots and heavy dashes."""

    DOT_DOT_DASH_HEAVY = (
        26,
        "dashDotDotHeavy",
        "An alternating heavy dot-dot-dash pattern.",
    )
    """An alternating heavy dot-dot-dash pattern."""

    WAVY_HEAVY = (27, "wavyHeavy", "A heavy wavy line.")
    """A heavy wavy line."""

    DASH_LONG = (39, "dashLong", "Long dashes.")
    """Long dashes."""

    WAVY_DOUBLE = (43, "wavyDouble", "A double wavy line.")
    """A double wavy line."""

    DASH_LONG_HEAVY = (55, "dashLongHeavy", "Long heavy dashes.")
    """Long heavy dashes."""


class WD_OUTLINELVL(BaseEnum):
    """Specifies the outline level of a paragraph.

    Maps to values 0..9 for heading levels and 10 for body text, matching
    the ``w:pPr/w:outlineLvl/@w:val`` attribute semantics. Aliases such as
    :attr:`LEVEL_1` through :attr:`LEVEL_9` may be used interchangeably with
    the bare integer values. :attr:`BODY_TEXT` is the sentinel ``10``.

    .. versionadded:: 2026.05.0
    """

    LEVEL_1 = (0, "Outline level 1 (e.g. Heading 1).")
    """Outline level 1 (e.g. Heading 1)."""

    LEVEL_2 = (1, "Outline level 2 (e.g. Heading 2).")
    """Outline level 2 (e.g. Heading 2)."""

    LEVEL_3 = (2, "Outline level 3 (e.g. Heading 3).")
    """Outline level 3 (e.g. Heading 3)."""

    LEVEL_4 = (3, "Outline level 4 (e.g. Heading 4).")
    """Outline level 4 (e.g. Heading 4)."""

    LEVEL_5 = (4, "Outline level 5 (e.g. Heading 5).")
    """Outline level 5 (e.g. Heading 5)."""

    LEVEL_6 = (5, "Outline level 6 (e.g. Heading 6).")
    """Outline level 6 (e.g. Heading 6)."""

    LEVEL_7 = (6, "Outline level 7 (e.g. Heading 7).")
    """Outline level 7 (e.g. Heading 7)."""

    LEVEL_8 = (7, "Outline level 8 (e.g. Heading 8).")
    """Outline level 8 (e.g. Heading 8)."""

    LEVEL_9 = (8, "Outline level 9 (e.g. Heading 9).")
    """Outline level 9 (e.g. Heading 9)."""

    LEVEL_10 = (9, "Outline level 10.")
    """Outline level 10."""

    BODY_TEXT = (10, "Body text (no outline level).")
    """Body text (no outline level)."""


class WD_NUMBER_FORMAT(BaseXmlEnum):
    """Specifies a numeric format used for numbering list items, footnotes, or endnotes.

    Maps to ``ST_NumberFormat`` values in OOXML. Used by:

    - ``w:numFmt`` child of ``w:footnotePr`` / ``w:endnotePr`` (footnote and endnote
      numbering style)
    - ``w:numFmt`` child of ``w:lvl`` within ``w:abstractNum`` (list item numbering)

    Only the most common members are exposed; the full OOXML enumeration is large
    and rarely needed. Use :meth:`from_xml` to convert a raw ``w:numFmt/@w:val``
    string to an enumeration member.

    .. versionadded:: 2026.05.0
    """

    DECIMAL = (0, "decimal", "Decimal numbers (1, 2, 3 ...).")
    """Decimal numbers (1, 2, 3 ...)."""

    ARABIC = DECIMAL
    """Alias for :attr:`DECIMAL` (Arabic numerals: 1, 2, 3 ...)."""

    UPPER_ROMAN = (1, "upperRoman", "Uppercase Roman numerals (I, II, III ...).")
    """Uppercase Roman numerals (I, II, III ...)."""

    LOWER_ROMAN = (2, "lowerRoman", "Lowercase Roman numerals (i, ii, iii ...).")
    """Lowercase Roman numerals (i, ii, iii ...)."""

    UPPER_LETTER = (3, "upperLetter", "Uppercase letters (A, B, C ...).")
    """Uppercase letters (A, B, C ...)."""

    LOWER_LETTER = (4, "lowerLetter", "Lowercase letters (a, b, c ...).")
    """Lowercase letters (a, b, c ...)."""

    ORDINAL = (5, "ordinal", "Ordinal numbers (1st, 2nd, 3rd ...).")
    """Ordinal numbers (1st, 2nd, 3rd ...)."""

    CARDINAL_TEXT = (6, "cardinalText", "Cardinal text (One, Two, Three ...).")
    """Cardinal text (One, Two, Three ...)."""

    ORDINAL_TEXT = (7, "ordinalText", "Ordinal text (First, Second ...).")
    """Ordinal text (First, Second ...)."""

    CHICAGO = (
        8,
        "chicago",
        "Chicago Manual of Style footnote marks (*, †, ‡, §).",
    )
    """Chicago Manual of Style footnote marks (*, †, ‡, §)."""

    BULLET = (23, "bullet", "Bullet character (not numbered).")
    """Bullet character (not numbered)."""

    NONE = (255, "none", "No number.")
    """No number."""


class WD_FOOTNOTE_RESTART(BaseXmlEnum):
    """Specifies when footnote numbering restarts.

    Maps to the ``w:numRestart`` child element of ``w:footnotePr``.

    .. versionadded:: 2026.05.0
    """

    CONTINUOUS = (0, "continuous", "Continuous numbering throughout the document.")
    """Continuous numbering throughout the document."""

    EACH_SECTION = (1, "eachSect", "Numbering restarts at the beginning of each section.")
    """Numbering restarts at the beginning of each section."""

    EACH_PAGE = (2, "eachPage", "Numbering restarts at the beginning of each page.")
    """Numbering restarts at the beginning of each page."""


class WD_FOOTNOTE_POSITION(BaseXmlEnum):
    """Specifies the position of footnotes on the page.

    Maps to the ``w:pos`` child element of ``w:footnotePr``.

    .. versionadded:: 2026.05.0
    """

    BOTTOM_OF_PAGE = (0, "pageBottom", "Footnotes appear at the bottom of the page.")
    """Footnotes appear at the bottom of the page."""

    BENEATH_TEXT = (
        1,
        "beneathText",
        "Footnotes appear immediately beneath the body text on the page.",
    )
    """Footnotes appear immediately beneath the body text on the page."""

    END_OF_SECTION = (
        2,
        "sectEnd",
        "Footnotes appear at the end of each section (section-end footnotes).",
    )
    """Footnotes appear at the end of each section (section-end footnotes)."""

    END_OF_DOCUMENT = (
        3,
        "docEnd",
        "Footnotes appear at the end of the document.",
    )
    """Footnotes appear at the end of the document."""


class WD_ENDNOTE_POSITION(BaseXmlEnum):
    """Specifies the position of endnotes in the document.

    Maps to the ``w:pos`` child element of ``w:endnotePr``.

    .. versionadded:: 2026.05.0
    """

    END_OF_DOCUMENT = (0, "docEnd", "Endnotes appear at the end of the document.")
    """Endnotes appear at the end of the document."""

    END_OF_SECTION = (1, "sectEnd", "Endnotes appear at the end of each section.")
    """Endnotes appear at the end of each section."""


class WD_VIEW(BaseXmlEnum):
    """Specifies the preferred view mode for displaying the document.

    Maps to the ``w:val`` attribute of the ``w:view`` child of ``w:settings``.

    .. versionadded:: 2026.05.0
    """

    NONE = (0, "none", "No view mode is specified.")
    """No view mode is specified."""

    PRINT = (1, "print", "Print layout view (Word's default editing view).")
    """Print layout view (Word's default editing view)."""

    OUTLINE = (2, "outline", "Outline view, showing document headings and hierarchy.")
    """Outline view, showing document headings and hierarchy."""

    MASTER_PAGES = (3, "masterPages", "Master-pages (master document) view.")
    """Master-pages (master document) view."""

    NORMAL = (4, "normal", "Normal (draft) view, emphasizing text flow over layout.")
    """Normal (draft) view, emphasizing text flow over layout."""

    WEB = (5, "web", "Web layout view, showing the document as it would appear in a browser.")
    """Web layout view, showing the document as it would appear in a browser."""

    READING = (6, "reading", "Full-screen reading view optimized for reading.")
    """Full-screen reading view optimized for reading."""


class WD_PROTECTION(BaseXmlEnum):
    """Specifies the document-protection editing mode.

    Maps to the ``w:edit`` attribute of ``w:documentProtection`` in the settings
    part. A document protection element with ``w:enforcement="1"`` prevents the
    user from editing the document in ways other than those permitted by the
    selected mode. For example, :attr:`COMMENTS` allows inserting comments but
    not modifying paragraph text.

    .. versionadded:: 2026.05.0
    """

    READ_ONLY = (0, "readOnly", "The document is read-only; no edits are permitted.")
    """The document is read-only; no edits are permitted."""

    COMMENTS = (1, "comments", "Only comments may be inserted or modified.")
    """Only comments may be inserted or modified."""

    TRACKED_CHANGES = (
        2,
        "trackedChanges",
        "Changes are permitted but recorded as tracked revisions.",
    )
    """Changes are permitted but recorded as tracked revisions."""

    FORMS = (3, "forms", "Only form-field content may be edited.")
    """Only form-field content may be edited."""


class WD_FRAME_H_ANCHOR(BaseXmlEnum):
    """Specifies the base from which a text frame's horizontal position is measured.

    Maps to the ``w:hAnchor`` attribute of ``w:framePr``.

    .. versionadded:: 2026.05.0
    """

    TEXT = (0, "text", "Horizontal position is relative to the text of the paragraph.")
    """Horizontal position is relative to the text of the paragraph."""

    MARGIN = (1, "margin", "Horizontal position is relative to the page margin.")
    """Horizontal position is relative to the page margin."""

    PAGE = (2, "page", "Horizontal position is relative to the page edge.")
    """Horizontal position is relative to the page edge."""


class WD_FRAME_V_ANCHOR(BaseXmlEnum):
    """Specifies the base from which a text frame's vertical position is measured.

    Maps to the ``w:vAnchor`` attribute of ``w:framePr``.

    .. versionadded:: 2026.05.0
    """

    TEXT = (0, "text", "Vertical position is relative to the text of the paragraph.")
    """Vertical position is relative to the text of the paragraph."""

    MARGIN = (1, "margin", "Vertical position is relative to the page margin.")
    """Vertical position is relative to the page margin."""

    PAGE = (2, "page", "Vertical position is relative to the page edge.")
    """Vertical position is relative to the page edge."""


class WD_FRAME_WRAP(BaseXmlEnum):
    """Specifies how text wraps around a text frame.

    Maps to the ``w:wrap`` attribute of ``w:framePr``.

    .. versionadded:: 2026.05.0
    """

    AUTO = (0, "auto", "Text wraps around the frame on all sides.")
    """Text wraps around the frame on all sides."""

    NOT_BESIDE = (1, "notBeside", "Text does not wrap beside the frame.")
    """Text does not wrap beside the frame."""

    AROUND = (2, "around", "Text wraps around the frame.")
    """Text wraps around the frame."""

    NONE = (3, "none", "Text does not wrap around the frame.")
    """Text does not wrap around the frame."""

    TIGHT = (4, "tight", "Text wraps tightly around the frame.")
    """Text wraps tightly around the frame."""

    THROUGH = (5, "through", "Text wraps through the frame.")
    """Text wraps through the frame."""


class WD_FRAME_DROP_CAP(BaseXmlEnum):
    """Specifies how a drop-cap text frame is positioned relative to the paragraph.

    Maps to the ``w:dropCap`` attribute of ``w:framePr``.

    .. versionadded:: 2026.05.0
    """

    NONE = (0, "none", "Not a drop-cap frame.")
    """Not a drop-cap frame."""

    DROP = (1, "drop", "Drop-cap frame dropped into the paragraph text.")
    """Drop-cap frame dropped into the paragraph text."""

    MARGIN = (2, "margin", "Drop-cap frame positioned in the margin.")
    """Drop-cap frame positioned in the margin."""


class WD_FRAME_H_ALIGN(BaseXmlEnum):
    """Specifies the horizontal alignment of a text frame.

    Maps to the ``w:xAlign`` attribute of ``w:framePr``.

    .. versionadded:: 2026.05.0
    """

    LEFT = (0, "left", "Frame is left-aligned.")
    """Frame is left-aligned."""

    CENTER = (1, "center", "Frame is center-aligned.")
    """Frame is center-aligned."""

    RIGHT = (2, "right", "Frame is right-aligned.")
    """Frame is right-aligned."""

    INSIDE = (3, "inside", "Frame is aligned to the inside of the page (for facing pages).")
    """Frame is aligned to the inside of the page (for facing pages)."""

    OUTSIDE = (4, "outside", "Frame is aligned to the outside of the page (for facing pages).")
    """Frame is aligned to the outside of the page (for facing pages)."""


class WD_FRAME_V_ALIGN(BaseXmlEnum):
    """Specifies the vertical alignment of a text frame.

    Maps to the ``w:yAlign`` attribute of ``w:framePr``.

    .. versionadded:: 2026.05.0
    """

    INLINE = (0, "inline", "Frame is positioned inline with the surrounding text.")
    """Frame is positioned inline with the surrounding text."""

    TOP = (1, "top", "Frame is top-aligned.")
    """Frame is top-aligned."""

    CENTER = (2, "center", "Frame is center-aligned vertically.")
    """Frame is center-aligned vertically."""

    BOTTOM = (3, "bottom", "Frame is bottom-aligned.")
    """Frame is bottom-aligned."""

    INSIDE = (4, "inside", "Frame is aligned to the inside of the page (for facing pages).")
    """Frame is aligned to the inside of the page (for facing pages)."""

    OUTSIDE = (5, "outside", "Frame is aligned to the outside of the page (for facing pages).")
    """Frame is aligned to the outside of the page (for facing pages)."""


class WD_BUILDING_BLOCK_GALLERY(BaseXmlEnum):
    """Specifies the Word building-block gallery a ``w:docPart`` belongs to.

    Maps to the ``w:val`` attribute of ``w:docPart/w:docPartPr/w:category/
    w:gallery``. The enumeration covers the galleries Word ships with plus a
    handful of "custom" buckets; uncommon or vendor-specific values land on
    :attr:`OTHER`. Use :meth:`from_xml_safe` to round-trip a value, returning
    |None| rather than raising when the string is not recognized.

    Example::

        from docx.enum.text import WD_BUILDING_BLOCK_GALLERY

        gallery = WD_BUILDING_BLOCK_GALLERY.from_xml_safe("quickParts")
        assert gallery is WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS

    .. versionadded:: 2026.05.0
    """

    PLACEHOLDER = (0, "placeholder", "Placeholder gallery.")
    """Placeholder gallery."""

    ANY = (1, "any", "Matches any gallery.")
    """Matches any gallery."""

    DEFAULT = (2, "default", "Default gallery.")
    """Default gallery."""

    DOC_PARTS = (3, "docParts", "Generic document-parts gallery.")
    """Generic document-parts gallery."""

    COVER_PAGES = (4, "coverPg", "Cover-page gallery.")
    """Cover-page gallery."""

    EQUATIONS = (5, "eq", "Equation gallery.")
    """Equation gallery."""

    FOOTERS = (6, "ftrs", "Footer gallery.")
    """Footer gallery."""

    HEADERS = (7, "hdrs", "Header gallery.")
    """Header gallery."""

    PAGE_NUMBERS = (8, "pgNum", "Page-number gallery.")
    """Page-number gallery."""

    TABLES = (9, "tbls", "Table gallery.")
    """Table gallery."""

    WATERMARKS = (10, "watermarks", "Watermark gallery.")
    """Watermark gallery."""

    AUTO_TEXT = (11, "autoTxt", "AutoText gallery.")
    """AutoText gallery."""

    TEXT_BOXES = (12, "txtBox", "Text-box gallery.")
    """Text-box gallery."""

    PAGE_NUMBERS_BOTTOM = (13, "pgNumB", "Page-number (bottom of page) gallery.")
    """Page-number (bottom of page) gallery."""

    PAGE_NUMBERS_TOP = (14, "pgNumT", "Page-number (top of page) gallery.")
    """Page-number (top of page) gallery."""

    BIBLIOGRAPHIES = (15, "bib", "Bibliography gallery.")
    """Bibliography gallery."""

    QUICK_PARTS = (16, "quickParts", "Quick Parts gallery.")
    """Quick Parts gallery."""

    CUSTOM_QUICK_PARTS = (17, "custQuickParts", "Custom Quick Parts gallery.")
    """Custom Quick Parts gallery."""

    CUSTOM_COVER_PAGES = (18, "custCoverPg", "Custom cover-page gallery.")
    """Custom cover-page gallery."""

    CUSTOM_EQUATIONS = (19, "custEq", "Custom equation gallery.")
    """Custom equation gallery."""

    CUSTOM_FOOTERS = (20, "custFtrs", "Custom footer gallery.")
    """Custom footer gallery."""

    CUSTOM_HEADERS = (21, "custHdrs", "Custom header gallery.")
    """Custom header gallery."""

    CUSTOM_PAGE_NUMBERS = (22, "custPgNum", "Custom page-number gallery.")
    """Custom page-number gallery."""

    CUSTOM_TABLES = (23, "custTbls", "Custom table gallery.")
    """Custom table gallery."""

    CUSTOM_WATERMARKS = (24, "custWatermarks", "Custom watermark gallery.")
    """Custom watermark gallery."""

    CUSTOM_AUTO_TEXT = (25, "custAutoTxt", "Custom AutoText gallery.")
    """Custom AutoText gallery."""

    CUSTOM_TEXT_BOXES = (26, "custTxtBox", "Custom text-box gallery.")
    """Custom text-box gallery."""

    CUSTOM_PAGE_NUMBERS_BOTTOM = (
        27,
        "custPgNumB",
        "Custom page-number (bottom of page) gallery.",
    )
    """Custom page-number (bottom of page) gallery."""

    CUSTOM_PAGE_NUMBERS_TOP = (
        28,
        "custPgNumT",
        "Custom page-number (top of page) gallery.",
    )
    """Custom page-number (top of page) gallery."""

    CUSTOM_BIBLIOGRAPHIES = (29, "custBib", "Custom bibliography gallery.")
    """Custom bibliography gallery."""

    CUSTOM_1 = (30, "custom1", "Generic custom gallery 1.")
    """Generic custom gallery 1."""

    CUSTOM_2 = (31, "custom2", "Generic custom gallery 2.")
    """Generic custom gallery 2."""

    CUSTOM_3 = (32, "custom3", "Generic custom gallery 3.")
    """Generic custom gallery 3."""

    CUSTOM_4 = (33, "custom4", "Generic custom gallery 4.")
    """Generic custom gallery 4."""

    CUSTOM_5 = (34, "custom5", "Generic custom gallery 5.")
    """Generic custom gallery 5."""

    @classmethod
    def from_xml_safe(
        cls, xml_value: str | None
    ) -> WD_BUILDING_BLOCK_GALLERY | None:
        """Return the enum member for `xml_value`, or |None| when unknown.

        Mirrors :meth:`BaseXmlEnum.from_xml` but never raises — callers that
        want permissive decoding of building-block galleries should use this
        entry point rather than the strict one.

        .. versionadded:: 2026.05.0
        """
        if xml_value is None:
            return None
        for member in cls:
            if member.xml_value == xml_value:
                return member
        return None


class WD_MAIL_MERGE_TYPE(BaseXmlEnum):
    """Specifies the mail-merge main document type.

    Maps to ``w:mailMerge/w:mainDocumentType/@w:val``.

    .. versionadded:: 2026.05.0
    """

    CATALOG = (0, "catalog", "Catalog-style merge (all records on one page).")
    """Catalog-style merge (all records on one page)."""

    ENVELOPES = (1, "envelopes", "Envelope printing merge.")
    """Envelope printing merge."""

    MAILING_LABELS = (2, "mailingLabels", "Mailing-label printing merge.")
    """Mailing-label printing merge."""

    FORM_LETTERS = (3, "formLetters", "Form-letter merge (one letter per record).")
    """Form-letter merge (one letter per record)."""

    EMAIL = (4, "email", "Email-message merge.")
    """Email-message merge."""

    FAX = (5, "fax", "Fax merge.")
    """Fax merge."""


class WD_MAIL_MERGE_DESTINATION(BaseXmlEnum):
    """Specifies where merged output is sent.

    Maps to ``w:mailMerge/w:destination/@w:val``.

    .. versionadded:: 2026.05.0
    """

    NEW_DOCUMENT = (0, "newDocument", "Produce a new Word document containing the merged output.")
    """Produce a new Word document containing the merged output."""

    PRINTER = (1, "printer", "Send output directly to the printer.")
    """Send output directly to the printer."""

    EMAIL = (2, "email", "Email each merged record.")
    """Email each merged record."""

    FAX = (3, "fax", "Fax each merged record.")
    """Fax each merged record."""


class WD_MAIL_MERGE_DATA_TYPE(BaseXmlEnum):
    """Specifies the data-source kind for a mail merge.

    Maps to ``w:mailMerge/w:dataType/@w:val``.

    .. versionadded:: 2026.05.0
    """

    TEXT_FILE = (0, "textFile", "Delimited text file (CSV / TSV).")
    """Delimited text file (CSV / TSV)."""

    DATABASE = (1, "database", "Microsoft Access or similar database.")
    """Microsoft Access or similar database."""

    SPREADSHEET = (2, "spreadsheet", "Excel spreadsheet.")
    """Excel spreadsheet."""

    QUERY = (3, "query", "Word query file.")
    """Word query file."""

    ODBC = (4, "odbc", "ODBC-connected data source.")
    """ODBC-connected data source."""

    NATIVE = (5, "native", "Native Word data source.")
    """Native Word data source."""


# -- Alias matching the ECMA "mainDocumentType" field name. ``WD_MAIL_MERGE_TYPE``
# -- (above) is the original, shorter name; ``WD_MAIL_MERGE_DOCUMENT_TYPE`` is the
# -- long form preferred by the Word object model and by recent roadmap work. Both
# -- names resolve to the same enum class, so existing code and tests continue to
# -- work unchanged.
WD_MAIL_MERGE_DOCUMENT_TYPE = WD_MAIL_MERGE_TYPE


class WD_ODSO_TYPE(BaseXmlEnum):
    """Specifies the ODSO data-source category.

    Maps to ``w:mailMerge/w:odso/w:type/@w:val``. Word uses this enum to tag
    the underlying storage technology for the data source manifest — whether
    it's a local database file, an Outlook address book, a legacy Office
    data source, or a native one-off table. python-docx preserves the value
    verbatim on round-trip; it doesn't drive any merge execution.

    .. versionadded:: 2026.05.10
    """

    DATABASE = (0, "database", "Local database file (Access, SQL, etc.).")
    """Local database file (Access, SQL, etc.)."""

    ADDRESS_BOOK = (1, "addressBook", "Outlook / Exchange address book.")
    """Outlook / Exchange address book."""

    DOCUMENT1 = (2, "document1", "Word document data source, format 1.")
    """Word document data source, format 1."""

    DOCUMENT2 = (3, "document2", "Word document data source, format 2.")
    """Word document data source, format 2."""

    TEXT = (4, "text", "Delimited text file.")
    """Delimited text file."""

    EMAIL = (5, "email", "Email data source.")
    """Email data source."""

    NATIVE = (6, "native", "Native Word data source.")
    """Native Word data source."""

    LEGACY = (7, "legacy", "Legacy (pre-Office 2007) data source.")
    """Legacy (pre-Office 2007) data source."""

    MASTER = (8, "master", "Master / header-less data source.")
    """Master / header-less data source."""
