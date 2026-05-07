"""Enumerations related to the main document in WordprocessingML files."""

from .base import BaseXmlEnum


class WD_BORDER_DISPLAY(BaseXmlEnum):
    """Specifies on which pages a page-border is displayed.

    Maps to the ``w:display`` attribute of the ``w:pgBorders`` element.

    Example::

        from docx.enum.section import WD_BORDER_DISPLAY

        section = document.sections[0]
        section.page_borders.display = WD_BORDER_DISPLAY.FIRST_PAGE

    .. versionadded:: 2026.05.0
    """

    ALL_PAGES = (0, "allPages", "Border is displayed on every page.")
    """Border is displayed on every page."""

    FIRST_PAGE = (1, "firstPage", "Border is displayed only on the first page.")
    """Border is displayed only on the first page."""

    NOT_FIRST_PAGE = (2, "notFirstPage", "Border is displayed on every page except the first.")
    """Border is displayed on every page except the first."""


class WD_BORDER_OFFSET_FROM(BaseXmlEnum):
    """Specifies the reference point used to position a page-border.

    Maps to the ``w:offsetFrom`` attribute of the ``w:pgBorders`` element.

    Example::

        from docx.enum.section import WD_BORDER_OFFSET_FROM

        section = document.sections[0]
        section.page_borders.offset_from = WD_BORDER_OFFSET_FROM.PAGE

    .. versionadded:: 2026.05.0
    """

    TEXT = (0, "text", "Border is positioned relative to the text extents.")
    """Border is positioned relative to the text extents."""

    PAGE = (1, "page", "Border is positioned relative to the page edge.")
    """Border is positioned relative to the page edge."""


class WD_LINE_NUMBERING_RESTART(BaseXmlEnum):
    """Specifies when line numbering restarts within a section.

    Maps to the ``w:restart`` attribute of the ``w:lnNumType`` element.

    Example::

        from docx.enum.section import WD_LINE_NUMBERING_RESTART

        section = document.sections[0]
        section.set_line_numbering(
            count_by=1, restart=WD_LINE_NUMBERING_RESTART.NEW_PAGE
        )

    .. versionadded:: 2026.05.0
    """

    CONTINUOUS = (0, "continuous", "Line numbering continues from the previous section.")
    """Line numbering continues from the previous section."""

    NEW_SECTION = (1, "newSection", "Line numbering restarts at the beginning of each section.")
    """Line numbering restarts at the beginning of each section."""

    NEW_PAGE = (2, "newPage", "Line numbering restarts at the beginning of each page.")
    """Line numbering restarts at the beginning of each page."""


class WD_CHAPTER_SEPARATOR(BaseXmlEnum):
    """Separator-glyph style between chapter number and page number.

    Maps to the ``w:chapSep`` attribute of the ``w:pgNumType`` element.

    Example::

        from docx.enum.section import WD_CHAPTER_SEPARATOR

        section = document.sections[0]
        section.set_page_numbering(chapter_style=1, chapter_separator=WD_CHAPTER_SEPARATOR.EM_DASH)

    .. versionadded:: 2026.05.3
    """

    HYPHEN = (0, "hyphen", "ASCII hyphen-minus separator.")
    """ASCII hyphen-minus separator (U+002D)."""

    PERIOD = (1, "period", "Period (full-stop) separator.")
    """Period separator (U+002E)."""

    COLON = (2, "colon", "Colon separator.")
    """Colon separator (U+003A)."""

    EM_DASH = (3, "emDash", "Em-dash separator.")
    """Em-dash separator (U+2014)."""

    EN_DASH = (4, "enDash", "En-dash separator.")
    """En-dash separator (U+2013)."""


class WD_DOC_GRID_TYPE(BaseXmlEnum):
    """Specifies the type of East Asian document character grid for a section.

    Maps to the ``w:type`` attribute of the ``w:docGrid`` element.

    Example::

        from docx.enum.section import WD_DOC_GRID_TYPE

        section = document.sections[0]
        section.set_document_grid(type=WD_DOC_GRID_TYPE.LINES_AND_CHARS)

    .. versionadded:: 2026.05.0
    """

    DEFAULT = (0, "default", "No document grid is applied.")
    """No document grid is applied."""

    LINES = (1, "lines", "Grid specifies lines per page only.")
    """Grid specifies lines per page only."""

    LINES_AND_CHARS = (
        2,
        "linesAndChars",
        "Grid specifies both lines per page and characters per line.",
    )
    """Grid specifies both lines per page and characters per line."""

    SNAP_TO_CHARS = (
        3,
        "snapToChars",
        "Characters snap to the grid; used when fixed character positions are required.",
    )
    """Characters snap to the grid; used when fixed character positions are required."""


class WD_HEADER_FOOTER_INDEX(BaseXmlEnum):
    """Alias: **WD_HEADER_FOOTER**

    Specifies one of the three possible header/footer definitions for a section.

    For internal use only; not part of the python-docx API.

    MS API name: `WdHeaderFooterIndex`
    URL: https://docs.microsoft.com/en-us/office/vba/api/word.wdheaderfooterindex
    """

    PRIMARY = (1, "default", "Header for odd pages or all if no even header.")
    """Header for odd pages or all if no even header."""

    FIRST_PAGE = (2, "first", "Header for first page of section.")
    """Header for first page of section."""

    EVEN_PAGE = (3, "even", "Header for even pages of recto/verso section.")
    """Header for even pages of recto/verso section."""


WD_HEADER_FOOTER = WD_HEADER_FOOTER_INDEX


class WD_ORIENTATION(BaseXmlEnum):
    """Alias: **WD_ORIENT**

    Specifies the page layout orientation.

    Example::

        from docx.enum.section import WD_ORIENT

        section = document.sections[-1] section.orientation = WD_ORIENT.LANDSCAPE

    MS API name: `WdOrientation`
    MS API URL: http://msdn.microsoft.com/en-us/library/office/ff837902.aspx
    """

    PORTRAIT = (0, "portrait", "Portrait orientation.")
    """Portrait orientation."""

    LANDSCAPE = (1, "landscape", "Landscape orientation.")
    """Landscape orientation."""


WD_ORIENT = WD_ORIENTATION


class WD_VERTICAL_ALIGNMENT(BaseXmlEnum):
    """Specifies the vertical alignment of text in a section.

    Maps to the ``w:val`` attribute of the ``w:vAlign`` child of a ``w:sectPr``
    element (ECMA-376 17.6.22, simple type ``ST_VerticalJc``).

    Example::

        from docx.enum.section import WD_VERTICAL_ALIGNMENT

        section = document.sections[0]
        section.vertical_alignment = WD_VERTICAL_ALIGNMENT.CENTER

    .. versionadded:: 2026.05.0
    """

    TOP = (0, "top", "Text is aligned to the top of the page.")
    """Text is aligned to the top of the page."""

    CENTER = (1, "center", "Text is centered vertically on the page.")
    """Text is centered vertically on the page."""

    BOTH = (
        2,
        "both",
        "Text is justified vertically so each line is evenly spaced between the top"
        " and bottom margins.",
    )
    """Text is justified vertically so each line is evenly spaced between the top and
    bottom margins."""

    BOTTOM = (3, "bottom", "Text is aligned to the bottom of the page.")
    """Text is aligned to the bottom of the page."""


class WD_SECTION_START(BaseXmlEnum):
    """Alias: **WD_SECTION**

    Specifies the start type of a section break.

    Example::

        from docx.enum.section import WD_SECTION

        section = document.sections[0] section.start_type = WD_SECTION.NEW_PAGE

    MS API name: `WdSectionStart`
    MS API URL: http://msdn.microsoft.com/en-us/library/office/ff840975.aspx
    """

    CONTINUOUS = (0, "continuous", "Continuous section break.")
    """Continuous section break."""

    NEW_COLUMN = (1, "nextColumn", "New column section break.")
    """New column section break."""

    NEW_PAGE = (2, "nextPage", "New page section break.")
    """New page section break."""

    EVEN_PAGE = (3, "evenPage", "Even pages section break.")
    """Even pages section break."""

    ODD_PAGE = (4, "oddPage", "Section begins on next odd page.")
    """Section begins on next odd page."""


WD_SECTION = WD_SECTION_START
