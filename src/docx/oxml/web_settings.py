"""Custom element classes related to the ``word/webSettings.xml`` part."""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from docx.oxml.simpletypes import ST_OnOff, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.shared import CT_OnOff


class CT_Encoding(BaseOxmlElement):
    """``<w:encoding>`` element within ``w:webSettings``.

    Records the text encoding used when the document is saved as a web
    page. The value is a free-form string (e.g. ``"utf-8"``).
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_OptimizeForBrowser(BaseOxmlElement):
    """``<w:optimizeForBrowser>`` element within ``w:webSettings``.

    Carries a ST_OnOff ``w:val`` attribute defaulting to ``True`` and an
    optional ``w:target`` attribute naming a target browser.
    """

    val: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_OnOff, default=True
    )
    target: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:target", ST_String
    )


class CT_Frameset(BaseOxmlElement):
    """``<w:frameset>`` element within ``w:webSettings``.

    The outer frameset of a framed HTML save. Contains an optional size
    descriptor, a splitbar, a frame layout, a title, and any number of
    nested ``<w:frameset>`` or ``<w:frame>`` children. This module
    models only the attributes python-docx needs to round-trip and enumerate
    the direct child ``<w:frame>`` entries.
    """

    frame_lst: list[CT_Frame]
    frameset_lst: list[CT_Frameset]

    frame = ZeroOrMore("w:frame")
    frameset = ZeroOrMore("w:frameset")


class CT_Frame(BaseOxmlElement):
    """``<w:frame>`` element within ``w:frameset``.

    Describes a single HTML frame (size, title, source, scrollbar
    policy). Only the readable children needed for round-trip and
    enumeration are modelled.
    """

    get_or_add_name: Callable[[], CT_Encoding]
    _remove_name: Callable[[], None]
    get_or_add_sourceFileName: Callable[[], BaseOxmlElement]
    _remove_sourceFileName: Callable[[], None]

    name: CT_Encoding | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:name",
        successors=(
            "w:title",
            "w:longDesc",
            "w:sourceFileName",
            "w:marW",
            "w:marH",
            "w:scrollbar",
            "w:noResizeAllowed",
            "w:linkedToFile",
        ),
    )


class CT_WebSettings(BaseOxmlElement):
    """``<w:webSettings>`` element, root of the web-settings part.

    Holds document-level settings relating to saving as a web page. The
    full ECMA-376 Part 1 §17.15.1.121 child sequence is modelled, although
    the high-level proxy only surfaces a subset. Unmodelled children are
    preserved bytewise on round-trip because lxml keeps unknown children
    in element order.
    """

    get_or_add_frameset: Callable[[], CT_Frameset]
    _remove_frameset: Callable[[], None]
    get_or_add_encoding: Callable[[], CT_Encoding]
    _remove_encoding: Callable[[], None]
    get_or_add_optimizeForBrowser: Callable[[], CT_OptimizeForBrowser]
    _remove_optimizeForBrowser: Callable[[], None]
    get_or_add_relyOnVML: Callable[[], "CT_OnOff"]
    _remove_relyOnVML: Callable[[], None]
    get_or_add_allowPNG: Callable[[], "CT_OnOff"]
    _remove_allowPNG: Callable[[], None]
    get_or_add_doNotSaveAsSingleFile: Callable[[], "CT_OnOff"]
    _remove_doNotSaveAsSingleFile: Callable[[], None]

    _tag_seq = (
        "w:frameset",
        "w:divs",
        "w:encoding",
        "w:optimizeForBrowser",
        "w:relyOnVML",
        "w:allowPNG",
        "w:doNotRelyOnCSS",
        "w:doNotSaveAsSingleFile",
        "w:doNotOrganizeInFolder",
        "w:doNotUseLongFileNames",
        "w:pixelsPerInch",
        "w:targetScreenSz",
        "w:saveSmartTagsAsXml",
    )

    frameset: CT_Frameset | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:frameset", successors=_tag_seq[1:]
    )
    encoding: CT_Encoding | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:encoding", successors=_tag_seq[3:]
    )
    optimizeForBrowser: CT_OptimizeForBrowser | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:optimizeForBrowser", successors=_tag_seq[4:]
    )
    relyOnVML: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:relyOnVML", successors=_tag_seq[5:]
    )
    allowPNG: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:allowPNG", successors=_tag_seq[6:]
    )
    doNotSaveAsSingleFile: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:doNotSaveAsSingleFile", successors=_tag_seq[8:]
    )
    del _tag_seq

    # -- value-attribute helpers -----------------------------------------------

    @property
    def encoding_val(self) -> str | None:
        """Value of ``w:encoding/@w:val`` or ``None`` if not present."""
        encoding = self.encoding
        if encoding is None:
            return None
        return encoding.val

    @property
    def optimizeForBrowser_val(self) -> bool:
        """True when ``w:optimizeForBrowser`` is present and not explicitly disabled."""
        optimizeForBrowser = self.optimizeForBrowser
        if optimizeForBrowser is None:
            return False
        return optimizeForBrowser.val

    @optimizeForBrowser_val.setter
    def optimizeForBrowser_val(self, value: bool | None):
        if value is None or value is False:
            self._remove_optimizeForBrowser()
            return
        self.get_or_add_optimizeForBrowser().val = value

    @property
    def relyOnVML_val(self) -> bool:
        """True when ``w:relyOnVML`` is present and not explicitly disabled."""
        relyOnVML = self.relyOnVML
        if relyOnVML is None:
            return False
        return relyOnVML.val

    @relyOnVML_val.setter
    def relyOnVML_val(self, value: bool | None):
        if value is None or value is False:
            self._remove_relyOnVML()
            return
        self.get_or_add_relyOnVML().val = value

    @property
    def allowPNG_val(self) -> bool:
        """True when ``w:allowPNG`` is present and not explicitly disabled."""
        allowPNG = self.allowPNG
        if allowPNG is None:
            return False
        return allowPNG.val

    @allowPNG_val.setter
    def allowPNG_val(self, value: bool | None):
        if value is None or value is False:
            self._remove_allowPNG()
            return
        self.get_or_add_allowPNG().val = value

    @property
    def doNotSaveAsSingleFile_val(self) -> bool:
        """True when ``w:doNotSaveAsSingleFile`` is present and not disabled."""
        doNotSaveAsSingleFile = self.doNotSaveAsSingleFile
        if doNotSaveAsSingleFile is None:
            return False
        return doNotSaveAsSingleFile.val

    @doNotSaveAsSingleFile_val.setter
    def doNotSaveAsSingleFile_val(self, value: bool | None):
        if value is None or value is False:
            self._remove_doNotSaveAsSingleFile()
            return
        self.get_or_add_doNotSaveAsSingleFile().val = value
