"""Custom element classes related to the ``word/webSettings.xml`` part."""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from docx.oxml.simpletypes import ST_OnOff, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne

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


class CT_WebSettings(BaseOxmlElement):
    """``<w:webSettings>`` element, root of the web-settings part.

    Holds document-level settings relating to saving as a web page. Only
    the child elements exposed by the thin Python proxy are modelled here;
    the ``w:frameset`` and ``w:divs`` children are structural and not
    included because this project does not expose frameset/div semantics.
    The remaining children are each optional and ST_OnOff flag wrappers
    or the string-valued ``w:encoding``.
    """

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
        "w:doNotSaveAsSingleFile",
        "w:doNotOrganizeInFolder",
        "w:doNotRelyOnCSS",
        "w:doNotUseLongFileNames",
        "w:pixelsPerInch",
        "w:targetScreenSz",
        "w:saveSmartTagsAsXml",
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
        "w:doNotSaveAsSingleFile", successors=_tag_seq[7:]
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
