"""|WebSettings| proxy for ``word/webSettings.xml``.

Provides read-only-ish access to a small subset of the OOXML web-settings
part: encoding, "optimize for browser", "allow PNG", and
"do not save as single file". The remaining schema children
(``w:frameset``, ``w:divs``, and several rarely-used flags) are not
exposed because they are unlikely to be useful from Python code and
add significant surface area.

Access via :attr:`docx.document.Document.web_settings`, which returns a
:class:`WebSettings` instance when the document has a ``webSettings``
relationship, or |None| otherwise.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.shared import ElementProxy

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.web_settings import CT_WebSettings
    from docx.oxml.xmlchemy import BaseOxmlElement


class WebSettings(ElementProxy):
    """Proxy for the ``w:webSettings`` root element of the web-settings part.

    Exposes a small, read-oriented slice of the OOXML web-settings
    schema. Boolean flag properties accept a setter that toggles the
    corresponding ``w:val`` child.
    """

    def __init__(
        self,
        element: "BaseOxmlElement",
        parent: "t.ProvidesXmlPart | None" = None,
    ):
        super().__init__(element, parent)
        self._web_settings = cast("CT_WebSettings", element)

    @property
    def encoding(self) -> str | None:
        """Value of ``w:encoding/@w:val`` or |None| if the element is absent.

        Read-only. Records the text encoding Word should use when the
        document is saved as a web page.
        """
        return self._web_settings.encoding_val

    @property
    def optimize_for_browser(self) -> bool:
        """True when ``w:optimizeForBrowser`` is present and not disabled.

        Read/write. Assigning ``False`` (or |None|) removes the element.
        """
        return self._web_settings.optimizeForBrowser_val

    @optimize_for_browser.setter
    def optimize_for_browser(self, value: bool | None):
        self._web_settings.optimizeForBrowser_val = value

    @property
    def allow_png(self) -> bool:
        """True when ``w:allowPNG`` is present and not disabled.

        Read/write. Assigning ``False`` (or |None|) removes the element.
        """
        return self._web_settings.allowPNG_val

    @allow_png.setter
    def allow_png(self, value: bool | None):
        self._web_settings.allowPNG_val = value

    @property
    def do_not_save_as_single_file(self) -> bool:
        """True when ``w:doNotSaveAsSingleFile`` is present and not disabled.

        Read/write. Assigning ``False`` (or |None|) removes the element.
        """
        return self._web_settings.doNotSaveAsSingleFile_val

    @do_not_save_as_single_file.setter
    def do_not_save_as_single_file(self, value: bool | None):
        self._web_settings.doNotSaveAsSingleFile_val = value
