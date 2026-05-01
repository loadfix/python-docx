"""|WebSettingsPart| and closely related objects.

Provides access to the ``word/webSettings.xml`` part of a document. This part
holds document-level web-publishing settings and is rarely edited from
Python; the proxy exposed here is intentionally read-oriented.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.opc.part import XmlPart
from docx.web_settings import WebSettings

if TYPE_CHECKING:
    from docx.opc.packuri import PackURI
    from docx.oxml.web_settings import CT_WebSettings
    from docx.package import Package


class WebSettingsPart(XmlPart):
    """Read-oriented proxy for the ``word/webSettings.xml`` part.

    The part is created by Word to persist web-publishing configuration
    such as preferred encoding and "save as single file" behaviour. A
    default part is not created on demand; ``web_settings`` returns
    ``None`` for documents that don't already have one.
    """

    def __init__(
        self,
        partname: "PackURI",
        content_type: str,
        element: "CT_WebSettings",
        package: "Package",
    ):
        super().__init__(partname, content_type, element, package)
        self._web_settings_elm = element

    @property
    def web_settings(self) -> WebSettings:
        """A |WebSettings| proxy for the ``w:webSettings`` root of this part."""
        return WebSettings(self._web_settings_elm, self)

    @property
    def web_settings_element(self) -> "CT_WebSettings":
        """The ``w:webSettings`` root element for this part."""
        return cast("CT_WebSettings", self._element)
