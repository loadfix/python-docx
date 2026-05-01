"""|ThemePart| and closely related objects.

Provides access to the ``word/theme/theme1.xml`` part of a document. The
theme is Word-authored (python-docx does not create a default theme on
demand); the proxy exposed here is intentionally read-only.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.opc.part import XmlPart
from docx.theme import Theme

if TYPE_CHECKING:
    from docx.opc.packuri import PackURI
    from docx.oxml.theme import CT_Theme
    from docx.package import Package


class ThemePart(XmlPart):
    """Read-only proxy for the ``word/theme/theme1.xml`` part.

    A default theme part is not created on demand; :attr:`docx.document.Document.theme`
    returns |None| for documents that do not already have a theme relationship.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_Theme,
        package: Package,
    ):
        super().__init__(partname, content_type, element, package)
        self._theme_elm = element

    @property
    def theme(self) -> Theme:
        """A |Theme| proxy for the ``a:theme`` root of this part."""
        return Theme(self._theme_elm, self)

    @property
    def theme_element(self) -> CT_Theme:
        """The ``a:theme`` root element for this part."""
        return cast("CT_Theme", self._element)
