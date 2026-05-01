"""|FontTablePart| and related objects.

Provides read-only access to the ``word/fontTable.xml`` part of a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.font_table import FontTable
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart

if TYPE_CHECKING:
    from docx.oxml.font_table import CT_Fonts
    from docx.package import Package


class FontTablePart(XmlPart):
    """Read-only proxy for the ``word/fontTable.xml`` part of a document.

    The font table records the fonts referenced by the document together with
    descriptive metadata. It is owned by Word, so there is no public API for
    adding or removing entries.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: "CT_Fonts",
        package: "Package",
    ):
        super().__init__(partname, content_type, element, package)
        self._fonts = element

    @property
    def font_table(self) -> "FontTable":
        """A |FontTable| proxy for the ``w:fonts`` root element of this part."""
        return FontTable(self._fonts, self)

    @property
    def font_table_element(self) -> "CT_Fonts":
        """The ``w:fonts`` root element for this part."""
        return cast("CT_Fonts", self._element)
