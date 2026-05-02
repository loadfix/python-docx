"""|FontTablePart| and related objects.

Provides access to the ``word/fontTable.xml`` part of a document and the
optional sibling font-data parts that hold embedded font binaries.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, cast

from docx.font_table import FontTable
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.oxml.font_table import CT_Fonts
    from docx.package import Package


_DEFAULT_FONT_TABLE_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b"<w:fonts "
    b'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    b"/>\n"
)


class FontTablePart(XmlPart):
    """Proxy for the ``word/fontTable.xml`` part of a document.

    The font table records the fonts referenced by the document together with
    descriptive metadata. Historically this part has been populated only by
    Word; as of python-docx 1.3.0 a small authoring surface is available via
    :meth:`docx.font_table.FontTable.add_embedded_font` so applications that
    need to embed a TrueType font can do so.
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

    @classmethod
    def default(cls, package: "Package") -> "FontTablePart":
        """Return a newly-created, empty font-table part.

        .. versionadded:: 1.3.0.dev0
        """
        partname = PackURI("/word/fontTable.xml")
        element = cast("CT_Fonts", parse_xml(_DEFAULT_FONT_TABLE_XML))
        return cls(partname, CT.WML_FONT_TABLE, element, package)

    @property
    def font_table(self) -> "FontTable":
        """A |FontTable| proxy for the ``w:fonts`` root element of this part."""
        return FontTable(self._fonts, self)

    @property
    def font_table_element(self) -> "CT_Fonts":
        """The ``w:fonts`` root element for this part."""
        return cast("CT_Fonts", self._element)

    def add_font_part(self, font_path: str | Path) -> "FontPart":
        """Return a new |FontPart| holding the contents of the file at `font_path`.

        The part is registered with this font-table part via an ``r:font``
        relationship so Word resolves the embedded font reference. The
        generated part-name has the form ``/word/fonts/fontN.fntdata`` where
        ``N`` is the next available integer across the package; the
        ``.fntdata`` extension is mapped to the ``application/x-fontdata``
        content type by the default content-types registration.

        .. versionadded:: 1.3.0.dev0
        """
        assert self._package is not None
        path = Path(font_path)
        blob = path.read_bytes()
        partname = self._package.next_partname("/word/fonts/font%d.fntdata")
        font_part = FontPart(partname, CT.X_FONTDATA, blob, self._package)
        return font_part


class FontPart(Part):
    """Binary part holding an embedded font referenced from the font table.

    Round-trip preservation is automatic — the OPC layer serialises the
    part's blob on save the same way it does for images. There is no public
    API for reading the binary payload; use :attr:`blob` for direct access
    when needed.

    .. versionadded:: 1.3.0.dev0
    """

    @classmethod
    def load(
        cls,
        partname: PackURI,
        content_type: str,
        blob: bytes,
        package: "Package",
    ) -> "FontPart":
        return cls(partname, content_type, blob, package)
