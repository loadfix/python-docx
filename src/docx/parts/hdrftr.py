"""Header and footer part objects."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from docx.opc.constants import CONTENT_TYPE as CT
from docx.oxml.parser import parse_xml
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.package import Package

class FooterPart(StoryPart):
    """Definition of a section footer."""

    @classmethod
    def new(cls, package: Package):
        """Return newly created footer part."""
        partname = package.next_partname("/word/footer%d.xml")
        content_type = CT.WML_FOOTER
        element = parse_xml(cls._default_footer_xml())
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_footer_xml(cls):
        """Return bytes containing XML for a default footer part."""
        path = Path(__file__).parent.parent / "templates" / "default-footer.xml"
        return path.read_bytes()

class HeaderPart(StoryPart):
    """Definition of a section header."""

    @classmethod
    def new(cls, package: Package):
        """Return newly created header part."""
        partname = package.next_partname("/word/header%d.xml")
        content_type = CT.WML_HEADER
        element = parse_xml(cls._default_header_xml())
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_header_xml(cls):
        """Return bytes containing XML for a default header part."""
        path = Path(__file__).parent.parent / "templates" / "default-header.xml"
        return path.read_bytes()
