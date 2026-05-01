"""Initialize `docx` package.

Export the `Document` constructor function and establish the mapping of part-type to
the part-classe that implements that type.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.api import Document

if TYPE_CHECKING:
    from docx.opc.part import Part

__version__ = "1.2.0"


__all__ = ["Document"]


# -- register custom Part classes with opc package reader --

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import PartFactory
from docx.opc.parts.coreprops import CorePropertiesPart
from docx.parts.chart import ChartPart
from docx.parts.comments import CommentsPart
from docx.parts.custom_properties import CustomPropertiesPart
from docx.parts.custom_xml import CustomXmlPart
from docx.parts.document import DocumentPart
from docx.parts.embedded_object import EmbeddedObjectPart
from docx.parts.endnotes import EndnotesPart
from docx.parts.font_table import FontTablePart
from docx.parts.footnotes import FootnotesPart
from docx.parts.glossary import GlossaryPart
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.image import ImagePart
from docx.parts.ink import InkPart
from docx.parts.numbering import NumberingPart
from docx.parts.settings import SettingsPart
from docx.parts.smart_art import DiagramDataPart
from docx.parts.styles import StylesPart
from docx.parts.theme import ThemePart
from docx.parts.web_settings import WebSettingsPart


def part_class_selector(content_type: str, reltype: str) -> type[Part] | None:
    if reltype == RT.IMAGE:
        return ImagePart
    return None


PartFactory.part_class_selector = part_class_selector
PartFactory.part_type_for[CT.DML_CHART] = ChartPart
PartFactory.part_type_for[CT.OFC_CUSTOM_PROPERTIES] = CustomPropertiesPart
PartFactory.part_type_for[CT.OPC_CORE_PROPERTIES] = CorePropertiesPart
PartFactory.part_type_for[CT.XML] = CustomXmlPart
PartFactory.part_type_for[CT.WML_COMMENTS] = CommentsPart
PartFactory.part_type_for[CT.WML_DOCUMENT_MAIN] = DocumentPart
PartFactory.part_type_for[CT.WML_DOCUMENT_MACRO] = DocumentPart
PartFactory.part_type_for[CT.WML_ENDNOTES] = EndnotesPart
PartFactory.part_type_for[CT.WML_FONT_TABLE] = FontTablePart
PartFactory.part_type_for[CT.WML_FOOTER] = FooterPart
PartFactory.part_type_for[CT.WML_FOOTNOTES] = FootnotesPart
PartFactory.part_type_for[CT.WML_DOCUMENT_GLOSSARY] = GlossaryPart
PartFactory.part_type_for[CT.WML_HEADER] = HeaderPart
PartFactory.part_type_for[CT.INKML] = InkPart
PartFactory.part_type_for[CT.OFC_OLE_OBJECT] = EmbeddedObjectPart
PartFactory.part_type_for[CT.WML_NUMBERING] = NumberingPart
PartFactory.part_type_for[CT.WML_SETTINGS] = SettingsPart
PartFactory.part_type_for[CT.WML_STYLES] = StylesPart
PartFactory.part_type_for[CT.DML_DIAGRAM_DATA] = DiagramDataPart
PartFactory.part_type_for[CT.OFC_THEME] = ThemePart
PartFactory.part_type_for[CT.WML_WEB_SETTINGS] = WebSettingsPart

del (
    CT,
    ChartPart,
    CorePropertiesPart,
    CommentsPart,
    CustomPropertiesPart,
    CustomXmlPart,
    DiagramDataPart,
    DocumentPart,
    EmbeddedObjectPart,
    EndnotesPart,
    FontTablePart,
    FooterPart,
    FootnotesPart,
    GlossaryPart,
    HeaderPart,
    InkPart,
    NumberingPart,
    PartFactory,
    SettingsPart,
    StylesPart,
    ThemePart,
    WebSettingsPart,
    part_class_selector,
)
