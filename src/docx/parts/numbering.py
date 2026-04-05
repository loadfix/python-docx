"""|NumberingPart| and closely related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, List, cast

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.numbering import CT_AbstractNum, CT_Lvl, CT_Num, CT_Numbering
from docx.oxml.parser import parse_xml
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from docx.package import Package


class NumberingPart(XmlPart):
    """Proxy for the numbering.xml part containing numbering definitions for a document
    or glossary."""

    @classmethod
    def new(cls) -> NumberingPart:
        """Newly created numbering part, containing only the root ``<w:numbering>``
        element."""
        partname = PackURI("/word/numbering.xml")
        content_type = CT.WML_NUMBERING
        element = cast(CT_Numbering, parse_xml(
            b'<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word'
            b'/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.c'
            b'om/office/mac/office/2008/main" xmlns:mc="http://schemas.openxml'
            b'formats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-mic'
            b'rosoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:of'
            b'fice" xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
            b'2006/relationships" xmlns:m="http://schemas.openxmlformats.org/o'
            b'fficeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml"'
            b' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/word'
            b'processingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:w'
            b'ord" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml'
            b'/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/'
            b'2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word'
            b'/2010/wordprocessingShape" mc:Ignorable="w14 wp14"/>'
        ))
        return cls(partname, content_type, element, None)

    @property
    def numbering_element(self) -> CT_Numbering:
        """The ``<w:numbering>`` root element of this part."""
        return cast(CT_Numbering, self._element)

    @lazyproperty
    def numbering_definitions(self):
        """The |_NumberingDefinitions| instance containing the numbering definitions
        (<w:num> element proxies) for this numbering part."""
        return _NumberingDefinitions(self._element)


class _NumberingDefinitions:
    """Collection of |_NumberingDefinition| instances corresponding to the ``<w:num>``
    elements in a numbering part."""

    def __init__(self, numbering_elm):
        super(_NumberingDefinitions, self).__init__()
        self._numbering = numbering_elm

    def __len__(self):
        return len(self._numbering.num_lst)
