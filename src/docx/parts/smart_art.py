"""|DiagramDataPart| and companion parts for a SmartArt diagram.

A SmartArt diagram is stored as up to five companion parts under
``/word/diagrams/``:

* ``dataN.xml`` — the logical node tree (``<dgm:dataModel>``)
* ``layoutN.xml`` — the layout algorithm (``<dgm:layoutDef>``)
* ``colorsN.xml`` — the colour mapping (``<dgm:colorsDef>``)
* ``quickStyleN.xml`` — the style mapping (``<dgm:styleDef>``)
* ``drawingN.xml`` — a pre-rendered fallback (``<dsp:drawing>``, MS extension)

Read-only access to the node text requires only the *data* part. This module
also carries the authoring helpers that materialise the four core parts from
vendored templates when :meth:`docx.document.Document.add_smart_art` is
called.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.oxml.smart_art import CT_DataModel
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.package import Package


_TEMPLATES_DIR = Path(__file__).parent.parent / "templates" / "smart_art"


def _read_template(relative: str) -> bytes:
    """Return the bytes of the vendored SmartArt template at *relative*."""
    return (_TEMPLATES_DIR / relative).read_bytes()


class DiagramDataPart(XmlPart):
    """Container part for a SmartArt diagram's data model.

    The root element is ``<dgm:dataModel>``, which holds the node list
    (``dgm:ptLst``) and the connection list (``dgm:cxnLst``). Loaded by
    :class:`~docx.opc.part.PartFactory` for content type
    ``application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml``.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_DataModel,
        package: Package,
    ):
        super().__init__(partname, content_type, element, package)

    @property
    def data_model(self) -> CT_DataModel:
        """The root ``<dgm:dataModel>`` element of this part."""
        return cast("CT_DataModel", self._element)

    @classmethod
    def new(cls, package: Package, layout_name: str) -> Self:
        """Return a newly created data part for a *layout_name*-family SmartArt.

        `layout_name` must be one of ``"list"``, ``"cycle"`` or ``"process"``.
        The newly-minted data part has a single ``type="doc"`` root point with
        the layout's canonical ``loTypeId`` URN; user content is added later
        via :meth:`docx.smart_art.SmartArt.add_node`.

        .. versionadded:: 2026.05.7
        """
        from docx.oxml.smart_art import CT_DataModel as _CT_DataModel

        partname = package.next_partname("/word/diagrams/data%d.xml")
        blob = _read_template(f"{layout_name}/data.xml")
        element = cast("_CT_DataModel", parse_xml(blob))
        return cls(partname, CT.DML_DIAGRAM_DATA, element, package)


class DiagramLayoutPart(XmlPart):
    """Container part for a SmartArt diagram's layout definition.

    The root element is ``<dgm:layoutDef>``. Content type is
    ``application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml``.
    python-docx embeds a vendored copy of Word's built-in ``process1`` layout
    for every authored SmartArt — Word uses its own built-in renderer for the
    ``loTypeId`` declared in ``data.xml``, so the embedded copy merely
    satisfies the package requirement.

    .. versionadded:: 2026.05.7
    """

    @classmethod
    def new(cls, package: Package) -> Self:
        partname = package.next_partname("/word/diagrams/layout%d.xml")
        blob = _read_template("layout1.xml")
        element = cast("BaseOxmlElement", parse_xml(blob))
        return cls(partname, CT.DML_DIAGRAM_LAYOUT, element, package)


class DiagramColorsPart(XmlPart):
    """Container part for a SmartArt diagram's colour definition.

    The root element is ``<dgm:colorsDef>``. Content type is
    ``application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml``.

    .. versionadded:: 2026.05.7
    """

    @classmethod
    def new(cls, package: Package) -> Self:
        partname = package.next_partname("/word/diagrams/colors%d.xml")
        blob = _read_template("colors1.xml")
        element = cast("BaseOxmlElement", parse_xml(blob))
        return cls(partname, CT.DML_DIAGRAM_COLORS, element, package)


class DiagramStylePart(XmlPart):
    """Container part for a SmartArt diagram's quick-style definition.

    The root element is ``<dgm:styleDef>``. Content type is
    ``application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml``.

    .. versionadded:: 2026.05.7
    """

    @classmethod
    def new(cls, package: Package) -> Self:
        partname = package.next_partname("/word/diagrams/quickStyle%d.xml")
        blob = _read_template("quickStyle1.xml")
        element = cast("BaseOxmlElement", parse_xml(blob))
        return cls(partname, CT.DML_DIAGRAM_STYLE, element, package)
