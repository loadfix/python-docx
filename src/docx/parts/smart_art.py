"""|DiagramDataPart| — read-only container for a SmartArt ``dataN.xml`` part.

A SmartArt diagram has four companion parts (``data``, ``layout``, ``colors``,
``quickStyle``). Read-only access to the node text requires only the *data*
part — the other three describe presentation. python-docx exposes these parts
read-only and makes no attempt to author or modify SmartArt content.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.opc.part import XmlPart

if TYPE_CHECKING:
    from docx.opc.packuri import PackURI
    from docx.oxml.smart_art import CT_DataModel
    from docx.package import Package


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
