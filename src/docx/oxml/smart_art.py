"""Custom element classes for SmartArt (DrawingML diagrams).

This module re-exports the authoritative SmartArt ``CT_*`` element
classes from :mod:`ooxml_smartart.oxml`. The shared package owns the
``dgm:`` element grammar for all loadfix OOXML parents (docx / pptx /
xlsx); this module preserves docx's historical public surface —
``CT_RelIds``, ``CT_DataModel``, ``CT_Pt``, ``CT_PtLst``, ``CT_Cxn`` —
so every ``from docx.oxml.smart_art import ...`` path keeps working.

Historical docx behaviour preserved via a thin subclass shim:

* :class:`CT_Pt` — docx treats ``@modelId`` as an optional attribute so
  partially-authored data parts that omit it read as ``None`` rather
  than raising :class:`InvalidXmlError`. The shared ``CT_Pt`` declares
  ``modelId`` as required (per the ECMA-376 XSD). The docx subclass
  overrides ``modelId`` to :class:`OptionalAttribute` and is re-
  registered against ``dgm:pt`` so the element-class lookup resolves
  to the docx variant.

SmartArt content is referenced from the document via a ``w:drawing``
containing ``dgm:relIds``, which carries four relationship ids pointing
at companion parts:

* ``r:dm`` → diagram *data* part (``word/diagrams/data1.xml``)
* ``r:lo`` → diagram *layout* part (``word/diagrams/layout1.xml``)
* ``r:qs`` → diagram *quickStyle* part (``word/diagrams/quickStyle1.xml``)
* ``r:cs`` → diagram *colors* part (``word/diagrams/colors1.xml``)

This module exposes read access (the tree of nodes and their text) plus
node-level authoring; the other three parts are owned by the parent
packaging layer and are not parsed further here.

.. versionchanged:: 2026.05.10
   Re-exported from the shared ``python-ooxml-smartart`` package.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

# ---------------------------------------------------------------------------
# Namespace-registry safety: importing ``ooxml_smartart.oxml`` appends the
# shared SmartArt registry to the process-global ``ooxml_xmlchemy``
# composite stack. The composite resolves lookups in reverse registration
# order (most-recent first), so docx's registry would be shadowed — e.g.
# ``OxmlElement("a:effectLst")`` would fall through to the shared smartart
# parser (which has an ``a:`` prefix but no ``effectLst`` class registered)
# and return a generic ``lxml.etree._Element``. Restore docx's registry at
# module-import-completion below (same pattern as ``docx.oxml.chart``).
# ---------------------------------------------------------------------------
from ooxml_smartart.authoring import (
    add_data_node,
    dgm_relIds_from_graphic_data,
    get_root_doc_pt_id,
)
from ooxml_smartart.oxml import (
    CT_Cxn,
    CT_DataModel,
    CT_RelIds,
    register_element_cls,
)
from ooxml_smartart.oxml.data_model import CT_Pt as _CT_Pt
from ooxml_smartart.oxml.data_model import CT_PtList as CT_PtLst
from ooxml_xmlchemy import OptionalAttribute
from ooxml_xmlchemy import configure_namespace_registry as _configure

from docx.oxml.ns import qn
from docx.oxml.parser import _DocxNamespaceRegistry as _DocxRegistry
from docx.oxml.simpletypes import ST_String

if TYPE_CHECKING:
    from docx.oxml.xmlchemy import BaseOxmlElement


__all__ = [
    "CT_Cxn",
    "CT_DataModel",
    "CT_Pt",
    "CT_PtLst",
    "CT_RelIds",
    "add_data_node",
    "dgm_relIds_from_drawing",
    "get_root_doc_pt_id",
    "new_smart_art_inline",
    "qn",
]


class CT_Pt(_CT_Pt):
    """docx-local ``<dgm:pt>`` shim — optional ``@modelId`` for defensive reads.

    The shared :class:`ooxml_smartart.oxml.data_model.CT_Pt` declares
    ``modelId`` as a required attribute (matching the ECMA-376 XSD).
    docx's :mod:`docx.smart_art` proxy layer reads partial data parts
    that may omit ``modelId`` (e.g. older Office builds, hand-authored
    fixtures) and relies on the attribute returning ``None`` rather
    than raising :class:`InvalidXmlError`. This subclass overrides the
    descriptor with :class:`OptionalAttribute` and re-registers against
    ``dgm:pt`` at import time so the element-class lookup resolves to
    the docx variant.

    .. versionadded:: 2026.05.10
    """

    modelId: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "modelId", ST_String
    )


# -- re-register ``dgm:pt`` so reads of docx drawings return the
# -- docx-local subclass with the optional-modelId descriptor. --
register_element_cls("dgm:pt", CT_Pt)


def dgm_relIds_from_drawing(drawing: "BaseOxmlElement") -> "CT_RelIds | None":
    """Return the ``dgm:relIds`` element nested inside `drawing`, or ``None``.

    Checks both inline and anchor variants of the ``w:drawing``
    envelope. Returns ``None`` when the drawing does not reference a
    diagram. Thin docx-local wrapper around the shared-package
    :func:`dgm_relIds_from_graphic_data` helper, adding the
    ``wp:inline`` / ``wp:anchor`` envelope traversal that is specific
    to the WordprocessingML drawing shape.
    """
    matches = drawing.xpath(
        "./wp:inline/a:graphic/a:graphicData/dgm:relIds"
        " | ./wp:anchor/a:graphic/a:graphicData/dgm:relIds"
    )
    if not matches:
        return None
    return cast("CT_RelIds", matches[0])


_SMART_ART_INLINE_XML = (
    '<wp:inline distT="0" distB="0" distL="0" distR="0"'
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    '<wp:extent cx="%(cx)d" cy="%(cy)d"/>'
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:docPr id="%(shape_id)d" name="Diagram %(shape_id)d"/>'
    '<wp:cNvGraphicFramePr>'
    '<a:graphicFrameLocks noChangeAspect="1"/>'
    '</wp:cNvGraphicFramePr>'
    '<a:graphic>'
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    '<dgm:relIds r:dm="%(dm_rId)s" r:lo="%(lo_rId)s" r:qs="%(qs_rId)s" r:cs="%(cs_rId)s"/>'
    '</a:graphicData>'
    '</a:graphic>'
    '</wp:inline>'
)


def new_smart_art_inline(
    shape_id: int,
    cx: int,
    cy: int,
    dm_rId: str,
    lo_rId: str,
    qs_rId: str,
    cs_rId: str,
) -> "BaseOxmlElement":
    """Return a new ``wp:inline`` element wrapping a SmartArt diagram.

    `shape_id` is the ``wp:docPr/@id`` value (a document-unique drawing id).
    `cx`, `cy` are the display extents in EMU. The four rId arguments are the
    relationships from the document part to the four companion diagram parts
    — dm=data, lo=layout, qs=quickStyle, cs=colors.

    .. versionadded:: 2026.05.7
    """
    from docx.oxml.parser import parse_xml

    xml = _SMART_ART_INLINE_XML % {
        "cx": cx,
        "cy": cy,
        "shape_id": shape_id,
        "dm_rId": dm_rId,
        "lo_rId": lo_rId,
        "qs_rId": qs_rId,
        "cs_rId": cs_rId,
    }
    return cast("BaseOxmlElement", parse_xml(xml.encode("utf-8")))


# -- restore docx's namespace registry as the most-recently-registered
# -- entry so docx's ``a:`` / ``w:`` / ``r:`` lookups take precedence
# -- over the shared smartart package's ``a:``-prefix entries. --
_configure(_DocxRegistry())
