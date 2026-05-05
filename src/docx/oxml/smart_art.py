"""Custom element classes for SmartArt (DrawingML diagrams).

SmartArt content is referenced from the document via a ``w:drawing`` containing
``dgm:relIds``, which carries four relationship ids pointing at companion parts:

* ``r:dm`` → diagram *data* part (``word/diagrams/data1.xml``)
* ``r:lo`` → diagram *layout* part (``word/diagrams/layout1.xml``)
* ``r:qs`` → diagram *quickStyle* part (``word/diagrams/quickStyle1.xml``)
* ``r:cs`` → diagram *colors* part (``word/diagrams/colors1.xml``)

This module parses only the data part (the tree of nodes and their text). The
other three parts are not needed for read-only extraction of the logical
content and are intentionally ignored.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute

if TYPE_CHECKING:
    pass


class CT_RelIds(BaseOxmlElement):
    """``<dgm:relIds>`` element inside ``a:graphicData``.

    Carries the four relationship ids (``r:dm``, ``r:lo``, ``r:qs``, ``r:cs``)
    that connect the document to the companion diagram parts. Only the data
    relationship (``r:dm``) is required for read-only node-text extraction.
    """

    dm_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:dm", ST_String
    )
    lo_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:lo", ST_String
    )
    qs_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:qs", ST_String
    )
    cs_rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:cs", ST_String
    )


class CT_DataModel(BaseOxmlElement):
    """``<dgm:dataModel>`` element — root of a diagram data part.

    Contains a ``dgm:ptLst`` (list of points / nodes) and a ``dgm:cxnLst``
    (list of connections / edges). Parent-child relationships among nodes are
    expressed by the connection list rather than nested XML structure.
    """

    @property
    def pt_lst(self) -> list[CT_Pt]:
        """All ``dgm:pt`` descendants of the ``dgm:ptLst`` child, in doc order."""
        return cast("list[CT_Pt]", self.xpath("./dgm:ptLst/dgm:pt"))

    @property
    def cxn_lst(self) -> list[CT_Cxn]:
        """All ``dgm:cxn`` descendants of the ``dgm:cxnLst`` child."""
        return cast("list[CT_Cxn]", self.xpath("./dgm:cxnLst/dgm:cxn"))


class CT_PtLst(BaseOxmlElement):
    """``<dgm:ptLst>`` — ordered list of ``dgm:pt`` nodes."""


class CT_Pt(BaseOxmlElement):
    """``<dgm:pt>`` — a single node in the diagram data tree.

    Carries a required ``modelId`` that serves as the node's identity (used by
    ``dgm:cxn`` to express parent-child relationships). The node's display text
    is stored inside a ``dgm:t`` child containing DrawingML-shaped paragraphs
    (``a:p``/``a:r``/``a:t``).
    """

    modelId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "modelId", ST_String
    )
    type: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "type", ST_String
    )

    @property
    def text(self) -> str:
        """Concatenated text of all ``a:t`` descendants of this node's ``dgm:t``.

        Returns the empty string when the node has no ``dgm:t`` child or when
        every text run is empty. Multiple paragraphs inside ``dgm:t`` are
        joined with a newline so callers see the same line-structure Word
        would render.
        """
        # -- `dgm:t`, `a:p`, `a:r`, `a:t` are not registered as custom classes,
        # -- so their lxml elements don't carry the docx namespace prefix map.
        # -- Doing all traversal via `self.xpath` keeps the prefix map active. --
        a_p_tag = qn("a:p")
        a_r_tag = qn("a:r")
        a_t_tag = qn("a:t")
        dgm_t_elems = self.xpath("./dgm:t")
        if not dgm_t_elems:
            return ""
        dgm_t = dgm_t_elems[0]
        paragraphs: list[str] = []
        for p in dgm_t.iterchildren(a_p_tag):
            parts: list[str] = []
            for r in p.iterchildren(a_r_tag):
                for t in r.iterchildren(a_t_tag):
                    if t.text:
                        parts.append(t.text)
            paragraphs.append("".join(parts))
        return "\n".join(paragraphs)


class CT_Cxn(BaseOxmlElement):
    """``<dgm:cxn>`` — one edge in the diagram data tree.

    Encodes a parent/child relationship between two ``dgm:pt`` nodes identified
    by ``srcId`` (parent) and ``destId`` (child). Connections with ``type`` other
    than ``parOf`` (for example ``presOf``) describe presentation links rather
    than logical hierarchy and are ignored by the hierarchy reconstructor.
    """

    type: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "type", ST_String
    )
    srcId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "srcId", ST_String
    )
    destId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "destId", ST_String
    )


def dgm_relIds_from_drawing(drawing: BaseOxmlElement) -> CT_RelIds | None:
    """Return the ``dgm:relIds`` element nested inside `drawing`, or None.

    Checks both inline and anchor variants. Returns ``None`` when the drawing
    does not reference a diagram.
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
) -> BaseOxmlElement:
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


_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


def _escape_xml_text(text: str) -> str:
    """Return `text` with XML special characters escaped for a text node."""
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )


def add_data_node(
    data_model: CT_DataModel,
    model_id: str,
    text: str,
    parent_id: str,
    src_ord: int,
) -> CT_Pt:
    """Append a content ``dgm:pt`` + ``parOf`` ``dgm:cxn`` to `data_model`.

    `model_id` is the new node's ``modelId`` (caller allocates; typically a GUID).
    `text` is the display text. `parent_id` is the ``modelId`` of the parent node
    (the root ``type="doc"`` point for top-level nodes). `src_ord` is the 0-based
    sibling index under the parent; it sets the ``srcOrd`` attribute on the
    connection so Word preserves insertion order.

    Returns the newly-added ``dgm:pt`` element.

    .. versionadded:: 2026.05.7
    """
    from docx.oxml.parser import parse_xml

    # -- locate or create the ptLst / cxnLst children --
    ptLst_matches = data_model.xpath("./dgm:ptLst")
    if not ptLst_matches:
        raise ValueError("data_model has no <dgm:ptLst> child")
    ptLst = ptLst_matches[0]

    cxnLst_matches = data_model.xpath("./dgm:cxnLst")
    if not cxnLst_matches:
        # -- insert a cxnLst directly after ptLst --
        cxnLst = parse_xml(
            f'<dgm:cxnLst xmlns:dgm="{_DGM_NS}"/>'
        )
        ptLst.addnext(cxnLst)
    else:
        cxnLst = cxnLst_matches[0]

    # -- new content node --
    pt_xml = (
        f'<dgm:pt xmlns:dgm="{_DGM_NS}" xmlns:a="{_A_NS}"'
        f' modelId="{model_id}">'
        f'<dgm:prSet phldrT="[Text]"/>'
        f'<dgm:spPr/>'
        f'<dgm:t>'
        f'<a:bodyPr/><a:lstStyle/>'
        f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{_escape_xml_text(text)}</a:t></a:r></a:p>'
        f'</dgm:t>'
        f'</dgm:pt>'
    )
    pt_el = parse_xml(pt_xml)
    ptLst.append(pt_el)

    # -- connection from parent to new node --
    cxn_xml = (
        f'<dgm:cxn xmlns:dgm="{_DGM_NS}"'
        f' modelId="{model_id}-cxn" type="parOf"'
        f' srcId="{parent_id}" destId="{model_id}"'
        f' srcOrd="{src_ord}" destOrd="0"/>'
    )
    cxn_el = parse_xml(cxn_xml)
    cxnLst.append(cxn_el)

    return cast("CT_Pt", pt_el)


def get_root_doc_pt_id(data_model: CT_DataModel) -> str:
    """Return the ``modelId`` of the ``type="doc"`` root point in `data_model`.

    Raises :class:`ValueError` when no such point is present.

    .. versionadded:: 2026.05.7
    """
    for pt in data_model.pt_lst:
        if pt.type == "doc" and pt.modelId is not None:
            return pt.modelId
    # -- doc-type points may also carry attributes via xpath when not registered --
    matches = data_model.xpath('./dgm:ptLst/dgm:pt[@type="doc"]/@modelId')
    if matches:
        return str(matches[0])
    raise ValueError("data_model has no <dgm:pt type='doc'> root point")


# Re-export qn so tests importing from this module can avoid deep paths.
__all__ = [
    "CT_Cxn",
    "CT_DataModel",
    "CT_Pt",
    "CT_PtLst",
    "CT_RelIds",
    "dgm_relIds_from_drawing",
    "qn",
]
