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
