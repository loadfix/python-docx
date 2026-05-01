"""Read-only proxy objects for SmartArt diagrams.

A SmartArt diagram is a DrawingML *diagram* embedded in the document via a
``w:drawing`` whose ``a:graphicData`` contains a ``dgm:relIds`` element. The
actual content (tree of nodes with text) lives in a companion
``word/diagrams/dataN.xml`` part that this module parses.

python-docx exposes SmartArt read-only. Callers can:

* Detect SmartArt on any :class:`Drawing` via :attr:`Drawing.is_smart_art`
  and :attr:`Drawing.smart_art`.
* Enumerate every SmartArt in the document body via
  :attr:`Document.smart_art`.
* Walk the parsed node tree via :attr:`SmartArt.nodes`, or fetch the full
  concatenated text via :attr:`SmartArt.text`.

Hierarchy reconstruction uses the ``dgm:cxnLst`` connection list — edges of
type ``parOf`` express parent/child relationships between nodes. When the
connection list is missing or does not describe a well-formed tree, the
proxy falls back to a flat list of every ``dgm:pt`` node in document order
(all at ``level == 0``).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.oxml.smart_art import (
    CT_DataModel,
    CT_Pt,
    CT_RelIds,
    dgm_relIds_from_drawing,
)

if TYPE_CHECKING:
    from docx.oxml.drawing import CT_Drawing
    from docx.parts.smart_art import DiagramDataPart


class SmartArtNode:
    """Read-only proxy for a single node in a SmartArt diagram.

    Each node wraps a ``<dgm:pt>`` element. ``text`` is the node's display
    text (possibly empty). ``level`` is the node's depth in the reconstructed
    hierarchy (``0`` for top-level nodes). ``children`` is the list of
    direct descendants, each itself a :class:`SmartArtNode`.
    """

    def __init__(
        self,
        pt: CT_Pt,
        level: int = 0,
        children: list[SmartArtNode] | None = None,
    ):
        self._pt = pt
        self._level = level
        self._children: list[SmartArtNode] = list(children) if children else []

    @property
    def children(self) -> list[SmartArtNode]:
        """Direct child nodes, in document (sibling) order."""
        return list(self._children)

    @property
    def level(self) -> int:
        """Depth of this node in the reconstructed hierarchy (0 = top-level)."""
        return self._level

    @property
    def model_id(self) -> str | None:
        """Value of the node's ``modelId`` attribute, or ``None`` if absent."""
        return self._pt.modelId

    @property
    def text(self) -> str:
        """Concatenated text of all runs inside this node's ``dgm:t``.

        Paragraphs inside ``dgm:t`` are joined with newlines; runs within a
        paragraph are concatenated without a separator.
        """
        return self._pt.text


class SmartArt:
    """Read-only proxy for a SmartArt diagram embedded in the document.

    Constructed from the ``dgm:relIds`` element inside a ``w:drawing`` plus the
    resolved :class:`~docx.parts.smart_art.DiagramDataPart` (which may be
    ``None`` when the relationship cannot be resolved). Exposes the parsed node
    list and a convenience ``text`` property that concatenates every node's
    text with indent-based formatting reflecting the hierarchy level.
    """

    def __init__(
        self,
        relIds: CT_RelIds,
        data_part: DiagramDataPart | None,
    ):
        self._relIds = relIds
        self._data_part = data_part

    @property
    def data_partname(self) -> str | None:
        """OPC partname of the related data part (e.g. ``/word/diagrams/data1.xml``).

        Returns ``None`` when the relationship referenced by ``r:dm`` could not
        be resolved (for example, when the referenced part is missing or is of
        an unexpected type).
        """
        if self._data_part is None:
            return None
        return str(self._data_part.partname)

    @property
    def dm_rId(self) -> str | None:
        """Value of the ``r:dm`` attribute on ``dgm:relIds``, or ``None``."""
        return self._relIds.dm_rId

    @property
    def nodes(self) -> list[SmartArtNode]:
        """Top-level :class:`SmartArtNode` objects parsed from the data part.

        Returns an empty list when the data part is missing or contains no
        ``dgm:pt`` nodes. When hierarchy reconstruction succeeds, the returned
        list contains only the roots and each root's children are reachable
        via :attr:`SmartArtNode.children`. When reconstruction is not possible
        (e.g. the connection list is missing or malformed), every node is
        returned flat at ``level == 0``.
        """
        if self._data_part is None:
            return []
        data_model = self._data_part.data_model
        return _build_nodes(data_model)

    @property
    def text(self) -> str:
        """Concatenated text of every node, one per line, with indent by level.

        Each node's text is prefixed with ``"  " * level`` (two spaces per level)
        so callers can inspect the logical structure at a glance. Empty nodes
        (no text) are still emitted as blank indented lines to preserve the
        shape of the tree. When the data part is missing the return value is
        the empty string.
        """
        lines: list[str] = []
        for root in self.nodes:
            _append_text_lines(root, lines)
        return "\n".join(lines)


def _build_nodes(data_model: CT_DataModel) -> list[SmartArtNode]:
    """Return top-level :class:`SmartArtNode` trees for *data_model*.

    Skips ``dgm:pt`` nodes whose ``type`` attribute is present and not
    ``"node"`` — those are presentation markers (for example ``parTrans``,
    ``sibTrans``, ``pres``) that do not carry user content. If the connection
    list cannot produce a clean tree, falls back to a flat list containing
    every remaining ``dgm:pt`` at ``level == 0``.
    """
    content_pts = [
        pt for pt in data_model.pt_lst
        if pt.type in (None, "node")
    ]
    if not content_pts:
        return []

    # -- map modelId -> pt for quick lookups --
    pt_by_id: dict[str, CT_Pt] = {}
    for pt in content_pts:
        mid = pt.modelId
        if mid is not None:
            pt_by_id[mid] = pt

    # -- collect parent -> [children] from `parOf` connections --
    parent_of: dict[str, list[str]] = {}
    has_parent: set[str] = set()
    for cxn in data_model.cxn_lst:
        if cxn.type not in (None, "parOf"):
            continue
        src, dst = cxn.srcId, cxn.destId
        if src is None or dst is None:
            continue
        # -- only consider connections between content nodes --
        if src not in pt_by_id or dst not in pt_by_id:
            continue
        parent_of.setdefault(src, []).append(dst)
        has_parent.add(dst)

    # -- if no usable connections exist, fall back to flat list --
    if not parent_of:
        return [SmartArtNode(pt, level=0) for pt in content_pts]

    def build(pt: CT_Pt, level: int, seen: set[str]) -> SmartArtNode:
        mid = pt.modelId
        child_ids = parent_of.get(mid or "", [])
        children: list[SmartArtNode] = []
        for cid in child_ids:
            if cid in seen:
                # -- defensive: avoid cycles from malformed data --
                continue
            child_pt = pt_by_id[cid]
            children.append(build(child_pt, level + 1, seen | {cid}))
        return SmartArtNode(pt, level=level, children=children)

    # -- roots are content nodes with no incoming parOf edge --
    roots = [pt for pt in content_pts if (pt.modelId or "") not in has_parent]
    if not roots:
        # -- every node claims a parent; fall back to flat --
        return [SmartArtNode(pt, level=0) for pt in content_pts]

    return [build(pt, 0, {pt.modelId or ""}) for pt in roots]


def _append_text_lines(node: SmartArtNode, lines: list[str]) -> None:
    """Recursively append ``"  " * level + text`` for *node* and its descendants."""
    prefix = "  " * node.level
    lines.append(prefix + node.text)
    for child in node.children:
        _append_text_lines(child, lines)


def smart_art_for_drawing(
    drawing: CT_Drawing,
    document_part,
) -> SmartArt | None:
    """Return a :class:`SmartArt` for *drawing* or ``None`` when not SmartArt.

    ``document_part`` must expose a ``related_parts`` mapping keyed by
    relationship id — :class:`~docx.parts.document.DocumentPart` satisfies
    this. When the ``dgm:relIds`` element is absent the return value is
    ``None`` (the drawing is not SmartArt). When ``dgm:relIds`` is present
    but the referenced data part cannot be resolved (missing, wrong type,
    etc.) a :class:`SmartArt` is still returned, with an empty node list.
    """
    from docx.parts.smart_art import DiagramDataPart

    relIds = dgm_relIds_from_drawing(drawing)
    if relIds is None:
        return None

    data_part: DiagramDataPart | None = None
    dm_rId = relIds.dm_rId
    if dm_rId:
        try:
            candidate = document_part.related_parts[dm_rId]
        except (KeyError, AttributeError):
            candidate = None
        if isinstance(candidate, DiagramDataPart):
            data_part = candidate

    return SmartArt(relIds, data_part)


__all__ = [
    "SmartArt",
    "SmartArtNode",
    "smart_art_for_drawing",
]
