"""Proxy objects for SmartArt diagrams.

A SmartArt diagram is a DrawingML *diagram* embedded in the document via a
``w:drawing`` whose ``a:graphicData`` contains a ``dgm:relIds`` element. The
actual content (tree of nodes with text) lives in a companion
``word/diagrams/dataN.xml`` part that this module parses.

python-docx exposes SmartArt for both reading and authoring. Callers can:

* Detect SmartArt on any :class:`Drawing` via :attr:`Drawing.is_smart_art`
  and :attr:`Drawing.smart_art`.
* Enumerate every SmartArt in the document body via
  :attr:`Document.smart_art`.
* Walk the parsed node tree via :attr:`SmartArt.nodes`, or fetch the full
  concatenated text via :attr:`SmartArt.text`.
* Append a new SmartArt via :meth:`docx.document.Document.add_smart_art` and
  populate it by calling :meth:`SmartArt.add_node` one text string at a time.

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
    from docx.document import Document
    from docx.oxml.drawing import CT_Drawing
    from docx.parts.smart_art import DiagramDataPart


class SmartArtNode:
    """Read-only proxy for a single node in a SmartArt diagram.

    Each node wraps a ``<dgm:pt>`` element. ``text`` is the node's display
    text (possibly empty). ``level`` is the node's depth in the reconstructed
    hierarchy (``0`` for top-level nodes). ``children`` is the list of
    direct descendants, each itself a :class:`SmartArtNode`.

    .. versionadded:: 2026.05.0
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
        """Direct child nodes, in document (sibling) order.

        .. versionadded:: 2026.05.0
        """
        return list(self._children)

    @property
    def level(self) -> int:
        """Depth of this node in the reconstructed hierarchy (0 = top-level).

        .. versionadded:: 2026.05.0
        """
        return self._level

    @property
    def model_id(self) -> str | None:
        """Value of the node's ``modelId`` attribute, or ``None`` if absent.

        .. versionadded:: 2026.05.0
        """
        return self._pt.modelId

    @property
    def text(self) -> str:
        """Concatenated text of all runs inside this node's ``dgm:t``.

        Paragraphs inside ``dgm:t`` are joined with newlines; runs within a
        paragraph are concatenated without a separator.

        .. versionadded:: 2026.05.0
        """
        return self._pt.text


class SmartArt:
    """Read-only proxy for a SmartArt diagram embedded in the document.

    Constructed from the ``dgm:relIds`` element inside a ``w:drawing`` plus the
    resolved :class:`~docx.parts.smart_art.DiagramDataPart` (which may be
    ``None`` when the relationship cannot be resolved). Exposes the parsed node
    list and a convenience ``text`` property that concatenates every node's
    text with indent-based formatting reflecting the hierarchy level.

    .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
        """
        if self._data_part is None:
            return None
        return str(self._data_part.partname)

    @property
    def dm_rId(self) -> str | None:
        """Value of the ``r:dm`` attribute on ``dgm:relIds``, or ``None``.

        .. versionadded:: 2026.05.0
        """
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

        .. versionadded:: 2026.05.0
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

        .. versionadded:: 2026.05.0
        """
        lines: list[str] = []
        for root in self.nodes:
            _append_text_lines(root, lines)
        return "\n".join(lines)

    def add_node(self, text: str) -> SmartArtNode:
        """Append a top-level content node carrying `text` and return its proxy.

        Nodes are appended in call order and become the children of the
        diagram's ``type="doc"`` root point. A fresh UUID-shaped ``modelId``
        is allocated for both the new ``dgm:pt`` and its parent ``parOf``
        connection. The ``srcOrd`` attribute of the new connection is set to
        the current number of top-level content nodes, so Word's layout
        algorithm preserves insertion order.

        Raises :class:`RuntimeError` when the SmartArt has no resolvable
        data part — the read-only wrapper returned for a drawing whose
        ``dgm:relIds`` does not resolve to a ``DiagramDataPart`` cannot
        be authored against.

        .. versionadded:: 2026.05.7
        """
        import uuid as _uuid

        from docx.oxml.smart_art import add_data_node, get_root_doc_pt_id

        if self._data_part is None:
            raise RuntimeError(
                "cannot add a node to a SmartArt whose data part did not resolve"
            )

        data_model = self._data_part.data_model
        parent_id = get_root_doc_pt_id(data_model)
        # -- src_ord: count of existing content nodes whose parent is the root --
        src_ord = _count_direct_children(data_model, parent_id)
        model_id = "{%s}" % str(_uuid.uuid4()).upper()
        pt_el = add_data_node(
            data_model, model_id, text, parent_id=parent_id, src_ord=src_ord
        )
        return SmartArtNode(pt_el, level=0)


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


def _count_direct_children(data_model: CT_DataModel, parent_id: str) -> int:
    """Return the number of ``parOf`` connections whose ``srcId`` is `parent_id`."""
    count = 0
    for cxn in data_model.cxn_lst:
        if cxn.type in (None, "parOf") and cxn.srcId == parent_id:
            count += 1
    return count


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

    .. versionadded:: 2026.05.0
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


_SUPPORTED_LAYOUTS = ("list", "cycle", "process")


def add_smart_art_to_document(
    document: "Document",
    layout_name: str,
    cx: int,
    cy: int,
) -> SmartArt:
    """Create the four SmartArt companion parts and append an inline drawing.

    ``document`` is the owning :class:`~docx.document.Document` proxy.
    ``layout_name`` must be one of ``"list"``, ``"cycle"`` or ``"process"``
    (case-insensitive). ``cx`` / ``cy`` are the EMU display dimensions for the
    inline drawing.

    Wires up:

    * a new ``word/diagrams/dataN.xml`` part (diagram data)
    * a new ``word/diagrams/layoutN.xml`` part (diagram layout)
    * a new ``word/diagrams/quickStyleN.xml`` part (diagram quick style)
    * a new ``word/diagrams/colorsN.xml`` part (diagram colours)
    * four relationships from the document part to those four companions
    * a new paragraph at the end of the body whose run carries a
      ``w:drawing`` referencing the four rIds

    The returned :class:`SmartArt` is fully authorable via :meth:`SmartArt.add_node`.

    .. versionadded:: 2026.05.7
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as _RT
    from docx.oxml.smart_art import new_smart_art_inline
    from docx.parts.smart_art import (
        DiagramColorsPart,
        DiagramDataPart,
        DiagramLayoutPart,
        DiagramStylePart,
    )

    layout_key = layout_name.lower()
    if layout_key not in _SUPPORTED_LAYOUTS:
        raise ValueError(
            f"unsupported SmartArt layout {layout_name!r}; "
            f"expected one of {_SUPPORTED_LAYOUTS!r}"
        )

    document_part = document.part
    package = document_part.package
    assert package is not None

    data_part = DiagramDataPart.new(package, layout_key)
    layout_part = DiagramLayoutPart.new(package)
    qs_part = DiagramStylePart.new(package)
    colors_part = DiagramColorsPart.new(package)

    dm_rId = document_part.relate_to(data_part, _RT.DIAGRAM_DATA)
    lo_rId = document_part.relate_to(layout_part, _RT.DIAGRAM_LAYOUT)
    qs_rId = document_part.relate_to(qs_part, _RT.DIAGRAM_QUICK_STYLE)
    cs_rId = document_part.relate_to(colors_part, _RT.DIAGRAM_COLORS)

    shape_id = document_part.next_id
    inline = new_smart_art_inline(
        shape_id=shape_id,
        cx=cx,
        cy=cy,
        dm_rId=dm_rId,
        lo_rId=lo_rId,
        qs_rId=qs_rId,
        cs_rId=cs_rId,
    )

    # -- append a new paragraph + run + drawing to the document body --
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    drawing = run._r.add_drawing(inline)  # pyright: ignore[reportPrivateUsage]

    # -- reach into the drawing to pull the dgm:relIds back out as a proxy --
    relIds = dgm_relIds_from_drawing(drawing)
    assert relIds is not None
    return SmartArt(relIds, data_part)


__all__ = [
    "SmartArt",
    "SmartArtNode",
    "add_smart_art_to_document",
    "smart_art_for_drawing",
]
