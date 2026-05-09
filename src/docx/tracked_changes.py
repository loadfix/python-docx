"""Proxy objects for tracked changes (revision marks) in a document."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.section import CT_SectPr
    from docx.oxml.table import CT_TblPr, CT_TcPr, CT_TrPr
    from docx.oxml.text.font import CT_RPr
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.run import CT_R
    from docx.oxml.tracked_changes import (
        CT_PPrChange,
        CT_RPrChange,
        CT_RunTrackChange,
        CT_SectPrChange,
        CT_TblPrChange,
        CT_TcPrChange,
        CT_TrackChange,
        CT_TrPrChange,
    )
    from docx.oxml.xmlchemy import BaseOxmlElement


class TrackedChange(ElementProxy):
    """Proxy for a single tracked change in a paragraph.

    Wraps a `<w:ins>`, `<w:del>`, `<w:moveFrom>`, or `<w:moveTo>` element
    containing one or more runs. For move revisions the :class:`MoveRevision`
    subclass exposes the additional `w:name` attribute and paired-peer lookup.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, element: CT_RunTrackChange):
        super().__init__(element)
        self._tc_element = element

    @property
    def author(self) -> str:
        """The author who made this change.

        .. versionadded:: 2026.05.0
        """
        return self._tc_element.author

    @property
    def date(self) -> dt.datetime | None:
        """The date and time when this change was made, or |None| if not recorded.

        .. versionadded:: 2026.05.0
        """
        return self._tc_element.date

    @property
    def text(self) -> str:
        """The textual content of this tracked change.

        .. versionadded:: 2026.05.0
        """
        return cast(str, self._tc_element.text)

    @property
    def type(self) -> str:
        """The type of this tracked change.

        One of ``"insertion"``, ``"deletion"``, ``"move_from"``, or ``"move_to"``.

        .. versionadded:: 2026.05.0
        """
        # -- check the move subclasses before their bases (CT_MoveFrom extends
        # -- CT_Del, CT_MoveTo extends CT_Ins) --
        from docx.oxml.tracked_changes import CT_Ins, CT_MoveFrom, CT_MoveTo

        if isinstance(self._tc_element, CT_MoveFrom):
            return "move_from"
        if isinstance(self._tc_element, CT_MoveTo):
            return "move_to"
        return "insertion" if isinstance(self._tc_element, CT_Ins) else "deletion"

    def accept(self) -> None:
        """Accept this tracked change.

        For an insertion, the `w:ins` wrapper is removed and its inserted runs remain
        in the paragraph. For a deletion, the `w:del` element and its deleted content
        are removed entirely. For a `w:moveFrom`, the source element and its content
        are removed (completing the move). For a `w:moveTo`, the wrapper is removed
        and its runs survive as live content.

        .. versionadded:: 2026.05.0
        """
        self._tc_element.accept()

    def reject(self) -> None:
        """Reject this tracked change.

        For an insertion, the `w:ins` element and its inserted content are removed
        entirely. For a deletion, the `w:del` wrapper is removed and its `w:delText`
        children are converted back to `w:t` so the content is restored as live text.
        For a `w:moveFrom`, the wrapper is unwound so the source text is restored in
        place. For a `w:moveTo`, the destination element and its content are removed
        (cancelling the move).

        .. versionadded:: 2026.05.0
        """
        self._tc_element.reject()


#: Alias of :class:`TrackedChange`, the common base for all four run-level
#: revision proxies. Prefer this name in new code — it matches the ECMA-376
#: terminology ("revision") and aligns with :attr:`Document.revisions` /
#: :attr:`Paragraph.revisions` / :attr:`Run.revisions`.
#:
#: .. versionadded:: 2026.05.11
Revision = TrackedChange


class Insertion(TrackedChange):
    """Proxy for an insertion revision (a `<w:ins>` element).

    :attr:`type` is always ``"insertion"``. :meth:`accept` unwraps the element
    (keeping the inserted runs as live content); :meth:`reject` removes the
    element and its content.

    .. versionadded:: 2026.05.11
    """


class Deletion(TrackedChange):
    """Proxy for a deletion revision (a `<w:del>` element).

    :attr:`type` is always ``"deletion"``. :meth:`accept` removes the element
    and its deleted runs; :meth:`reject` unwraps the element and converts
    `w:delText` descendants back to `w:t` so the deleted text is restored.

    .. versionadded:: 2026.05.11
    """


class Move(TrackedChange):
    """Proxy for a move revision — a `<w:moveFrom>` or `<w:moveTo>` element.

    In addition to the common author/date/text surface inherited from
    :class:`TrackedChange`, a move revision carries a ``name`` that pairs the
    source (`w:moveFrom`) with the destination (`w:moveTo`). The :attr:`peer`
    property resolves the counterpart element anywhere in the same XML tree by
    matching `@w:name`.

    :attr:`type` is ``"move_from"`` for the source and ``"move_to"`` for the
    destination.

    Note on the paragraph-level range markers `w:moveFromRangeStart/End` and
    `w:moveToRangeStart/End`: those bracket cross-paragraph moves rather than
    wrap run content, so no proxy type is exposed for them. They survive a
    round-trip unchanged; callers that need to work with them can iterate the
    underlying XML.

    .. versionadded:: 2026.05.11
    """

    @property
    def name(self) -> str | None:
        """The `@w:name` attribute pairing this move half with its peer, or |None|.

        Well-formed move-revision XML always includes a name, but the attribute
        is declared optional per ECMA-376 so callers must handle |None|.

        .. versionadded:: 2026.05.0
        """
        return self._tc_element.get(qn("w:name"))

    @property
    def peer(self) -> "Move | None":
        """The paired `w:moveFrom`/`w:moveTo` on the other side of the move.

        Looks up the first element (other than ``self``) in the same tree whose
        local tag matches the opposite side and whose `@w:name` equals this
        element's name. Returns |None| if the name is unset, if there is no
        tree root (detached element), or if no peer is found.

        .. versionadded:: 2026.05.0
        """
        from docx.oxml.tracked_changes import CT_MoveFrom, CT_MoveTo

        name = self.name
        if not name:
            return None

        # -- walk up to the document root (or nearest ancestor) and search from
        # -- there; this handles both attached and fragment-rooted elements --
        root = self._tc_element
        while root.getparent() is not None:
            root = cast("CT_RunTrackChange", root.getparent())

        if isinstance(self._tc_element, CT_MoveFrom):
            peer_xpath = ".//w:moveTo"
            peer_cls: type = CT_MoveTo
        else:
            peer_xpath = ".//w:moveFrom"
            peer_cls = CT_MoveFrom

        for candidate in root.xpath(peer_xpath):
            if candidate is self._tc_element:
                continue
            if not isinstance(candidate, peer_cls):
                continue
            if candidate.get(qn("w:name")) == name:
                return Move(candidate)
        return None


#: Back-compatibility alias. New code should prefer :class:`Move`.
#:
#: .. versionadded:: 2026.05.0
#: .. versionchanged:: 2026.05.11
#:    Renamed to :class:`Move`; ``MoveRevision`` is now an alias of the new name.
MoveRevision = Move


class FormattingChange(ElementProxy):
    """Proxy for a formatting revision mark (`w:rPrChange`, `w:pPrChange`,
    `w:sectPrChange`).

    Records the author and date of a formatting edit and provides access to the
    previous formatting via :attr:`old_properties`, which returns the inner
    `w:rPr`, `w:pPr`, or `w:sectPr` element holding the pre-edit values.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, element: CT_TrackChange):
        super().__init__(element)
        self._fc_element = element

    @property
    def author(self) -> str:
        """The author who made this formatting change.

        .. versionadded:: 2026.05.0
        """
        return self._fc_element.author

    @property
    def date(self) -> dt.datetime | None:
        """When this formatting change was made, or |None| if not recorded.

        .. versionadded:: 2026.05.0
        """
        return self._fc_element.date

    @property
    def old_properties(
        self,
    ) -> CT_RPr | CT_PPr | CT_SectPr | CT_TcPr | CT_TrPr | CT_TblPr | None:
        """The nested properties element holding the prior formatting.

        Returns the inner `w:rPr`, `w:pPr`, `w:sectPr`, `w:tcPr`, `w:trPr`, or
        `w:tblPr` element for the corresponding change type.

        |None| if the change element has no inner properties element (malformed or
        "no prior formatting" case).

        .. versionadded:: 2026.05.0
        """
        from docx.oxml.tracked_changes import (
            CT_PPrChange,
            CT_RPrChange,
            CT_SectPrChange,
            CT_TblPrChange,
            CT_TcPrChange,
            CT_TrPrChange,
        )

        if isinstance(self._fc_element, CT_RPrChange):
            return self._fc_element.rPr
        if isinstance(self._fc_element, CT_PPrChange):
            return self._fc_element.pPr
        if isinstance(self._fc_element, CT_SectPrChange):
            return self._fc_element.sectPr
        if isinstance(self._fc_element, CT_TcPrChange):
            return self._fc_element.tcPr
        if isinstance(self._fc_element, CT_TrPrChange):
            return self._fc_element.trPr
        if isinstance(self._fc_element, CT_TblPrChange):
            return self._fc_element.tblPr
        return None


def _wrap_revision(elm: "CT_RunTrackChange") -> TrackedChange:
    """Return the proxy subclass matching `elm`'s element type.

    Maps `w:ins` to :class:`Insertion`, `w:del` to :class:`Deletion`, and
    `w:moveFrom`/`w:moveTo` to :class:`Move`. The move subclasses must be
    checked before their bases (`CT_MoveFrom` extends `CT_Del`,
    `CT_MoveTo` extends `CT_Ins`).

    .. versionadded:: 2026.05.11
    """
    from docx.oxml.tracked_changes import CT_Ins, CT_MoveFrom, CT_MoveTo

    if isinstance(elm, (CT_MoveFrom, CT_MoveTo)):
        return Move(elm)
    if isinstance(elm, CT_Ins):
        return Insertion(elm)
    return Deletion(elm)


def _render_paragraph_marks(
    p_elm: CT_P,
    open_ins: str = "[+",
    close_ins: str = "+]",
    open_del: str = "[-",
    close_del: str = "-]",
) -> str:
    """Render `p_elm` as text with insertion/deletion revision markers.

    Walks the paragraph's children in document order. Plain runs contribute their
    text; `w:ins` and `w:del` wrappers contribute their inner text wrapped with the
    corresponding open/close markers. `w:hyperlink` elements are recursed into so
    track-change wrappers inside them are rendered in place. Other inner-content
    elements (`w:fldSimple`, `w:sdt`) contribute their plain text.

    When the paragraph has no tracked changes the returned string matches
    `paragraph.text`.
    """
    parts: list[str] = []
    _append_container_text(
        p_elm, parts, open_ins, close_ins, open_del, close_del
    )
    return "".join(parts)


def _append_container_text(
    container: BaseOxmlElement,
    parts: list[str],
    open_ins: str,
    close_ins: str,
    open_del: str,
    close_del: str,
) -> None:
    """Walk direct children of `container` and append rendered text into `parts`."""
    ins_tag = qn("w:ins")
    del_tag = qn("w:del")
    r_tag = qn("w:r")
    hyperlink_tag = qn("w:hyperlink")
    fldSimple_tag = qn("w:fldSimple")
    sdt_tag = qn("w:sdt")

    for child in container:
        tag = child.tag
        if tag == r_tag:
            parts.append(child.text or "")  # CT_R.text
        elif tag == ins_tag:
            parts.append(open_ins)
            _append_container_text(
                cast("BaseOxmlElement", child),
                parts, open_ins, close_ins, open_del, close_del,
            )
            parts.append(close_ins)
        elif tag == del_tag:
            parts.append(open_del)
            # -- `w:del` contains `w:r` children whose text sits in `w:delText` --
            for delText in child.xpath(".//w:delText"):
                parts.append(delText.text or "")
            parts.append(close_del)
        elif tag == hyperlink_tag:
            _append_container_text(
                cast("BaseOxmlElement", child),
                parts, open_ins, close_ins, open_del, close_del,
            )
        elif tag == fldSimple_tag or tag == sdt_tag:
            # -- defer to the element's own `.text` for fields and SDTs --
            parts.append(child.text or "")


def _resolve_all_changes(root: BaseOxmlElement, *, accept: bool) -> int:
    """Accept or reject every tracked change beneath `root`.

    Processes run-level track changes (`w:ins`, `w:del`, `w:moveFrom`,
    `w:moveTo`), formatting track changes (`w:rPrChange`, `w:pPrChange`,
    `w:sectPrChange`, `w:tcPrChange`, `w:trPrChange`, `w:tblPrChange`), and
    cell-level revisions (`w:cellIns`, `w:cellDel`). Returns the count of
    change elements resolved.

    Nested changes (e.g. a `w:ins` inside a `w:del`) are handled by processing
    innermost elements first so outer wrappers see stable children.
    """
    from docx.oxml.tracked_changes import (
        CT_Del,
        CT_Ins,
        accept_formatting_change,
        reject_formatting_change,
    )

    run_changes: list[BaseOxmlElement] = root.xpath(
        ".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"
    )
    run_changes.sort(key=lambda e: len(list(e.iterancestors())), reverse=True)
    count = 0
    for elm in run_changes:
        if elm.getparent() is None:
            continue
        # -- CT_MoveFrom is a CT_Del and CT_MoveTo is a CT_Ins, so this check
        # -- covers all four element types without listing the move classes --
        if isinstance(elm, (CT_Ins, CT_Del)):
            elm.accept() if accept else elm.reject()
            count += 1

    # -- cell-level revisions. Resolve before formatting changes so a
    # -- `w:tcPrChange` inside a cell being deleted is only processed once (the
    # -- enclosing `w:tc` is removed here if needed). --
    cell_changes: list[BaseOxmlElement] = root.xpath(".//w:cellIns | .//w:cellDel")
    for elm in cell_changes:
        if elm.getparent() is None:
            continue
        count += _resolve_cell_change(elm, accept=accept)

    fmt_changes: list[BaseOxmlElement] = root.xpath(
        ".//w:rPrChange | .//w:pPrChange | .//w:sectPrChange"
        " | .//w:tcPrChange | .//w:trPrChange | .//w:tblPrChange"
    )
    for elm in fmt_changes:
        if elm.getparent() is None:
            continue
        if accept:
            accept_formatting_change(elm)
        else:
            reject_formatting_change(elm)
        count += 1

    return count


def _resolve_cell_change(elm: BaseOxmlElement, *, accept: bool) -> int:
    """Accept or reject a `w:cellIns` or `w:cellDel` revision marker.

    - Accept `w:cellIns` -> the insertion is accepted; the marker is removed but
      the cell is kept.
    - Reject `w:cellIns` -> the insertion is rejected; the whole enclosing cell
      is removed.
    - Accept `w:cellDel` -> the deletion is accepted; the whole enclosing cell
      is removed.
    - Reject `w:cellDel` -> the deletion is rejected; the marker is removed but
      the cell is kept.

    Returns 1 if the change was processed, 0 if the marker was orphaned or its
    surrounding structure was unexpected.
    """
    from docx.oxml.tracked_changes import CT_CellDel, CT_CellIns

    tcPr = elm.getparent()
    if tcPr is None:
        return 0
    tc = tcPr.getparent()  # -- the enclosing `w:tc`
    if tc is None:
        # -- detached `w:tcPr`; just remove the marker --
        tcPr.remove(elm)
        return 1

    is_insertion = isinstance(elm, CT_CellIns)
    is_deletion = isinstance(elm, CT_CellDel)
    if not (is_insertion or is_deletion):
        # -- unexpected element class; remove marker defensively --
        tcPr.remove(elm)
        return 1

    remove_cell = (is_deletion and accept) or (is_insertion and not accept)
    if remove_cell:
        row = tc.getparent()
        if row is not None:
            row.remove(tc)
        return 1

    # -- keep the cell, just remove the marker --
    tcPr.remove(elm)
    return 1


# -- Track-changes writer helpers -------------------------------------------
#
# These helpers are used by `BlockItemContainer.add_paragraph` and
# `Paragraph.add_run` when the document-level `Document.tracked_changes(...)`
# context manager is active (or when the `track_author=` keyword argument is
# passed to either of those methods).


def _next_revision_id(root: BaseOxmlElement) -> int:
    """Return the next unused integer revision id within `root`.

    Scans every `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, `w:rPrChange`,
    `w:pPrChange`, `w:sectPrChange`, `w:tcPrChange`, `w:trPrChange`,
    `w:tblPrChange`, `w:cellIns`, and `w:cellDel` descendant for a `w:id`
    attribute and returns ``max(existing) + 1``. Returns ``1`` when no
    revision element is present.
    """
    ids: list[int] = []
    for el in root.xpath(
        ".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"
        " | .//w:rPrChange | .//w:pPrChange | .//w:sectPrChange"
        " | .//w:tcPrChange | .//w:trPrChange | .//w:tblPrChange"
        " | .//w:cellIns | .//w:cellDel"
    ):
        raw = el.get(qn("w:id"))
        if raw is None:
            continue
        try:
            ids.append(int(raw))
        except ValueError:
            continue
    return (max(ids) + 1) if ids else 1


def wrap_run_in_ins(
    r: CT_R,
    author: str,
    date: dt.datetime | None = None,
    change_id: int | None = None,
) -> CT_RunTrackChange:
    """Replace `r` in its parent with a `w:ins` wrapper and return the wrapper.

    The newly-created `w:ins` element is positioned where `r` sat among its
    siblings and `r` becomes its sole child. If `change_id` is omitted an id
    is allocated by scanning `r`'s document root for the next unused id. If
    `date` is omitted the current UTC time is used (normalised to whole
    seconds for deterministic XML).

    Returns the new `w:ins` element.

    .. versionadded:: 2026.05.0
    """
    parent = r.getparent()
    if parent is None:
        raise ValueError("cannot wrap a detached run in a w:ins element")
    if change_id is None:
        # -- walk up to document root to scope id allocation --
        root = r
        while root.getparent() is not None:
            root = cast("CT_R", root.getparent())
        change_id = _next_revision_id(cast("BaseOxmlElement", root))
    if date is None:
        date = dt.datetime.now(dt.timezone.utc).replace(microsecond=0)

    ins = cast(
        "CT_RunTrackChange",
        OxmlElement(
            "w:ins",
            attrs={
                qn("w:id"): str(change_id),
                qn("w:author"): author,
                qn("w:date"): date.strftime("%Y-%m-%dT%H:%M:%SZ"),
            },
        ),
    )
    index = parent.index(r)
    parent.remove(r)
    ins.append(r)
    parent.insert(index, ins)
    return ins


class _TrackedChangesCtx:
    """Context-manager for tracked-change writes on a :class:`Document`.

    Created by :meth:`Document.tracked_changes`. While active, every call to
    :meth:`Document.add_paragraph`, :meth:`BlockItemContainer.add_paragraph`,
    or :meth:`Paragraph.add_run` wraps the freshly-inserted `w:r` in a
    `w:ins` element whose `w:author` and `w:date` come from this context.

    Contexts can be nested; the innermost active context supplies the
    author/date. Passing an explicit ``track_author=`` keyword argument to
    `add_paragraph` / `add_run` overrides the context and works even when no
    context is active.

    .. versionadded:: 2026.05.0
    """

    def __init__(
        self, document, author: str, date: dt.datetime | None = None
    ):
        self._document = document
        self._author = author
        self._date = date

    @property
    def author(self) -> str:
        """Author string applied to each tracked insertion."""
        return self._author

    @property
    def date(self) -> dt.datetime | None:
        """Timestamp applied to each tracked insertion, or |None| to use now()."""
        return self._date

    def __enter__(self):
        self._document._tracked_changes_stack.append(self)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        stack = self._document._tracked_changes_stack
        # -- pop the top frame; be defensive if the caller has nested wrongly --
        if stack and stack[-1] is self:
            stack.pop()
        elif self in stack:
            stack.remove(self)


def _active_track_author(part) -> tuple[str, dt.datetime | None] | None:
    """Return ``(author, date)`` from the active tracked-changes context.

    Looks up the innermost `_TrackedChangesCtx` registered on the Document
    proxy that owns `part`. Returns |None| when no context is active or the
    part does not belong to a Document proxy (e.g. header/footer story
    parts, which don't carry the stack).
    """
    doc_proxy = getattr(part, "_track_changes_doc_proxy", None)
    if doc_proxy is None:
        return None
    stack = getattr(doc_proxy, "_tracked_changes_stack", None)
    # -- Require a real list; `Mock.getattr` returns another Mock which is
    # -- truthy but not a valid stack. Reject anything that isn't a list. --
    if not isinstance(stack, list) or not stack:
        return None
    top = stack[-1]
    if not isinstance(top, _TrackedChangesCtx):
        return None
    return top.author, top.date
