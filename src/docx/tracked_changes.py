"""Proxy objects for tracked changes (revision marks) in a document."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.section import CT_SectPr
    from docx.oxml.text.font import CT_RPr
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.tracked_changes import (
        CT_PPrChange,
        CT_RPrChange,
        CT_RunTrackChange,
        CT_SectPrChange,
        CT_TrackChange,
    )
    from docx.oxml.xmlchemy import BaseOxmlElement


class TrackedChange(ElementProxy):
    """Proxy for a single tracked change in a paragraph.

    Wraps a `<w:ins>`, `<w:del>`, `<w:moveFrom>`, or `<w:moveTo>` element
    containing one or more runs. For move revisions the :class:`MoveRevision`
    subclass exposes the additional `w:name` attribute and paired-peer lookup.
    """

    def __init__(self, element: CT_RunTrackChange):
        super().__init__(element)
        self._tc_element = element

    @property
    def author(self) -> str:
        """The author who made this change."""
        return self._tc_element.author

    @property
    def date(self) -> dt.datetime | None:
        """The date and time when this change was made, or |None| if not recorded."""
        return self._tc_element.date

    @property
    def text(self) -> str:
        """The textual content of this tracked change."""
        return cast(str, self._tc_element.text)

    @property
    def type(self) -> str:
        """The type of this tracked change.

        One of ``"insertion"``, ``"deletion"``, ``"move_from"``, or ``"move_to"``.
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
        """
        self._tc_element.reject()


class MoveRevision(TrackedChange):
    """Proxy for a move revision — a `<w:moveFrom>` or `<w:moveTo>` element.

    In addition to the common author/date/text surface inherited from
    :class:`TrackedChange`, a move revision carries a ``name`` that pairs the
    source (`w:moveFrom`) with the destination (`w:moveTo`). The :attr:`peer`
    property resolves the counterpart element anywhere in the same XML tree by
    matching `@w:name`.

    Note on the paragraph-level range markers `w:moveFromRangeStart/End` and
    `w:moveToRangeStart/End`: those bracket cross-paragraph moves rather than
    wrap run content, so no proxy type is exposed for them. They survive a
    round-trip unchanged; callers that need to work with them can iterate the
    underlying XML.
    """

    @property
    def name(self) -> str | None:
        """The `@w:name` attribute pairing this move half with its peer, or |None|.

        Well-formed move-revision XML always includes a name, but the attribute
        is declared optional per ECMA-376 so callers must handle |None|.
        """
        return self._tc_element.get(qn("w:name"))

    @property
    def peer(self) -> MoveRevision | None:
        """The paired `w:moveFrom`/`w:moveTo` on the other side of the move.

        Looks up the first element (other than ``self``) in the same tree whose
        local tag matches the opposite side and whose `@w:name` equals this
        element's name. Returns |None| if the name is unset, if there is no
        tree root (detached element), or if no peer is found.
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
                return MoveRevision(candidate)
        return None


class FormattingChange(ElementProxy):
    """Proxy for a formatting revision mark (`w:rPrChange`, `w:pPrChange`,
    `w:sectPrChange`).

    Records the author and date of a formatting edit and provides access to the
    previous formatting via :attr:`old_properties`, which returns the inner
    `w:rPr`, `w:pPr`, or `w:sectPr` element holding the pre-edit values.
    """

    def __init__(self, element: CT_TrackChange):
        super().__init__(element)
        self._fc_element = element

    @property
    def author(self) -> str:
        """The author who made this formatting change."""
        return self._fc_element.author

    @property
    def date(self) -> dt.datetime | None:
        """When this formatting change was made, or |None| if not recorded."""
        return self._fc_element.date

    @property
    def old_properties(self) -> CT_RPr | CT_PPr | CT_SectPr | None:
        """The nested `w:rPr`, `w:pPr`, or `w:sectPr` holding prior formatting.

        |None| if the change element has no inner properties element (malformed or
        "no prior formatting" case).
        """
        from docx.oxml.tracked_changes import CT_PPrChange, CT_RPrChange, CT_SectPrChange

        if isinstance(self._fc_element, CT_RPrChange):
            return self._fc_element.rPr
        if isinstance(self._fc_element, CT_PPrChange):
            return self._fc_element.pPr
        if isinstance(self._fc_element, CT_SectPrChange):
            return self._fc_element.sectPr
        return None


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
    `w:moveTo`) and formatting track changes (`w:rPrChange`, `w:pPrChange`,
    `w:sectPrChange`). Returns the count of change elements resolved.

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

    fmt_changes: list[BaseOxmlElement] = root.xpath(
        ".//w:rPrChange | .//w:pPrChange | .//w:sectPrChange"
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
