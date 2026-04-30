"""Custom element classes related to tracked changes (revisions)."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, List

from lxml import etree  # noqa: F401  -- used for QName

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_DateTime, ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.text.run import CT_R


class CT_RunTrackChange(BaseOxmlElement):
    """Base for `<w:ins>` and `<w:del>` elements wrapping runs in a paragraph.

    Both share the same attribute set: `w:id`, `w:author`, and `w:date`.
    """

    r_lst: List[CT_R]

    r = ZeroOrMore("w:r", successors=())

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    author: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:author", ST_String
    )
    date: dt.datetime | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:date", ST_DateTime
    )

    def accept(self) -> None:
        """Accept this tracked change, resolving it in the document.

        Implementation-specific behavior is provided by subclasses.
        """
        raise NotImplementedError  # pragma: no cover

    def reject(self) -> None:
        """Reject this tracked change, restoring prior state.

        Implementation-specific behavior is provided by subclasses.
        """
        raise NotImplementedError  # pragma: no cover

    def _remove_from_parent(self) -> None:
        """Detach this element from its parent, discarding the element and its content."""
        parent = self.getparent()
        if parent is not None:
            parent.remove(self)

    def _unwrap(self) -> None:
        """Replace this element in its parent with its children, in place.

        Children keep their original order. This element is removed from the tree.
        """
        parent = self.getparent()
        if parent is None:
            return
        index = parent.index(self)
        for i, child in enumerate(list(self)):
            parent.insert(index + i, child)
        parent.remove(self)


class CT_Ins(CT_RunTrackChange):
    """`<w:ins>` element, containing runs that were inserted."""

    @property
    def text(self) -> str:
        """The textual content of the inserted runs."""
        return "".join(r.text for r in self.r_lst)

    def accept(self) -> None:
        """Accept this insertion: keep the content, remove the `w:ins` wrapper."""
        self._unwrap()

    def reject(self) -> None:
        """Reject this insertion: remove the `w:ins` element and its contents."""
        self._remove_from_parent()


class CT_Del(CT_RunTrackChange):
    """`<w:del>` element, containing runs that were deleted."""

    @property
    def text(self) -> str:
        """The textual content of the deleted runs.

        Deleted runs use `w:delText` elements rather than `w:t`.
        """
        return "".join(
            str(e) for e in self.xpath("w:r/w:delText")
        )

    def accept(self) -> None:
        """Accept this deletion: remove the `w:del` element and its contents."""
        self._remove_from_parent()

    def reject(self) -> None:
        """Reject this deletion: keep the content, remove the `w:del` wrapper.

        `w:delText` descendants are converted back to `w:t` so the restored text renders
        as normal run content.
        """
        for delText in self.xpath(".//w:delText"):
            delText.tag = qn("w:t")
        self._unwrap()


class CT_DelText(BaseOxmlElement):
    """`<w:delText>` element, containing text in a deleted run."""

    def __str__(self) -> str:
        """Text contained in this element, the empty string if it has no content."""
        return self.text or ""


def accept_formatting_change(change_elm: BaseOxmlElement) -> None:
    """Accept a `w:rPrChange`, `w:pPrChange`, or `w:sectPrChange` element.

    Accepting a formatting change discards the record of the prior formatting while
    leaving the current (new) properties in place. The change element is detached from
    its parent.
    """
    parent = change_elm.getparent()
    if parent is not None:
        parent.remove(change_elm)


def reject_formatting_change(change_elm: BaseOxmlElement) -> None:
    """Reject a `w:rPrChange`, `w:pPrChange`, or `w:sectPrChange` element.

    Rejecting restores the prior formatting: the inner `w:rPr` / `w:pPr` / `w:sectPr`
    holds the old properties, and its children replace the current children of the
    parent properties element. Parent attributes are preserved.
    """
    parent = change_elm.getparent()
    if parent is None:
        return
    local = etree.QName(change_elm).localname  # e.g. "rPrChange"
    old_local = local[: -len("Change")]  # "rPr"
    old_elm = change_elm.find(qn(f"w:{old_local}"))
    for child in list(parent):
        parent.remove(child)
    if old_elm is not None:
        for child in list(old_elm):
            parent.append(child)
