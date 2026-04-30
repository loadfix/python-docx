"""Proxy objects for tracked changes (revision marks) in a document."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, cast

from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.tracked_changes import CT_RunTrackChange
    from docx.oxml.xmlchemy import BaseOxmlElement


class TrackedChange(ElementProxy):
    """Proxy for a single tracked change (insertion or deletion) in a paragraph.

    Wraps a `<w:ins>` or `<w:del>` element that contains one or more runs.
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
        """The type of this tracked change, either ``"insertion"`` or ``"deletion"``."""
        from docx.oxml.tracked_changes import CT_Ins

        return "insertion" if isinstance(self._tc_element, CT_Ins) else "deletion"

    def accept(self) -> None:
        """Accept this tracked change.

        For an insertion, the `w:ins` wrapper is removed and its inserted runs remain
        in the paragraph. For a deletion, the `w:del` element and its deleted content
        are removed entirely.
        """
        self._tc_element.accept()

    def reject(self) -> None:
        """Reject this tracked change.

        For an insertion, the `w:ins` element and its inserted content are removed
        entirely. For a deletion, the `w:del` wrapper is removed and its `w:delText`
        children are converted back to `w:t` so the content is restored as live text.
        """
        self._tc_element.reject()


def _resolve_all_changes(root: BaseOxmlElement, *, accept: bool) -> int:
    """Accept or reject every tracked change beneath `root`.

    Processes run-level track changes (`w:ins`, `w:del`) and formatting track changes
    (`w:rPrChange`, `w:pPrChange`, `w:sectPrChange`). Returns the count of change
    elements resolved.

    Nested changes (e.g. a `w:ins` inside a `w:del`) are handled by processing
    innermost elements first so outer wrappers see stable children.
    """
    from docx.oxml.tracked_changes import (
        CT_Del,
        CT_Ins,
        accept_formatting_change,
        reject_formatting_change,
    )

    run_changes: list[BaseOxmlElement] = root.xpath(".//w:ins | .//w:del")
    run_changes.sort(key=lambda e: len(list(e.iterancestors())), reverse=True)
    count = 0
    for elm in run_changes:
        if elm.getparent() is None:
            continue
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
