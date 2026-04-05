"""Proxy objects for tracked changes (revision marks) in a document."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING

from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.tracked_changes import CT_RunTrackChange


class TrackedChange(ElementProxy):
    """Proxy for a single tracked change (insertion or deletion) in a paragraph.

    Wraps a `<w:ins>` or `<w:del>` element that contains one or more runs.
    """

    def __init__(self, element: CT_RunTrackChange):
        super().__init__(element)

    @property
    def author(self) -> str:
        """The author who made this change."""
        return self._element.author

    @property
    def date(self) -> dt.datetime | None:
        """The date and time when this change was made, or |None| if not recorded."""
        return self._element.date

    @property
    def text(self) -> str:
        """The textual content of this tracked change."""
        return self._element.text

    @property
    def type(self) -> str:
        """The type of this tracked change, either ``"insertion"`` or ``"deletion"``."""
        from docx.oxml.tracked_changes import CT_Ins

        return "insertion" if isinstance(self._element, CT_Ins) else "deletion"
