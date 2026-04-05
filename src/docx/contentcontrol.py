"""Proxy objects for content controls (structured document tags)."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.contentcontrol import WD_CONTENT_CONTROL_TYPE
from docx.shared import StoryChild

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.sdt import CT_Sdt


class _ContentControlBase(StoryChild):
    """Base class for block-level and inline content control proxies.

    Provides shared property implementations for `w:sdt` element access.
    """

    def __init__(self, sdt: CT_Sdt, parent: t.ProvidesStoryPart):
        super().__init__(parent)
        self._sdt = self._element = sdt

    @property
    def checked(self) -> bool | None:
        """True/False for checkbox content controls, None for other types."""
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        if sdtPr.control_type != WD_CONTENT_CONTROL_TYPE.CHECKBOX:
            return None
        checkbox = sdtPr.checkbox
        if checkbox is None:
            return None
        return checkbox.checked

    @checked.setter
    def checked(self, value: bool) -> None:
        sdtPr = self._sdt.get_or_add_sdtPr()
        if sdtPr.control_type != WD_CONTENT_CONTROL_TYPE.CHECKBOX:
            raise ValueError("checked can only be set on checkbox content controls")
        checkbox = sdtPr.checkbox
        if checkbox is None:
            raise ValueError("checkbox element not found in sdtPr")
        checkbox.checked = value

    @property
    def tag(self) -> str | None:
        """The tag value of this content control, or None if not set."""
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        return sdtPr.tag_val

    @tag.setter
    def tag(self, value: str | None) -> None:
        self._sdt.get_or_add_sdtPr().tag_val = value

    @property
    def text(self) -> str:
        """The text content of this content control."""
        sdtContent = self._sdt.sdtContent
        if sdtContent is None:
            return ""
        return sdtContent.text

    @property
    def title(self) -> str | None:
        """The title (alias) of this content control, or None if not set."""
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        return sdtPr.title

    @title.setter
    def title(self, value: str | None) -> None:
        self._sdt.get_or_add_sdtPr().title = value

    @property
    def type(self) -> WD_CONTENT_CONTROL_TYPE:
        """The type of this content control."""
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return WD_CONTENT_CONTROL_TYPE.RICH_TEXT
        return sdtPr.control_type


class BlockContentControl(_ContentControlBase):
    """Proxy for a block-level `w:sdt` element (content control).

    A block-level content control appears as a direct child of `w:body` and contains
    paragraphs and tables.
    """


class InlineContentControl(_ContentControlBase):
    """Proxy for an inline `w:sdt` element (content control).

    An inline content control appears as a child of `w:p` and contains runs.
    """

    @_ContentControlBase.text.setter
    def text(self, value: str) -> None:
        """Set the text content of this inline content control."""
        sdtContent = self._sdt.get_or_add_sdtContent()
        # -- remove existing runs --
        for r in sdtContent.r_lst:
            sdtContent.remove(r)
        # -- add new run with text --
        if value:
            r = sdtContent.add_r()
            r.text = value
