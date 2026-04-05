"""Content control (structured document tag) proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.enum.sdt import WD_CONTENT_CONTROL_TYPE
from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.sdt import CT_Sdt


class ContentControl(ElementProxy):
    """Proxy for a `<w:sdt>` element, either block-level or inline.

    Provides access to the tag, title, type, text content, and (for checkbox SDTs) the
    checked state.
    """

    def __init__(self, sdt: CT_Sdt, parent: object | None = None):
        super().__init__(sdt, parent)
        self._sdt = sdt

    @property
    def tag(self) -> str | None:
        """The tag value of this content control, or None if not set."""
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        return sdtPr.tag_val

    @tag.setter
    def tag(self, value: str | None) -> None:
        sdtPr = self._sdt.get_or_add_sdtPr()
        sdtPr.tag_val = value

    @property
    def title(self) -> str | None:
        """The title (alias) of this content control, or None if not set."""
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        return sdtPr.alias_val

    @title.setter
    def title(self, value: str | None) -> None:
        sdtPr = self._sdt.get_or_add_sdtPr()
        sdtPr.alias_val = value

    @property
    def type(self) -> WD_CONTENT_CONTROL_TYPE:
        """The type of this content control."""
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return WD_CONTENT_CONTROL_TYPE.RICH_TEXT
        type_str = sdtPr.sdt_type
        return WD_CONTENT_CONTROL_TYPE(type_str)

    @property
    def text(self) -> str:
        """The text content of this content control.

        For block-level SDTs, paragraph boundaries are indicated with newlines. For
        inline SDTs, the text of all runs is concatenated.
        """
        sdtContent = self._sdt.sdtContent
        if sdtContent is None:
            return ""
        return sdtContent.text

    @text.setter
    def text(self, value: str) -> None:
        """Set the text content of this content control.

        Replaces any existing content with a single run (inline) or single paragraph
        (block-level) containing the specified text.
        """
        from docx.oxml.ns import qn
        from docx.oxml.parser import OxmlElement

        sdtContent = self._sdt.get_or_add_sdtContent()

        # -- remove existing content --
        for child in list(sdtContent):
            sdtContent.remove(child)

        if self._sdt.is_block_level:
            p = sdtContent.add_p()
            r = p.add_r()
        else:
            r = sdtContent.add_r()

        t = OxmlElement("w:t")
        t.text = value
        if value and (value[0] == " " or value[-1] == " "):
            t.set(qn("xml:space"), "preserve")
        r.append(t)

    @property
    def checked(self) -> bool | None:
        """True if this checkbox content control is checked, False if unchecked.

        Returns None if this is not a checkbox content control.
        """
        if self.type != WD_CONTENT_CONTROL_TYPE.CHECKBOX:
            return None
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        checkbox = sdtPr.checkbox
        if checkbox is None:
            return False
        return checkbox.checked

    @checked.setter
    def checked(self, value: bool) -> None:
        """Set the checked state of this checkbox content control.

        Raises ValueError if this is not a checkbox content control.
        """
        if self.type != WD_CONTENT_CONTROL_TYPE.CHECKBOX:
            raise ValueError("can only set checked on a checkbox content control")
        sdtPr = self._sdt.get_or_add_sdtPr()
        checkbox = sdtPr.checkbox
        if checkbox is None:
            raise ValueError("checkbox element not found in sdtPr")
        checkbox.checked = value
