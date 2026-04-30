"""Content control (structured document tag) proxy types."""

from __future__ import annotations

import enum
import random
from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
    from docx.oxml.content_controls import CT_Sdt


class ContentControlType(enum.Enum):
    """Enumerates the kinds of content controls (structured document tags).

    The type is determined by the presence of specific children of `w:sdtPr`.
    """

    RICH_TEXT = "richText"
    """Rich-text content control. Default when no specific type marker is present."""

    PLAIN_TEXT = "text"
    """Plain-text content control (`w:text` marker)."""

    CHECKBOX = "checkbox"
    """Checkbox content control (`w14:checkbox` marker)."""

    COMBO_BOX = "comboBox"
    """Combo-box content control (`w:comboBox` marker)."""

    DROPDOWN = "dropDownList"
    """Drop-down-list content control (`w:dropDownList` marker)."""

    DATE = "date"
    """Date content control (`w:date` marker)."""

    PICTURE = "picture"
    """Picture content control (`w:picture` marker). Picture manipulation itself is
    not yet supported - the type is surfaced for introspection only."""


# -- map SDT child-tag to type --
_MARKER_TYPE_MAP = {
    "w:text": ContentControlType.PLAIN_TEXT,
    "w14:checkbox": ContentControlType.CHECKBOX,
    "w:comboBox": ContentControlType.COMBO_BOX,
    "w:dropDownList": ContentControlType.DROPDOWN,
    "w:date": ContentControlType.DATE,
    "w:picture": ContentControlType.PICTURE,
    "w:richText": ContentControlType.RICH_TEXT,
}

# -- reverse map, type to marker tag (None for RICH_TEXT which has no explicit marker) --
_TYPE_MARKER_MAP = {
    ContentControlType.PLAIN_TEXT: "w:text",
    ContentControlType.CHECKBOX: "w14:checkbox",
    ContentControlType.COMBO_BOX: "w:comboBox",
    ContentControlType.DROPDOWN: "w:dropDownList",
    ContentControlType.DATE: "w:date",
    ContentControlType.PICTURE: "w:picture",
    ContentControlType.RICH_TEXT: None,
}


class ContentControl:
    """Proxy object for a `w:sdt` element (a structured document tag / content control).

    Usage is the same whether the SDT is block-level or inline. A :class:`ContentControl`
    exposes common metadata (tag, title, type, id) as well as read/write access to the
    text inside the SDT's `w:sdtContent`.
    """

    def __init__(self, sdt: "CT_Sdt"):
        self._sdt = sdt

    @property
    def element(self) -> "CT_Sdt":
        """The underlying `w:sdt` lxml element."""
        return self._sdt

    # -- tag (metadata) ------------------------------------------------------

    @property
    def tag(self) -> str | None:
        """Programmatic tag value (`w:sdtPr/w:tag/@w:val`), or |None| if not set."""
        return self._sdt.tag_val

    @tag.setter
    def tag(self, value: str | None) -> None:
        self._sdt.tag_val = value

    # -- title / alias -------------------------------------------------------

    @property
    def title(self) -> str | None:
        """Friendly title (`w:sdtPr/w:alias/@w:val`), or |None| if not set."""
        return self._sdt.alias_val

    @title.setter
    def title(self, value: str | None) -> None:
        self._sdt.alias_val = value

    # -- type ----------------------------------------------------------------

    @property
    def type(self) -> ContentControlType:
        """A :class:`ContentControlType` member describing this content control.

        Returns :attr:`ContentControlType.RICH_TEXT` when no specific marker is
        present in `w:sdtPr` (which is how rich-text content controls are identified
        in OOXML).
        """
        marker = self._sdt.type_marker_tag()
        if marker is None:
            return ContentControlType.RICH_TEXT
        return _MARKER_TYPE_MAP.get(marker, ContentControlType.RICH_TEXT)

    # -- id ------------------------------------------------------------------

    @property
    def sdt_id(self) -> int | None:
        """Integer `w:sdtPr/w:id/@w:val` value, or |None| if not present."""
        return self._sdt.sdt_id

    # -- text ----------------------------------------------------------------

    @property
    def text(self) -> str:
        """Concatenated textual content of this content control."""
        return self._sdt.text

    @text.setter
    def text(self, value: str) -> None:
        """Replace this control's content with a single paragraph or run containing
        `value`.

        For block-level SDTs (whose `sdtContent` contains `w:p` children), the content
        is replaced with a single `w:p` holding one `w:r` and `w:t` with `value`. For
        inline SDTs (whose `sdtContent` contains `w:r` children), the content is
        replaced with a single `w:r` holding a `w:t` with `value`.
        """
        sdtContent = self._sdt.get_or_add_sdtContent()
        # -- detect inline vs block by looking at existing children --
        is_inline = self._is_inline()
        # -- clear existing children --
        for child in list(sdtContent):
            sdtContent.remove(child)

        if is_inline:
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            if value != value.strip():
                t.set(qn("xml:space"), "preserve")
            t.text = value
            r.append(t)
            sdtContent.append(r)
        else:
            p = OxmlElement("w:p")
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            if value != value.strip():
                t.set(qn("xml:space"), "preserve")
            t.text = value
            r.append(t)
            p.append(r)
            sdtContent.append(p)

    def _is_inline(self) -> bool:
        """Return True if this SDT is inline (i.e. a child of a paragraph).

        Falls back to inspecting existing sdtContent children when the SDT has
        not yet been attached to a parent.
        """
        parent = self._sdt.getparent()
        if parent is not None:
            if parent.tag == qn("w:p"):
                return True
            if parent.tag == qn("w:body") or parent.tag == qn("w:tc"):
                return False
        # -- fall back to inspecting existing content --
        sdtContent = self._sdt.sdtContent
        if sdtContent is not None:
            for child in sdtContent:
                if child.tag == qn("w:p"):
                    return False
                if child.tag == qn("w:r"):
                    return True
        # -- default: treat as inline (tighter scope for `.text` assignment) --
        return True

    # -- checkbox ------------------------------------------------------------

    @property
    def checked(self) -> bool | None:
        """Value of `w14:checkbox/w14:checked/@w14:val` for checkbox SDTs.

        Returns |None| if this is not a checkbox SDT or no `w14:checked` child exists.
        """
        return self._sdt.checked

    @checked.setter
    def checked(self, value: bool) -> None:
        self._sdt.checked = value


# ---------------------------------------------------------------------------
# factory helpers


def _new_sdt_id() -> int:
    """Return a random positive 32-bit integer suitable for a `w:sdtPr/w:id/@w:val`."""
    return random.randint(1, 2**31 - 1)


def new_sdt(
    content_control_type: ContentControlType,
    tag: str | None = None,
    title: str | None = None,
    inline: bool = False,
) -> "CT_Sdt":
    """Create and return a new `w:sdt` element with the requested type/tag/title.

    When `inline` is True, the sdtContent is initialized with an empty `w:r`. When
    False (block-level), the sdtContent is initialized with an empty `w:p`.
    """
    from docx.oxml.content_controls import CT_Sdt  # local import to avoid cycles

    sdt = cast(CT_Sdt, OxmlElement("w:sdt"))
    sdtPr = sdt.get_or_add_sdtPr()
    # -- w:alias first (per schema friendliness), then w:tag, then w:id --
    if title is not None:
        sdt.alias_val = title
    if tag is not None:
        sdt.tag_val = tag
    sdt.sdt_id = _new_sdt_id()

    # -- set the type marker, if any --
    marker = _TYPE_MARKER_MAP.get(content_control_type)
    if marker is not None:
        sdt.set_type_marker(marker)

    sdtContent = sdt.get_or_add_sdtContent()
    if inline:
        sdtContent.append(OxmlElement("w:r"))
    else:
        sdtContent.append(OxmlElement("w:p"))

    return sdt
