"""Content control (structured document tag) proxy types."""

from __future__ import annotations

import enum
import random
from typing import TYPE_CHECKING, Iterator, Union, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
    from docx.oxml.content_controls import (
        CT_DataBinding,
        CT_Sdt,
        CT_SdtComboBox,
        CT_SdtDate,
        CT_SdtDocPart,
        CT_SdtDropDownList,
    )
    from docx.text.paragraph import Paragraph


class ContentControlType(enum.Enum):
    """Enumerates the kinds of content controls (structured document tags).

    The type is determined by the presence of specific children of `w:sdtPr`.

    .. versionadded:: 2026.05.0
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

    REPEATING_SECTION = "repeatingSection"
    """Repeating-section content control (`w15:repeatingSection` marker, MS Word
    2013+ extension). Wraps a block or table-row region that users can duplicate
    via the "Insert New Item" UI. `[Added in 2026.05.10]`."""

    BUILDING_BLOCK = "docPartObj"
    """Building-block gallery content control (`w:docPartObj`/`w:docPartList`
    marker). Offers the user a choice of preset content from a glossary
    document gallery. `[Added in 2026.05.10]`."""


# -- map SDT child-tag to type --
_MARKER_TYPE_MAP = {
    "w:text": ContentControlType.PLAIN_TEXT,
    "w14:checkbox": ContentControlType.CHECKBOX,
    "w:comboBox": ContentControlType.COMBO_BOX,
    "w:dropDownList": ContentControlType.DROPDOWN,
    "w:date": ContentControlType.DATE,
    "w:picture": ContentControlType.PICTURE,
    "w:richText": ContentControlType.RICH_TEXT,
    "w15:repeatingSection": ContentControlType.REPEATING_SECTION,
    "w:docPartObj": ContentControlType.BUILDING_BLOCK,
    "w:docPartList": ContentControlType.BUILDING_BLOCK,
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
    ContentControlType.REPEATING_SECTION: "w15:repeatingSection",
    ContentControlType.BUILDING_BLOCK: "w:docPartObj",
}


class ContentControl:
    """Proxy object for a `w:sdt` element (a structured document tag / content control).

    Usage is the same whether the SDT is block-level or inline. A :class:`ContentControl`
    exposes common metadata (tag, title, type, id) as well as read/write access to the
    text inside the SDT's `w:sdtContent`.

    .. versionadded:: 2026.05.0
    """

    # -- populated by the module-level dispatcher block below --
    proxy_for: "staticmethod[[CT_Sdt], ContentControl]"

    def __init__(self, sdt: "CT_Sdt"):
        self._sdt = sdt

    @property
    def element(self) -> "CT_Sdt":
        """The underlying `w:sdt` lxml element.

        .. versionadded:: 2026.05.0
        """
        return self._sdt

    # -- tag (metadata) ------------------------------------------------------

    @property
    def tag(self) -> str | None:
        """Programmatic tag value (`w:sdtPr/w:tag/@w:val`), or |None| if not set.

        .. versionadded:: 2026.05.0
        """
        return self._sdt.tag_val

    @tag.setter
    def tag(self, value: str | None) -> None:
        self._sdt.tag_val = value

    # -- title / alias -------------------------------------------------------

    @property
    def title(self) -> str | None:
        """Friendly title (`w:sdtPr/w:alias/@w:val`), or |None| if not set.

        .. versionadded:: 2026.05.0
        """
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

        .. versionadded:: 2026.05.0
        """
        marker = self._sdt.type_marker_tag()
        if marker is None:
            return ContentControlType.RICH_TEXT
        return _MARKER_TYPE_MAP.get(marker, ContentControlType.RICH_TEXT)

    # -- lock ----------------------------------------------------------------

    @property
    def lock(self) -> str | None:
        """Value of `w:sdtPr/w:lock/@w:val`, or |None| when no lock is set.

        One of the :class:`docx.oxml.simpletypes.ST_Lock` members — ``"unlocked"``,
        ``"sdtContentLocked"``, ``"sdtLocked"``, ``"contentLocked"``.

        .. versionadded:: 2026.05.10
        """
        return self._sdt.lock_val

    @lock.setter
    def lock(self, value: str | None) -> None:
        self._sdt.lock_val = value

    # -- id ------------------------------------------------------------------

    @property
    def sdt_id(self) -> int | None:
        """Integer `w:sdtPr/w:id/@w:val` value, or |None| if not present.

        .. versionadded:: 2026.05.0
        """
        return self._sdt.sdt_id

    # -- text ----------------------------------------------------------------

    @property
    def text(self) -> str:
        """Concatenated textual content of this content control.

        .. versionadded:: 2026.05.0
        """
        return self._sdt.text

    @text.setter
    def text(self, value: str) -> None:
        """Replace this control's content with a single paragraph or run containing
        `value`.

        For block-level SDTs (whose `sdtContent` contains `w:p` children), the content
        is replaced with a single `w:p` holding one `w:r` and `w:t` with `value`. For
        inline SDTs (whose `sdtContent` contains `w:r` children), the content is
        replaced with a single `w:r` holding a `w:t` with `value`.

        .. versionadded:: 2026.05.0
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

    # -- data binding --------------------------------------------------------

    @property
    def data_binding(self) -> "DataBinding | None":
        """The |DataBinding| for this content control, or |None| if unbound.

        A content control is "data-bound" when its `w:sdtPr` contains a
        `w:dataBinding` child. The binding ties the SDT's displayed text to an
        XPath expression over a custom XML data part (``/customXml/itemN.xml``).
        python-docx surfaces the binding metadata only — it does not evaluate
        the XPath.

        .. versionadded:: 2026.05.0
        """
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        dataBinding = sdtPr.dataBinding
        if dataBinding is None:
            return None
        return DataBinding(dataBinding)

    def set_data_binding(
        self,
        xpath: str,
        prefix_mappings: str = "",
        store_item_id: str | None = None,
    ) -> "DataBinding":
        """Create or update this content control's `w:dataBinding`.

        `xpath` is the XPath expression the binding points at. `prefix_mappings`
        is a whitespace-separated list of namespace declarations used to
        resolve prefixes in `xpath` (e.g.
        ``"xmlns:ns0='http://example.com/ns'"``). `store_item_id` is the
        ``{GUID}``-formatted id of the target custom XML data part; |None|
        leaves the `@w:storeItemID` attribute unset.

        Returns the resulting |DataBinding|.

        .. versionadded:: 2026.05.0
        """
        sdtPr = self._sdt.get_or_add_sdtPr()
        dataBinding = sdtPr.get_or_add_dataBinding()
        dataBinding.xpath_val = xpath
        dataBinding.prefixMappings = prefix_mappings
        dataBinding.storeItemID = store_item_id
        return DataBinding(dataBinding)

    def remove_data_binding(self) -> None:
        """Remove the `w:dataBinding` child, if present.

        Does nothing when this content control has no data binding.

        .. versionadded:: 2026.05.0
        """
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return
        sdtPr._remove_dataBinding()  # pyright: ignore[reportPrivateUsage]

    # -- checkbox ------------------------------------------------------------

    @property
    def checked(self) -> bool | None:
        """Value of `w14:checkbox/w14:checked/@w14:val` for checkbox SDTs.

        Returns |None| if this is not a checkbox SDT or no `w14:checked` child exists.

        .. versionadded:: 2026.05.0
        """
        return self._sdt.checked

    @checked.setter
    def checked(self, value: bool) -> None:
        self._sdt.checked = value


class DataBinding:
    """Read/write proxy for the `w:dataBinding` child of a content control's `w:sdtPr`.

    A data binding ties a content control to an XPath expression over a custom
    XML data part in the package (``/customXml/itemN.xml``). python-docx
    exposes the binding metadata only — it does not evaluate the XPath or
    synchronize the control's displayed text with the bound value.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, dataBinding: "CT_DataBinding"):
        self._dataBinding = dataBinding

    @property
    def element(self) -> "CT_DataBinding":
        """The underlying `w:dataBinding` lxml element.

        .. versionadded:: 2026.05.0
        """
        return self._dataBinding

    @property
    def prefix_mappings(self) -> str:
        """Value of `@w:prefixMappings` — namespace declarations for `xpath`.

        Returns the empty string when the attribute is not present, matching
        Word's behavior of omitting the attribute when no namespace prefixes
        are required.

        .. versionadded:: 2026.05.0
        """
        value = self._dataBinding.prefixMappings
        return value if value is not None else ""

    @prefix_mappings.setter
    def prefix_mappings(self, value: str) -> None:
        self._dataBinding.prefixMappings = value if value else None

    @property
    def xpath(self) -> str:
        """Value of `@w:xpath` — the XPath expression for this binding.

        Returns the empty string when the attribute is not present.

        .. versionadded:: 2026.05.0
        """
        value = self._dataBinding.xpath_val
        return value if value is not None else ""

    @xpath.setter
    def xpath(self, value: str) -> None:
        self._dataBinding.xpath_val = value if value else None

    @property
    def store_item_id(self) -> str | None:
        """Value of `@w:storeItemID` — `{GUID}` of the target custom XML part.

        |None| when the attribute is not present.

        .. versionadded:: 2026.05.0
        """
        return self._dataBinding.storeItemID

    @store_item_id.setter
    def store_item_id(self, value: str | None) -> None:
        self._dataBinding.storeItemID = value


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

    .. versionadded:: 2026.05.0
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


# ---------------------------------------------------------------------------
# Type-specific proxy subclasses
#
# Each subclass corresponds to a :class:`ContentControlType` value and exposes
# the type-specific ``w:sdtPr`` accessors (list items, date format, checkbox
# state, repeating-section rows, etc.) that are meaningful only for that
# flavour of SDT. Construct them directly (e.g. ``RichTextControl(sdt)``)
# when the type is known in advance, or use :meth:`ContentControl.proxy_for`
# for a typed dispatch. All subclasses inherit :class:`ContentControl`'s
# type-agnostic surface (``.tag``, ``.title``, ``.sdt_id``, ``.text``,
# ``.lock``, data-binding helpers) and can be used interchangeably with the
# base class when only those common members are needed.


class RichTextControl(ContentControl):
    """Rich-text content control proxy.

    Rich-text is the default SDT flavour — ``<w:sdtPr>`` needs no type
    marker. This subclass exists so downstream code can pattern-match on the
    control's Python type instead of inspecting :attr:`ContentControl.type`.

    .. versionadded:: 2026.05.10
    """


class PlainTextControl(ContentControl):
    """Plain-text content control proxy (``<w:text>`` marker).

    Plain-text controls allow only a single run of text; no inline hyperlinks,
    fields, drawings, or nested SDTs. The ``multi_line`` property surfaces the
    ``<w:text>@w:multiLine`` attribute which, when true, permits the user to
    insert soft line breaks.

    .. versionadded:: 2026.05.10
    """

    @property
    def multi_line(self) -> bool:
        """Value of `w:sdtPr/w:text/@w:multiLine` (False when absent).

        .. versionadded:: 2026.05.10
        """
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return False
        text_elm = sdtPr.find(qn("w:text"))
        if text_elm is None:
            return False
        val = text_elm.get(qn("w:multiLine"))
        if val is None:
            return False
        return val in ("1", "true")

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        sdtPr = self._sdt.get_or_add_sdtPr()
        text_elm = sdtPr.find(qn("w:text"))
        if text_elm is None:
            # -- installing a text marker if missing (promotes the SDT to plain text) --
            self._sdt.set_type_marker("w:text")
            text_elm = sdtPr.find(qn("w:text"))
            assert text_elm is not None
        text_elm.set(qn("w:multiLine"), "1" if value else "0")


class PictureControl(ContentControl):
    """Picture content control proxy (``<w:picture>`` marker).

    A picture SDT restricts its contents to a single inline image. python-docx
    does not yet supply image-manipulation helpers on this proxy; the
    subclass is provided for type-identification and round-trip fidelity.

    .. versionadded:: 2026.05.10
    """


class CheckboxControl(ContentControl):
    """Checkbox content control proxy (``<w14:checkbox>`` marker).

    Microsoft Word 2010 extension (namespace
    ``http://schemas.microsoft.com/office/word/2010/wordml``). The
    :attr:`checked` property reflects the ``<w14:checked>@w14:val`` value and
    defaults to |False| when the marker is present without an explicit
    checked-state child.

    .. versionadded:: 2026.05.10
    """


class DateControl(ContentControl):
    """Date-picker content control proxy (``<w:date>`` marker).

    Exposes the ``@w:fullDate`` attribute (ISO-8601 date/datetime) and the
    optional ``<w:dateFormat>`` child that governs display formatting.
    Neither accessor parses or validates the value — both carry verbatim
    string content.

    .. versionadded:: 2026.05.10
    """

    def _date_elm(self) -> "CT_SdtDate | None":
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        return cast("CT_SdtDate | None", sdtPr.find(qn("w:date")))

    def _get_or_add_date_elm(self) -> "CT_SdtDate":
        sdtPr = self._sdt.get_or_add_sdtPr()
        date_elm = cast("CT_SdtDate | None", sdtPr.find(qn("w:date")))
        if date_elm is None:
            self._sdt.set_type_marker("w:date")
            date_elm = cast("CT_SdtDate", sdtPr.find(qn("w:date")))
        return date_elm

    @property
    def full_date(self) -> str | None:
        """Value of `w:sdtPr/w:date/@w:fullDate` — the currently-selected date in
        ISO-8601 form — or |None| when absent.

        .. versionadded:: 2026.05.10
        """
        date_elm = self._date_elm()
        if date_elm is None:
            return None
        return date_elm.fullDate

    @full_date.setter
    def full_date(self, value: str | None) -> None:
        date_elm = self._get_or_add_date_elm()
        date_elm.fullDate = value

    @property
    def date_format(self) -> str | None:
        """Value of `w:sdtPr/w:date/w:dateFormat/@w:val`, or |None| when unset.

        The string is an OOXML date-format specifier (e.g. ``"yyyy-MM-dd"``).

        .. versionadded:: 2026.05.10
        """
        date_elm = self._date_elm()
        if date_elm is None or date_elm.dateFormat is None:
            return None
        return date_elm.dateFormat.get(qn("w:val"))

    @date_format.setter
    def date_format(self, value: str | None) -> None:
        date_elm = self._get_or_add_date_elm()
        df = date_elm.get_or_add_dateFormat()
        if value is None:
            # -- remove @w:val but leave the child (Word tolerates empty dateFormat) --
            if qn("w:val") in df.attrib:
                del df.attrib[qn("w:val")]
            return
        df.set(qn("w:val"), value)


class _ListItemControlMixin:
    """Shared behaviour for drop-down list and combo-box proxies.

    Both SDT flavours carry a ``<w:listItem>`` sequence under a dedicated
    marker element (``<w:dropDownList>`` or ``<w:comboBox>``); this mixin
    centralises the list-access logic so the subclasses only need to name
    their marker tag.
    """

    _sdt: "CT_Sdt"  # populated by ContentControl.__init__
    _marker_tag: str

    def _marker_elm(self) -> "CT_SdtDropDownList | CT_SdtComboBox | None":
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        return cast(
            "CT_SdtDropDownList | CT_SdtComboBox | None",
            sdtPr.find(qn(self._marker_tag)),
        )

    def _get_or_add_marker_elm(self) -> "CT_SdtDropDownList | CT_SdtComboBox":
        sdtPr = self._sdt.get_or_add_sdtPr()
        marker = sdtPr.find(qn(self._marker_tag))
        if marker is None:
            self._sdt.set_type_marker(self._marker_tag)
            marker = sdtPr.find(qn(self._marker_tag))
        return cast("CT_SdtDropDownList | CT_SdtComboBox", marker)

    @property
    def items(self) -> list[str]:
        """The list of ``@w:displayText`` values for this control's ``<w:listItem>``
        children, in document order. Items missing a ``@w:displayText`` fall back
        to their ``@w:value`` attribute; items missing both are represented as
        the empty string.

        .. versionadded:: 2026.05.10
        """
        marker = self._marker_elm()
        if marker is None:
            return []
        result: list[str] = []
        for item in marker.listItem_lst:
            display = item.displayText
            if display is None:
                display = item.value
            result.append(display if display is not None else "")
        return result

    @items.setter
    def items(self, values: list[str]) -> None:
        marker = self._get_or_add_marker_elm()
        # -- clear existing list items --
        for existing in list(marker.findall(qn("w:listItem"))):
            marker.remove(existing)
        for text in values:
            item = marker.add_listItem()
            item.displayText = text
            item.value = text

    def add_item(self, display_text: str, value: str | None = None) -> None:
        """Append a `<w:listItem>` with `display_text` and `value` (defaulting to
        `display_text`).

        .. versionadded:: 2026.05.10
        """
        marker = self._get_or_add_marker_elm()
        item = marker.add_listItem()
        item.displayText = display_text
        item.value = display_text if value is None else value


class DropDownListControl(_ListItemControlMixin, ContentControl):
    """Drop-down-list content control proxy (``<w:dropDownList>`` marker).

    Drop-down lists forbid free-text entry: the user must pick one of the
    predefined ``<w:listItem>`` values.

    .. versionadded:: 2026.05.10
    """

    _marker_tag = "w:dropDownList"


class ComboBoxControl(_ListItemControlMixin, ContentControl):
    """Combo-box content control proxy (``<w:comboBox>`` marker).

    Combo boxes behave like drop-down lists but additionally allow the user
    to type a value that is not in the list; Word records the last free-text
    entry in the ``@w:lastValue`` attribute.

    .. versionadded:: 2026.05.10
    """

    _marker_tag = "w:comboBox"

    @property
    def last_value(self) -> str | None:
        """Value of `w:sdtPr/w:comboBox/@w:lastValue`, or |None| when absent.

        .. versionadded:: 2026.05.10
        """
        marker = cast("CT_SdtComboBox | None", self._marker_elm())
        if marker is None:
            return None
        return marker.lastValue

    @last_value.setter
    def last_value(self, value: str | None) -> None:
        marker = cast("CT_SdtComboBox", self._get_or_add_marker_elm())
        marker.lastValue = value


class BuildingBlockControl(ContentControl):
    """Building-block content control proxy (``<w:docPartObj>`` marker).

    A building-block SDT lets the user pick a preset fragment from a named
    glossary-document gallery. ``gallery``, ``category``, and ``unique`` map
    to the three ``<w:docPartGallery>`` / ``<w:docPartCategory>`` /
    ``<w:docPartUnique>`` children of the marker.

    .. versionadded:: 2026.05.10
    """

    def _doc_part_elm(self) -> "CT_SdtDocPart | None":
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        elm = sdtPr.find(qn("w:docPartObj"))
        if elm is None:
            elm = sdtPr.find(qn("w:docPartList"))
        return cast("CT_SdtDocPart | None", elm)

    def _get_or_add_doc_part_elm(self) -> "CT_SdtDocPart":
        sdtPr = self._sdt.get_or_add_sdtPr()
        elm = sdtPr.find(qn("w:docPartObj"))
        if elm is None:
            elm = sdtPr.find(qn("w:docPartList"))
        if elm is None:
            self._sdt.set_type_marker("w:docPartObj")
            elm = sdtPr.find(qn("w:docPartObj"))
        return cast("CT_SdtDocPart", elm)

    @property
    def gallery(self) -> str | None:
        """Value of `w:docPartObj/w:docPartGallery/@w:val`, or |None| when unset.

        .. versionadded:: 2026.05.10
        """
        elm = self._doc_part_elm()
        if elm is None or elm.docPartGallery is None:
            return None
        return elm.docPartGallery.get(qn("w:val"))

    @gallery.setter
    def gallery(self, value: str | None) -> None:
        elm = self._get_or_add_doc_part_elm()
        if value is None:
            if elm.docPartGallery is not None:
                elm.remove(elm.docPartGallery)
            return
        gallery = elm.get_or_add_docPartGallery()
        gallery.set(qn("w:val"), value)

    @property
    def category(self) -> str | None:
        """Value of `w:docPartObj/w:docPartCategory/@w:val`, or |None| when unset.

        .. versionadded:: 2026.05.10
        """
        elm = self._doc_part_elm()
        if elm is None or elm.docPartCategory is None:
            return None
        return elm.docPartCategory.get(qn("w:val"))

    @category.setter
    def category(self, value: str | None) -> None:
        elm = self._get_or_add_doc_part_elm()
        if value is None:
            if elm.docPartCategory is not None:
                elm.remove(elm.docPartCategory)
            return
        category = elm.get_or_add_docPartCategory()
        category.set(qn("w:val"), value)

    @property
    def unique(self) -> bool:
        """Whether `w:docPartObj/w:docPartUnique` is present.

        Word sets this flag when each glossary-entry instance must appear at
        most once in the document.

        .. versionadded:: 2026.05.10
        """
        elm = self._doc_part_elm()
        if elm is None:
            return False
        return elm.docPartUnique is not None

    @unique.setter
    def unique(self, value: bool) -> None:
        elm = self._get_or_add_doc_part_elm()
        if value:
            elm.get_or_add_docPartUnique()
        else:
            if elm.docPartUnique is not None:
                elm.remove(elm.docPartUnique)


class RepeatingSectionControl(ContentControl):
    """Repeating-section content control proxy (``<w15:repeatingSection>`` marker).

    Microsoft Word 2013+ extension. Wraps a block region (typically a table
    row or an inner block-level SDT) that users can duplicate via Word's
    "Insert New Item" button. Each duplicated instance is itself a
    ``<w:sdt>`` bearing a ``<w15:repeatingSectionItem>`` marker; python-docx
    surfaces those child SDTs via :attr:`rows`.

    .. versionadded:: 2026.05.10
    """

    @property
    def section_title(self) -> str | None:
        """Value of `w:sdtPr/w15:repeatingSection/@w15:sectionTitle`, or |None|.

        .. versionadded:: 2026.05.10
        """
        sdtPr = self._sdt.sdtPr
        if sdtPr is None:
            return None
        marker = sdtPr.find(qn("w15:repeatingSection"))
        if marker is None:
            return None
        return marker.get(qn("w15:sectionTitle"))

    @section_title.setter
    def section_title(self, value: str | None) -> None:
        sdtPr = self._sdt.get_or_add_sdtPr()
        marker = sdtPr.find(qn("w15:repeatingSection"))
        if marker is None:
            self._sdt.set_type_marker("w15:repeatingSection")
            marker = sdtPr.find(qn("w15:repeatingSection"))
            assert marker is not None
        if value is None:
            if qn("w15:sectionTitle") in marker.attrib:
                del marker.attrib[qn("w15:sectionTitle")]
            return
        marker.set(qn("w15:sectionTitle"), value)

    def _iter_row_sdts(self) -> Iterator["CT_Sdt"]:
        sdtContent = self._sdt.sdtContent
        if sdtContent is None:
            return
        # -- each instance is a <w:sdt> carrying <w15:repeatingSectionItem> --
        for child in sdtContent:
            if child.tag != qn("w:sdt"):
                continue
            inner_sdtPr = child.find(qn("w:sdtPr"))
            if inner_sdtPr is None:
                continue
            if inner_sdtPr.find(qn("w15:repeatingSectionItem")) is not None:
                yield cast("CT_Sdt", child)

    @property
    def rows(self) -> list["ContentControl"]:
        """List of per-row |ContentControl| instances, one per
        ``<w15:repeatingSectionItem>`` child SDT.

        Each row is itself a content control whose proxy class is determined
        by its own ``<w:sdtPr>`` marker, so callers can descend into nested
        typed controls (e.g. a per-row :class:`DateControl`).

        .. versionadded:: 2026.05.10
        """
        return [ContentControl.proxy_for(sdt) for sdt in self._iter_row_sdts()]

    def add_row(self) -> "ContentControl":
        """Append a new repeating-section row and return its proxy.

        A fresh ``<w:sdt>`` carrying a ``<w15:repeatingSectionItem>`` marker is
        created and inserted as the last child of this control's
        ``<w:sdtContent>``. The row has an empty ``<w:sdtContent>/<w:p>``
        body that callers can populate.

        .. versionadded:: 2026.05.10
        """
        sdtContent = self._sdt.get_or_add_sdtContent()
        inner = cast("CT_Sdt", OxmlElement("w:sdt"))
        inner_sdtPr = inner.get_or_add_sdtPr()
        inner_sdtPr.append(OxmlElement("w15:repeatingSectionItem"))
        inner.sdt_id = _new_sdt_id()
        inner_sdtContent = inner.get_or_add_sdtContent()
        inner_sdtContent.append(OxmlElement("w:p"))
        sdtContent.append(inner)
        return ContentControl.proxy_for(inner)


# ---------------------------------------------------------------------------
# dispatch helper
#
# ``ContentControl.proxy_for`` is the public entry point that callers use
# when they have a raw ``CT_Sdt`` and want a typed proxy selected by the
# SDT's ``w:sdtPr`` marker. It is implemented at module scope (rather than
# inside the class body) to keep the subclass map textually close to the
# subclass definitions.


_TYPE_PROXY_MAP: "dict[ContentControlType, type[ContentControl]]" = {
    ContentControlType.RICH_TEXT: RichTextControl,
    ContentControlType.PLAIN_TEXT: PlainTextControl,
    ContentControlType.PICTURE: PictureControl,
    ContentControlType.CHECKBOX: CheckboxControl,
    ContentControlType.DATE: DateControl,
    ContentControlType.DROPDOWN: DropDownListControl,
    ContentControlType.COMBO_BOX: ComboBoxControl,
    ContentControlType.REPEATING_SECTION: RepeatingSectionControl,
    ContentControlType.BUILDING_BLOCK: BuildingBlockControl,
}


def _proxy_for(sdt: "CT_Sdt") -> ContentControl:
    """Return the most specific :class:`ContentControl` subclass for `sdt`.

    The subclass is chosen by inspecting the SDT's ``w:sdtPr`` type marker;
    when no marker is present (the schema default) the rich-text subclass is
    returned.
    """
    marker = sdt.type_marker_tag()
    if marker is None:
        return RichTextControl(sdt)
    type_ = _MARKER_TYPE_MAP.get(marker, ContentControlType.RICH_TEXT)
    cls = _TYPE_PROXY_MAP.get(type_, ContentControl)
    return cls(sdt)


# -- expose the dispatcher as a classmethod on ContentControl --
ContentControl.proxy_for = staticmethod(_proxy_for)  # type: ignore[attr-defined]


# ---------------------------------------------------------------------------
# public re-export of the typed subclasses — importing from this module is
# the supported API surface. ``AnyControl`` is a convenient union type for
# callers that want to annotate collections or return values without pinning
# to the base class.

AnyControl = Union[
    ContentControl,
    RichTextControl,
    PlainTextControl,
    PictureControl,
    CheckboxControl,
    DateControl,
    DropDownListControl,
    ComboBoxControl,
    BuildingBlockControl,
    RepeatingSectionControl,
]
