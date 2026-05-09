"""Custom element classes related to structured document tags (content controls).

Structured document tags (SDTs), commonly referred to as "content controls", come in
both block-level and inline flavors. The oxml element `w:sdt` represents both; whether
it is block-level or inline is determined by its parent (direct child of `w:body`, a
table cell, etc. for block-level; child of a paragraph for inline).

The XML shape is::

    <w:sdt>
      <w:sdtPr>
        <w:tag w:val="name"/>
        <w:alias w:val="title"/>
        <w:id w:val="12345"/>
        ... optional type marker ...
      </w:sdtPr>
      <w:sdtContent>
        <w:p>...</w:p>     <!-- block-level -->
        <w:r>...</w:r>     <!-- inline -->
      </w:sdtContent>
    </w:sdt>
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast
from collections.abc import Callable

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.simpletypes import (
    ST_OnOff,
    ST_SdtDateMappingType,
    ST_String,
)
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.shared import CT_OnOff as CT_OnOffElement
    from docx.oxml.shared import CT_String as CT_StringElement
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.run import CT_R


class CT_Sdt(BaseOxmlElement):
    """``<w:sdt>`` element - a structured document tag (content control)."""

    sdtPr: "CT_SdtPr | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtPr", successors=("w:sdtEndPr", "w:sdtContent")
    )
    sdtContent: "CT_SdtContent | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtContent", successors=()
    )

    get_or_add_sdtPr: Callable[[], "CT_SdtPr"]
    get_or_add_sdtContent: Callable[[], "CT_SdtContent"]

    # NOTE: the names `tag_val` and `alias_val` are used instead of `tag` / `alias`
    # because `tag` is an lxml `_Element` built-in attribute that MUST NOT be shadowed.
    # The :class:`ContentControl` proxy re-exposes these under the friendlier names
    # `tag` and `title`.

    @property
    def tag_val(self) -> str | None:
        """Value of `w:sdtPr/w:tag/@w:val`, or |None| if not present."""
        sdtPr = self.sdtPr
        if sdtPr is None:
            return None
        return sdtPr.tag_val

    @tag_val.setter
    def tag_val(self, value: str | None) -> None:
        sdtPr = self.get_or_add_sdtPr()
        sdtPr.tag_val = value

    @property
    def alias_val(self) -> str | None:
        """Value of `w:sdtPr/w:alias/@w:val`, or |None| if not present."""
        sdtPr = self.sdtPr
        if sdtPr is None:
            return None
        return sdtPr.alias_val

    @alias_val.setter
    def alias_val(self, value: str | None) -> None:
        sdtPr = self.get_or_add_sdtPr()
        sdtPr.alias_val = value

    @property
    def sdt_id(self) -> int | None:
        """Value of `w:sdtPr/w:id/@w:val`, or |None| if not present."""
        sdtPr = self.sdtPr
        if sdtPr is None:
            return None
        id_elm = sdtPr.find(qn("w:id"))
        if id_elm is None:
            return None
        val = id_elm.get(qn("w:val"))
        if val is None:
            return None
        try:
            return int(val)
        except (TypeError, ValueError):
            return None

    @sdt_id.setter
    def sdt_id(self, value: int | None) -> None:
        sdtPr = self.get_or_add_sdtPr()
        sdtPr.set_id(value)

    @property
    def text(self) -> str:
        """Concatenated text from this SDT's `w:sdtContent`.

        Includes text from paragraph children (block-level) and run children (inline).
        """
        sdtContent = self.sdtContent
        if sdtContent is None:
            return ""
        return sdtContent.text

    @property
    def checked(self) -> bool | None:
        """Value of `w:sdtPr/w14:checkbox/w14:checked/@w14:val` for checkbox SDTs.

        Returns |None| if the SDT has no checkbox marker or no `w14:checked` element.
        """
        sdtPr = self.sdtPr
        if sdtPr is None:
            return None
        checkbox = sdtPr.find(qn("w14:checkbox"))
        if checkbox is None:
            return None
        checked_elm = checkbox.find(qn("w14:checked"))
        if checked_elm is None:
            return None
        val = checked_elm.get(qn("w14:val"))
        if val is None:
            # -- presence of `w14:checked` without @val implies checked --
            return True
        return val in ("1", "true")

    @checked.setter
    def checked(self, value: bool) -> None:
        sdtPr = self.get_or_add_sdtPr()
        checkbox = sdtPr.find(qn("w14:checkbox"))
        if checkbox is None:
            # -- create the checkbox marker too --
            checkbox = OxmlElement("w14:checkbox")
            sdtPr.append(checkbox)
        checked_elm = checkbox.find(qn("w14:checked"))
        if checked_elm is None:
            checked_elm = OxmlElement("w14:checked")
            checkbox.append(checked_elm)
        checked_elm.set(qn("w14:val"), "1" if value else "0")

    def set_type_marker(self, marker_tag: str) -> None:
        """Unconditionally set a type-marker child element on `sdtPr`.

        `marker_tag` is a namespace-prefixed tag name like 'w:text', 'w:comboBox',
        'w:dropDownList', 'w:date', 'w:picture', or 'w14:checkbox'.
        """
        sdtPr = self.get_or_add_sdtPr()
        # -- remove any existing known type markers --
        for mtag in (
            "w:text",
            "w:comboBox",
            "w:dropDownList",
            "w:date",
            "w:picture",
            "w14:checkbox",
            "w:richText",
        ):
            for existing in sdtPr.findall(qn(mtag)):
                sdtPr.remove(existing)
        sdtPr.append(OxmlElement(marker_tag))

    def type_marker_tag(self) -> str | None:
        """Return namespace-prefixed tag of the first known type-marker child of sdtPr.

        Returns |None| if no known type marker is present (which is valid - it implies
        rich-text by default).
        """
        sdtPr = self.sdtPr
        if sdtPr is None:
            return None
        for mtag in (
            "w14:checkbox",
            "w:text",
            "w:comboBox",
            "w:dropDownList",
            "w:date",
            "w:picture",
            "w:richText",
        ):
            if sdtPr.find(qn(mtag)) is not None:
                return mtag
        return None


class CT_SdtPr(BaseOxmlElement):
    """``<w:sdtPr>`` element - properties for a structured document tag."""

    @property
    def tag_val(self) -> str | None:
        """Value of `w:tag/@w:val` child element, or |None| if not present."""
        tag_elm = self.find(qn("w:tag"))
        if tag_elm is None:
            return None
        return tag_elm.get(qn("w:val"))

    @tag_val.setter
    def tag_val(self, value: str | None) -> None:
        tag_elm = self.find(qn("w:tag"))
        if value is None:
            if tag_elm is not None:
                self.remove(tag_elm)
            return
        if tag_elm is None:
            tag_elm = OxmlElement("w:tag")
            # -- insert at start so it comes before type markers (schema tolerant) --
            self.insert(0, tag_elm)
        tag_elm.set(qn("w:val"), value)

    @property
    def alias_val(self) -> str | None:
        """Value of `w:alias/@w:val` child element, or |None| if not present."""
        alias_elm = self.find(qn("w:alias"))
        if alias_elm is None:
            return None
        return alias_elm.get(qn("w:val"))

    @alias_val.setter
    def alias_val(self, value: str | None) -> None:
        alias_elm = self.find(qn("w:alias"))
        if value is None:
            if alias_elm is not None:
                self.remove(alias_elm)
            return
        if alias_elm is None:
            alias_elm = OxmlElement("w:alias")
            self.insert(0, alias_elm)
        alias_elm.set(qn("w:val"), value)

    def set_id(self, value: int | None) -> None:
        """Set value of `w:id/@w:val` child element, removing it when `value` is None."""
        id_elm = self.find(qn("w:id"))
        if value is None:
            if id_elm is not None:
                self.remove(id_elm)
            return
        if id_elm is None:
            id_elm = OxmlElement("w:id")
            self.append(id_elm)
        id_elm.set(qn("w:val"), str(value))

    @property
    def dataBinding(self) -> "CT_DataBinding | None":
        """The `w:dataBinding` child element, or |None| when not present."""
        return cast("CT_DataBinding | None", self.find(qn("w:dataBinding")))

    def get_or_add_dataBinding(self) -> "CT_DataBinding":
        """Return the `w:dataBinding` child, creating it if not already present."""
        dataBinding = self.dataBinding
        if dataBinding is None:
            dataBinding = cast("CT_DataBinding", OxmlElement("w:dataBinding"))
            # -- append: place after other sdtPr children. Word tolerates
            #    w:dataBinding anywhere in sdtPr, though the XSD places it
            #    near the end (before any type marker). --
            self.append(dataBinding)
        return dataBinding

    def _remove_dataBinding(self) -> None:
        """Remove the `w:dataBinding` child element, if present."""
        dataBinding = self.dataBinding
        if dataBinding is not None:
            self.remove(dataBinding)


class CT_DataBinding(BaseOxmlElement):
    """``<w:dataBinding>`` element — ties an SDT to an XPath over a custom XML part.

    The attributes are defined by ECMA-376 as:

    - ``@w:prefixMappings`` — a whitespace-separated list of namespace declarations used
      to resolve prefixes in ``@w:xpath``.
    - ``@w:xpath`` — the XPath expression (required by the schema but we treat it as
      optional here for read resiliency).
    - ``@w:storeItemID`` — the ``{GUID}``-formatted identifier of the target custom XML
      data part.

    Live evaluation of ``@w:xpath`` against the referenced data part is **not** in
    scope for this class; it carries the metadata verbatim.

    NOTE: The Python attribute exposing ``@w:xpath`` is named ``xpath_val`` to avoid
    shadowing :meth:`BaseOxmlElement.xpath`, which is lxml's XPath query method.
    """

    prefixMappings: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:prefixMappings", ST_String
    )
    xpath_val: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:xpath", ST_String
    )
    storeItemID: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:storeItemID", ST_String
    )


class CT_SdtContent(BaseOxmlElement):
    """``<w:sdtContent>`` element - the current content of a structured document tag."""

    p = cast(
        "Callable[[], list[CT_P]]",
        property(lambda self: self.findall(qn("w:p"))),
    )
    r = cast(
        "Callable[[], list[CT_R]]",
        property(lambda self: self.findall(qn("w:r"))),
    )

    @property
    def p_lst(self) -> list["CT_P"]:
        """List of `w:p` children of this sdtContent element."""
        return self.findall(qn("w:p"))

    @property
    def r_lst(self) -> list["CT_R"]:
        """List of `w:r` children of this sdtContent element."""
        return self.findall(qn("w:r"))

    @property
    def text(self) -> str:
        """Concatenated textual content of this sdtContent.

        Combines text of child paragraphs (joined with newlines between them) and
        text from direct run children (inline case). The implementation concatenates
        child text in document order.
        """
        parts: list[str] = []
        for child in self:
            tag = child.tag
            if tag == qn("w:p"):
                parts.append(child.text)  # CT_P.text
            elif tag == qn("w:r"):
                parts.append(child.text)  # CT_R.text
            elif tag == qn("w:sdt"):
                # -- nested SDT --
                parts.append(cast("CT_Sdt", child).text)
        return "".join(parts)


# ---------------------------------------------------------------------------
# SDT property-value types
#
# These classes model the *value-carrying* children of ``<w:sdtPr>`` that
# distinguish the different content-control flavours (date picker, combo
# box, drop-down list, doc-part gallery, plain text). They correspond to
# ``CT_SdtListItem``, ``CT_SdtDate``, ``CT_SdtComboBox``,
# ``CT_SdtDocPart``, ``CT_SdtDropDownList``, ``CT_SdtText``,
# ``CT_SdtDateMappingType``, and ``CT_SdtEndPr`` in ``wml.xsd``.
#
# Each class is a verbatim port of the XSD complexType: attribute names,
# child-element names, and sequence ordering match the schema exactly so
# that ``successors`` tuples are authoritative for insertion.


class CT_SdtListItem(BaseOxmlElement):
    """``<w:listItem>`` element — a single item in a combo-box / drop-down list.

    Carries ``@w:displayText`` (what the user sees) and ``@w:value`` (what is
    written into the bound data part when the item is selected). Both are
    optional per the ECMA-376 schema.
    """

    displayText: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:displayText", ST_String
    )
    value: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:value", ST_String
    )


class CT_SdtDateMappingType(BaseOxmlElement):
    """``<w:storeMappedDataAs>`` element — bound-data serialisation format.

    Used inside ``<w:date>`` to tell Word whether to write the user-entered
    date back into the bound custom-XML part as plain text, ``xsd:date``, or
    ``xsd:dateTime``.
    """

    val: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_SdtDateMappingType
    )


class CT_SdtDate(BaseOxmlElement):
    """``<w:date>`` element — date-picker type marker and properties.

    Its presence under ``<w:sdtPr>`` turns the SDT into a date-picker.
    Carries the optional ``@w:fullDate`` attribute (the currently-chosen
    date, in ISO-8601) plus optional ``dateFormat``, ``lid``,
    ``storeMappedDataAs`` and ``calendar`` children.
    """

    dateFormat: "CT_StringElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:dateFormat",
        successors=("w:lid", "w:storeMappedDataAs", "w:calendar"),
    )
    lid: "CT_StringElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:lid", successors=("w:storeMappedDataAs", "w:calendar")
    )
    storeMappedDataAs: "CT_SdtDateMappingType | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:storeMappedDataAs", successors=("w:calendar",)
    )
    calendar: "CT_StringElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:calendar", successors=()
    )

    fullDate: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:fullDate", ST_String
    )


class CT_SdtComboBox(BaseOxmlElement):
    """``<w:comboBox>`` element — combo-box type marker and item list.

    Carries zero-or-more ``<w:listItem>`` children plus an optional
    ``@w:lastValue`` that records the most recent free-text input (combo
    boxes allow the user to type a value that is not in the list).
    """

    listItem = ZeroOrMore("w:listItem", successors=())
    listItem_lst: list[CT_SdtListItem]
    add_listItem: Callable[..., CT_SdtListItem]

    lastValue: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:lastValue", ST_String
    )


class CT_SdtDocPart(BaseOxmlElement):
    """``<w:docPartObj>`` / ``<w:docPartList>`` element — doc-part gallery marker.

    Both ``docPartObj`` and ``docPartList`` share this type. Carries the
    optional ``docPartGallery``, ``docPartCategory`` and ``docPartUnique``
    child elements that constrain which glossary entries the control offers.
    """

    docPartGallery: "CT_StringElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartGallery", successors=("w:docPartCategory", "w:docPartUnique")
    )
    docPartCategory: "CT_StringElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartCategory", successors=("w:docPartUnique",)
    )
    docPartUnique: "CT_OnOffElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartUnique", successors=()
    )


class CT_SdtDropDownList(BaseOxmlElement):
    """``<w:dropDownList>`` element — drop-down list type marker and items.

    Same shape as :class:`CT_SdtComboBox` but drop-down lists forbid free-text
    input: the user must pick one of the ``<w:listItem>`` values.
    """

    listItem = ZeroOrMore("w:listItem", successors=())
    listItem_lst: list[CT_SdtListItem]
    add_listItem: Callable[..., CT_SdtListItem]

    lastValue: "str | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:lastValue", ST_String
    )


class CT_SdtText(BaseOxmlElement):
    """``<w:text>`` element — plain-text type marker.

    Presence under ``<w:sdtPr>`` restricts the SDT's editable surface to a
    single run of plain text. Carries the optional ``@w:multiLine`` attribute
    (default ``false``) that lets the user insert line breaks.
    """

    multiLine: "bool | None" = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:multiLine", ST_OnOff
    )


class CT_SdtEndPr(BaseOxmlElement):
    """``<w:sdtEndPr>`` element — run-property marker applied at SDT end.

    Holds zero or more ``<w:rPr>`` children that Word uses to reset the
    effective run-property stack when the content control closes.
    """

    rPr = ZeroOrMore("w:rPr", successors=())
