"""Custom element classes for structured document tags (content controls)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.run import CT_R


class CT_SdtCheckbox(BaseOxmlElement):
    """`w14:checkbox` element, specifying checkbox state."""

    @property
    def checked(self) -> bool:
        """True if the checkbox is checked."""
        checked_elms = self.xpath("./w14:checked/@w14:val")
        if checked_elms:
            return checked_elms[0] in ("1", "true")
        return False

    @checked.setter
    def checked(self, value: bool) -> None:
        checked_elms = self.xpath("./w14:checked")
        if checked_elms:
            checked_elms[0].set(qn("w14:val"), "1" if value else "0")
        else:
            checked_elm = OxmlElement("w14:checked")
            checked_elm.set(qn("w14:val"), "1" if value else "0")
            self.append(checked_elm)


class CT_SdtListItem(BaseOxmlElement):
    """`w:listItem` element within a combo box or drop-down list."""

    displayText: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:displayText", ST_String
    )
    value: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:value", ST_String
    )


class CT_SdtComboBox(BaseOxmlElement):
    """`w:comboBox` element."""

    listItem = ZeroOrMore("w:listItem", successors=())


class CT_SdtDropDownList(BaseOxmlElement):
    """`w:dropDownList` element."""

    listItem = ZeroOrMore("w:listItem", successors=())


class CT_SdtDate(BaseOxmlElement):
    """`w:date` element."""

    fullDate: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:fullDate", ST_String
    )


class CT_SdtPr(BaseOxmlElement):
    """`w:sdtPr` element, containing properties for a structured document tag."""

    @property
    def tag_val(self) -> str | None:
        """The value of the `w:tag` child's `w:val` attribute, or None."""
        tags = self.xpath("./w:tag/@w:val")
        return tags[0] if tags else None

    @tag_val.setter
    def tag_val(self, value: str | None) -> None:
        for t in self.xpath("./w:tag"):
            self.remove(t)
        if value is not None:
            tag_elm = OxmlElement("w:tag")
            tag_elm.set(qn("w:val"), value)
            self.append(tag_elm)

    @property
    def alias_val(self) -> str | None:
        """The value of the `w:alias` child's `w:val` attribute, or None (title)."""
        aliases = self.xpath("./w:alias/@w:val")
        return aliases[0] if aliases else None

    @alias_val.setter
    def alias_val(self, value: str | None) -> None:
        for a in self.xpath("./w:alias"):
            self.remove(a)
        if value is not None:
            alias_elm = OxmlElement("w:alias")
            alias_elm.set(qn("w:val"), value)
            self.append(alias_elm)

    @property
    def sdt_type(self) -> str:
        """The type of this SDT, determined by which type-specific child is present."""
        if self.xpath("./w14:checkbox"):
            return "checkbox"
        if self.xpath("./w:comboBox"):
            return "comboBox"
        if self.xpath("./w:dropDownList"):
            return "dropDown"
        if self.xpath("./w:date"):
            return "date"
        if self.xpath("./w:picture"):
            return "picture"
        if self.xpath("./w:text"):
            return "plainText"
        return "richText"

    @property
    def checkbox(self) -> CT_SdtCheckbox | None:
        """The `w14:checkbox` child element, or None."""
        results = self.xpath("./w14:checkbox")
        return results[0] if results else None


class CT_SdtContent(BaseOxmlElement):
    """`w:sdtContent` element, the container for SDT content."""

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())
    r = ZeroOrMore("w:r", successors=())

    add_p: Callable[[], CT_P]
    add_r: Callable[[], CT_R]
    p_lst: List[CT_P]
    r_lst: List[CT_R]

    @property
    def text(self) -> str:
        """The text content of this SDT content element."""
        paras = self.xpath("./w:p")
        if paras:
            return "\n".join(p.text for p in paras)
        # -- inline SDT: collect text from runs --
        runs = self.xpath("./w:r")
        return "".join(r.text for r in runs)


class CT_Sdt(BaseOxmlElement):
    """`w:sdt` element, used for both block-level and inline content controls.

    lxml registers one class per tag name, so this unified class handles both the
    block-level case (child of `w:body`) and the inline case (child of `w:p`).
    """

    sdtPr: CT_SdtPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtPr", successors=("w:sdtEndPr", "w:sdtContent")
    )
    sdtContent: CT_SdtContent | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtContent", successors=()
    )

    get_or_add_sdtPr: Callable[[], CT_SdtPr]
    get_or_add_sdtContent: Callable[[], CT_SdtContent]

    @classmethod
    def new_block(cls, sdt_type: str, tag: str | None = None, title: str | None = None) -> CT_Sdt:
        """Return a new block-level `w:sdt` element with a paragraph in its content."""
        xml = (
            f"<w:sdt {nsdecls('w', 'w14')}>"
            f"  <w:sdtPr/>"
            f"  <w:sdtContent>"
            f"    <w:p/>"
            f"  </w:sdtContent>"
            f"</w:sdt>"
        )
        sdt = cast(CT_Sdt, parse_xml(xml))
        sdtPr = sdt.sdtPr
        assert sdtPr is not None
        _configure_sdtPr(sdtPr, sdt_type, tag, title)
        return sdt

    @classmethod
    def new_inline(cls, sdt_type: str, tag: str | None = None, title: str | None = None) -> CT_Sdt:
        """Return a new inline `w:sdt` element with a run in its content."""
        xml = (
            f"<w:sdt {nsdecls('w', 'w14')}>"
            f"  <w:sdtPr/>"
            f"  <w:sdtContent>"
            f"    <w:r/>"
            f"  </w:sdtContent>"
            f"</w:sdt>"
        )
        sdt = cast(CT_Sdt, parse_xml(xml))
        sdtPr = sdt.sdtPr
        assert sdtPr is not None
        _configure_sdtPr(sdtPr, sdt_type, tag, title)
        return sdt

    @property
    def is_block_level(self) -> bool:
        """True if this SDT is block-level (contains paragraphs/tables)."""
        from docx.oxml.document import CT_Body

        parent = self.getparent()
        return isinstance(parent, CT_Body)

    @property
    def inner_content_elements(self) -> list[CT_P]:
        """All `w:p` and `w:tbl` elements inside this SDT's content."""
        return self.xpath("./w:sdtContent/w:p | ./w:sdtContent/w:tbl")


def _configure_sdtPr(sdtPr: CT_SdtPr, sdt_type: str, tag: str | None, title: str | None) -> None:
    """Configure the `w:sdtPr` element with the given type, tag, and title."""
    if tag is not None:
        sdtPr.tag_val = tag
    if title is not None:
        sdtPr.alias_val = title

    if sdt_type == "plainText":
        sdtPr.append(OxmlElement("w:text"))
    elif sdt_type == "checkbox":
        checkbox_xml = (
            f"<w14:checkbox {nsdecls('w14')}>"
            f'  <w14:checked w14:val="0"/>'
            f'  <w14:checkedState w14:val="2612"/>'
            f'  <w14:uncheckedState w14:val="2610"/>'
            f"</w14:checkbox>"
        )
        sdtPr.append(parse_xml(checkbox_xml))
    elif sdt_type == "comboBox":
        sdtPr.append(OxmlElement("w:comboBox"))
    elif sdt_type == "dropDown":
        sdtPr.append(OxmlElement("w:dropDownList"))
    elif sdt_type == "date":
        sdtPr.append(OxmlElement("w:date"))
    elif sdt_type == "picture":
        sdtPr.append(OxmlElement("w:picture"))
    # richText is the default (no extra child element needed)
