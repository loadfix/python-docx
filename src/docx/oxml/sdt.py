"""Custom element classes related to structured document tags (content controls)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.enum.contentcontrol import WD_CONTENT_CONTROL_TYPE
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.run import CT_R


class CT_SdtCheckbox(BaseOxmlElement):
    """`w14:checkbox` element, specifying that the SDT is a checkbox."""

    @property
    def checked(self) -> bool:
        """True if the checkbox is checked."""
        checked_elm = self.find(qn("w14:checked"))
        if checked_elm is None:
            return False
        val = checked_elm.get(qn("w14:val"))
        return val in ("1", "true")

    @checked.setter
    def checked(self, value: bool) -> None:
        checked_elm = self.find(qn("w14:checked"))
        if checked_elm is None:
            checked_elm = OxmlElement("w14:checked")
            self.insert(0, checked_elm)
        checked_elm.set(qn("w14:val"), "1" if value else "0")


class CT_SdtPr(BaseOxmlElement):
    """`w:sdtPr` element, containing properties for a structured document tag."""

    tag_elm: CT_SdtTag | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:tag",
        successors=(
            "w:lock",
            "w:placeholder",
            "w:showingPlcHdr",
            "w:dataBinding",
            "w:temporary",
            "w:text",
            "w14:checkbox",
            "w:comboBox",
            "w:dropDownList",
            "w:date",
            "w:picture",
        ),
    )
    alias: CT_SdtAlias | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:alias",
        successors=(
            "w:tag",
            "w:lock",
            "w:placeholder",
            "w:showingPlcHdr",
            "w:dataBinding",
            "w:temporary",
            "w:text",
            "w14:checkbox",
            "w:comboBox",
            "w:dropDownList",
            "w:date",
            "w:picture",
        ),
    )

    get_or_add_tag_elm: Callable[[], CT_SdtTag]
    _remove_tag_elm: Callable[[], None]
    get_or_add_alias: Callable[[], CT_SdtAlias]
    _remove_alias: Callable[[], None]

    @property
    def tag_val(self) -> str | None:
        """Value of `w:tag/@w:val` or None."""
        tag = self.tag_elm
        if tag is None:
            return None
        return tag.val

    @tag_val.setter
    def tag_val(self, value: str | None) -> None:
        if value is None:
            self._remove_tag_elm()
        else:
            self.get_or_add_tag_elm().val = value

    @property
    def title(self) -> str | None:
        """Value of `w:alias/@w:val` or None."""
        alias_elm = self.alias
        if alias_elm is None:
            return None
        return alias_elm.val

    @title.setter
    def title(self, value: str | None) -> None:
        if value is None:
            self._remove_alias()
        else:
            self.get_or_add_alias().val = value

    @property
    def control_type(self) -> WD_CONTENT_CONTROL_TYPE:
        """The type of this content control."""
        if self.find(qn("w:text")) is not None:
            return WD_CONTENT_CONTROL_TYPE.PLAIN_TEXT
        if self.find(qn("w14:checkbox")) is not None:
            return WD_CONTENT_CONTROL_TYPE.CHECKBOX
        if self.find(qn("w:comboBox")) is not None:
            return WD_CONTENT_CONTROL_TYPE.COMBO_BOX
        if self.find(qn("w:dropDownList")) is not None:
            return WD_CONTENT_CONTROL_TYPE.DROP_DOWN
        if self.find(qn("w:date")) is not None:
            return WD_CONTENT_CONTROL_TYPE.DATE
        if self.find(qn("w:picture")) is not None:
            return WD_CONTENT_CONTROL_TYPE.PICTURE
        return WD_CONTENT_CONTROL_TYPE.RICH_TEXT

    @property
    def checkbox(self) -> CT_SdtCheckbox | None:
        """The `w14:checkbox` child element or None."""
        return self.find(qn("w14:checkbox"))  # pyright: ignore[reportReturnType]


class CT_SdtTag(BaseOxmlElement):
    """`w:tag` element within `w:sdtPr`."""

    val: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String, default=""
    )


class CT_SdtAlias(BaseOxmlElement):
    """`w:alias` element within `w:sdtPr` (title/alias of the SDT)."""

    val: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String, default=""
    )


class CT_SdtContent(BaseOxmlElement):
    """`w:sdtContent` element, used for both block-level and inline SDTs.

    When block-level, contains paragraphs and tables.
    When inline, contains runs.
    """

    add_p: Callable[[], CT_P]
    add_r: Callable[[], CT_R]
    p_lst: List[CT_P]
    r_lst: List[CT_R]
    tbl_lst: List[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    p = ZeroOrMore("w:p", successors=())
    r = ZeroOrMore("w:r", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        return self.xpath("./w:p | ./w:tbl")

    @property
    def text(self) -> str:
        """The concatenated text of all runs (for inline SDTs) or paragraphs (for block)."""
        # -- inline case: has runs directly --
        runs = self.r_lst
        if runs:
            return "".join(r.text for r in runs)
        # -- block case: has paragraphs --
        paragraphs = self.p_lst
        if paragraphs:
            return "\n".join(p.text for p in paragraphs)
        return ""


class CT_Sdt(BaseOxmlElement):
    """`w:sdt` element, a structured document tag (content control).

    Used for both block-level (child of `w:body`) and inline (child of `w:p`) SDTs.
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
    def new_block(
        cls,
        control_type: WD_CONTENT_CONTROL_TYPE = WD_CONTENT_CONTROL_TYPE.RICH_TEXT,
        tag: str | None = None,
        title: str | None = None,
    ) -> CT_Sdt:
        """Return a new block-level `w:sdt` element with a paragraph in its content."""
        type_xml = _type_element_xml(control_type)
        tag_xml = f'<w:tag {nsdecls("w")} w:val="{tag}"/>' if tag else ""
        alias_xml = f'<w:alias {nsdecls("w")} w:val="{title}"/>' if title else ""
        xml = (
            f"<w:sdt {nsdecls('w', 'w14')}>"
            f"  <w:sdtPr>{alias_xml}{tag_xml}{type_xml}</w:sdtPr>"
            f"  <w:sdtContent><w:p/></w:sdtContent>"
            f"</w:sdt>"
        )
        return cast(CT_Sdt, parse_xml(xml))

    @classmethod
    def new_inline(
        cls,
        control_type: WD_CONTENT_CONTROL_TYPE = WD_CONTENT_CONTROL_TYPE.RICH_TEXT,
        tag: str | None = None,
        title: str | None = None,
    ) -> CT_Sdt:
        """Return a new inline `w:sdt` element with a run in its content."""
        type_xml = _type_element_xml(control_type)
        tag_xml = f'<w:tag {nsdecls("w")} w:val="{tag}"/>' if tag else ""
        alias_xml = f'<w:alias {nsdecls("w")} w:val="{title}"/>' if title else ""
        xml = (
            f"<w:sdt {nsdecls('w', 'w14')}>"
            f"  <w:sdtPr>{alias_xml}{tag_xml}{type_xml}</w:sdtPr>"
            f"  <w:sdtContent><w:r><w:t/></w:r></w:sdtContent>"
            f"</w:sdt>"
        )
        return cast(CT_Sdt, parse_xml(xml))


def _type_element_xml(control_type: WD_CONTENT_CONTROL_TYPE) -> str:
    """Return the XML fragment for the type-specific element in `w:sdtPr`."""
    nsdecls_w = nsdecls("w")
    nsdecls_w14 = nsdecls("w14")
    type_map = {
        WD_CONTENT_CONTROL_TYPE.PLAIN_TEXT: f"<w:text {nsdecls_w}/>",
        WD_CONTENT_CONTROL_TYPE.CHECKBOX: (
            f"<w14:checkbox {nsdecls_w14}>"
            f'<w14:checked w14:val="0"/>'
            f'<w14:checkedState w14:val="2612"/>'
            f'<w14:uncheckedState w14:val="2610"/>'
            f"</w14:checkbox>"
        ),
        WD_CONTENT_CONTROL_TYPE.COMBO_BOX: f"<w:comboBox {nsdecls_w}/>",
        WD_CONTENT_CONTROL_TYPE.DROP_DOWN: f"<w:dropDownList {nsdecls_w}/>",
        WD_CONTENT_CONTROL_TYPE.DATE: f"<w:date {nsdecls_w}/>",
        WD_CONTENT_CONTROL_TYPE.PICTURE: f"<w:picture {nsdecls_w}/>",
        WD_CONTENT_CONTROL_TYPE.RICH_TEXT: "",
    }
    return type_map.get(control_type, "")
