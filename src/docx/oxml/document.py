"""Custom element classes that correspond to the document part, e.g. <w:document>."""

from __future__ import annotations

from typing import TYPE_CHECKING
from collections.abc import Callable

from docx.oxml.ns import qn
from docx.oxml.section import CT_SectPr
from docx.oxml.simpletypes import ST_HexColor, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.shared import CT_OnOff
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.shared import RGBColor


class CT_Background(BaseOxmlElement):
    """``<w:background>`` element, the document-wide page background.

    Appears as the first child of `w:document`. Holds the document background
    color in its ``w:color`` attribute (hex RGB).
    """

    color: RGBColor | str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:color", ST_HexColor
    )


class CT_Document(BaseOxmlElement):
    """``<w:document>`` element, the root element of a document.xml file."""

    get_or_add_background: Callable[[], CT_Background]
    _remove_background: Callable[[], None]

    _tag_seq = ("w:background", "w:body")
    background: CT_Background | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:background", successors=("w:body",)
    )
    body: CT_Body = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:body", successors=()
    )
    del _tag_seq

    @property
    def sectPr_lst(self) -> list[CT_SectPr]:
        """All `w:sectPr` elements directly accessible from document element.

        Note this does not include a `sectPr` child in a paragraphs wrapped in
        revision marks or other intervening layer, perhaps `w:sdt` or customXml
        elements.

        `w:sectPr` elements appear in document order. The last one is always
        `w:body/w:sectPr`, all preceding are `w:p/w:pPr/w:sectPr`.
        """
        xpath = "./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr"
        return self.xpath(xpath)


class CT_AltChunkPr(BaseOxmlElement):
    """`w:altChunkPr` element, properties of a `w:altChunk` import reference.

    Optional child of `w:altChunk`. Currently carries a single boolean child
    `w:matchSrc` that requests Word attempt to match the character formatting
    of the source payload when importing (see ECMA-376 §17.17.2.3).
    """

    get_or_add_matchSrc: Callable[[], "CT_OnOff"]
    _remove_matchSrc: Callable[[], None]

    matchSrc: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:matchSrc", successors=()
    )


class CT_AltChunk(BaseOxmlElement):
    """`w:altChunk` element, an "alternate chunk" import reference.

    Points at an external-format payload part (HTML, RTF, XHTML, text, etc.)
    by OPC relationship id. Word substitutes the referenced part's contents
    for the ``w:altChunk`` element at render time. Relationships carry the
    ``aFChunk`` reltype and the target part's content-type declares the
    payload format (e.g. ``text/html``). See ECMA-376 §17.17.

    May optionally carry a child ``w:altChunkPr`` element with import options
    (currently only ``w:matchSrc``).
    """

    get_or_add_altChunkPr: Callable[[], CT_AltChunkPr]
    _remove_altChunkPr: Callable[[], None]

    altChunkPr: CT_AltChunkPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:altChunkPr", successors=()
    )
    rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:id", ST_String
    )

    @property
    def match_src(self) -> bool | None:
        """Resolved value of ``w:altChunkPr/w:matchSrc/@w:val``.

        |None| when no ``w:altChunkPr`` or ``w:matchSrc`` child is present.
        Otherwise the boolean value of the ``w:val`` attribute (default
        ``True`` when the attribute is absent on a bare ``<w:matchSrc/>``).
        """
        pr = self.altChunkPr
        if pr is None:
            return None
        match = pr.matchSrc
        if match is None:
            return None
        return bool(match.val)

    @match_src.setter
    def match_src(self, value: bool | None) -> None:
        if value is None:
            # -- drop both the child and the wrapper when the wrapper is empty --
            pr = self.altChunkPr
            if pr is None:
                return
            pr._remove_matchSrc()
            # -- if altChunkPr is now empty, drop it too --
            if len(pr) == 0:
                self._remove_altChunkPr()
            return
        pr = self.get_or_add_altChunkPr()
        match = pr.get_or_add_matchSrc()
        match.val = bool(value)


class CT_Body(BaseOxmlElement):
    """`w:body`, the container element for the main document story in `document.xml`."""

    add_p: Callable[[], CT_P]
    get_or_add_sectPr: Callable[[], CT_SectPr]
    p_lst: list[CT_P]
    tbl_lst: list[CT_Tbl]

    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    p = ZeroOrMore("w:p", successors=("w:sectPr",))
    tbl = ZeroOrMore("w:tbl", successors=("w:sectPr",))
    sdt = ZeroOrMore("w:sdt", successors=("w:sectPr",))
    altChunk = ZeroOrMore("w:altChunk", successors=("w:sectPr",))
    altChunk_lst: list[CT_AltChunk]
    sectPr: CT_SectPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sectPr", successors=()
    )

    def add_altChunk(self, rId: str) -> CT_AltChunk:
        """Append a new `w:altChunk` element with the given relationship id.

        .. versionadded:: 2026.05.0
        """
        altChunk = self._add_altChunk()  # type: ignore[attr-defined]
        altChunk.set(qn("r:id"), rId)
        return altChunk

    def add_section_break(self) -> CT_SectPr:
        """Return `w:sectPr` element for new section added at end of document.

        The last `w:sectPr` becomes the second-to-last, with the new `w:sectPr` being an
        exact clone of the previous one, except that all header and footer references
        are removed (and are therefore now "inherited" from the prior section).

        A copy of the previously-last `w:sectPr` will now appear in a new `w:p` at the
        end of the document. The returned `w:sectPr` is the sentinel `w:sectPr` for the
        document (and as implemented, `is` the prior sentinel `w:sectPr` with headers
        and footers removed).
        """
        # ---get the sectPr at file-end, which controls last section (sections[-1])---
        sentinel_sectPr = self.get_or_add_sectPr()
        # ---add exact copy to new `w:p` element; that is now second-to last section---
        self.add_p().set_sectPr(sentinel_sectPr.clone())
        # ---remove any header or footer references from "new" last section---
        for hdrftr_ref in sentinel_sectPr.xpath("w:headerReference|w:footerReference"):
            sentinel_sectPr.remove(hdrftr_ref)
        # ---the sentinel `w:sectPr` now controls the new last section---
        return sentinel_sectPr

    def clear_content(self):
        """Remove all content child elements from this <w:body> element.

        Leave the <w:sectPr> element if it is present.
        """
        for content_elm in self.xpath("./*[not(self::w:sectPr)]"):
            self.remove(content_elm)

    @property
    def inner_content_elements(self) -> list[CT_P | CT_Tbl]:
        """Generate all `w:p` and `w:tbl` elements in this document-body.

        Elements appear in document order. Elements shaded by nesting in a `w:ins` or
        other "wrapper" element will not be included.
        """
        return self.xpath("./w:p | ./w:tbl")
