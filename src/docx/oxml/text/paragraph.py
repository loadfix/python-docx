# pyright: reportPrivateUsage=false

"""Custom element classes related to paragraphs (CT_P)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.fields import CT_FldSimple
    from docx.oxml.section import CT_SectPr
    from docx.oxml.text.hyperlink import CT_Hyperlink
    from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.run import CT_R
    from docx.oxml.tracked_changes import CT_Del, CT_Ins
    from docx.oxml.text.font import CT_RPr


class CT_P(BaseOxmlElement):
    """`<w:p>` element, containing the properties and text for a paragraph."""

    add_r: Callable[[], CT_R]
    get_or_add_pPr: Callable[[], CT_PPr]
    hyperlink_lst: List[CT_Hyperlink]
    r_lst: List[CT_R]
    fldSimple_lst: List[CT_FldSimple]

    pPr: CT_PPr | None = ZeroOrOne("w:pPr")  # pyright: ignore[reportAssignmentType]
    hyperlink = ZeroOrMore("w:hyperlink")
    r = ZeroOrMore("w:r")
    fldSimple = ZeroOrMore("w:fldSimple")

    def add_hyperlink(
        self, rId: str | None, anchor: str | None, text: str, rPr: CT_RPr | None
    ) -> CT_Hyperlink:
        """Return a newly appended `CT_Hyperlink` child element.

        `rId` is the relationship id for an external URL (or None for internal links).
        `anchor` is a bookmark name for internal links (or None for external links).
        `text` is the visible text of the hyperlink.
        `rPr` is an optional run-properties element to apply to the hyperlink run.
        """
        from docx.oxml.text.hyperlink import CT_Hyperlink

        hyperlink = cast(CT_Hyperlink, OxmlElement("w:hyperlink"))
        if rId is not None:
            hyperlink.rId = rId
        if anchor is not None:
            hyperlink.anchor = anchor
        hyperlink.history = True
        r = hyperlink.add_r()
        if rPr is not None:
            r.insert(0, rPr)
        r.add_t(text)
        self.append(hyperlink)
        return hyperlink

    def add_bookmark(self, bookmark_id: int, name: str) -> None:
        """Add bookmarkStart/bookmarkEnd pair to this paragraph.

        When no specific run positions are given, the bookmark wraps the entire
        paragraph content (all runs).
        """
        bookmarkStart = OxmlElement(
            "w:bookmarkStart",
            attrs={qn("w:id"): str(bookmark_id), qn("w:name"): name},
        )
        bookmarkEnd = OxmlElement(
            "w:bookmarkEnd",
            attrs={qn("w:id"): str(bookmark_id)},
        )
        # -- insert bookmarkStart after pPr (or at beginning) and bookmarkEnd at end --
        if self.pPr is not None:
            self.pPr.addnext(bookmarkStart)
        else:
            self.insert(0, bookmarkStart)
        self.append(bookmarkEnd)

    def add_p_before(self) -> CT_P:
        """Return a new `<w:p>` element inserted directly prior to this one."""
        new_p = cast(CT_P, OxmlElement("w:p"))
        self.addprevious(new_p)
        return new_p

    def add_fldSimple(self, instr: str, text: str | None = None) -> "CT_FldSimple":
        """Append a `<w:fldSimple>` child with `instr` and a result-text run.

        `text` is the current rendered result; when provided it is added as a
        single `<w:r><w:t>` child of the new `<w:fldSimple>` element. The new
        `<w:fldSimple>` element is returned.
        """
        from docx.oxml.fields import CT_FldSimple as _CT_FldSimple

        fldSimple = cast(_CT_FldSimple, OxmlElement("w:fldSimple"))
        fldSimple.instr = instr
        if text is not None and text != "":
            r = fldSimple.add_r()
            r.add_t(text)
        self.append(fldSimple)
        return fldSimple

    def add_complex_field(self, instr: str, result_text: str | None = None) -> CT_R:
        """Append the run sequence for a complex field and return the ``begin`` run.

        Appends the five-run sequence: ``begin`` marker, ``instrText``,
        ``separate`` marker, optional result-text run, and ``end`` marker. The
        first run (the one containing the ``begin`` fldChar) is returned for
        reference.
        """
        from docx.oxml.ns import qn

        # -- begin marker --
        r_begin = self.add_r()
        fldChar_begin = OxmlElement("w:fldChar")
        fldChar_begin.set(qn("w:fldCharType"), "begin")
        r_begin.append(fldChar_begin)

        # -- instrText --
        r_instr = self.add_r()
        instrText = OxmlElement("w:instrText")
        instrText.text = instr
        # -- preserve whitespace whenever the instruction contains any, to match
        #    Word's own behavior and avoid trim-happy consumers --
        if any(ch.isspace() for ch in instr):
            instrText.set(qn("xml:space"), "preserve")
        r_instr.append(instrText)

        # -- separate marker --
        r_sep = self.add_r()
        fldChar_sep = OxmlElement("w:fldChar")
        fldChar_sep.set(qn("w:fldCharType"), "separate")
        r_sep.append(fldChar_sep)

        # -- optional result-text run --
        if result_text is not None and result_text != "":
            r_result = self.add_r()
            r_result.add_t(result_text)

        # -- end marker --
        r_end = self.add_r()
        fldChar_end = OxmlElement("w:fldChar")
        fldChar_end.set(qn("w:fldCharType"), "end")
        r_end.append(fldChar_end)

        return r_begin

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """The value of the `<w:jc>` grandchild element or |None| if not present."""
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.jc_val

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        pPr = self.get_or_add_pPr()
        pPr.jc_val = value

    def clear_content(self):
        """Remove all child elements, except the `<w:pPr>` element if present."""
        for child in self.xpath("./*[not(self::w:pPr)]"):
            self.remove(child)

    @property
    def inner_content_elements(self) -> List[CT_R | CT_Hyperlink]:
        """Run and hyperlink children of the `w:p` element, in document order."""
        return self.xpath("./w:r | ./w:hyperlink")

    def iter_field_elements(self):
        """Yield ``(kind, element)`` pairs for each field in this paragraph.

        ``kind`` is either ``"simple"`` (and `element` is the `<w:fldSimple>`
        element) or ``"complex"`` (and `element` is the ``begin`` run that
        opens the complex field).

        Fields are yielded in document order. Nested complex fields are not
        separately reported; only the outer ``begin`` marker surfaces. This is
        intentional: nested fields are uncommon, and handling them here would
        complicate the iteration. Callers that need to traverse nested
        structure can walk the run sequence themselves starting from the
        returned ``begin`` run.
        """
        for el in self.xpath(
            "./w:fldSimple | ./w:r[w:fldChar[@w:fldCharType='begin']]"
        ):
            tag = el.tag.rsplit("}", 1)[-1]
            if tag == "fldSimple":
                yield "simple", el
            else:
                yield "complex", el

    @property
    def lastRenderedPageBreaks(self) -> List[CT_LastRenderedPageBreak]:
        """All `w:lastRenderedPageBreak` descendants of this paragraph.

        Rendered page-breaks commonly occur in a run but can also occur in a run inside
        a hyperlink. This returns both.
        """
        return self.xpath(
            "./w:r/w:lastRenderedPageBreak | ./w:hyperlink/w:r/w:lastRenderedPageBreak"
        )

    def set_sectPr(self, sectPr: CT_SectPr):
        """Unconditionally replace or add `sectPr` as grandchild in correct sequence."""
        pPr = self.get_or_add_pPr()
        pPr._remove_sectPr()
        pPr._insert_sectPr(sectPr)

    @property
    def style(self) -> str | None:
        """String contained in `w:val` attribute of `./w:pPr/w:pStyle` grandchild.

        |None| if not present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.style

    @style.setter
    def style(self, style: str | None):
        pPr = self.get_or_add_pPr()
        pPr.style = style

    @property
    def text(self):  # pyright: ignore[reportIncompatibleMethodOverride]
        """The textual content of this paragraph.

        Inner-content child elements like `w:r`, `w:hyperlink`, and `w:fldSimple` are
        translated to their text equivalent.
        """
        return "".join(e.text for e in self.xpath("w:r | w:hyperlink | w:fldSimple"))

    @property
    def tracked_change_elements(self) -> List[CT_Ins | CT_Del]:
        """`w:ins` and `w:del` children of this paragraph, in document order."""
        return self.xpath("./w:ins | ./w:del")

    def _insert_pPr(self, pPr: CT_PPr) -> CT_PPr:
        self.insert(0, pPr)
        return pPr
