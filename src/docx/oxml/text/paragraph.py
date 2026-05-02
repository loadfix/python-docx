# pyright: reportPrivateUsage=false

"""Custom element classes related to paragraphs (CT_P)."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast
from collections.abc import Callable

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.fields import CT_FldSimple
    from docx.oxml.section import CT_SectPr
    from docx.oxml.text.hyperlink import CT_Hyperlink
    from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.run import CT_R
    from docx.oxml.tracked_changes import CT_Del, CT_Ins, CT_MoveFrom, CT_MoveTo
    from docx.oxml.text.font import CT_RPr


_TRANSPARENT_WRAPPERS = frozenset(
    {
        qn("w:smartTag"),
        qn("w:customXml"),
        # -- tracked-insertion / move-destination are present in the final doc --
        qn("w:ins"),
        qn("w:moveTo"),
    }
)
_RUN_LIKE_TAGS = frozenset(
    {qn("w:r"), qn("w:hyperlink"), qn("w:fldSimple"), qn("w:sdt")}
)

# -- pre-computed Clark names used by _iter_run_like for sdt/AlternateContent --
_W_SDT = qn("w:sdt")
_W_SDT_CONTENT = qn("w:sdtContent")
_MC_ALTERNATE_CONTENT = qn("mc:AlternateContent")
_MC_CHOICE = qn("mc:Choice")
_MC_FALLBACK = qn("mc:Fallback")


def _has_run_like_descendant(container: BaseOxmlElement) -> bool:
    """Return True when ``container`` yields any run-like element."""
    for _ in _iter_run_like(container):
        return True
    return False


def _iter_run_like(container: BaseOxmlElement):
    """Yield run-like children of ``container`` in document order.

    ``w:r``, ``w:hyperlink``, ``w:fldSimple``, and ``w:sdt`` are yielded
    directly. Several wrapper elements are transparent — the iterator descends
    into them recursively and yields their run-like descendants in place so
    callers see a flat run-list regardless of the wrapper scaffolding:

    - ``w:smartTag`` / ``w:customXml`` — generic run wrappers.
    - ``w:ins`` / ``w:moveTo`` — tracked-insertion and move-destination;
      their run content *is* present in the final document and must be
      visible to Find/Replace and text accessors.
    - ``w:sdt`` — yielded directly so that higher-level iterators that want
      content-control boundaries can see them; but note ``CT_P.text`` descends
      one level further so the inner text surfaces (see handling below).
    - ``mc:AlternateContent`` — the choose-one multi-markup element. Prefer
      ``mc:Choice``; fall back to ``mc:Fallback`` when Choice has no run-like
      descendants.

    This is the fix for upstream #932 / #225 (smartTag), #1327 / #1389 / #335
    / PR#1538 / PR#734 (sdt / AlternateContent / ins / moveTo).
    """
    for child in container:
        tag = child.tag
        if tag in _TRANSPARENT_WRAPPERS:
            yield from _iter_run_like(child)
        elif tag == _MC_ALTERNATE_CONTENT:
            # -- prefer the first Choice that has run-like content; otherwise
            # -- fall back to Fallback. This matches how Word resolves
            # -- mc:AlternateContent when opening documents that contain
            # -- alternative renderings for newer features. --
            chosen = None
            fallback = None
            for branch in child:
                btag = branch.tag
                if btag == _MC_CHOICE and chosen is None:
                    if _has_run_like_descendant(branch):
                        chosen = branch
                elif btag == _MC_FALLBACK and fallback is None:
                    fallback = branch
            target = chosen if chosen is not None else fallback
            if target is not None:
                yield from _iter_run_like(target)
        elif tag == _W_SDT:
            # -- yield the w:sdt itself (callers treat it as run-like);
            #    its inner text still surfaces via CT_Sdt.text --
            yield child
        elif tag in _RUN_LIKE_TAGS:
            yield child


def _iter_sdt_content(sdt: BaseOxmlElement):
    """Yield run-like descendants inside a ``w:sdt/w:sdtContent`` element.

    Skips ``w:sdtPr`` / ``w:sdtEndPr`` siblings which are property metadata
    rather than visible content.
    """
    for child in sdt:
        if child.tag == _W_SDT_CONTENT:
            yield from _iter_run_like(child)


def _run_is_field_code_only(r: BaseOxmlElement) -> bool:
    """Return True when a ``w:r`` contains only field-code (`w:instrText`).

    Such runs carry the field instruction (the code), not the rendered text
    the user sees, and should be excluded from text-visible iteration used by
    Find/Replace.
    """
    # -- a run is "code-only" when all its children are instrText or rPr,
    #    *and* at least one instrText is present --
    has_instr = False
    for child in r:
        tag = child.tag
        if tag == qn("w:instrText"):
            has_instr = True
            continue
        if tag == qn("w:rPr"):
            continue
        # -- any other child means there is visible content too --
        return False
    return has_instr


def _iter_r_descendants(container: BaseOxmlElement):
    """Yield visible ``w:r`` elements under ``container``.

    Descends through ``w:hyperlink``, ``w:fldSimple``, transparent wrappers,
    and nested ``w:sdt/w:sdtContent``. Skips ``w:r`` elements that carry only
    ``w:instrText`` (the field code).
    """
    for child in _iter_run_like(container):
        tag = child.tag
        if tag == qn("w:r"):
            if not _run_is_field_code_only(child):
                yield child
        elif tag in (qn("w:hyperlink"), qn("w:fldSimple")):
            yield from _iter_r_descendants(child)
        elif tag == _W_SDT:
            yield from _iter_all_r_elements_in(child)


def _iter_all_r_elements_in(sdt: BaseOxmlElement):
    """Yield visible ``w:r`` elements nested inside an ``w:sdt`` element."""
    for inner in _iter_sdt_content(sdt):
        itag = inner.tag
        if itag == qn("w:r"):
            if not _run_is_field_code_only(inner):
                yield inner
        elif itag in (qn("w:hyperlink"), qn("w:fldSimple")):
            yield from _iter_r_descendants(inner)
        elif itag == _W_SDT:
            yield from _iter_all_r_elements_in(inner)


class CT_P(BaseOxmlElement):
    """`<w:p>` element, containing the properties and text for a paragraph."""

    add_r: Callable[[], CT_R]
    get_or_add_pPr: Callable[[], CT_PPr]
    hyperlink_lst: list[CT_Hyperlink]
    r_lst: list[CT_R]
    fldSimple_lst: list[CT_FldSimple]

    pPr: CT_PPr | None = ZeroOrOne("w:pPr")  # pyright: ignore[reportAssignmentType]
    hyperlink = ZeroOrMore("w:hyperlink")
    r = ZeroOrMore("w:r")
    fldSimple = ZeroOrMore("w:fldSimple")
    sdt = ZeroOrMore("w:sdt")

    rsidR: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:rsidR", ST_String
    )

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

    def add_permission_range(
        self,
        perm_id: int,
        edit_group: str | None = None,
        user: str | None = None,
    ) -> None:
        """Add permStart/permEnd pair wrapping this paragraph's run content.

        At least one of `edit_group` or `user` should be provided for the range
        to be meaningful to consumers, but the XSD permits both to be absent.
        """
        start_attrs: dict[str, str] = {qn("w:id"): str(perm_id)}
        if edit_group is not None:
            start_attrs[qn("w:edGrp")] = edit_group
        if user is not None:
            start_attrs[qn("w:ed")] = user

        permStart = OxmlElement("w:permStart", attrs=start_attrs)
        permEnd = OxmlElement("w:permEnd", attrs={qn("w:id"): str(perm_id)})

        # -- insert permStart after pPr (or at beginning) and permEnd at end --
        if self.pPr is not None:
            self.pPr.addnext(permStart)
        else:
            self.insert(0, permStart)
        self.append(permEnd)

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
    def inner_content_elements(self) -> list[CT_R | CT_Hyperlink]:
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
    def lastRenderedPageBreaks(self) -> list[CT_LastRenderedPageBreak]:
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

        Inner-content child elements like `w:r`, `w:hyperlink`, `w:fldSimple`, and
        `w:sdt` (structured document tag / content control) are translated to their
        text equivalent. Runs wrapped in ``w:smartTag`` or ``w:customXml``
        elements are descended into transparently so their text is included
        in document order (upstream #932, #225).
        """
        return "".join(e.text for e in _iter_run_like(self))

    def iter_r_elements(self):
        """Yield ``w:r`` descendants in document order, transparent to wrappers.

        ``w:smartTag`` and ``w:customXml`` wrappers around runs are descended
        into; the runs inside hyperlinks, fldSimple, or sdt are *not* yielded
        (those are yielded by higher-level iterators as ``CT_Hyperlink``,
        ``CT_FldSimple``, and ``CT_Sdt`` respectively). See upstream #932.
        """
        for el in _iter_run_like(self):
            if el.tag == qn("w:r"):
                yield el

    def iter_all_r_elements(self):
        """Yield every visible ``w:r`` descendant, including those nested inside
        ``w:hyperlink``, ``w:fldSimple``, and ``w:sdt/w:sdtContent`` wrappers.

        This is the content-visible run iterator used by Find/Replace and
        formatting loops (upstream #1370, #1021). ``w:instrText`` content —
        the field *code*, not the rendered result — is intentionally *not*
        yielded: runs whose sole content is ``w:instrText`` are skipped.

        Runs appearing between a complex-field ``separate`` marker and its
        ``end`` marker *are* yielded because they contain the rendered result
        that the user sees and would expect to search.
        """
        for el in _iter_run_like(self):
            tag = el.tag
            if tag == qn("w:r"):
                if _run_is_field_code_only(el):
                    continue
                yield el
            elif tag == qn("w:hyperlink"):
                yield from _iter_r_descendants(el)
            elif tag == qn("w:fldSimple"):
                yield from _iter_r_descendants(el)
            elif tag == _W_SDT:
                for inner in _iter_sdt_content(el):
                    itag = inner.tag
                    if itag == qn("w:r"):
                        if _run_is_field_code_only(inner):
                            continue
                        yield inner
                    elif itag in (qn("w:hyperlink"), qn("w:fldSimple")):
                        yield from _iter_r_descendants(inner)
                    elif itag == _W_SDT:
                        # -- nested sdt: recurse via temporary iterator --
                        yield from _iter_all_r_elements_in(inner)

    @property
    def tracked_change_elements(
        self,
    ) -> list[CT_Ins | CT_Del | CT_MoveFrom | CT_MoveTo]:
        """Run-level track-change children of this paragraph, in document order.

        Includes `w:ins`, `w:del`, `w:moveFrom`, and `w:moveTo`. The paragraph-
        level range-start / range-end elements (`w:moveFromRangeStart`,
        `w:moveFromRangeEnd`, `w:moveToRangeStart`, `w:moveToRangeEnd`) used to
        bracket cross-paragraph moves are intentionally excluded — they mark
        ranges rather than wrap run content.
        """
        return self.xpath("./w:ins | ./w:del | ./w:moveFrom | ./w:moveTo")

    def _insert_pPr(self, pPr: CT_PPr) -> CT_PPr:
        self.insert(0, pPr)
        return pPr
