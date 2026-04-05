"""Unit test suite for the docx.oxml.endnotes module."""

from __future__ import annotations

from typing import cast

from docx.oxml.endnotes import CT_Endnote, CT_Endnotes
from docx.oxml.ns import qn

from ..unitutil.cxml import element


class DescribeCT_Endnotes:
    """Unit test suite for `docx.oxml.endnotes.CT_Endnotes` objects."""

    def it_provides_access_to_its_endnote_children(self):
        endnotes = cast(
            CT_Endnotes,
            element("w:endnotes/(w:endnote{w:id=0},w:endnote{w:id=1})"),
        )

        assert len(endnotes.endnote_lst) == 2

    def it_can_add_an_endnote(self):
        endnotes = cast(
            CT_Endnotes,
            element(
                "w:endnotes/(w:endnote{w:id=0,w:type=separator}"
                ",w:endnote{w:id=1,w:type=continuationSeparator})"
            ),
        )

        endnote = endnotes.add_endnote()

        assert endnote.id == 2
        # -- the endnote has a paragraph with EndnoteText style --
        assert len(endnote.p_lst) == 1
        p = endnote.p_lst[0]
        assert p.style == "EndnoteText"
        # -- the paragraph has a run with EndnoteReference style and endnoteRef --
        assert len(p.r_lst) == 1
        r = p.r_lst[0]
        assert r.style == "EndnoteReference"
        assert r[-1].tag == qn("w:endnoteRef")

    def it_assigns_sequential_ids_to_added_endnotes(self):
        endnotes = cast(
            CT_Endnotes,
            element(
                "w:endnotes/(w:endnote{w:id=0,w:type=separator}"
                ",w:endnote{w:id=1,w:type=continuationSeparator})"
            ),
        )

        en1 = endnotes.add_endnote()
        en2 = endnotes.add_endnote()

        assert en1.id == 2
        assert en2.id == 3

    def it_skips_used_ids_when_assigning(self):
        endnotes = cast(
            CT_Endnotes,
            element(
                "w:endnotes/(w:endnote{w:id=0,w:type=separator}"
                ",w:endnote{w:id=1,w:type=continuationSeparator}"
                ",w:endnote{w:id=2})"
            ),
        )

        endnote = endnotes.add_endnote()

        assert endnote.id == 3


class DescribeCT_Endnote:
    """Unit test suite for `docx.oxml.endnotes.CT_Endnote` objects."""

    def it_provides_access_to_its_id(self):
        endnote = cast(CT_Endnote, element("w:endnote{w:id=42}"))

        assert endnote.id == 42

    def it_provides_access_to_its_type(self):
        endnote = cast(CT_Endnote, element("w:endnote{w:id=0,w:type=separator}"))

        assert endnote.type == "separator"

    def it_returns_None_for_type_when_not_present(self):
        endnote = cast(CT_Endnote, element("w:endnote{w:id=2}"))

        assert endnote.type is None

    def it_can_clear_its_content(self):
        endnote = cast(
            CT_Endnote,
            element('w:endnote{w:id=2}/(w:p/w:r/w:t"Para one",w:p/w:r/w:t"Para two")'),
        )
        assert len(endnote.p_lst) == 2

        endnote.clear_content()

        assert len(endnote.p_lst) == 1
        p = endnote.p_lst[0]
        assert p.style == "EndnoteText"
        # -- the paragraph has an endnoteRef run to preserve the auto-number mark --
        assert len(p.r_lst) == 1
        r = p.r_lst[0]
        assert r.style == "EndnoteReference"
        assert r[-1].tag == qn("w:endnoteRef")

    def it_provides_access_to_its_inner_content_elements(self):
        endnote = cast(
            CT_Endnote,
            element("w:endnote{w:id=2}/(w:p,w:tbl,w:p)"),
        )

        content = endnote.inner_content_elements
        assert len(content) == 3
