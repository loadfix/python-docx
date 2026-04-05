"""Unit test suite for the docx.oxml.footnotes module."""

from __future__ import annotations

from typing import cast

from docx.oxml.footnotes import CT_Footnote, CT_Footnotes
from docx.oxml.ns import qn

from ..unitutil.cxml import element


class DescribeCT_Footnotes:
    """Unit test suite for `docx.oxml.footnotes.CT_Footnotes` objects."""

    def it_provides_access_to_its_footnote_children(self):
        footnotes = cast(
            CT_Footnotes,
            element("w:footnotes/(w:footnote{w:id=0},w:footnote{w:id=1})"),
        )

        assert len(footnotes.footnote_lst) == 2

    def it_can_add_a_footnote(self):
        footnotes = cast(
            CT_Footnotes,
            element(
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator})"
            ),
        )

        footnote = footnotes.add_footnote()

        assert footnote.id == 2
        # -- the footnote has a paragraph with FootnoteText style --
        assert len(footnote.p_lst) == 1
        p = footnote.p_lst[0]
        assert p.style == "FootnoteText"
        # -- the paragraph has a run with FootnoteReference style and footnoteRef --
        assert len(p.r_lst) == 1
        r = p.r_lst[0]
        assert r.style == "FootnoteReference"
        assert r[-1].tag == qn("w:footnoteRef")

    def it_assigns_sequential_ids_to_added_footnotes(self):
        footnotes = cast(
            CT_Footnotes,
            element(
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator})"
            ),
        )

        fn1 = footnotes.add_footnote()
        fn2 = footnotes.add_footnote()

        assert fn1.id == 2
        assert fn2.id == 3

    def it_skips_used_ids_when_assigning(self):
        footnotes = cast(
            CT_Footnotes,
            element(
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator}"
                ",w:footnote{w:id=2})"
            ),
        )

        footnote = footnotes.add_footnote()

        assert footnote.id == 3


class DescribeCT_Footnote:
    """Unit test suite for `docx.oxml.footnotes.CT_Footnote` objects."""

    def it_provides_access_to_its_id(self):
        footnote = cast(CT_Footnote, element("w:footnote{w:id=42}"))

        assert footnote.id == 42

    def it_provides_access_to_its_type(self):
        footnote = cast(CT_Footnote, element("w:footnote{w:id=0,w:type=separator}"))

        assert footnote.type == "separator"

    def it_returns_None_for_type_when_not_present(self):
        footnote = cast(CT_Footnote, element("w:footnote{w:id=2}"))

        assert footnote.type is None

    def it_can_clear_its_content(self):
        footnote = cast(
            CT_Footnote,
            element('w:footnote{w:id=2}/(w:p/w:r/w:t"Para one",w:p/w:r/w:t"Para two")'),
        )
        assert len(footnote.p_lst) == 2

        footnote.clear_content()

        assert len(footnote.p_lst) == 1
        p = footnote.p_lst[0]
        assert p.style == "FootnoteText"
        # -- the paragraph has a footnoteRef run to preserve the auto-number mark --
        assert len(p.r_lst) == 1
        r = p.r_lst[0]
        assert r.style == "FootnoteReference"
        assert r[-1].tag == qn("w:footnoteRef")

    def it_provides_access_to_its_inner_content_elements(self):
        footnote = cast(
            CT_Footnote,
            element("w:footnote{w:id=2}/(w:p,w:tbl,w:p)"),
        )

        content = footnote.inner_content_elements
        assert len(content) == 3
