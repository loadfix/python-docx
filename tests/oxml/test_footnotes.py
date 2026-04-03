"""Unit test suite for the docx.oxml.footnotes module."""

from __future__ import annotations

from typing import cast

from docx.oxml.footnotes import CT_Footnote, CT_Footnotes

from ..unitutil.cxml import element


class DescribeCT_Footnotes:
    """Unit test suite for `docx.oxml.footnotes.CT_Footnotes` objects."""

    def it_provides_access_to_its_footnote_children(self):
        footnotes = cast(
            CT_Footnotes,
            element("w:footnotes/(w:footnote{w:id=0},w:footnote{w:id=1})"),
        )

        assert len(footnotes.footnote_lst) == 2

    def it_can_determine_the_next_available_footnote_id(self):
        footnotes = cast(
            CT_Footnotes,
            element("w:footnotes/(w:footnote{w:id=0},w:footnote{w:id=1})"),
        )

        assert footnotes._next_available_footnote_id() == 2

    def it_returns_2_as_minimum_next_id(self):
        footnotes = cast(CT_Footnotes, element("w:footnotes"))

        assert footnotes._next_available_footnote_id() == 2

    def it_skips_used_ids(self):
        footnotes = cast(
            CT_Footnotes,
            element(
                "w:footnotes/(w:footnote{w:id=0},w:footnote{w:id=1},"
                "w:footnote{w:id=2},w:footnote{w:id=3})"
            ),
        )

        assert footnotes._next_available_footnote_id() == 4


class DescribeCT_Footnote:
    """Unit test suite for `docx.oxml.footnotes.CT_Footnote` objects."""

    def it_provides_access_to_its_id(self):
        footnote = cast(CT_Footnote, element("w:footnote{w:id=42}"))

        assert footnote.id == 42

    def it_provides_access_to_its_inner_content_elements(self):
        footnote = cast(
            CT_Footnote,
            element("w:footnote{w:id=2}/(w:p,w:tbl,w:p)"),
        )

        content = footnote.inner_content_elements
        assert len(content) == 3
