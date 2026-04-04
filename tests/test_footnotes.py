# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.footnotes` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.footnotes import Footnote, Footnotes
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.footnotes import CT_Footnote, CT_Footnotes
from docx.oxml.ns import qn
from docx.oxml.text.run import CT_R
from docx.package import Package
from docx.parts.footnotes import FootnotesPart
from docx.text.run import Run

from .unitutil.cxml import element
from .unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeFootnotes:
    """Unit-test suite for `docx.footnotes.Footnotes` objects."""

    @pytest.mark.parametrize(
        ("cxml", "count"),
        [
            # -- empty footnotes (only separators) --
            (
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator})",
                0,
            ),
            # -- one user footnote --
            (
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator}"
                ",w:footnote{w:id=2})",
                1,
            ),
            # -- two user footnotes --
            (
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator}"
                ",w:footnote{w:id=2},w:footnote{w:id=3})",
                2,
            ),
        ],
    )
    def it_knows_how_many_footnotes_it_contains(self, cxml: str, count: int, package_: Mock):
        footnotes_elm = cast(CT_Footnotes, element(cxml))
        footnotes_part = FootnotesPart(
            PackURI("/word/footnotes.xml"), CT.WML_FOOTNOTES, footnotes_elm, package_
        )
        footnotes = Footnotes(footnotes_elm, footnotes_part)

        assert len(footnotes) == count

    def it_is_iterable_over_user_footnotes(self, package_: Mock):
        footnotes_elm = cast(
            CT_Footnotes,
            element(
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator}"
                ",w:footnote{w:id=2},w:footnote{w:id=3})"
            ),
        )
        footnotes_part = FootnotesPart(
            PackURI("/word/footnotes.xml"), CT.WML_FOOTNOTES, footnotes_elm, package_
        )
        footnotes = Footnotes(footnotes_elm, footnotes_part)

        footnote_iter = iter(footnotes)

        fn1 = next(footnote_iter)
        assert type(fn1) is Footnote
        assert fn1.footnote_id == 2
        fn2 = next(footnote_iter)
        assert type(fn2) is Footnote
        assert fn2.footnote_id == 3
        with pytest.raises(StopIteration):
            next(footnote_iter)

    def it_can_add_a_footnote(self, package_: Mock):
        footnotes_elm = cast(
            CT_Footnotes,
            element(
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator})"
            ),
        )
        footnotes_part = FootnotesPart(
            PackURI("/word/footnotes.xml"), CT.WML_FOOTNOTES, footnotes_elm, package_
        )
        footnotes = Footnotes(footnotes_elm, footnotes_part)

        # -- create a run to anchor the footnote reference --
        para_elm = element("w:p/w:r")
        r_elm = cast(CT_R, para_elm[0])
        run = Run(r_elm, footnotes_part)

        footnote = footnotes.add(run)

        # -- a Footnote is returned --
        assert isinstance(footnote, Footnote)
        assert footnote.footnote_id == 2
        # -- the footnote part is linked --
        assert footnote.part is footnotes_part
        # -- the footnote has a single paragraph with FootnoteText style --
        assert len(footnote.paragraphs) == 1
        assert footnote.paragraphs[0]._p.style == "FootnoteText"
        # -- a footnoteReference was inserted into the run --
        ref_elms = r_elm.xpath("./w:footnoteReference")
        assert len(ref_elms) == 1
        assert ref_elms[0].get(qn("w:id")) == "2"
        # -- the run has FootnoteReference character style --
        assert r_elm.style == "FootnoteReference"

    def it_can_add_a_footnote_with_text(self, package_: Mock):
        footnotes_elm = cast(
            CT_Footnotes,
            element(
                "w:footnotes/(w:footnote{w:id=0,w:type=separator}"
                ",w:footnote{w:id=1,w:type=continuationSeparator})"
            ),
        )
        footnotes_part = FootnotesPart(
            PackURI("/word/footnotes.xml"), CT.WML_FOOTNOTES, footnotes_elm, package_
        )
        footnotes = Footnotes(footnotes_elm, footnotes_part)

        para_elm = element("w:p/w:r")
        r_elm = cast(CT_R, para_elm[0])
        run = Run(r_elm, footnotes_part)

        footnote = footnotes.add(run, text="This is a footnote.")

        # -- the first paragraph has the text after the footnote ref run --
        first_para = footnote.paragraphs[0]
        assert len(first_para._p.r_lst) == 2
        assert first_para._p.r_lst[1].text == "This is a footnote."

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def package_(self, request: FixtureRequest):
        return instance_mock(request, Package)


class DescribeFootnote:
    """Unit-test suite for `docx.footnotes.Footnote`."""

    def it_knows_its_footnote_id(self, footnotes_part_: Mock):
        footnote_elm = cast(CT_Footnote, element("w:footnote{w:id=42}"))
        footnote = Footnote(footnote_elm, footnotes_part_)

        assert footnote.footnote_id == 42

    def it_provides_access_to_the_paragraphs_it_contains(self, footnotes_part_: Mock):
        footnote_elm = cast(
            CT_Footnote,
            element('w:footnote{w:id=2}/(w:p/w:r/w:t"First para",w:p/w:r/w:t"Second para")'),
        )
        footnote = Footnote(footnote_elm, footnotes_part_)

        paragraphs = footnote.paragraphs

        assert len(paragraphs) == 2
        assert [para.text for para in paragraphs] == ["First para", "Second para"]

    def it_can_add_a_paragraph(self, footnotes_part_: Mock):
        footnote_elm = cast(CT_Footnote, element("w:footnote{w:id=2}/w:p"))
        footnote = Footnote(footnote_elm, footnotes_part_)

        paragraph = footnote.add_paragraph("New paragraph text")

        assert len(footnote.paragraphs) == 2
        assert footnote.paragraphs[1].text == "New paragraph text"
        # -- default style is FootnoteText --
        assert paragraph._p.style == "FootnoteText"

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def footnotes_part_(self, request: FixtureRequest):
        return instance_mock(request, FootnotesPart)
