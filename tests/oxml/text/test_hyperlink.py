"""Test suite for the docx.oxml.text.hyperlink module."""

from typing import cast

import pytest

from docx.oxml.text.hyperlink import CT_Hyperlink
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R

from ...unitutil.cxml import element


class DescribeCT_P_Hyperlink:
    """Unit-test suite for CT_P.add_hyperlink()."""

    def it_can_add_an_external_hyperlink(self):
        p = cast(CT_P, element("w:p"))

        hyperlink = p.add_hyperlink(rId="rId7")

        assert isinstance(hyperlink, CT_Hyperlink)
        assert hyperlink.rId == "rId7"
        assert hyperlink.anchor is None
        assert len(p.hyperlink_lst) == 1

    def it_can_add_an_internal_hyperlink(self):
        p = cast(CT_P, element("w:p"))

        hyperlink = p.add_hyperlink(anchor="_Toc123")

        assert isinstance(hyperlink, CT_Hyperlink)
        assert hyperlink.rId is None
        assert hyperlink.anchor == "_Toc123"

    def it_can_add_a_hyperlink_with_both_rId_and_anchor(self):
        p = cast(CT_P, element("w:p"))

        hyperlink = p.add_hyperlink(rId="rId7", anchor="section1")

        assert hyperlink.rId == "rId7"
        assert hyperlink.anchor == "section1"


class DescribeCT_Hyperlink:
    """Unit-test suite for the CT_Hyperlink (<w:hyperlink>) element."""

    def it_has_a_relationship_that_contains_the_hyperlink_address(self):
        cxml = 'w:hyperlink{r:id=rId6}/w:r/w:t"post"'
        hyperlink = cast(CT_Hyperlink, element(cxml))

        rId = hyperlink.rId

        assert rId == "rId6"

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            # -- default (when omitted) is True, somewhat surprisingly --
            ("w:hyperlink{r:id=rId6}", True),
            ("w:hyperlink{r:id=rId6,w:history=0}", False),
            ("w:hyperlink{r:id=rId6,w:history=1}", True),
        ],
    )
    def it_knows_whether_it_has_been_clicked_on_aka_visited(self, cxml: str, expected_value: bool):
        hyperlink = cast(CT_Hyperlink, element(cxml))
        assert hyperlink.history is expected_value

    def it_has_zero_or_more_runs_containing_the_hyperlink_text(self):
        cxml = 'w:hyperlink{r:id=rId6,w:history=1}/(w:r/w:t"blog",w:r/w:t" post")'
        hyperlink = cast(CT_Hyperlink, element(cxml))

        rs = hyperlink.r_lst

        assert [type(r) for r in rs] == [CT_R, CT_R]
        assert rs[0].text == "blog"
        assert rs[1].text == " post"

    def it_can_add_a_run_with_text(self):
        hyperlink = cast(CT_Hyperlink, element("w:hyperlink{r:id=rId6}"))

        r = hyperlink.add_r_with_text("click here")

        assert isinstance(r, CT_R)
        assert r.text == "click here"
        assert len(hyperlink.r_lst) == 1

    def it_can_add_a_run_with_text_and_style(self):
        hyperlink = cast(CT_Hyperlink, element("w:hyperlink{r:id=rId6}"))

        r = hyperlink.add_r_with_text("click here", style_id="Hyperlink")

        assert r.text == "click here"
        assert r.style == "Hyperlink"
