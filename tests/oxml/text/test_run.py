"""Test suite for the docx.oxml.text.run module."""

from typing import cast

import pytest

from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R

from ...unitutil.cxml import element, xml


class DescribeCT_R:
    """Unit-test suite for the CT_R (run, <w:r>) element."""

    @pytest.mark.parametrize(
        ("initial_cxml", "text", "expected_cxml"),
        [
            ("w:r", "foobar", 'w:r/w:t"foobar"'),
            ("w:r", "foobar ", 'w:r/w:t{xml:space=preserve}"foobar "'),
            (
                "w:r/(w:rPr/w:rStyle{w:val=emphasis}, w:cr)",
                "foobar",
                'w:r/(w:rPr/w:rStyle{w:val=emphasis}, w:cr, w:t"foobar")',
            ),
        ],
    )
    def it_can_add_a_t_preserving_edge_whitespace(
        self, initial_cxml: str, text: str, expected_cxml: str
    ):
        r = cast(CT_R, element(initial_cxml))
        expected_xml = xml(expected_cxml)

        r.add_t(text)

        assert r.xml == expected_xml

    def it_can_assemble_the_text_in_the_run(self):
        cxml = 'w:r/(w:br,w:cr,w:noBreakHyphen,w:ptab,w:t"foobar",w:tab)'
        r = cast(CT_R, element(cxml))

        assert r.text == "\n\n-\tfoobar\t"

    @pytest.mark.parametrize(
        ("p_cxml", "offset", "expected_left_text", "expected_right_text"),
        [
            ('w:p/w:r/w:t"abcdef"', 3, "abc", "def"),
            ('w:p/w:r/w:t"abcdef"', 1, "a", "bcdef"),
            ('w:p/w:r/w:t"abcdef"', 5, "abcde", "f"),
        ],
    )
    def it_can_split_at_a_character_position(
        self,
        p_cxml: str,
        offset: int,
        expected_left_text: str,
        expected_right_text: str,
    ):
        p = cast(CT_P, element(p_cxml))
        r = p.r_lst[0]

        left_r, right_r = r.split(offset)

        assert left_r is r
        assert left_r.text == expected_left_text
        assert right_r.text == expected_right_text
        # -- right run is now next sibling in the paragraph --
        assert p.r_lst[0] is left_r
        assert p.r_lst[1] is right_r
        assert len(p.r_lst) == 2

    def it_preserves_formatting_on_split(self):
        p = cast(CT_P, element('w:p/w:r/(w:rPr/(w:b,w:i),w:t"abcdef")'))
        r = p.r_lst[0]

        left_r, right_r = r.split(3)

        assert left_r.xml == xml('w:r/(w:rPr/(w:b,w:i),w:t"abc")')
        assert right_r.xml == xml('w:r/(w:rPr/(w:b,w:i),w:t"def")')

    def it_preserves_run_style_on_split(self):
        p = cast(CT_P, element('w:p/w:r/(w:rPr/w:rStyle{w:val=Emphasis},w:t"abcdef")'))
        r = p.r_lst[0]

        left_r, right_r = r.split(2)

        assert left_r.xml == xml('w:r/(w:rPr/w:rStyle{w:val=Emphasis},w:t"ab")')
        assert right_r.xml == xml('w:r/(w:rPr/w:rStyle{w:val=Emphasis},w:t"cdef")')

    @pytest.mark.parametrize("offset", [0, 6, -1, 10])
    def it_raises_on_invalid_split_offset(self, offset: int):
        p = cast(CT_P, element('w:p/w:r/w:t"abcdef"'))
        r = p.r_lst[0]

        with pytest.raises(ValueError, match="offset .* not in range"):
            r.split(offset)

    def it_splits_a_run_without_formatting(self):
        p = cast(CT_P, element('w:p/w:r/w:t"abcdef"'))
        r = p.r_lst[0]

        left_r, right_r = r.split(3)

        assert left_r.xml == xml('w:r/w:t"abc"')
        assert right_r.xml == xml('w:r/w:t"def"')

    def it_inserts_the_new_run_in_the_right_position(self):
        p = cast(CT_P, element('w:p/(w:r/w:t"abc",w:r/w:t"def")'))
        r = p.r_lst[0]

        left_r, right_r = r.split(2)

        assert len(p.r_lst) == 3
        assert p.r_lst[0].text == "ab"
        assert p.r_lst[1].text == "c"
        assert p.r_lst[2].text == "def"
