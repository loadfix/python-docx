"""Test suite for docx.text.parfmt module, containing the ParagraphFormat object."""

from __future__ import annotations

import pytest

from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_FRAME_DROP_CAP,
    WD_FRAME_H_ALIGN,
    WD_FRAME_H_ANCHOR,
    WD_FRAME_V_ALIGN,
    WD_FRAME_V_ANCHOR,
    WD_FRAME_WRAP,
    WD_LINE_SPACING,
)
from docx.shared import Pt, Twips
from docx.text.parfmt import DropCap, ParagraphFormat, TextFrame
from docx.text.tabstops import TabStops

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class DescribeParagraphFormat:
    def it_knows_its_alignment_value(self, alignment_get_fixture):
        paragraph_format, expected_value = alignment_get_fixture
        assert paragraph_format.alignment == expected_value

    def it_can_change_its_alignment_value(self, alignment_set_fixture):
        paragraph_format, value, expected_xml = alignment_set_fixture
        paragraph_format.alignment = value
        assert paragraph_format._element.xml == expected_xml

    def it_knows_its_space_before(self, space_before_get_fixture):
        paragraph_format, expected_value = space_before_get_fixture
        assert paragraph_format.space_before == expected_value

    def it_can_change_its_space_before(self, space_before_set_fixture):
        paragraph_format, value, expected_xml = space_before_set_fixture
        paragraph_format.space_before = value
        assert paragraph_format._element.xml == expected_xml

    def it_knows_its_space_after(self, space_after_get_fixture):
        paragraph_format, expected_value = space_after_get_fixture
        assert paragraph_format.space_after == expected_value

    def it_can_change_its_space_after(self, space_after_set_fixture):
        paragraph_format, value, expected_xml = space_after_set_fixture
        paragraph_format.space_after = value
        assert paragraph_format._element.xml == expected_xml

    def it_knows_its_line_spacing(self, line_spacing_get_fixture):
        paragraph_format, expected_value = line_spacing_get_fixture
        assert paragraph_format.line_spacing == expected_value

    def it_can_change_its_line_spacing(self, line_spacing_set_fixture):
        paragraph_format, value, expected_xml = line_spacing_set_fixture
        paragraph_format.line_spacing = value
        assert paragraph_format._element.xml == expected_xml

    def it_knows_its_line_spacing_rule(self, line_spacing_rule_get_fixture):
        paragraph_format, expected_value = line_spacing_rule_get_fixture
        assert paragraph_format.line_spacing_rule == expected_value

    def it_can_change_its_line_spacing_rule(self, line_spacing_rule_set_fixture):
        paragraph_format, value, expected_xml = line_spacing_rule_set_fixture
        paragraph_format.line_spacing_rule = value
        assert paragraph_format._element.xml == expected_xml

    def it_knows_its_first_line_indent(self, first_indent_get_fixture):
        paragraph_format, expected_value = first_indent_get_fixture
        assert paragraph_format.first_line_indent == expected_value

    def it_can_change_its_first_line_indent(self, first_indent_set_fixture):
        paragraph_format, value, expected_xml = first_indent_set_fixture
        paragraph_format.first_line_indent = value
        assert paragraph_format._element.xml == expected_xml

    def it_knows_its_left_indent(self, left_indent_get_fixture):
        paragraph_format, expected_value = left_indent_get_fixture
        assert paragraph_format.left_indent == expected_value

    def it_can_change_its_left_indent(self, left_indent_set_fixture):
        paragraph_format, value, expected_xml = left_indent_set_fixture
        paragraph_format.left_indent = value
        assert paragraph_format._element.xml == expected_xml

    def it_knows_its_right_indent(self, right_indent_get_fixture):
        paragraph_format, expected_value = right_indent_get_fixture
        assert paragraph_format.right_indent == expected_value

    def it_can_change_its_right_indent(self, right_indent_set_fixture):
        paragraph_format, value, expected_xml = right_indent_set_fixture
        paragraph_format.right_indent = value
        assert paragraph_format._element.xml == expected_xml

    def it_knows_its_on_off_prop_values(self, on_off_get_fixture):
        paragraph_format, prop_name, expected_value = on_off_get_fixture
        assert getattr(paragraph_format, prop_name) == expected_value

    def it_can_change_its_on_off_props(self, on_off_set_fixture):
        paragraph_format, prop_name, value, expected_xml = on_off_set_fixture
        setattr(paragraph_format, prop_name, value)
        assert paragraph_format._element.xml == expected_xml

    def it_provides_access_to_its_tab_stops(self, tab_stops_fixture):
        paragraph_format, TabStops_, pPr, tab_stops_ = tab_stops_fixture
        tab_stops = paragraph_format.tab_stops
        TabStops_.assert_called_once_with(pPr)
        assert tab_stops is tab_stops_

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:jc{w:val=center}", WD_ALIGN_PARAGRAPH.CENTER),
        ]
    )
    def alignment_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, expected_value

    @pytest.fixture(
        params=[
            ("w:p", WD_ALIGN_PARAGRAPH.LEFT, "w:p/w:pPr/w:jc{w:val=left}"),
            ("w:p/w:pPr", WD_ALIGN_PARAGRAPH.CENTER, "w:p/w:pPr/w:jc{w:val=center}"),
            (
                "w:p/w:pPr/w:jc{w:val=center}",
                WD_ALIGN_PARAGRAPH.RIGHT,
                "w:p/w:pPr/w:jc{w:val=right}",
            ),
            ("w:p/w:pPr/w:jc{w:val=right}", None, "w:p/w:pPr"),
            ("w:p", None, "w:p/w:pPr"),
        ]
    )
    def alignment_set_fixture(self, request):
        p_cxml, value, expected_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_cxml)
        return paragraph_format, value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:ind", None),
            ("w:p/w:pPr/w:ind{w:firstLine=240}", Pt(12)),
            ("w:p/w:pPr/w:ind{w:hanging=240}", Pt(-12)),
        ]
    )
    def first_indent_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, expected_value

    @pytest.fixture(
        params=[
            ("w:p", Pt(36), "w:p/w:pPr/w:ind{w:firstLine=720}"),
            ("w:p", Pt(-36), "w:p/w:pPr/w:ind{w:hanging=720}"),
            ("w:p", 0, "w:p/w:pPr/w:ind{w:firstLine=0}"),
            ("w:p", None, "w:p/w:pPr"),
            ("w:p/w:pPr/w:ind{w:firstLine=240}", None, "w:p/w:pPr/w:ind"),
            (
                "w:p/w:pPr/w:ind{w:firstLine=240}",
                Pt(-18),
                "w:p/w:pPr/w:ind{w:hanging=360}",
            ),
            (
                "w:p/w:pPr/w:ind{w:hanging=240}",
                Pt(18),
                "w:p/w:pPr/w:ind{w:firstLine=360}",
            ),
        ]
    )
    def first_indent_set_fixture(self, request):
        p_cxml, value, expected_p_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_p_cxml)
        return paragraph_format, value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:ind", None),
            ("w:p/w:pPr/w:ind{w:left=120}", Pt(6)),
            ("w:p/w:pPr/w:ind{w:left=-06.3pt}", Pt(-6.3)),
        ]
    )
    def left_indent_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, expected_value

    @pytest.fixture(
        params=[
            ("w:p", Pt(36), "w:p/w:pPr/w:ind{w:left=720}"),
            ("w:p", Pt(-3), "w:p/w:pPr/w:ind{w:left=-60}"),
            ("w:p", 0, "w:p/w:pPr/w:ind{w:left=0}"),
            ("w:p", None, "w:p/w:pPr"),
            ("w:p/w:pPr/w:ind{w:left=240}", None, "w:p/w:pPr/w:ind"),
        ]
    )
    def left_indent_set_fixture(self, request):
        p_cxml, value, expected_p_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_p_cxml)
        return paragraph_format, value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:spacing", None),
            ("w:p/w:pPr/w:spacing{w:line=420}", 1.75),
            ("w:p/w:pPr/w:spacing{w:line=840,w:lineRule=exact}", Pt(42)),
            ("w:p/w:pPr/w:spacing{w:line=840,w:lineRule=atLeast}", Pt(42)),
        ]
    )
    def line_spacing_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, expected_value

    @pytest.fixture(
        params=[
            ("w:p", 1, "w:p/w:pPr/w:spacing{w:line=240,w:lineRule=auto}"),
            ("w:p", 2.0, "w:p/w:pPr/w:spacing{w:line=480,w:lineRule=auto}"),
            ("w:p", Pt(42), "w:p/w:pPr/w:spacing{w:line=840,w:lineRule=exact}"),
            ("w:p/w:pPr", 2, "w:p/w:pPr/w:spacing{w:line=480,w:lineRule=auto}"),
            (
                "w:p/w:pPr/w:spacing{w:line=360}",
                1,
                "w:p/w:pPr/w:spacing{w:line=240,w:lineRule=auto}",
            ),
            (
                "w:p/w:pPr/w:spacing{w:line=240,w:lineRule=exact}",
                1.75,
                "w:p/w:pPr/w:spacing{w:line=420,w:lineRule=auto}",
            ),
            (
                "w:p/w:pPr/w:spacing{w:line=240,w:lineRule=atLeast}",
                Pt(42),
                "w:p/w:pPr/w:spacing{w:line=840,w:lineRule=atLeast}",
            ),
            (
                "w:p/w:pPr/w:spacing{w:line=240,w:lineRule=exact}",
                None,
                "w:p/w:pPr/w:spacing",
            ),
            ("w:p/w:pPr", None, "w:p/w:pPr"),
        ]
    )
    def line_spacing_set_fixture(self, request):
        p_cxml, value, expected_p_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_p_cxml)
        return paragraph_format, value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:spacing", None),
            ("w:p/w:pPr/w:spacing{w:line=240}", WD_LINE_SPACING.SINGLE),
            ("w:p/w:pPr/w:spacing{w:line=360}", WD_LINE_SPACING.ONE_POINT_FIVE),
            ("w:p/w:pPr/w:spacing{w:line=480}", WD_LINE_SPACING.DOUBLE),
            ("w:p/w:pPr/w:spacing{w:line=420}", WD_LINE_SPACING.MULTIPLE),
            ("w:p/w:pPr/w:spacing{w:lineRule=auto}", WD_LINE_SPACING.MULTIPLE),
            ("w:p/w:pPr/w:spacing{w:lineRule=exact}", WD_LINE_SPACING.EXACTLY),
            ("w:p/w:pPr/w:spacing{w:lineRule=atLeast}", WD_LINE_SPACING.AT_LEAST),
        ]
    )
    def line_spacing_rule_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, expected_value

    @pytest.fixture(
        params=[
            (
                "w:p",
                WD_LINE_SPACING.SINGLE,
                "w:p/w:pPr/w:spacing{w:line=240,w:lineRule=auto}",
            ),
            (
                "w:p",
                WD_LINE_SPACING.ONE_POINT_FIVE,
                "w:p/w:pPr/w:spacing{w:line=360,w:lineRule=auto}",
            ),
            (
                "w:p",
                WD_LINE_SPACING.DOUBLE,
                "w:p/w:pPr/w:spacing{w:line=480,w:lineRule=auto}",
            ),
            ("w:p", WD_LINE_SPACING.MULTIPLE, "w:p/w:pPr/w:spacing{w:lineRule=auto}"),
            ("w:p", WD_LINE_SPACING.EXACTLY, "w:p/w:pPr/w:spacing{w:lineRule=exact}"),
            (
                "w:p/w:pPr/w:spacing{w:line=280,w:lineRule=exact}",
                WD_LINE_SPACING.AT_LEAST,
                "w:p/w:pPr/w:spacing{w:line=280,w:lineRule=atLeast}",
            ),
        ]
    )
    def line_spacing_rule_set_fixture(self, request):
        p_cxml, value, expected_p_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_p_cxml)
        return paragraph_format, value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", "keep_together", None),
            ("w:p/w:pPr/w:keepLines{w:val=on}", "keep_together", True),
            ("w:p/w:pPr/w:keepLines{w:val=0}", "keep_together", False),
            ("w:p", "keep_with_next", None),
            ("w:p/w:pPr/w:keepNext{w:val=1}", "keep_with_next", True),
            ("w:p/w:pPr/w:keepNext{w:val=false}", "keep_with_next", False),
            ("w:p", "page_break_before", None),
            ("w:p/w:pPr/w:pageBreakBefore", "page_break_before", True),
            ("w:p/w:pPr/w:pageBreakBefore{w:val=0}", "page_break_before", False),
            ("w:p", "widow_control", None),
            ("w:p/w:pPr/w:widowControl{w:val=true}", "widow_control", True),
            ("w:p/w:pPr/w:widowControl{w:val=off}", "widow_control", False),
        ]
    )
    def on_off_get_fixture(self, request):
        p_cxml, prop_name, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, prop_name, expected_value

    @pytest.fixture(
        params=[
            ("w:p", "keep_together", True, "w:p/w:pPr/w:keepLines"),
            ("w:p", "keep_with_next", True, "w:p/w:pPr/w:keepNext"),
            ("w:p", "page_break_before", True, "w:p/w:pPr/w:pageBreakBefore"),
            ("w:p", "widow_control", True, "w:p/w:pPr/w:widowControl"),
            (
                "w:p/w:pPr/w:keepLines",
                "keep_together",
                False,
                "w:p/w:pPr/w:keepLines{w:val=0}",
            ),
            (
                "w:p/w:pPr/w:keepNext",
                "keep_with_next",
                False,
                "w:p/w:pPr/w:keepNext{w:val=0}",
            ),
            (
                "w:p/w:pPr/w:pageBreakBefore",
                "page_break_before",
                False,
                "w:p/w:pPr/w:pageBreakBefore{w:val=0}",
            ),
            (
                "w:p/w:pPr/w:widowControl",
                "widow_control",
                False,
                "w:p/w:pPr/w:widowControl{w:val=0}",
            ),
            ("w:p/w:pPr/w:keepLines{w:val=0}", "keep_together", None, "w:p/w:pPr"),
            ("w:p/w:pPr/w:keepNext{w:val=0}", "keep_with_next", None, "w:p/w:pPr"),
            (
                "w:p/w:pPr/w:pageBreakBefore{w:val=0}",
                "page_break_before",
                None,
                "w:p/w:pPr",
            ),
            ("w:p/w:pPr/w:widowControl{w:val=0}", "widow_control", None, "w:p/w:pPr"),
        ]
    )
    def on_off_set_fixture(self, request):
        p_cxml, prop_name, value, expected_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_cxml)
        return paragraph_format, prop_name, value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:ind", None),
            ("w:p/w:pPr/w:ind{w:right=160}", Pt(8)),
            ("w:p/w:pPr/w:ind{w:right=-4.2pt}", Pt(-4.2)),
        ]
    )
    def right_indent_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, expected_value

    @pytest.fixture(
        params=[
            ("w:p", Pt(36), "w:p/w:pPr/w:ind{w:right=720}"),
            ("w:p", Pt(-3), "w:p/w:pPr/w:ind{w:right=-60}"),
            ("w:p", 0, "w:p/w:pPr/w:ind{w:right=0}"),
            ("w:p", None, "w:p/w:pPr"),
            ("w:p/w:pPr/w:ind{w:right=240}", None, "w:p/w:pPr/w:ind"),
        ]
    )
    def right_indent_set_fixture(self, request):
        p_cxml, value, expected_p_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_p_cxml)
        return paragraph_format, value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:spacing", None),
            ("w:p/w:pPr/w:spacing{w:after=240}", Pt(12)),
        ]
    )
    def space_after_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, expected_value

    @pytest.fixture(
        params=[
            ("w:p", Pt(12), "w:p/w:pPr/w:spacing{w:after=240}"),
            ("w:p", None, "w:p/w:pPr"),
            ("w:p/w:pPr", Pt(12), "w:p/w:pPr/w:spacing{w:after=240}"),
            ("w:p/w:pPr", None, "w:p/w:pPr"),
            ("w:p/w:pPr/w:spacing", Pt(12), "w:p/w:pPr/w:spacing{w:after=240}"),
            ("w:p/w:pPr/w:spacing", None, "w:p/w:pPr/w:spacing"),
            (
                "w:p/w:pPr/w:spacing{w:after=240}",
                Pt(42),
                "w:p/w:pPr/w:spacing{w:after=840}",
            ),
            ("w:p/w:pPr/w:spacing{w:after=840}", None, "w:p/w:pPr/w:spacing"),
        ]
    )
    def space_after_set_fixture(self, request):
        p_cxml, value, expected_p_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_p_cxml)
        return paragraph_format, value, expected_xml

    @pytest.fixture(
        params=[
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:spacing", None),
            ("w:p/w:pPr/w:spacing{w:before=420}", Pt(21)),
        ]
    )
    def space_before_get_fixture(self, request):
        p_cxml, expected_value = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        return paragraph_format, expected_value

    @pytest.fixture(
        params=[
            ("w:p", Pt(12), "w:p/w:pPr/w:spacing{w:before=240}"),
            ("w:p", None, "w:p/w:pPr"),
            ("w:p/w:pPr", Pt(12), "w:p/w:pPr/w:spacing{w:before=240}"),
            ("w:p/w:pPr", None, "w:p/w:pPr"),
            ("w:p/w:pPr/w:spacing", Pt(12), "w:p/w:pPr/w:spacing{w:before=240}"),
            ("w:p/w:pPr/w:spacing", None, "w:p/w:pPr/w:spacing"),
            (
                "w:p/w:pPr/w:spacing{w:before=240}",
                Pt(42),
                "w:p/w:pPr/w:spacing{w:before=840}",
            ),
            ("w:p/w:pPr/w:spacing{w:before=840}", None, "w:p/w:pPr/w:spacing"),
        ]
    )
    def space_before_set_fixture(self, request):
        p_cxml, value, expected_p_cxml = request.param
        paragraph_format = ParagraphFormat(element(p_cxml))
        expected_xml = xml(expected_p_cxml)
        return paragraph_format, value, expected_xml

    @pytest.fixture
    def tab_stops_fixture(self, TabStops_, tab_stops_):
        p = element("w:p/w:pPr")
        pPr = p.pPr
        paragraph_format = ParagraphFormat(p, None)
        return paragraph_format, TabStops_, pPr, tab_stops_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def TabStops_(self, request, tab_stops_):
        return class_mock(request, "docx.text.parfmt.TabStops", return_value=tab_stops_)

    @pytest.fixture
    def tab_stops_(self, request):
        return instance_mock(request, TabStops)


class DescribeParagraphFormat_EastAsianTypography:
    """Unit-test suite for ``ParagraphFormat.kinsoku`` and ``.word_wrap``."""

    @pytest.mark.parametrize(
        ("p_cxml", "expected_value"),
        [
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:kinsoku", True),
            ("w:p/w:pPr/w:kinsoku{w:val=1}", True),
            ("w:p/w:pPr/w:kinsoku{w:val=true}", True),
            ("w:p/w:pPr/w:kinsoku{w:val=on}", True),
            ("w:p/w:pPr/w:kinsoku{w:val=0}", False),
            ("w:p/w:pPr/w:kinsoku{w:val=false}", False),
            ("w:p/w:pPr/w:kinsoku{w:val=off}", False),
        ],
    )
    def it_knows_its_kinsoku_value(self, p_cxml: str, expected_value: bool | None):
        paragraph_format = ParagraphFormat(element(p_cxml))
        assert paragraph_format.kinsoku is expected_value

    @pytest.mark.parametrize(
        ("p_cxml", "value", "expected_cxml"),
        [
            ("w:p", True, "w:p/w:pPr/w:kinsoku"),
            ("w:p/w:pPr", True, "w:p/w:pPr/w:kinsoku"),
            ("w:p", False, "w:p/w:pPr/w:kinsoku{w:val=0}"),
            ("w:p/w:pPr/w:kinsoku", False, "w:p/w:pPr/w:kinsoku{w:val=0}"),
            ("w:p/w:pPr/w:kinsoku{w:val=0}", True, "w:p/w:pPr/w:kinsoku"),
            ("w:p/w:pPr/w:kinsoku", None, "w:p/w:pPr"),
            ("w:p", None, "w:p/w:pPr"),
        ],
    )
    def it_can_change_its_kinsoku_value(
        self, p_cxml: str, value: bool | None, expected_cxml: str
    ):
        paragraph_format = ParagraphFormat(element(p_cxml))
        paragraph_format.kinsoku = value
        assert paragraph_format._element.xml == xml(expected_cxml)

    def it_preserves_schema_order_when_inserting_kinsoku(self):
        """w:kinsoku must precede w:bidi in CT_PPr (schema order)."""
        paragraph_format = ParagraphFormat(element("w:p/w:pPr/w:bidi"))
        paragraph_format.kinsoku = True
        assert paragraph_format._element.xml == xml(
            "w:p/w:pPr/(w:kinsoku,w:bidi)"
        )

    @pytest.mark.parametrize(
        ("p_cxml", "expected_value"),
        [
            ("w:p", None),
            ("w:p/w:pPr", None),
            ("w:p/w:pPr/w:wordWrap", True),
            ("w:p/w:pPr/w:wordWrap{w:val=1}", True),
            ("w:p/w:pPr/w:wordWrap{w:val=true}", True),
            ("w:p/w:pPr/w:wordWrap{w:val=on}", True),
            ("w:p/w:pPr/w:wordWrap{w:val=0}", False),
            ("w:p/w:pPr/w:wordWrap{w:val=false}", False),
            ("w:p/w:pPr/w:wordWrap{w:val=off}", False),
        ],
    )
    def it_knows_its_word_wrap_value(self, p_cxml: str, expected_value: bool | None):
        paragraph_format = ParagraphFormat(element(p_cxml))
        assert paragraph_format.word_wrap is expected_value

    @pytest.mark.parametrize(
        ("p_cxml", "value", "expected_cxml"),
        [
            ("w:p", True, "w:p/w:pPr/w:wordWrap"),
            ("w:p/w:pPr", True, "w:p/w:pPr/w:wordWrap"),
            ("w:p", False, "w:p/w:pPr/w:wordWrap{w:val=0}"),
            ("w:p/w:pPr/w:wordWrap", False, "w:p/w:pPr/w:wordWrap{w:val=0}"),
            ("w:p/w:pPr/w:wordWrap{w:val=0}", True, "w:p/w:pPr/w:wordWrap"),
            ("w:p/w:pPr/w:wordWrap", None, "w:p/w:pPr"),
            ("w:p", None, "w:p/w:pPr"),
        ],
    )
    def it_can_change_its_word_wrap_value(
        self, p_cxml: str, value: bool | None, expected_cxml: str
    ):
        paragraph_format = ParagraphFormat(element(p_cxml))
        paragraph_format.word_wrap = value
        assert paragraph_format._element.xml == xml(expected_cxml)

    def it_preserves_schema_order_for_both_kinsoku_and_word_wrap(self):
        paragraph_format = ParagraphFormat(element("w:p"))
        paragraph_format.word_wrap = False
        paragraph_format.kinsoku = True
        # kinsoku precedes wordWrap in schema order
        assert paragraph_format._element.xml == xml(
            "w:p/w:pPr/(w:kinsoku,w:wordWrap{w:val=0})"
        )


class DescribeParagraphFormat_right_to_left:
    """Unit-test suite for `ParagraphFormat.right_to_left`."""

    @pytest.mark.parametrize(
        ("p_cxml", "expected_value"),
        [
            ("w:p", False),
            ("w:p/w:pPr", False),
            ("w:p/w:pPr/w:bidi", True),
            ("w:p/w:pPr/w:bidi{w:val=1}", True),
            ("w:p/w:pPr/w:bidi{w:val=true}", True),
            ("w:p/w:pPr/w:bidi{w:val=on}", True),
            ("w:p/w:pPr/w:bidi{w:val=0}", False),
            ("w:p/w:pPr/w:bidi{w:val=false}", False),
            ("w:p/w:pPr/w:bidi{w:val=off}", False),
        ],
    )
    def it_knows_whether_it_is_right_to_left(
        self, p_cxml: str, expected_value: bool
    ):
        paragraph_format = ParagraphFormat(element(p_cxml))
        assert paragraph_format.right_to_left is expected_value

    @pytest.mark.parametrize(
        ("p_cxml", "value", "expected_cxml"),
        [
            ("w:p", True, "w:p/w:pPr/w:bidi"),
            ("w:p/w:pPr", True, "w:p/w:pPr/w:bidi"),
            ("w:p/w:pPr/w:bidi", False, "w:p/w:pPr"),
            ("w:p/w:pPr/w:bidi", None, "w:p/w:pPr"),
            ("w:p/w:pPr/w:bidi{w:val=off}", True, "w:p/w:pPr/w:bidi"),
            ("w:p", False, "w:p/w:pPr"),
        ],
    )
    def it_can_change_whether_it_is_right_to_left(
        self, p_cxml: str, value: bool | None, expected_cxml: str
    ):
        paragraph_format = ParagraphFormat(element(p_cxml))
        paragraph_format.right_to_left = value
        assert paragraph_format._element.xml == xml(expected_cxml)


class DescribeParagraphFormat_Frame:
    """Unit-test suite for the text-frame (``w:framePr``) API on ParagraphFormat."""

    def it_returns_None_when_pPr_is_absent(self):
        pf = ParagraphFormat(element("w:p"))
        assert pf.frame is None

    def it_returns_None_when_framePr_is_absent(self):
        pf = ParagraphFormat(element("w:p/w:pPr"))
        assert pf.frame is None

    def it_returns_TextFrame_when_framePr_is_present(self):
        pf = ParagraphFormat(element("w:p/w:pPr/w:framePr"))
        assert isinstance(pf.frame, TextFrame)

    def it_can_set_frame_creating_pPr_and_framePr(self):
        pf = ParagraphFormat(element("w:p"))

        frame = pf.set_frame(width=Twips(1440), wrap=WD_FRAME_WRAP.AROUND)

        assert isinstance(frame, TextFrame)
        expected_xml = xml(
            "w:p/w:pPr/w:framePr{w:w=1440,w:wrap=around}"
        )
        assert pf._element.xml == expected_xml

    def it_updates_existing_framePr_incrementally(self):
        pf = ParagraphFormat(element("w:p/w:pPr/w:framePr{w:w=1440}"))

        pf.set_frame(height=Twips(2880), horizontal_anchor=WD_FRAME_H_ANCHOR.PAGE)

        expected_xml = xml(
            "w:p/w:pPr/w:framePr{w:w=1440,w:h=2880,w:hAnchor=page}"
        )
        assert pf._element.xml == expected_xml

    def it_sets_every_attribute(self):
        pf = ParagraphFormat(element("w:p"))

        pf.set_frame(
            width=Twips(1440),
            height=Twips(2880),
            horizontal_position=Twips(720),
            vertical_position=Twips(360),
            horizontal_anchor=WD_FRAME_H_ANCHOR.PAGE,
            vertical_anchor=WD_FRAME_V_ANCHOR.MARGIN,
            wrap=WD_FRAME_WRAP.AROUND,
            drop_cap=WD_FRAME_DROP_CAP.DROP,
            lines=3,
            horizontal_alignment=WD_FRAME_H_ALIGN.CENTER,
            vertical_alignment=WD_FRAME_V_ALIGN.TOP,
        )

        frame = pf.frame
        assert frame is not None
        assert frame.width == Twips(1440)
        assert frame.height == Twips(2880)
        assert frame.horizontal_position == Twips(720)
        assert frame.vertical_position == Twips(360)
        assert frame.horizontal_anchor == WD_FRAME_H_ANCHOR.PAGE
        assert frame.vertical_anchor == WD_FRAME_V_ANCHOR.MARGIN
        assert frame.wrap == WD_FRAME_WRAP.AROUND
        assert frame.drop_cap == WD_FRAME_DROP_CAP.DROP
        assert frame.lines == 3
        assert frame.horizontal_alignment == WD_FRAME_H_ALIGN.CENTER
        assert frame.vertical_alignment == WD_FRAME_V_ALIGN.TOP

    def it_can_remove_frame(self):
        pf = ParagraphFormat(element("w:p/w:pPr/w:framePr{w:w=1440}"))

        pf.remove_frame()

        assert pf.frame is None
        expected_xml = xml("w:p/w:pPr")
        assert pf._element.xml == expected_xml

    def it_remove_frame_is_noop_when_pPr_absent(self):
        pf = ParagraphFormat(element("w:p"))
        pf.remove_frame()  # should not raise
        expected_xml = xml("w:p")
        assert pf._element.xml == expected_xml

    def it_remove_frame_is_noop_when_framePr_absent(self):
        pf = ParagraphFormat(element("w:p/w:pPr"))
        pf.remove_frame()  # should not raise
        expected_xml = xml("w:p/w:pPr")
        assert pf._element.xml == expected_xml


class DescribeTextFrame:
    """Unit-test suite for the TextFrame proxy class."""

    @pytest.mark.parametrize(
        ("cxml", "prop", "expected"),
        [
            ("w:framePr{w:w=1440}", "width", Twips(1440)),
            ("w:framePr{w:h=2880}", "height", Twips(2880)),
            ("w:framePr{w:x=720}", "horizontal_position", Twips(720)),
            ("w:framePr{w:y=-360}", "vertical_position", Twips(-360)),
            ("w:framePr{w:hAnchor=page}", "horizontal_anchor", WD_FRAME_H_ANCHOR.PAGE),
            ("w:framePr{w:vAnchor=margin}", "vertical_anchor", WD_FRAME_V_ANCHOR.MARGIN),
            ("w:framePr{w:wrap=around}", "wrap", WD_FRAME_WRAP.AROUND),
            ("w:framePr{w:dropCap=drop}", "drop_cap", WD_FRAME_DROP_CAP.DROP),
            ("w:framePr{w:lines=4}", "lines", 4),
            ("w:framePr{w:xAlign=center}", "horizontal_alignment", WD_FRAME_H_ALIGN.CENTER),
            ("w:framePr{w:yAlign=top}", "vertical_alignment", WD_FRAME_V_ALIGN.TOP),
            ("w:framePr", "width", None),
            ("w:framePr", "height", None),
            ("w:framePr", "horizontal_position", None),
            ("w:framePr", "vertical_position", None),
            ("w:framePr", "horizontal_anchor", None),
            ("w:framePr", "vertical_anchor", None),
            ("w:framePr", "wrap", None),
            ("w:framePr", "drop_cap", None),
            ("w:framePr", "lines", None),
            ("w:framePr", "horizontal_alignment", None),
            ("w:framePr", "vertical_alignment", None),
        ],
    )
    def it_reads_attributes(self, cxml, prop, expected):
        framePr = element(cxml)
        frame = TextFrame(framePr)
        assert getattr(frame, prop) == expected

    @pytest.mark.parametrize(
        ("prop", "value", "expected_cxml"),
        [
            ("width", Twips(1440), "w:framePr{w:w=1440}"),
            ("height", Twips(2880), "w:framePr{w:h=2880}"),
            ("horizontal_position", Twips(720), "w:framePr{w:x=720}"),
            ("vertical_position", Twips(-360), "w:framePr{w:y=-360}"),
            ("horizontal_anchor", WD_FRAME_H_ANCHOR.PAGE, "w:framePr{w:hAnchor=page}"),
            ("vertical_anchor", WD_FRAME_V_ANCHOR.MARGIN, "w:framePr{w:vAnchor=margin}"),
            ("wrap", WD_FRAME_WRAP.TIGHT, "w:framePr{w:wrap=tight}"),
            ("drop_cap", WD_FRAME_DROP_CAP.DROP, "w:framePr{w:dropCap=drop}"),
            ("lines", 5, "w:framePr{w:lines=5}"),
            ("horizontal_alignment", WD_FRAME_H_ALIGN.CENTER, "w:framePr{w:xAlign=center}"),
            ("vertical_alignment", WD_FRAME_V_ALIGN.BOTTOM, "w:framePr{w:yAlign=bottom}"),
        ],
    )
    def it_writes_attributes(self, prop, value, expected_cxml):
        framePr = element("w:framePr")
        frame = TextFrame(framePr)
        setattr(frame, prop, value)
        assert framePr.xml == xml(expected_cxml)

    def it_clears_attribute_when_set_to_None(self):
        framePr = element("w:framePr{w:w=1440,w:wrap=around}")
        frame = TextFrame(framePr)

        frame.width = None

        assert framePr.xml == xml("w:framePr{w:wrap=around}")


class DescribeParagraphFormatNewFields:
    """Unit tests for Phase C field additions on ParagraphFormat."""

    # -- outline_level --------------------------------------------------

    def it_reads_outline_level_as_None_when_not_present(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        assert ParagraphFormat(p).outline_level is None

    def it_reads_outline_level_value(self):
        from docx.enum.text import WD_OUTLINELVL
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:outlineLvl{w:val=3}")
        assert ParagraphFormat(p).outline_level == WD_OUTLINELVL.LEVEL_4

    def it_writes_outline_level(self):
        from docx.enum.text import WD_OUTLINELVL
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        ParagraphFormat(p).outline_level = WD_OUTLINELVL.LEVEL_1
        assert p.xml == xml("w:p/w:pPr/w:outlineLvl{w:val=0}")

    def it_writes_body_text_outline_level(self):
        from docx.enum.text import WD_OUTLINELVL
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        ParagraphFormat(p).outline_level = WD_OUTLINELVL.BODY_TEXT
        assert p.xml == xml("w:p/w:pPr/w:outlineLvl{w:val=10}")

    def it_clears_outline_level(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:outlineLvl{w:val=0}")
        ParagraphFormat(p).outline_level = None
        assert p.xml == xml("w:p/w:pPr")

    def it_raises_on_invalid_outline_level(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        with pytest.raises(ValueError):
            ParagraphFormat(p).outline_level = 42

    # -- contextual_spacing ---------------------------------------------

    def it_reads_contextual_spacing(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:contextualSpacing")
        assert ParagraphFormat(p).contextual_spacing is True

    def it_writes_contextual_spacing(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        ParagraphFormat(p).contextual_spacing = True
        assert p.xml == xml("w:p/w:pPr/w:contextualSpacing")

    def it_clears_contextual_spacing(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:contextualSpacing")
        ParagraphFormat(p).contextual_spacing = None
        assert p.xml == xml("w:p/w:pPr")

    # -- first_line_chars ------------------------------------------------

    def it_reads_first_line_chars(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:ind{w:firstLineChars=200}")
        assert ParagraphFormat(p).first_line_chars == 200

    def it_reads_first_line_chars_as_None(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        assert ParagraphFormat(p).first_line_chars is None

    def it_writes_first_line_chars(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        ParagraphFormat(p).first_line_chars = 300
        assert p.xml == xml("w:p/w:pPr/w:ind{w:firstLineChars=300}")

    def it_clears_first_line_chars(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:ind{w:firstLineChars=200}")
        ParagraphFormat(p).first_line_chars = None
        assert p.xml == xml("w:p/w:pPr/w:ind")

    # -- auto_space_de / auto_space_dn ----------------------------------

    def it_reads_auto_space_de(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:autoSpaceDE{w:val=0}")
        assert ParagraphFormat(p).auto_space_de is False

    def it_writes_auto_space_de(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        ParagraphFormat(p).auto_space_de = False
        assert p.xml == xml("w:p/w:pPr/w:autoSpaceDE{w:val=0}")

    def it_reads_auto_space_dn(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:autoSpaceDN{w:val=0}")
        assert ParagraphFormat(p).auto_space_dn is False

    def it_writes_auto_space_dn(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        ParagraphFormat(p).auto_space_dn = True
        assert p.xml == xml("w:p/w:pPr/w:autoSpaceDN")

    # -- shading_color --------------------------------------------------

    def it_reads_shading_color(self):
        from docx.shared import RGBColor
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:shd{w:val=clear,w:fill=FFFF00}")
        assert ParagraphFormat(p).shading_color == RGBColor(0xFF, 0xFF, 0x00)

    def it_reads_shading_color_as_None_when_absent(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        assert ParagraphFormat(p).shading_color is None

    def it_writes_shading_color(self):
        from docx.shared import RGBColor
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p")
        ParagraphFormat(p).shading_color = RGBColor(0xAB, 0xCD, 0xEF)
        assert p.xml == xml("w:p/w:pPr/w:shd{w:val=clear,w:fill=ABCDEF}")

    def it_clears_shading_color(self):
        from docx.text.parfmt import ParagraphFormat

        p = element("w:p/w:pPr/w:shd{w:val=clear,w:fill=FFFF00}")
        ParagraphFormat(p).shading_color = None
        assert p.xml == xml("w:p/w:pPr")


class DescribeDropCap:
    """Unit-test suite for the DropCap proxy class."""

    @pytest.mark.parametrize(
        ("cxml", "prop", "expected"),
        [
            ("w:framePr{w:dropCap=drop}", "mode", WD_FRAME_DROP_CAP.DROP),
            ("w:framePr{w:dropCap=margin}", "mode", WD_FRAME_DROP_CAP.MARGIN),
            ("w:framePr{w:lines=3}", "lines", 3),
            ("w:framePr{w:x=720}", "x", Twips(720)),
            ("w:framePr{w:y=-360}", "y", Twips(-360)),
            ("w:framePr{w:w=1440}", "width", Twips(1440)),
            ("w:framePr{w:h=2880}", "height", Twips(2880)),
            ("w:framePr{w:wrap=around}", "wrap", WD_FRAME_WRAP.AROUND),
            (
                "w:framePr{w:hAnchor=page}",
                "horizontal_anchor",
                WD_FRAME_H_ANCHOR.PAGE,
            ),
            (
                "w:framePr{w:vAnchor=margin}",
                "vertical_anchor",
                WD_FRAME_V_ANCHOR.MARGIN,
            ),
            ("w:framePr", "mode", None),
            ("w:framePr", "lines", None),
            ("w:framePr", "x", None),
            ("w:framePr", "y", None),
            ("w:framePr", "width", None),
            ("w:framePr", "height", None),
            ("w:framePr", "wrap", None),
            ("w:framePr", "horizontal_anchor", None),
            ("w:framePr", "vertical_anchor", None),
        ],
    )
    def it_reads_attributes(self, cxml, prop, expected):
        framePr = element(cxml)
        assert getattr(DropCap(framePr), prop) == expected

    @pytest.mark.parametrize(
        ("prop", "value", "expected_cxml"),
        [
            ("mode", WD_FRAME_DROP_CAP.DROP, "w:framePr{w:dropCap=drop}"),
            ("mode", WD_FRAME_DROP_CAP.MARGIN, "w:framePr{w:dropCap=margin}"),
            ("lines", 4, "w:framePr{w:lines=4}"),
            ("x", Twips(720), "w:framePr{w:x=720}"),
            ("y", Twips(-180), "w:framePr{w:y=-180}"),
            ("width", Twips(1440), "w:framePr{w:w=1440}"),
            ("height", Twips(2880), "w:framePr{w:h=2880}"),
            ("wrap", WD_FRAME_WRAP.AROUND, "w:framePr{w:wrap=around}"),
            (
                "horizontal_anchor",
                WD_FRAME_H_ANCHOR.PAGE,
                "w:framePr{w:hAnchor=page}",
            ),
            (
                "vertical_anchor",
                WD_FRAME_V_ANCHOR.MARGIN,
                "w:framePr{w:vAnchor=margin}",
            ),
        ],
    )
    def it_writes_attributes(self, prop, value, expected_cxml):
        framePr = element("w:framePr")
        setattr(DropCap(framePr), prop, value)
        assert framePr.xml == xml(expected_cxml)

    def it_clears_attribute_when_set_to_None(self):
        framePr = element("w:framePr{w:dropCap=drop,w:lines=3}")
        dc = DropCap(framePr)

        dc.lines = None

        assert framePr.xml == xml("w:framePr{w:dropCap=drop}")
