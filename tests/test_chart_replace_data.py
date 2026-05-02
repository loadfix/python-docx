# pyright: reportPrivateUsage=false

"""Unit-test suite for `Chart.replace_data`."""

from __future__ import annotations

from typing import cast

import pytest

from docx.chart import Chart
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.chart import CT_ChartSpace
from docx.oxml.ns import qn
from docx.package import Package
from docx.parts.chart import ChartPart, _rewrite_ser

from .unitutil.cxml import element


def _make_chart(cxml: str) -> Chart:
    chartSpace = cast(CT_ChartSpace, element(cxml))
    package = Package()
    part = ChartPart(
        PackURI("/word/charts/chart1.xml"), CT.DML_CHART, chartSpace, package
    )
    return Chart(part)


class DescribeChart_replace_data:
    """Unit-test suite for `docx.chart.Chart.replace_data`."""

    def it_rewrites_categories_and_values_on_a_single_series(self):
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:barChart/"
            "(c:barDir{val=col},c:ser/("
            "c:idx{val=0},c:order{val=0},"
            'c:tx/c:v"Old",'
            'c:cat/c:strRef/c:strCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"a"),'
            "c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"1\")"
            "))"
        )
        chart = _make_chart(cxml)

        chart.replace_data(["Q1", "Q2", "Q3"], {"Revenue": [10.0, 20.0, 30.0]})

        assert chart.categories == ["Q1", "Q2", "Q3"]
        assert chart.series[0].name == "Revenue"
        assert chart.series[0].values == [10.0, 20.0, 30.0]

    def it_preserves_non_data_styling_children_on_series(self):
        # -- include a c:spPr (styling) child that replace_data must preserve --
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:barChart/"
            "(c:barDir{val=col},c:ser/("
            "c:idx{val=0},c:order{val=0},"
            'c:tx/c:v"Old",'
            "c:spPr,"
            'c:cat/c:strRef/c:strCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"a"),'
            "c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"1\")"
            "))"
        )
        chart = _make_chart(cxml)
        chart.replace_data(["A"], {"New": [5.0]})
        ser = chart.part.chartSpace.xpath(".//c:ser")[0]
        assert ser.find(qn("c:spPr")) is not None

    def it_clones_the_last_series_when_adding_more(self):
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:barChart/"
            "(c:barDir{val=col},c:ser/("
            "c:idx{val=0},c:order{val=0},"
            'c:tx/c:v"S0",'
            "c:spPr,"
            'c:cat/c:strRef/c:strCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"a"),'
            "c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"1\")"
            "))"
        )
        chart = _make_chart(cxml)

        chart.replace_data(
            ["Q1", "Q2"],
            {"A": [1.0, 2.0], "B": [3.0, 4.0]},
        )

        assert [s.name for s in chart.series] == ["A", "B"]
        assert [s.values for s in chart.series] == [[1.0, 2.0], [3.0, 4.0]]
        # -- cloned series keeps c:spPr from the template --
        for ser in chart.part.chartSpace.xpath(".//c:ser"):
            assert ser.find(qn("c:spPr")) is not None

    def it_removes_excess_series_when_shrinking(self):
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:barChart/"
            "(c:barDir{val=col},"
            "c:ser/("
            "c:idx{val=0},c:order{val=0},"
            'c:tx/c:v"S0",'
            'c:cat/c:strRef/c:strCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"a"),'
            "c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"1\")),"
            "c:ser/("
            "c:idx{val=1},c:order{val=1},"
            'c:tx/c:v"S1",'
            'c:cat/c:strRef/c:strCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"a"),'
            "c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"2\"))"
            ")"
        )
        chart = _make_chart(cxml)

        chart.replace_data(["X"], {"OnlyOne": [99.0]})

        assert len(chart.series) == 1
        assert chart.series[0].name == "OnlyOne"
        assert chart.series[0].values == [99.0]

    def it_raises_when_series_length_mismatches_categories(self):
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:barChart/"
            "(c:barDir{val=col},c:ser/("
            "c:idx{val=0},c:order{val=0},"
            'c:tx/c:v"S0",'
            'c:cat/c:strRef/c:strCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"a"),'
            "c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"1\")"
            "))"
        )
        chart = _make_chart(cxml)
        with pytest.raises(ValueError, match="3 categories"):
            chart.replace_data(["a", "b", "c"], {"X": [1.0, 2.0]})

    def it_raises_when_no_existing_series(self):
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:barChart/c:barDir{val=col}"
        )
        chart = _make_chart(cxml)
        with pytest.raises(ValueError, match="at least one existing c:ser"):
            chart.replace_data(["a"], {"X": [1.0]})

    def it_preserves_chart_type_across_replacement(self):
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:lineChart/"
            "c:ser/("
            "c:idx{val=0},c:order{val=0},"
            'c:tx/c:v"S0",'
            'c:cat/c:strRef/c:strCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"a"),'
            "c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"1\")"
            ")"
        )
        chart = _make_chart(cxml)
        before = chart.chart_type
        chart.replace_data(["A", "B"], {"New": [1.0, 2.0]})
        assert chart.chart_type is before


class DescribeRewriteSer:
    """Direct tests for the `_rewrite_ser` helper."""

    def it_preserves_non_data_siblings_in_order(self):
        cxml = (
            "c:ser/("
            "c:idx{val=0},c:order{val=0},"
            'c:tx/c:v"Old",'
            "c:spPr,"
            "c:smooth"
            ")"
        )
        ser = element(cxml)
        _rewrite_ser(ser, idx=3, name="New", categories=["a"], values=[42.0])
        # -- spPr and smooth remain after the rewritten data elements --
        tags = [child.tag for child in list(ser)]
        assert qn("c:idx") in tags
        assert qn("c:order") in tags
        assert qn("c:spPr") in tags
        assert qn("c:smooth") in tags
        # -- idx is updated --
        assert ser.find(qn("c:idx")).get("val") == "3"
