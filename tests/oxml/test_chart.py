# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.chart` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.chart import CT_BarChart, CT_Chart, CT_ChartSpace, CT_PlotArea, CT_Ser

from ..unitutil.cxml import element


class DescribeCT_ChartSpace:
    def it_provides_access_to_its_chart_child(self):
        cs = cast(CT_ChartSpace, element("c:chartSpace/c:chart"))
        assert cs.chart is not None
        assert isinstance(cs.chart, CT_Chart)

    def and_returns_None_when_chart_is_absent(self):
        cs = cast(CT_ChartSpace, element("c:chartSpace"))
        assert cs.chart is None


class DescribeCT_Chart:
    def it_provides_access_to_its_plotArea(self):
        chart = cast(CT_Chart, element("c:chart/c:plotArea"))
        assert chart.plotArea is not None
        assert isinstance(chart.plotArea, CT_PlotArea)

    def it_extracts_the_title_text(self):
        cxml = 'c:chart/c:title/c:tx/c:rich/a:p/(a:r/a:t"Foo",a:r/a:t" Bar")'
        chart = cast(CT_Chart, element(cxml))
        assert chart.title_text == "Foo Bar"

    def it_returns_None_when_no_title_present(self):
        chart = cast(CT_Chart, element("c:chart/c:plotArea"))
        assert chart.title_text is None

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("c:chart/c:plotArea", False),
            ("c:chart/(c:plotArea,c:legend)", True),
        ],
    )
    def it_knows_whether_it_has_a_legend(self, cxml: str, expected: bool):
        chart = cast(CT_Chart, element(cxml))
        assert chart.has_legend == expected


class DescribeCT_PlotArea:
    @pytest.mark.parametrize(
        ("child_tag",),
        [
            ("c:barChart",),
            ("c:lineChart",),
            ("c:pieChart",),
            ("c:doughnutChart",),
            ("c:scatterChart",),
            ("c:areaChart",),
        ],
    )
    def it_finds_its_chart_kind_element(self, child_tag: str):
        plotArea = cast(CT_PlotArea, element(f"c:plotArea/(c:layout,{child_tag})"))
        kind = plotArea.chart_kind_element
        assert kind is not None
        assert kind.tag.endswith(child_tag.split(":")[1])

    def it_returns_None_when_no_kind_child_present(self):
        plotArea = cast(CT_PlotArea, element("c:plotArea/c:layout"))
        assert plotArea.chart_kind_element is None

    def it_lists_its_series(self):
        cxml = "c:plotArea/c:barChart/(c:ser,c:ser,c:ser)"
        plotArea = cast(CT_PlotArea, element(cxml))
        assert len(plotArea.ser_lst) == 3


class DescribeCT_BarChart:
    @pytest.mark.parametrize(
        ("direction",),
        [("bar",), ("col",)],
    )
    def it_reads_its_bar_direction(self, direction: str):
        bar = cast(
            CT_BarChart,
            element(f"c:barChart/c:barDir{{val={direction}}}"),
        )
        assert bar.bar_dir == direction

    def it_reads_its_grouping(self):
        bar = cast(CT_BarChart, element("c:barChart/c:grouping{val=stacked}"))
        assert bar.grouping == "stacked"


class DescribeCT_Ser:
    def it_reads_its_name_from_strCache(self):
        cxml = (
            "c:ser/c:tx/c:strRef/c:strCache/c:pt{idx=0}"
            '/c:v"Revenue"'
        )
        ser = cast(CT_Ser, element(cxml))
        assert ser.tx_name == "Revenue"

    def it_reads_its_name_from_literal_v(self):
        cxml = 'c:ser/c:tx/c:v"Inline Name"'
        ser = cast(CT_Ser, element(cxml))
        assert ser.tx_name == "Inline Name"

    def its_name_is_None_when_no_tx(self):
        ser = cast(CT_Ser, element("c:ser"))
        assert ser.tx_name is None

    def it_reads_categories_from_strCache(self):
        cxml = (
            "c:ser/c:cat/c:strRef/c:strCache/"
            '(c:pt{idx=0}/c:v"Q1",c:pt{idx=1}/c:v"Q2",c:pt{idx=2}/c:v"Q3")'
        )
        ser = cast(CT_Ser, element(cxml))
        assert ser.cat_values == ["Q1", "Q2", "Q3"]

    def it_returns_empty_categories_when_absent(self):
        ser = cast(CT_Ser, element("c:ser"))
        assert ser.cat_values == []

    def it_reads_values_from_numCache(self):
        cxml = (
            "c:ser/c:val/c:numRef/c:numCache/"
            '(c:pt{idx=0}/c:v"1.5",c:pt{idx=1}/c:v"2.0",c:pt{idx=2}/c:v"3.25")'
        )
        ser = cast(CT_Ser, element(cxml))
        assert ser.val_values == [1.5, 2.0, 3.25]

    def it_falls_back_to_numLit_for_values(self):
        cxml = (
            "c:ser/c:val/c:numLit/"
            '(c:pt{idx=0}/c:v"7",c:pt{idx=1}/c:v"8")'
        )
        ser = cast(CT_Ser, element(cxml))
        assert ser.val_values == [7.0, 8.0]

    def it_returns_empty_values_when_absent(self):
        ser = cast(CT_Ser, element("c:ser"))
        assert ser.val_values == []

    def it_skips_unparseable_value_points(self):
        cxml = (
            "c:ser/c:val/c:numRef/c:numCache/"
            '(c:pt{idx=0}/c:v"1.0",c:pt{idx=1}/c:v"not-a-number",c:pt{idx=2}/c:v"3.0")'
        )
        ser = cast(CT_Ser, element(cxml))
        assert ser.val_values == [1.0, 3.0]
