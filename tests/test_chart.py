# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.chart` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.chart import Chart, ChartSeries, WD_CHART_TYPE
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.chart import CT_ChartSpace, CT_Ser
from docx.package import Package
from docx.parts.chart import ChartPart

from .unitutil.cxml import element


def _chart_part(cxml: str) -> ChartPart:
    chartSpace = cast(CT_ChartSpace, element(cxml))
    package = Package()
    return ChartPart(PackURI("/word/charts/chart1.xml"), CT.DML_CHART, chartSpace, package)


class DescribeWD_CHART_TYPE:
    def it_is_an_enum_with_expected_members(self):
        assert WD_CHART_TYPE.BAR.value == "bar"
        assert WD_CHART_TYPE.BAR_STACKED.value == "barStacked"
        assert WD_CHART_TYPE.COLUMN.value == "column"
        assert WD_CHART_TYPE.COLUMN_STACKED.value == "columnStacked"
        assert WD_CHART_TYPE.LINE.value == "line"
        assert WD_CHART_TYPE.PIE.value == "pie"
        assert WD_CHART_TYPE.DOUGHNUT.value == "doughnut"
        assert WD_CHART_TYPE.SCATTER.value == "scatter"
        assert WD_CHART_TYPE.AREA.value == "area"


class DescribeChart:
    @pytest.mark.parametrize(
        ("kind_cxml", "expected"),
        [
            ("c:barChart/c:barDir{val=bar}", WD_CHART_TYPE.BAR),
            (
                "c:barChart/(c:barDir{val=bar},c:grouping{val=stacked})",
                WD_CHART_TYPE.BAR_STACKED,
            ),
            ("c:barChart/c:barDir{val=col}", WD_CHART_TYPE.COLUMN),
            (
                "c:barChart/(c:barDir{val=col},c:grouping{val=stacked})",
                WD_CHART_TYPE.COLUMN_STACKED,
            ),
            ("c:lineChart", WD_CHART_TYPE.LINE),
            ("c:pieChart", WD_CHART_TYPE.PIE),
            ("c:doughnutChart", WD_CHART_TYPE.DOUGHNUT),
            ("c:scatterChart", WD_CHART_TYPE.SCATTER),
            ("c:areaChart", WD_CHART_TYPE.AREA),
        ],
    )
    def it_identifies_its_chart_type(self, kind_cxml: str, expected: WD_CHART_TYPE):
        cxml = f"c:chartSpace/c:chart/c:plotArea/{kind_cxml}"
        part = _chart_part(cxml)
        chart = Chart(part)
        assert chart.chart_type is expected

    def its_chart_type_is_None_when_no_kind_element(self):
        part = _chart_part("c:chartSpace/c:chart/c:plotArea")
        assert Chart(part).chart_type is None

    def it_reads_its_title(self):
        cxml = (
            'c:chartSpace/c:chart/(c:title/c:tx/c:rich/a:p/a:r/a:t"Sales",'
            "c:plotArea/c:barChart/c:barDir{val=bar})"
        )
        part = _chart_part(cxml)
        assert Chart(part).title == "Sales"

    def its_title_is_None_when_absent(self):
        cxml = "c:chartSpace/c:chart/c:plotArea/c:barChart/c:barDir{val=bar}"
        part = _chart_part(cxml)
        assert Chart(part).title is None

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("c:chartSpace/c:chart/c:plotArea", False),
            ("c:chartSpace/c:chart/(c:plotArea,c:legend)", True),
        ],
    )
    def it_knows_whether_it_has_a_legend(self, cxml: str, expected: bool):
        part = _chart_part(cxml)
        assert Chart(part).has_legend is expected

    def it_provides_access_to_its_series(self):
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:barChart"
            '/(c:barDir{val=bar},c:ser/c:tx/c:v"S1",c:ser/c:tx/c:v"S2")'
        )
        part = _chart_part(cxml)
        chart = Chart(part)
        names = [s.name for s in chart.series]
        assert names == ["S1", "S2"]

    def it_returns_empty_series_when_no_plotArea(self):
        part = _chart_part("c:chartSpace/c:chart")
        assert Chart(part).series == []

    def it_returns_categories_from_the_first_series(self):
        cxml = (
            "c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser/"
            "c:cat/c:strRef/c:strCache/"
            '(c:pt{idx=0}/c:v"A",c:pt{idx=1}/c:v"B")'
        )
        part = _chart_part(cxml)
        assert Chart(part).categories == ["A", "B"]

    def it_returns_empty_categories_when_no_series(self):
        cxml = "c:chartSpace/c:chart/c:plotArea/c:barChart"
        part = _chart_part(cxml)
        assert Chart(part).categories == []


class DescribeChartSeries:
    def it_exposes_name_values_and_categories(self):
        cxml = (
            "c:ser/"
            '(c:tx/c:v"Rev",'
            "c:cat/c:strRef/c:strCache/"
            '(c:pt{idx=0}/c:v"Q1",c:pt{idx=1}/c:v"Q2"),'
            "c:val/c:numRef/c:numCache/"
            '(c:pt{idx=0}/c:v"10",c:pt{idx=1}/c:v"20"))'
        )
        ser = cast(CT_Ser, element(cxml))
        series = ChartSeries(ser)

        assert series.name == "Rev"
        assert series.categories == ["Q1", "Q2"]
        assert series.values == [10.0, 20.0]

    def its_name_is_empty_string_when_not_set(self):
        ser = cast(CT_Ser, element("c:ser"))
        assert ChartSeries(ser).name == ""
