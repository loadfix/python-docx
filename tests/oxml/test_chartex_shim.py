# pyright: reportPrivateUsage=false

"""Smoke tests for the `docx.oxml.chartex` re-export shim.

Confirms that the Office 2016+ ChartEx (``cx:`` namespace) ``CT_Cx*``
element classes shipped by ``python-ooxml-chart`` 0.3.0 resolve under
their conventional docx path. Kept narrow — the shared package owns
the exhaustive layout-by-layout coverage (see
``python-ooxml-chart/tests/unit/test_oxml_chartex.py``); this suite
guards docx-side adoption only.

docx's public ``Chart`` proxy is classic-chart-only and does not
promote ChartEx parts onto its read surface; this shim supports
downstream consumers that want typed lxml access to a parsed
``cx:chartSpace`` tree without importing the shared package directly.

.. versionadded:: 2026.05.0
   Landed alongside chart 0.3.0 adoption (round2/chart-0.3-adoption).
"""

from __future__ import annotations

from typing import cast

from ooxml_chart.oxml import nsdecls, parse_xml

from docx.oxml.chartex import (
    CT_CxBoxWhiskerLayout,
    CT_CxChart,
    CT_CxChartData,
    CT_CxChartSpace,
    CT_CxClusteredColumnLayout,
    CT_CxFunnelLayout,
    CT_CxSeries,
    CT_CxSunburstLayout,
    CT_CxTreemapLayout,
    CT_CxWaterfallLayout,
)


class DescribeChartExShim:
    """The shim is a thin re-export — only verify identity + parse."""

    def it_reexports_all_six_core_layoutId_dispatch_classes(self):
        # -- one class per core ChartEx layoutId (per 0.3.0 changelog) --
        classes = [
            CT_CxTreemapLayout,
            CT_CxSunburstLayout,
            CT_CxFunnelLayout,
            CT_CxWaterfallLayout,
            CT_CxClusteredColumnLayout,  # histogram / Pareto
            CT_CxBoxWhiskerLayout,
        ]
        for cls in classes:
            assert cls.__module__ == "ooxml_chart.oxml.chartex", (
                f"{cls.__name__} resolved through {cls.__module__}"
                " but should have come from the shared ooxml_chart package"
            )

    def it_parses_a_minimal_chartex_funnel_tree(self):
        xml = (
            "<cx:chartSpace %s>"
            '<cx:chartData><cx:data id="0"/></cx:chartData>'
            "<cx:chart><cx:plotArea><cx:plotAreaRegion>"
            '<cx:series layoutId="funnel"/>'
            "</cx:plotAreaRegion></cx:plotArea></cx:chart>"
            "</cx:chartSpace>" % nsdecls("cx")
        ).encode()

        chartSpace = cast(CT_CxChartSpace, parse_xml(xml))

        assert isinstance(chartSpace, CT_CxChartSpace)
        assert isinstance(chartSpace.chartData, CT_CxChartData)
        assert isinstance(chartSpace.chart, CT_CxChart)

        series_list = chartSpace.xpath(".//cx:series")
        assert series_list, "expected one cx:series child"
        series = cast(CT_CxSeries, series_list[0])
        assert series.get("layoutId") == "funnel"
