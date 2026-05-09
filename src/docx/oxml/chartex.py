"""Custom element classes for Office 2016+ ChartEx (``cx:``) elements.

Re-exports the ``CT_Cx*`` classes from :mod:`ooxml_chart.oxml.chartex`.
The shared package holds the authoritative definitions — docx keeps
this module as a thin shim so a ``from docx.oxml.chartex import
CT_Cx*`` import resolves to the shared-package class.

ChartEx is the Microsoft-extension chart vocabulary introduced with
Office 2016 and specified in ``[MS-ODRAWXML]`` (not in ECMA-376). It
covers eight ``cx:series/@layoutId`` dispatches — ``treemap``,
``sunburst``, ``funnel``, ``waterfall``, ``clusteredColumn`` (histogram
/ Pareto, driven by ``cx:binning``), ``boxWhisker``, ``regionMap``,
and ``paretoLine``.

docx-local divergences: **none required**. ChartEx does not reuse
docx's existing chart-proxy enum surface (``WD_CHART_TYPE`` covers
only the classic ``c:barChart`` / ``c:lineChart`` / etc. types). The
shared-package classes expose plain attribute types — every ``ST_*``
on the ChartEx side is a string-enumeration scalar — so no docx
subclassing is needed. Descriptor-driven child creation that resolves
to classic ``a:`` / ``c:`` elements underneath a ``cx:chartSpace``
inherits docx's registered classes through the composite namespace
registry (see the registry-restore pattern in :mod:`docx.oxml.chart`).

docx's public :class:`docx.chart.Chart` proxy is classic-chart-only.
Word-authored ChartEx parts land in the zip but are not promoted to
the :class:`Chart` surface; this shim supports downstream consumers
that want typed lxml access to a parsed ``cx:chartSpace`` tree
without importing the shared package directly.

.. versionadded:: 2026.05.0
   Introduced alongside the ``python-ooxml-chart`` 0.3.0 ChartEx
   layer. Namespace / content-type / relationship-type constants
   (``NS_CX``, ``CONTENT_TYPE_CHARTEX``,
   ``RELATIONSHIP_TYPE_CHARTEX``) are re-exported from
   :mod:`ooxml_chart` alongside these ``CT_Cx*`` classes.
"""

from __future__ import annotations

from ooxml_chart.oxml.chartex import (
    CT_CxAxis,
    CT_CxBinning,
    CT_CxBoxWhiskerLayout,
    CT_CxChart,
    CT_CxChartData,
    CT_CxChartSpace,
    CT_CxClusteredColumnLayout,
    CT_CxData,
    CT_CxDataLabel,
    CT_CxDataLabels,
    CT_CxDataPoint,
    CT_CxExternalData,
    CT_CxFormula,
    CT_CxFunnelLayout,
    CT_CxLayoutPr,
    CT_CxLegend,
    CT_CxNumDim,
    CT_CxNumericFormula,
    CT_CxParetoLineLayout,
    CT_CxPlotArea,
    CT_CxPlotAreaRegion,
    CT_CxPlotSurface,
    CT_CxRegionMapLayout,
    CT_CxSeries,
    CT_CxStrDim,
    CT_CxStringLevel,
    CT_CxSunburstLayout,
    CT_CxTitle,
    CT_CxTreemapLayout,
    CT_CxTx,
    CT_CxWaterfallLayout,
)

__all__ = [
    "CT_CxAxis",
    "CT_CxBinning",
    "CT_CxBoxWhiskerLayout",
    "CT_CxChart",
    "CT_CxChartData",
    "CT_CxChartSpace",
    "CT_CxClusteredColumnLayout",
    "CT_CxData",
    "CT_CxDataLabel",
    "CT_CxDataLabels",
    "CT_CxDataPoint",
    "CT_CxExternalData",
    "CT_CxFormula",
    "CT_CxFunnelLayout",
    "CT_CxLayoutPr",
    "CT_CxLegend",
    "CT_CxNumDim",
    "CT_CxNumericFormula",
    "CT_CxParetoLineLayout",
    "CT_CxPlotArea",
    "CT_CxPlotAreaRegion",
    "CT_CxPlotSurface",
    "CT_CxRegionMapLayout",
    "CT_CxSeries",
    "CT_CxStrDim",
    "CT_CxStringLevel",
    "CT_CxSunburstLayout",
    "CT_CxTitle",
    "CT_CxTreemapLayout",
    "CT_CxTx",
    "CT_CxWaterfallLayout",
]
