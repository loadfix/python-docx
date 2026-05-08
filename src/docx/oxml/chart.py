"""Custom element classes for DrawingML chart-related elements.

Re-exports the CT_* chart element classes from the shared
:mod:`ooxml_chart.oxml` package and adds docx-local thin subclasses
where docx's historically narrow read API had different (None-returning)
semantics or extra convenience properties than the shared pptx-anchored
superset.

.. versionchanged:: 2026.05.0
   Superseded by re-exports from :mod:`ooxml_chart.oxml`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

# ---------------------------------------------------------------------------
# Namespace-registry safety: importing ``ooxml_chart.oxml`` appends the
# shared chart registry to the process-global ``ooxml_xmlchemy`` composite
# stack. The composite resolves lookups in reverse registration order
# (most-recent first), so docx's registry would be shadowed — e.g.
# ``OxmlElement("a:effectLst")`` would fall through to the shared chart
# parser (which has an ``a:`` prefix but no ``effectLst`` class registered)
# and return a generic ``lxml.etree._Element`` instead of
# :class:`docx.oxml.shape.CT_EffectList`. Restore docx's registry at
# module-import-completion below (see the same pattern in
# ``pptx.oxml.extprops``).
# ---------------------------------------------------------------------------
from ooxml_chart.oxml.chart import CT_PlotArea as _SharedCT_PlotArea
from ooxml_chart.oxml.plot import (
    CT_AreaChart,
    CT_BarChart as _SharedCT_BarChart,
    CT_DoughnutChart,
    CT_LineChart,
    CT_PieChart,
    CT_ScatterChart,
)
from ooxml_chart.oxml.series import CT_SeriesComposite
from ooxml_xmlchemy import configure_namespace_registry as _configure

from docx.oxml.ns import qn
from docx.oxml.parser import _DocxNamespaceRegistry as _DocxRegistry
from docx.oxml.xmlchemy import BaseOxmlElement

if TYPE_CHECKING:
    pass


__all__ = [
    "CT_AreaChart",
    "CT_BarChart",
    "CT_Chart",
    "CT_ChartSpace",
    "CT_DoughnutChart",
    "CT_LineChart",
    "CT_PieChart",
    "CT_PlotArea",
    "CT_ScatterChart",
    "CT_Ser",
]


def _numeric_values_from_ref_or_lit(parent: BaseOxmlElement) -> list[float]:
    """Return list[float] from a `c:numRef/c:numCache` or `c:numLit` child.

    Returns an empty list when neither is present or when no usable data
    points can be parsed.
    """
    # -- prefer cached values in c:numRef/c:numCache --
    cache = parent.find(qn("c:numRef") + "/" + qn("c:numCache"))
    if cache is None:
        cache = parent.find(qn("c:numLit"))
    if cache is None:
        return []
    values: list[float] = []
    for pt in cache.findall(qn("c:pt")):
        v = pt.find(qn("c:v"))
        if v is None or v.text is None:
            continue
        try:
            values.append(float(v.text))
        except (TypeError, ValueError):
            continue
    return values


def _string_values_from_ref_or_lit(parent: BaseOxmlElement) -> list[str]:
    """Return list[str] from a `c:strRef/c:strCache` or `c:strLit` child."""
    cache = parent.find(qn("c:strRef") + "/" + qn("c:strCache"))
    if cache is None:
        cache = parent.find(qn("c:strLit"))
    if cache is None:
        return []
    values: list[str] = []
    for pt in cache.findall(qn("c:pt")):
        v = pt.find(qn("c:v"))
        if v is None:
            continue
        values.append(v.text or "")
    return values


class CT_ChartSpace(BaseOxmlElement):
    """`<c:chartSpace>` root element of a chart part.

    docx-local subclass — the shared ``CT_ChartSpace`` declares ``c:chart``
    as ``OneAndOnlyOne`` (raises on missing), whereas docx's read API
    historically returned |None| for a chartSpace that lacks a chart
    grandchild. This subclass restores the |None|-returning behaviour.
    """

    @property
    def chart(self) -> "CT_Chart | None":
        return cast("CT_Chart | None", self.find(qn("c:chart")))


class CT_Chart(BaseOxmlElement):
    """`<c:chart>` element, the chart proper inside a chartSpace.

    docx-local subclass — adds ``title_text`` and ``has_legend`` helpers
    and overrides ``plotArea`` to return |None| (rather than raising) when
    the ``c:plotArea`` grandchild is absent.
    """

    @property
    def plotArea(self) -> "CT_PlotArea | None":
        return cast("CT_PlotArea | None", self.find(qn("c:plotArea")))

    @property
    def title_text(self) -> str | None:
        """Concatenated text from `c:title/c:tx/c:rich//a:t`, or None if absent."""
        if self.find(qn("c:title")) is None:
            return None
        # -- use the BaseOxmlElement's xpath (which supplies namespaces) from
        #    the root; avoids issues when the intermediate element isn't a
        #    registered CT_ subclass that would carry the nsmap. --
        texts = self.xpath("./c:title//a:t")
        if not texts:
            return None
        return "".join(t.text or "" for t in texts)

    @property
    def has_legend(self) -> bool:
        return self.find(qn("c:legend")) is not None


class CT_PlotArea(_SharedCT_PlotArea):
    """`<c:plotArea>` element, the container for chart-type-specific child(ren).

    docx-local subclass — re-adds the ``chart_kind_element`` and
    ``ser_lst`` convenience accessors the docx chart proxy relies on.
    """

    @property
    def chart_kind_element(self) -> BaseOxmlElement | None:
        """Return the first chart-kind child element (barChart, lineChart, etc.)."""
        kinds = {
            qn("c:barChart"),
            qn("c:lineChart"),
            qn("c:pieChart"),
            qn("c:doughnutChart"),
            qn("c:scatterChart"),
            qn("c:areaChart"),
        }
        for child in self:
            if child.tag in kinds:
                return cast(BaseOxmlElement, child)
        return None

    @property
    def ser_lst(self) -> "list[CT_Ser]":
        """All `c:ser` descendants of the plot area."""
        return cast("list[CT_Ser]", self.xpath(".//c:ser"))


class CT_BarChart(_SharedCT_BarChart):
    """`<c:barChart>` element.

    docx-local subclass — exposes ``bar_dir`` and ``grouping`` as plain
    strings (or |None|) rather than the shared superset's child-element
    objects, matching docx's historical read API.

    Can represent either a vertical (column) or horizontal (bar) chart
    depending on the value of `c:barDir/@val` ("col" -> column, "bar" -> bar).
    """

    @property
    def bar_dir(self) -> str | None:  # pyright: ignore[reportIncompatibleVariableOverride]
        barDir = self.find(qn("c:barDir"))
        if barDir is None:
            return None
        return barDir.get("val")

    @property
    def grouping(self) -> str | None:  # pyright: ignore[reportIncompatibleVariableOverride]
        grp = self.find(qn("c:grouping"))
        if grp is None:
            return None
        return grp.get("val")


class CT_Ser(CT_SeriesComposite):
    """`<c:ser>` element, a chart series.

    docx-local subclass — the shared ``CT_SeriesComposite`` supplies the
    authoring surface; docx adds the read-side helpers its ``ChartSeries``
    proxy relies on.
    """

    @property
    def tx_name(self) -> str | None:
        """Series name from `c:tx/c:strRef/c:strCache/c:pt/c:v` or `c:tx/c:v`."""
        tx = self.find(qn("c:tx"))
        if tx is None:
            return None
        strCache = tx.find(qn("c:strRef") + "/" + qn("c:strCache"))
        if strCache is not None:
            pt = strCache.find(qn("c:pt"))
            if pt is not None:
                v = pt.find(qn("c:v"))
                if v is not None:
                    return v.text or ""
        v = tx.find(qn("c:v"))
        if v is not None:
            return v.text or ""
        return None

    @property
    def cat_values(self) -> list[str]:
        cat = self.find(qn("c:cat"))
        if cat is None:
            return []
        # -- category axis may hold numeric or string data; stringify either --
        strs = _string_values_from_ref_or_lit(cat)
        if strs:
            return strs
        # -- fall back to numeric cache rendered as strings --
        nums = _numeric_values_from_ref_or_lit(cat)
        return [str(int(n)) if n.is_integer() else str(n) for n in nums]

    @property
    def val_values(self) -> list[float]:
        val = self.find(qn("c:val"))
        if val is None:
            return []
        return _numeric_values_from_ref_or_lit(val)


# -- restore docx's namespace registry as the most-recently-registered entry
# -- so docx's ``a:`` / ``w:`` / ``r:`` lookups take precedence over the
# -- shared chart package's ``a:``-prefix entries. --
_configure(_DocxRegistry())
