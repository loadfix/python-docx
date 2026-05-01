"""Custom element classes for DrawingML chart-related elements.

Only the minimum subset of the chartML (`c:`) vocabulary required by the
read API (chart_type, title, categories, series name/values, legend) and
the minimal create templates (bar, column, line, pie) is modeled here.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.oxml.xmlchemy import BaseOxmlElement

if TYPE_CHECKING:
    pass


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
    """`<c:chartSpace>` root element of a chart part."""

    @property
    def chart(self) -> CT_Chart | None:
        return cast("CT_Chart | None", self.find(qn("c:chart")))


class CT_Chart(BaseOxmlElement):
    """`<c:chart>` element, the chart proper inside a chartSpace."""

    @property
    def plotArea(self) -> CT_PlotArea | None:
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


class CT_PlotArea(BaseOxmlElement):
    """`<c:plotArea>` element, the container for chart-type-specific child(ren)."""

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
    def ser_lst(self) -> list[CT_Ser]:
        """All `c:ser` descendants of the plot area."""
        return cast("list[CT_Ser]", self.xpath(".//c:ser"))


class CT_BarChart(BaseOxmlElement):
    """`<c:barChart>` element.

    Can represent either a vertical (column) or horizontal (bar) chart depending
    on the value of `c:barDir/@val` ("col" → column, "bar" → bar).
    """

    @property
    def bar_dir(self) -> str | None:
        barDir = self.find(qn("c:barDir"))
        if barDir is None:
            return None
        return barDir.get("val")

    @property
    def grouping(self) -> str | None:
        grp = self.find(qn("c:grouping"))
        if grp is None:
            return None
        return grp.get("val")


class CT_LineChart(BaseOxmlElement):
    """`<c:lineChart>` element."""


class CT_PieChart(BaseOxmlElement):
    """`<c:pieChart>` element."""


class CT_DoughnutChart(BaseOxmlElement):
    """`<c:doughnutChart>` element."""


class CT_ScatterChart(BaseOxmlElement):
    """`<c:scatterChart>` element."""


class CT_AreaChart(BaseOxmlElement):
    """`<c:areaChart>` element."""


class CT_Ser(BaseOxmlElement):
    """`<c:ser>` element, a chart series."""

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
