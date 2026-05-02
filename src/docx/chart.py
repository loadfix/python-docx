"""|Chart| and |ChartSeries| proxy objects plus the `WD_CHART_TYPE` enum.

This module provides read-side access to charts embedded in a document and
the minimal create-side support needed for building a new chart from
categories and a mapping of series names to values.
"""

from __future__ import annotations

import enum
from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.oxml.chart import CT_ChartSpace, CT_Ser
    from docx.parts.chart import ChartPart


class WD_CHART_TYPE(enum.Enum):
    """Subset of Word's `WdChartType` enumeration.

    Only the chart types supported by the create API (and a superset for
    reads) are included.

    .. versionadded:: 1.3.0.dev0
    """

    BAR = "bar"
    BAR_STACKED = "barStacked"
    COLUMN = "column"
    COLUMN_STACKED = "columnStacked"
    LINE = "line"
    PIE = "pie"
    DOUGHNUT = "doughnut"
    SCATTER = "scatter"
    AREA = "area"


def _chart_type_for(chartSpace: CT_ChartSpace) -> WD_CHART_TYPE | None:
    """Return the `WD_CHART_TYPE` corresponding to `chartSpace`, or None."""
    chart = chartSpace.chart
    if chart is None:
        return None
    plotArea = chart.plotArea
    if plotArea is None:
        return None
    kind_elm = plotArea.chart_kind_element
    if kind_elm is None:
        return None
    tag = kind_elm.tag
    if tag == qn("c:barChart"):
        # -- distinguish bar (horizontal) vs column (vertical) via c:barDir --
        from docx.oxml.chart import CT_BarChart

        bar_dir = kind_elm.bar_dir if isinstance(kind_elm, CT_BarChart) else None
        grouping = kind_elm.grouping if isinstance(kind_elm, CT_BarChart) else None
        is_stacked = grouping == "stacked"
        if bar_dir == "bar":
            return WD_CHART_TYPE.BAR_STACKED if is_stacked else WD_CHART_TYPE.BAR
        return WD_CHART_TYPE.COLUMN_STACKED if is_stacked else WD_CHART_TYPE.COLUMN
    if tag == qn("c:lineChart"):
        return WD_CHART_TYPE.LINE
    if tag == qn("c:pieChart"):
        return WD_CHART_TYPE.PIE
    if tag == qn("c:doughnutChart"):
        return WD_CHART_TYPE.DOUGHNUT
    if tag == qn("c:scatterChart"):
        return WD_CHART_TYPE.SCATTER
    if tag == qn("c:areaChart"):
        return WD_CHART_TYPE.AREA
    return None


class ChartSeries:
    """Read-only proxy for a single series (`c:ser`) within a chart.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, ser: CT_Ser):
        self._ser = ser

    @property
    def name(self) -> str:
        """Series name; empty string when not set.

        .. versionadded:: 1.3.0.dev0
        """
        value = self._ser.tx_name
        return value or ""

    @property
    def values(self) -> list[float]:
        """Series values as a list of floats (empty if none cached).

        .. versionadded:: 1.3.0.dev0
        """
        return self._ser.val_values

    @property
    def categories(self) -> list[str]:
        """Category labels associated with this series.

        .. versionadded:: 1.3.0.dev0
        """
        return self._ser.cat_values


class Chart:
    """Read-only proxy for a chart embedded in a document.

    The chart is backed by a `docx.parts.chart.ChartPart` which owns the
    `c:chartSpace` XML tree.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, chart_part: ChartPart):
        from docx.oxml.chart import CT_ChartSpace

        self._chart_part = chart_part
        self._chartSpace = cast("CT_ChartSpace", chart_part.element)

    @property
    def part(self) -> ChartPart:
        return self._chart_part

    @property
    def chart_type(self) -> WD_CHART_TYPE | None:
        """The chart's type, or |None| if not recognized.

        .. versionadded:: 1.3.0.dev0
        """
        return _chart_type_for(self._chartSpace)

    @property
    def title(self) -> str | None:
        """Chart title text, or None if no title is set.

        .. versionadded:: 1.3.0.dev0
        """
        chart = self._chartSpace.chart
        if chart is None:
            return None
        return chart.title_text

    @property
    def has_legend(self) -> bool:
        """True when the chart has a `c:legend` element.

        .. versionadded:: 1.3.0.dev0
        """
        chart = self._chartSpace.chart
        if chart is None:
            return False
        return chart.has_legend

    @property
    def series(self) -> list[ChartSeries]:
        """All `ChartSeries` for this chart, in document order.

        .. versionadded:: 1.3.0.dev0
        """
        chart = self._chartSpace.chart
        if chart is None:
            return []
        plotArea = chart.plotArea
        if plotArea is None:
            return []
        return [ChartSeries(ser) for ser in plotArea.ser_lst]

    @property
    def categories(self) -> list[str]:
        """Categories from the first series, or empty list when none.

        .. versionadded:: 1.3.0.dev0
        """
        ser_list = self.series
        if not ser_list:
            return []
        return ser_list[0].categories
