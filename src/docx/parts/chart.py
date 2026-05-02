"""|ChartPart| — container for a DrawingML chart payload.

A chart is stored as a separate XML part (typically ``word/charts/chartN.xml``)
with content type ``application/vnd.openxmlformats-officedocument.drawingml.chart+xml``.
It is referenced from a ``c:chart`` element inside a ``w:drawing/a:graphic/a:graphicData``
container. python-docx exposes charts read/minimal-create only.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.chart import CT_ChartSpace
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.chart import WD_CHART_TYPE
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.package import Package


_CHART_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
_MAIN_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
_REL_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def _escape_xml(text: str) -> str:
    """Return `text` with XML special characters escaped."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _ser_xml(idx: int, name: str, categories: list[str], values: list[float]) -> str:
    """Return the XML fragment for a single `c:ser`."""
    cat_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{_escape_xml(c)}</c:v></c:pt>'
        for i, c in enumerate(categories)
    )
    val_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values)
    )
    return (
        "<c:ser>"
        f'<c:idx val="{idx}"/>'
        f'<c:order val="{idx}"/>'
        f"<c:tx><c:v>{_escape_xml(name)}</c:v></c:tx>"
        "<c:cat><c:strRef>"
        "<c:strCache>"
        f'<c:ptCount val="{len(categories)}"/>'
        f"{cat_pts}"
        "</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef>"
        "<c:numCache>"
        '<c:formatCode>General</c:formatCode>'
        f'<c:ptCount val="{len(values)}"/>'
        f"{val_pts}"
        "</c:numCache></c:numRef></c:val>"
        "</c:ser>"
    )


_CHART_NS_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _rewrite_ser(  # pyright: ignore[reportUnusedFunction]
    ser: BaseOxmlElement,
    idx: int,
    name: str,
    categories: list[str],
    values: list[float],
) -> None:
    """Rewrite the ``c:idx``, ``c:order``, ``c:tx``, ``c:cat`` and ``c:val``
    children of `ser` in place.

    Other children (``c:spPr``, ``c:marker``, ``c:dLbls``, ``c:smooth`` etc.)
    are left untouched so chart styling is preserved. Only the data payload
    and the series label change.
    """
    from docx.oxml.ns import qn as _qn
    from docx.oxml.parser import parse_xml as _parse_xml

    # -- drop children we are about to rewrite --
    for tag_local in ("idx", "order", "tx", "cat", "val"):
        for child in list(ser.findall(_qn("c:%s" % tag_local))):
            ser.remove(child)

    cat_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{_escape_xml(c)}</c:v></c:pt>'
        for i, c in enumerate(categories)
    )
    val_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values)
    )
    wrapper_xml = (
        f'<c:root xmlns:c="{_CHART_NS_URI}">'
        f'<c:idx val="{idx}"/>'
        f'<c:order val="{idx}"/>'
        f"<c:tx><c:v>{_escape_xml(name)}</c:v></c:tx>"
        "<c:cat><c:strRef>"
        "<c:strCache>"
        f'<c:ptCount val="{len(categories)}"/>'
        f"{cat_pts}"
        "</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef>"
        "<c:numCache>"
        '<c:formatCode>General</c:formatCode>'
        f'<c:ptCount val="{len(values)}"/>'
        f"{val_pts}"
        "</c:numCache></c:numRef></c:val>"
        "</c:root>"
    )
    wrapper = _parse_xml(wrapper_xml.encode("utf-8"))
    # -- insert the new children at the top of ser (so schema order is
    # -- preserved: idx, order, tx then the rest). --
    for i, new_child in enumerate(list(wrapper)):
        ser.insert(i, new_child)


def _chart_kind_xml(
    chart_type: WD_CHART_TYPE,
    categories: list[str],
    series_data: dict[str, list[float]],
) -> str:
    """Return the XML fragment containing the chart-kind element (`c:barChart`, etc.)."""
    from docx.chart import WD_CHART_TYPE as WCT

    series_xml = "".join(
        _ser_xml(i, name, categories, values)
        for i, (name, values) in enumerate(series_data.items())
    )

    if chart_type in (WCT.BAR, WCT.BAR_STACKED):
        bar_dir = "bar"
        grouping = "stacked" if chart_type is WCT.BAR_STACKED else "clustered"
        overlap = '<c:overlap val="100"/>' if grouping == "stacked" else ""
        return (
            "<c:barChart>"
            f'<c:barDir val="{bar_dir}"/>'
            f'<c:grouping val="{grouping}"/>'
            '<c:varyColors val="0"/>'
            f"{series_xml}"
            f"{overlap}"
            "</c:barChart>"
        )
    if chart_type in (WCT.COLUMN, WCT.COLUMN_STACKED):
        bar_dir = "col"
        grouping = "stacked" if chart_type is WCT.COLUMN_STACKED else "clustered"
        overlap = '<c:overlap val="100"/>' if grouping == "stacked" else ""
        return (
            "<c:barChart>"
            f'<c:barDir val="{bar_dir}"/>'
            f'<c:grouping val="{grouping}"/>'
            '<c:varyColors val="0"/>'
            f"{series_xml}"
            f"{overlap}"
            "</c:barChart>"
        )
    if chart_type is WCT.LINE:
        return (
            "<c:lineChart>"
            '<c:grouping val="standard"/>'
            '<c:varyColors val="0"/>'
            f"{series_xml}"
            "<c:marker val=\"1\"/>"
            "</c:lineChart>"
        )
    if chart_type is WCT.PIE:
        return (
            "<c:pieChart>"
            '<c:varyColors val="1"/>'
            f"{series_xml}"
            "</c:pieChart>"
        )
    raise ValueError(f"unsupported chart_type for creation: {chart_type!r}")


def _chartSpace_xml(
    chart_type: WD_CHART_TYPE,
    categories: list[str],
    series_data: dict[str, list[float]],
) -> bytes:
    """Return the full XML bytes for a minimal `c:chartSpace` of the requested type."""
    kind_xml = _chart_kind_xml(chart_type, categories, series_data)
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<c:chartSpace {_CHART_NS} {_MAIN_NS} {_REL_NS}>"
        "<c:chart>"
        "<c:plotArea>"
        "<c:layout/>"
        f"{kind_xml}"
        "</c:plotArea>"
        "<c:plotVisOnly val=\"1\"/>"
        "<c:dispBlanksAs val=\"gap\"/>"
        "</c:chart>"
        "</c:chartSpace>"
    )
    return xml.encode("utf-8")


class ChartPart(XmlPart):
    """A DrawingML chart part.

    Corresponds to the target part of a relationship whose ``content-type`` is
    ``application/vnd.openxmlformats-officedocument.drawingml.chart+xml``. The
    contents are a ``<c:chartSpace>`` XML tree.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_ChartSpace,
        package: Package,
    ):
        super().__init__(partname, content_type, element, package)
        self._chartSpace = element

    @property
    def chartSpace(self) -> CT_ChartSpace:
        """The `<c:chartSpace>` root element of this part."""
        return self._chartSpace

    @classmethod
    def new(
        cls,
        package: Package,
        chart_type: WD_CHART_TYPE,
        categories: list[str],
        series_data: dict[str, list[float]],
    ) -> Self:
        """Return a newly created chart part populated with the supplied data.

        `chart_type` selects which chart kind is authored (BAR, COLUMN, LINE, or
        PIE — see :class:`docx.chart.WD_CHART_TYPE`). `categories` is a list of
        category labels (used as x-axis labels / pie slice labels). `series_data`
        is a dict mapping each series name to its list of values; all value
        lists must be the same length as `categories`.
        """
        for name, values in series_data.items():
            if len(values) != len(categories):
                raise ValueError(
                    f"series {name!r} has {len(values)} values but "
                    f"{len(categories)} categories were given"
                )

        partname = package.next_partname("/word/charts/chart%d.xml")
        content_type = CT.DML_CHART
        blob = _chartSpace_xml(chart_type, categories, series_data)
        element = cast(CT_ChartSpace, parse_xml(blob))
        return cls(partname, content_type, element, package)
