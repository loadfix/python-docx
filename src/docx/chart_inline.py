# pyright: reportPrivateUsage=false

"""Ergonomic chart authoring — :meth:`Document.add_chart_inline` (issue #76).

Synthesises ``c:chartSpace`` XML from dict / list-of-dicts / DataFrame
input across 13 v1 chart kinds.

Chart-kind decision tree::

    Compare values across categories     -> bar / column
    Stacked totals share a band          -> stacked-bar / stacked-column
    Trend over a continuous x-axis       -> line / area
    Trend with stacked totals            -> stacked-area
    Whole-of-100% breakdown              -> pie / donut
    Two numeric vars, no time order      -> scatter
    Three numeric vars (x, y, size)      -> bubble
    Different y-scales same chart        -> combo (with secondary_axis)
    Tiny inline trend, no axes/labels    -> sparkline

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterable, List, Mapping, Sequence, Tuple, Union

if TYPE_CHECKING:
    from docx.chart import Chart
    from docx.document import Document
    from docx.shared import Length


ChartKind = str
DataT = Union[Mapping[str, float], Sequence[Mapping[str, Any]], Any]
Size = Union["Length", Tuple[float, float], List[float]]


_VALID_KINDS = (
    "bar",
    "column",
    "line",
    "area",
    "pie",
    "donut",
    "scatter",
    "bubble",
    "combo",
    "stacked-bar",
    "stacked-column",
    "stacked-area",
    "sparkline",
    # also accept the ``grouped-`` aliases callers reach for
    "grouped-bar",
    "grouped-column",
)


def _is_dataframe(obj: Any) -> bool:
    """Return |True| when `obj` quacks like a ``pandas.DataFrame``."""
    try:
        import pandas as pd  # noqa: F401
    except ImportError:
        return False
    return type(obj).__name__ == "DataFrame" and hasattr(obj, "to_dict")


def _normalise_dataframe(
    df: Any,
    x: Union[str, None],
    y: Union[str, Sequence[str], None],
) -> Tuple[List[str], "dict[str, list[float]]"]:
    if x is None:
        raise ValueError("DataFrame input requires `x=<column-name>`")
    if x not in df.columns:
        raise ValueError(f"x column {x!r} not found in DataFrame")

    categories = [str(v) for v in df[x].tolist()]

    y_cols: List[str]
    if y is None:
        y_cols = [c for c in df.columns if c != x]
    elif isinstance(y, str):
        y_cols = [y]
    else:
        y_cols = list(y)
    for col in y_cols:
        if col not in df.columns:
            raise ValueError(f"y column {col!r} not found in DataFrame")

    series_data: "dict[str, list[float]]" = {}
    for col in y_cols:
        series_data[str(col)] = [float(v) for v in df[col].tolist()]
    return categories, series_data


def _normalise_records(
    rows: Sequence[Mapping[str, Any]],
    x: Union[str, None],
    y: Union[str, Sequence[str], None],
) -> Tuple[List[str], "dict[str, list[float]]"]:
    if not rows:
        return [], {}
    if x is None:
        raise ValueError("list-of-dicts input requires `x=<key>`")

    sample = rows[0]
    if x not in sample:
        raise ValueError(f"x key {x!r} not found in first record")

    if y is None:
        y_cols = [k for k in sample.keys() if k != x]
    elif isinstance(y, str):
        y_cols = [y]
    else:
        y_cols = list(y)

    categories = [str(row[x]) for row in rows]
    series_data: "dict[str, list[float]]" = {col: [] for col in y_cols}
    for row in rows:
        for col in y_cols:
            if col not in row:
                raise ValueError(f"row missing y key {col!r}")
            series_data[col].append(float(row[col]))
    return categories, series_data


def _normalise_data(
    data: DataT,
    x: Union[str, None],
    y: Union[str, Sequence[str], None],
) -> Tuple[List[str], "dict[str, list[float]]"]:
    """Return ``(categories, series_data)`` for any input shape."""
    if _is_dataframe(data):
        return _normalise_dataframe(data, x, y)
    if isinstance(data, Mapping):
        # single-series dict: {category: value}
        categories = [str(k) for k in data.keys()]
        values = [float(v) for v in data.values()]
        series_name = (
            y if isinstance(y, str) else (y[0] if isinstance(y, (list, tuple)) and y else "Series 1")
        )
        return categories, {str(series_name): values}
    if isinstance(data, Sequence):
        return _normalise_records(list(data), x, y)
    raise TypeError(
        f"unsupported `data` shape {type(data).__name__}; "
        "expected dict, list-of-dicts, or pandas.DataFrame"
    )


_CHART_NS_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_CHART_NS = f'xmlns:c="{_CHART_NS_URI}"'
_MAIN_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
_REL_NS = (
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _ser_xml_categorical(
    idx: int,
    name: str,
    categories: Sequence[str],
    values: Sequence[float],
    *,
    smooth: bool = False,
) -> str:
    """Return `<c:ser>` XML for axis-bound charts."""
    cat_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{_xml_escape(c)}</c:v></c:pt>'
        for i, c in enumerate(categories)
    )
    val_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values)
    )
    smooth_xml = '<c:smooth val="0"/>' if smooth else ""
    return (
        "<c:ser>"
        f'<c:idx val="{idx}"/>'
        f'<c:order val="{idx}"/>'
        f"<c:tx><c:v>{_xml_escape(name)}</c:v></c:tx>"
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
        f"{smooth_xml}"
        "</c:ser>"
    )


def _ser_xml_xy(
    idx: int,
    name: str,
    x_values: Sequence[float],
    y_values: Sequence[float],
) -> str:
    """Return `<c:ser>` XML for scatter."""
    x_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(x_values)
    )
    y_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(y_values)
    )
    return (
        "<c:ser>"
        f'<c:idx val="{idx}"/>'
        f'<c:order val="{idx}"/>'
        f"<c:tx><c:v>{_xml_escape(name)}</c:v></c:tx>"
        "<c:xVal><c:numRef><c:numCache>"
        '<c:formatCode>General</c:formatCode>'
        f'<c:ptCount val="{len(x_values)}"/>'
        f"{x_pts}"
        "</c:numCache></c:numRef></c:xVal>"
        "<c:yVal><c:numRef><c:numCache>"
        '<c:formatCode>General</c:formatCode>'
        f'<c:ptCount val="{len(y_values)}"/>'
        f"{y_pts}"
        "</c:numCache></c:numRef></c:yVal>"
        '<c:smooth val="0"/>'
        "</c:ser>"
    )


def _ser_xml_bubble(
    idx: int,
    name: str,
    x_values: Sequence[float],
    y_values: Sequence[float],
    sizes: Sequence[float],
) -> str:
    """Return `<c:ser>` XML for bubble."""
    x_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(x_values)
    )
    y_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(y_values)
    )
    sz_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(sizes)
    )
    return (
        "<c:ser>"
        f'<c:idx val="{idx}"/>'
        f'<c:order val="{idx}"/>'
        f"<c:tx><c:v>{_xml_escape(name)}</c:v></c:tx>"
        "<c:xVal><c:numRef><c:numCache>"
        '<c:formatCode>General</c:formatCode>'
        f'<c:ptCount val="{len(x_values)}"/>'
        f"{x_pts}"
        "</c:numCache></c:numRef></c:xVal>"
        "<c:yVal><c:numRef><c:numCache>"
        '<c:formatCode>General</c:formatCode>'
        f'<c:ptCount val="{len(y_values)}"/>'
        f"{y_pts}"
        "</c:numCache></c:numRef></c:yVal>"
        "<c:bubbleSize><c:numRef><c:numCache>"
        '<c:formatCode>General</c:formatCode>'
        f'<c:ptCount val="{len(sizes)}"/>'
        f"{sz_pts}"
        "</c:numCache></c:numRef></c:bubbleSize>"
        "</c:ser>"
    )


# Stable axis-IDs — any positive int works; Word treats them as opaque
# pointers between the chart-kind element's c:axId children and the
# c:catAx / c:valAx siblings under c:plotArea.
_CAT_AX_ID = 111111111
_VAL_AX_ID = 222222222
_VAL_AX_ID_2 = 333333333


def _cat_value_axes_xml(*, secondary: bool, sparkline: bool = False) -> str:
    """Return `<c:catAx>` + `<c:valAx>` siblings."""
    delete = '<c:delete val="1"/>' if sparkline else '<c:delete val="0"/>'
    cat_ax = (
        "<c:catAx>"
        f'<c:axId val="{_CAT_AX_ID}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        f"{delete}"
        '<c:axPos val="b"/>'
        f'<c:crossAx val="{_VAL_AX_ID}"/>'
        "</c:catAx>"
    )
    val_ax = (
        "<c:valAx>"
        f'<c:axId val="{_VAL_AX_ID}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        f"{delete}"
        '<c:axPos val="l"/>'
        f'<c:crossAx val="{_CAT_AX_ID}"/>'
        "</c:valAx>"
    )
    if not secondary:
        return cat_ax + val_ax
    val_ax_2 = (
        "<c:valAx>"
        f'<c:axId val="{_VAL_AX_ID_2}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="r"/>'
        f'<c:crossAx val="{_CAT_AX_ID}"/>'
        '<c:crosses val="max"/>'
        "</c:valAx>"
    )
    return cat_ax + val_ax + val_ax_2


def _xy_value_axes_xml() -> str:
    """Return `<c:valAx>` x + y for scatter / bubble."""
    return (
        "<c:valAx>"
        f'<c:axId val="{_CAT_AX_ID}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="b"/>'
        f'<c:crossAx val="{_VAL_AX_ID}"/>'
        "</c:valAx>"
        "<c:valAx>"
        f'<c:axId val="{_VAL_AX_ID}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="l"/>'
        f'<c:crossAx val="{_CAT_AX_ID}"/>'
        "</c:valAx>"
    )


def _ax_id_pair(secondary_axis: bool = False) -> str:
    """Return c:axId children for an axis-using c:<kind>Chart."""
    if secondary_axis:
        return f'<c:axId val="{_CAT_AX_ID}"/><c:axId val="{_VAL_AX_ID_2}"/>'
    return f'<c:axId val="{_CAT_AX_ID}"/><c:axId val="{_VAL_AX_ID}"/>'


def _bar_chart_xml(
    *,
    bar_dir: str,
    grouping: str,
    series_xml: str,
    secondary: bool = False,
) -> str:
    overlap = '<c:overlap val="100"/>' if grouping == "stacked" else ""
    return (
        "<c:barChart>"
        f'<c:barDir val="{bar_dir}"/>'
        f'<c:grouping val="{grouping}"/>'
        '<c:varyColors val="0"/>'
        f"{series_xml}"
        f"{overlap}"
        f"{_ax_id_pair(secondary)}"
        "</c:barChart>"
    )


def _line_chart_xml(
    *, grouping: str, series_xml: str, secondary: bool = False, smooth: bool = False
) -> str:
    smooth_attr = '<c:smooth val="1"/>' if smooth else ""
    return (
        "<c:lineChart>"
        f'<c:grouping val="{grouping}"/>'
        '<c:varyColors val="0"/>'
        f"{series_xml}"
        '<c:marker val="1"/>'
        f"{smooth_attr}"
        f"{_ax_id_pair(secondary)}"
        "</c:lineChart>"
    )


def _area_chart_xml(*, grouping: str, series_xml: str, secondary: bool = False) -> str:
    return (
        "<c:areaChart>"
        f'<c:grouping val="{grouping}"/>'
        '<c:varyColors val="0"/>'
        f"{series_xml}"
        f"{_ax_id_pair(secondary)}"
        "</c:areaChart>"
    )


def _pie_chart_xml(*, series_xml: str, hole_size: int = 0) -> str:
    if hole_size:
        return (
            "<c:doughnutChart>"
            '<c:varyColors val="1"/>'
            f"{series_xml}"
            '<c:firstSliceAng val="0"/>'
            f'<c:holeSize val="{hole_size}"/>'
            "</c:doughnutChart>"
        )
    return (
        "<c:pieChart>"
        '<c:varyColors val="1"/>'
        f"{series_xml}"
        "</c:pieChart>"
    )


def _scatter_chart_xml(*, series_xml: str) -> str:
    return (
        "<c:scatterChart>"
        '<c:scatterStyle val="lineMarker"/>'
        '<c:varyColors val="0"/>'
        f"{series_xml}"
        f'<c:axId val="{_CAT_AX_ID}"/><c:axId val="{_VAL_AX_ID}"/>'
        "</c:scatterChart>"
    )


def _bubble_chart_xml(*, series_xml: str) -> str:
    return (
        "<c:bubbleChart>"
        '<c:varyColors val="0"/>'
        f"{series_xml}"
        '<c:bubbleScale val="100"/>'
        '<c:showNegBubbles val="0"/>'
        f'<c:axId val="{_CAT_AX_ID}"/><c:axId val="{_VAL_AX_ID}"/>'
        "</c:bubbleChart>"
    )


def _split_secondary_series(
    series_data: "dict[str, list[float]]",
    secondary_axis: Union[Sequence[str], None],
) -> Tuple["dict[str, list[float]]", "dict[str, list[float]]"]:
    """Split the series dict into (primary, secondary)."""
    if not secondary_axis:
        return series_data, {}
    secondary_set = set(secondary_axis)
    primary = {k: v for k, v in series_data.items() if k not in secondary_set}
    secondary = {k: v for k, v in series_data.items() if k in secondary_set}
    return primary, secondary


def _build_chart_xml(
    kind: str,
    categories: List[str],
    series_data: "dict[str, list[float]]",
    *,
    title: Union[str, None],
    subtitle: Union[str, None],
    show_values: bool,
    show_legend: Union[bool, str],
    secondary_axis: Union[Sequence[str], None],
) -> bytes:
    """Return the bytes payload for `c:chartSpace` for the given kind."""
    kind = kind.lower()
    if kind not in _VALID_KINDS:
        raise ValueError(
            f"unsupported chart kind {kind!r}; expected one of {sorted(_VALID_KINDS)}"
        )

    # Series-count sanity check (axis-bound charts only — scatter/bubble
    # have their own value-shape requirements handled by the builders).
    primary_series, secondary_series = _split_secondary_series(
        series_data, secondary_axis
    )
    if not primary_series and not secondary_series:
        raise ValueError("at least one series is required")

    # Per-kind value-shape checks — categorical kinds require equal-length
    # value lists (caller checks too, but we re-check defensively so the
    # XML synthesis cannot emit an inconsistent c:ptCount).
    if kind not in ("scatter", "bubble"):
        for name, values in series_data.items():
            if len(values) != len(categories):
                raise ValueError(
                    f"series {name!r} has {len(values)} values but "
                    f"{len(categories)} categories"
                )

    title_xml = _title_xml(title, subtitle)
    legend_xml = _legend_xml(show_legend, series_count=len(series_data))
    sparkline = kind == "sparkline"

    # ----- bar / column variants -----
    if kind in ("bar", "column", "stacked-bar", "stacked-column", "grouped-bar", "grouped-column"):
        bar_dir = "bar" if kind in ("bar", "stacked-bar", "grouped-bar") else "col"
        grouping = "stacked" if kind.startswith("stacked-") else "clustered"
        ser_blocks = []
        idx = 0
        for name, vals in primary_series.items():
            ser_blocks.append(
                _ser_xml_categorical(idx, name, categories, vals, smooth=False)
            )
            idx += 1
        for name, vals in secondary_series.items():
            ser_blocks.append(
                _ser_xml_categorical(idx, name, categories, vals, smooth=False)
            )
            idx += 1
        kind_xml = _bar_chart_xml(
            bar_dir=bar_dir,
            grouping=grouping,
            series_xml="".join(ser_blocks),
            secondary=False,
        )
        plot_xml = _wrap_plot_area(
            [kind_xml],
            axes_xml=_cat_value_axes_xml(secondary=bool(secondary_series)),
        )
        return _wrap_chartSpace(plot_xml, title_xml, legend_xml, show_values)

    # ----- line -----
    if kind == "line":
        ser_blocks = [
            _ser_xml_categorical(i, n, categories, v)
            for i, (n, v) in enumerate(series_data.items())
        ]
        kind_xml = _line_chart_xml(
            grouping="standard", series_xml="".join(ser_blocks)
        )
        plot_xml = _wrap_plot_area([kind_xml], axes_xml=_cat_value_axes_xml(secondary=False))
        return _wrap_chartSpace(plot_xml, title_xml, legend_xml, show_values)

    # ----- area / stacked-area -----
    if kind in ("area", "stacked-area"):
        grouping = "stacked" if kind == "stacked-area" else "standard"
        ser_blocks = [
            _ser_xml_categorical(i, n, categories, v)
            for i, (n, v) in enumerate(series_data.items())
        ]
        kind_xml = _area_chart_xml(grouping=grouping, series_xml="".join(ser_blocks))
        plot_xml = _wrap_plot_area([kind_xml], axes_xml=_cat_value_axes_xml(secondary=False))
        return _wrap_chartSpace(plot_xml, title_xml, legend_xml, show_values)

    # ----- pie / donut -----
    if kind in ("pie", "donut"):
        # Only the first series is rendered; mirrors PowerPoint behaviour
        # for pie charts authored with multi-row data.
        first_name, first_vals = next(iter(series_data.items()))
        ser_xml = _ser_xml_categorical(0, first_name, categories, first_vals)
        kind_xml = _pie_chart_xml(
            series_xml=ser_xml, hole_size=50 if kind == "donut" else 0
        )
        plot_xml = _wrap_plot_area([kind_xml], axes_xml="")
        return _wrap_chartSpace(plot_xml, title_xml, legend_xml, show_values)

    # ----- scatter -----
    if kind == "scatter":
        # Categories are interpreted as x-values; coerce them to floats.
        try:
            x_values = [float(c) for c in categories]
        except ValueError as exc:
            raise ValueError(
                "scatter requires numeric x-values; got non-numeric categories"
            ) from exc
        ser_blocks = [
            _ser_xml_xy(i, n, x_values, v)
            for i, (n, v) in enumerate(series_data.items())
        ]
        kind_xml = _scatter_chart_xml(series_xml="".join(ser_blocks))
        plot_xml = _wrap_plot_area([kind_xml], axes_xml=_xy_value_axes_xml())
        return _wrap_chartSpace(plot_xml, title_xml, legend_xml, show_values)

    # ----- bubble -----
    if kind == "bubble":
        # Bubble needs paired (x, y, size) triples per series.
        # Minimum input: list-of-dicts with x, y, size keys -- but the
        # normalised series_data has only a single column per series.
        # Convention: when len(series_data) >= 2 and a column named 'size'
        # exists, treat the first series as y, second-or-named as size.
        try:
            x_values = [float(c) for c in categories]
        except ValueError as exc:
            raise ValueError(
                "bubble requires numeric x-values; got non-numeric categories"
            ) from exc
        items = list(series_data.items())
        if len(items) < 2:
            raise ValueError(
                "bubble requires at least two y/size series (e.g. y=['Value', 'Size'])"
            )
        # First series is y; named 'size' / second series is bubble-size.
        size_name = next(
            (n for n in series_data if n.lower() == "size"), items[1][0]
        )
        sizes = series_data[size_name]
        y_items = [(n, v) for n, v in items if n != size_name]
        ser_blocks = [
            _ser_xml_bubble(i, n, x_values, v, sizes)
            for i, (n, v) in enumerate(y_items)
        ]
        kind_xml = _bubble_chart_xml(series_xml="".join(ser_blocks))
        plot_xml = _wrap_plot_area([kind_xml], axes_xml=_xy_value_axes_xml())
        return _wrap_chartSpace(plot_xml, title_xml, legend_xml, show_values)

    # ----- combo (column for primary, line for secondary) -----
    if kind == "combo":
        primary_blocks = [
            _ser_xml_categorical(i, n, categories, v)
            for i, (n, v) in enumerate(primary_series.items())
        ]
        primary_xml = _bar_chart_xml(
            bar_dir="col",
            grouping="clustered",
            series_xml="".join(primary_blocks),
            secondary=False,
        )
        secondary_blocks = [
            _ser_xml_categorical(
                i + len(primary_blocks), n, categories, v
            )
            for i, (n, v) in enumerate(secondary_series.items())
        ]
        secondary_xml = (
            _line_chart_xml(
                grouping="standard",
                series_xml="".join(secondary_blocks),
                secondary=True,
            )
            if secondary_blocks
            else ""
        )
        children = [primary_xml] + ([secondary_xml] if secondary_xml else [])
        plot_xml = _wrap_plot_area(
            children, axes_xml=_cat_value_axes_xml(secondary=bool(secondary_blocks))
        )
        return _wrap_chartSpace(plot_xml, title_xml, legend_xml, show_values)

    # ----- sparkline -----
    if kind == "sparkline":
        # Tiny inline trend; line chart with axes hidden, no legend, no title.
        ser_blocks = [
            _ser_xml_categorical(i, n, categories, v)
            for i, (n, v) in enumerate(series_data.items())
        ]
        kind_xml = _line_chart_xml(
            grouping="standard", series_xml="".join(ser_blocks)
        )
        plot_xml = _wrap_plot_area(
            [kind_xml], axes_xml=_cat_value_axes_xml(secondary=False, sparkline=True)
        )
        # Force no legend / no title for sparkline ergonomics.
        return _wrap_chartSpace(plot_xml, title_xml="", legend_xml="", show_values=False)

    # Defensive — _VALID_KINDS catches this above.
    raise ValueError(f"unsupported kind: {kind!r}")  # pragma: no cover


def _wrap_plot_area(kind_children_xml: Iterable[str], *, axes_xml: str) -> str:
    return (
        "<c:plotArea>"
        "<c:layout/>"
        + "".join(kind_children_xml)
        + axes_xml
        + "</c:plotArea>"
    )


def _wrap_chartSpace(
    plot_area_xml: str,
    title_xml: str,
    legend_xml: str,
    show_values: bool,
) -> bytes:
    plot_vis = '<c:plotVisOnly val="1"/>'
    disp_blanks = '<c:dispBlanksAs val="gap"/>'
    auto_title_deleted = (
        '<c:autoTitleDeleted val="0"/>' if title_xml else '<c:autoTitleDeleted val="1"/>'
    )
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<c:chartSpace {_CHART_NS} {_MAIN_NS} {_REL_NS}>"
        "<c:chart>"
        f"{title_xml}"
        f"{auto_title_deleted}"
        f"{plot_area_xml}"
        f"{legend_xml}"
        f"{plot_vis}"
        f"{disp_blanks}"
        "</c:chart>"
        "</c:chartSpace>"
    )
    # `show_values` reserved for a future per-series c:dLbls hook; the
    # current 0.5.x ooxml-chart proxy already exposes `dLbls` if a caller
    # needs it.  We keep the parameter so the signature stays stable.
    del show_values
    return xml.encode("utf-8")


def _title_xml(title: Union[str, None], subtitle: Union[str, None]) -> str:
    if not title and not subtitle:
        return ""
    parts: List[str] = []
    if title:
        parts.append(
            "<a:p><a:pPr><a:defRPr/></a:pPr>"
            f"<a:r><a:rPr lang=\"en-US\"/><a:t>{_xml_escape(title)}</a:t></a:r>"
            "</a:p>"
        )
    if subtitle:
        parts.append(
            "<a:p><a:pPr><a:defRPr/></a:pPr>"
            f"<a:r><a:rPr lang=\"en-US\" sz=\"1200\"/><a:t>{_xml_escape(subtitle)}</a:t></a:r>"
            "</a:p>"
        )
    return (
        "<c:title>"
        "<c:tx><c:rich>"
        '<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>'
        "<a:lstStyle/>"
        + "".join(parts)
        + "</c:rich></c:tx>"
        '<c:overlay val="0"/>'
        "</c:title>"
    )


def _legend_xml(show_legend: Union[bool, str], *, series_count: int) -> str:
    if show_legend == "auto":
        # Word convention: skip the legend on single-series charts.
        return _legend_xml_inner() if series_count > 1 else ""
    if show_legend is True:
        return _legend_xml_inner()
    return ""


def _legend_xml_inner() -> str:
    return (
        "<c:legend>"
        '<c:legendPos val="r"/>'
        '<c:overlay val="0"/>'
        "</c:legend>"
    )


def add_chart_inline(
    document: "Document",
    *,
    kind: ChartKind = "column",
    data: DataT,
    x: Union[str, None] = None,
    y: Union[str, Sequence[str], None] = None,
    title: Union[str, None] = None,
    subtitle: Union[str, None] = None,
    size: Union[Size, None] = None,
    show_values: bool = False,
    show_legend: Union[bool, str] = "auto",
    secondary_axis: Union[Sequence[str], None] = None,
) -> "Chart":
    """Implementation of :meth:`Document.add_chart_inline`."""
    from docx.chart import Chart
    from docx.opc.constants import CONTENT_TYPE as CT
    from docx.opc.constants import RELATIONSHIP_TYPE as _RT
    from docx.oxml.parser import parse_xml
    from docx.oxml.shape import CT_Inline
    from docx.parts.chart import ChartPart

    categories, series_data = _normalise_data(data, x, y)

    if not categories:
        raise ValueError("`data` produced zero categories")
    # Per-kind value-shape checks happen inside ``_build_chart_xml``;
    # categorical kinds enforce equal-length lists here for the obvious
    # input-validation message.
    if kind not in ("scatter", "bubble"):
        for name, values in series_data.items():
            if len(values) != len(categories):
                raise ValueError(
                    f"series {name!r} has {len(values)} values but "
                    f"{len(categories)} categories"
                )

    blob = _build_chart_xml(
        kind,
        categories,
        series_data,
        title=title,
        subtitle=subtitle,
        show_values=show_values,
        show_legend=show_legend,
        secondary_axis=secondary_axis,
    )

    # -- create the chart part directly so we can use a custom blob --
    package = document._part.package
    assert package is not None
    partname = package.next_partname("/word/charts/chart%d.xml")
    chart_elem = parse_xml(blob)
    chart_part = ChartPart(partname, CT.DML_CHART, chart_elem, package)
    rId = document._part.relate_to(chart_part, _RT.CHART)

    # -- decide on size --
    cx, cy = _resolve_size(size)

    shape_id = document._part.next_id
    inline = CT_Inline.new_chart_inline(shape_id, rId, cx, cy)

    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    run._r.add_drawing(inline)

    return Chart(chart_part)


def _resolve_size(
    size: Union[Size, None],
) -> Tuple["Length", "Length"]:
    """Return ``(cx, cy)`` |Emu| values for the inline chart bounding box."""
    from docx.shared import Emu, Inches, Length

    if size is None:
        return Emu(int(Inches(6))), Emu(int(Inches(4)))
    if isinstance(size, Length):
        # Single dimension treated as a square — uncommon, but stable.
        return Emu(int(size)), Emu(int(size))
    if isinstance(size, (tuple, list)):
        if len(size) != 2:
            raise ValueError("size tuple must have exactly two elements (cx, cy)")
        cx_in, cy_in = size
        if isinstance(cx_in, Length):
            cx = Emu(int(cx_in))
        else:
            cx = Emu(int(Inches(float(cx_in))))
        if isinstance(cy_in, Length):
            cy = Emu(int(cy_in))
        else:
            cy = Emu(int(Inches(float(cy_in))))
        return cx, cy
    raise TypeError(f"unsupported size {type(size).__name__}")
