# pyright: reportPrivateUsage=false

"""DataFrame -> styled Word table helper (issue #40).

Implements :meth:`docx.document.Document.add_dataframe`. Pandas is
**not** a hard dependency — DataFrame input is sniffed via duck-typing
(matching :mod:`docx.chart_inline`); when pandas is missing the helper
raises :class:`ImportError` with an actionable message. Four built-in
styles (``executive`` / ``minimal`` / ``boxed`` / ``striped``) plus
alternating-row tints, theme-aware header colours, per-column
alignment, per-column number-format DSL, and total-row aggregation
all sit on the existing :class:`docx.table.Table` authoring API.

The number-format DSL accepts the standard Python format-spec mini
language (``$,.1f``, ``0.0%``, ``,d`` …) for numeric columns plus a
small set of date tokens (``YYYY``, ``YY``, ``MMMM``, ``MMM``, ``MM``,
``DD``, ``HH``, ``mm``, ``ss``) for date / datetime columns,
translated to ``strftime`` directives.

Closes #40.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import datetime as _dt
from typing import TYPE_CHECKING, Any, List, Mapping, Optional, Sequence, Tuple, Union

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table, _Cell


# ---------------------------------------------------------------------------
# Public surface
# ---------------------------------------------------------------------------


_BUILTIN_STYLES = ("executive", "minimal", "boxed", "striped")


_AggregatorName = str  # "sum" | "mean" | "count" | "none"

TotalRowSpec = Union[bool, _AggregatorName, Mapping[str, _AggregatorName]]


def _is_dataframe(obj: Any) -> bool:
    """Return |True| when `obj` quacks like a ``pandas.DataFrame``.

    Mirrors the runtime sniff used in :mod:`docx.chart_inline`. Returns
    |False| when pandas is not installed so callers can raise a more
    specific error.
    """
    try:
        import pandas as pd  # noqa: F401
    except ImportError:
        return False
    return type(obj).__name__ == "DataFrame" and hasattr(obj, "to_dict")


def _require_pandas() -> Any:
    """Import and return the ``pandas`` module, or raise |ImportError|.

    Matches the contract documented on
    :meth:`Document.add_dataframe` — pandas is optional, but the
    function cannot do anything without it.
    """
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover - defensive
        raise ImportError(
            "Document.add_dataframe(...) requires the optional 'pandas' "
            "dependency. Install it with `pip install pandas`."
        ) from exc
    return pd


# ---------------------------------------------------------------------------
# Style descriptors
# ---------------------------------------------------------------------------


class _StyleSpec:
    """Inline style spec for one of the four built-in DataFrame styles."""

    __slots__ = (
        "name",
        "header_fill",
        "header_text",
        "alt_row_fill",
        "row_text",
        "border",
        "monospace_numbers",
        "header_underline_only",
        "total_row_top_border",
    )

    def __init__(
        self,
        name: str,
        *,
        header_fill: Optional[RGBColor],
        header_text: Optional[RGBColor],
        alt_row_fill: Optional[RGBColor],
        row_text: Optional[RGBColor],
        border: str,  # "all" | "none" | "horizontal" | "header_underline"
        monospace_numbers: bool = False,
        header_underline_only: bool = False,
        total_row_top_border: bool = True,
    ):
        self.name = name
        self.header_fill = header_fill
        self.header_text = header_text
        self.alt_row_fill = alt_row_fill
        self.row_text = row_text
        self.border = border
        self.monospace_numbers = monospace_numbers
        self.header_underline_only = header_underline_only
        self.total_row_top_border = total_row_top_border


def _resolve_style(name: str) -> _StyleSpec:
    if name == "executive":
        return _StyleSpec(
            name="executive",
            header_fill=RGBColor(0x1F, 0x49, 0x7D),
            header_text=RGBColor(0xFF, 0xFF, 0xFF),
            alt_row_fill=RGBColor(0xF2, 0xF2, 0xF2),
            row_text=None,
            border="horizontal",
        )
    if name == "minimal":
        return _StyleSpec(
            name="minimal",
            header_fill=None,
            header_text=None,
            alt_row_fill=None,
            row_text=None,
            border="header_underline",
            monospace_numbers=True,
            header_underline_only=True,
        )
    if name == "boxed":
        return _StyleSpec(
            name="boxed",
            header_fill=RGBColor(0xEE, 0xEE, 0xEE),
            header_text=None,
            alt_row_fill=None,
            row_text=None,
            border="all",
        )
    if name == "striped":
        return _StyleSpec(
            name="striped",
            header_fill=None,
            header_text=None,
            alt_row_fill=RGBColor(0xF2, 0xF2, 0xF2),
            row_text=None,
            border="none",
        )
    raise ValueError(
        "unknown style %r; expected one of %s" % (name, ", ".join(_BUILTIN_STYLES))
    )


# ---------------------------------------------------------------------------
# Number-format DSL
# ---------------------------------------------------------------------------

# Date tokens (longest first so substring matches don't shadow each other).
_DATE_TOKENS: Tuple[Tuple[str, str], ...] = (
    ("YYYY", "%Y"),
    ("YY", "%y"),
    ("MMMM", "%B"),
    ("MMM", "%b"),
    ("MM", "%m"),
    ("DD", "%d"),
    ("HH", "%H"),
    ("mm", "%M"),
    ("ss", "%S"),
)


def _looks_like_date_format(spec: str) -> bool:
    """Return |True| when ``spec`` carries any of the date DSL tokens."""
    for token, _ in _DATE_TOKENS:
        if token in spec:
            return True
    return False


def _date_spec_to_strftime(spec: str) -> str:
    """Translate ``spec`` from the date DSL to a ``strftime`` template.

    Tokens are replaced longest-first so e.g. ``MMMM`` does not collide
    with ``MMM``.
    """
    placeholders: List[Tuple[str, str]] = []
    out = spec
    for i, (token, repl) in enumerate(_DATE_TOKENS):
        sentinel = "\x00%d\x00" % i
        out = out.replace(token, sentinel)
        placeholders.append((sentinel, repl))
    for sentinel, repl in placeholders:
        out = out.replace(sentinel, repl)
    return out


def _format_number_spec(value: Any, spec: str) -> str:
    """Format a numeric ``value`` against a Python format-spec.

    Strips a leading ``$`` so ``$,.1f`` renders ``$1,234.5`` (Python
    rejects ``$`` inside the actual format spec). Anything we cannot
    coerce to ``float`` is rendered via ``str(value)``. Integer-only
    specs (``"d"`` / ``",d"`` / ``"+,d"``) coerce the value to int
    first so callers don't have to manually narrow their column dtype.
    """
    prefix = ""
    spec = spec.strip()
    if spec.startswith("$"):
        prefix = "$"
        spec = spec[1:]
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    # Integer-only specs need an int operand
    if spec.endswith(("d", "b", "o", "x", "X", "n")):
        try:
            return prefix + format(int(f), spec)
        except (ValueError, TypeError):
            pass
    try:
        return prefix + format(f, spec)
    except (ValueError, TypeError):
        return prefix + str(f)


def _format_value(value: Any, spec: Optional[str]) -> str:
    """Render ``value`` as a string using the ``spec`` mini-language.

    ``spec`` may be |None| (use the default rendering), a date DSL
    template, or a Python format spec. Pandas null sentinels (``NaN``,
    ``NaT``) become an empty string so the resulting cell is blank.
    """
    if value is None:
        return ""
    # -- detect pandas/numpy NaN / NaT without importing pandas eagerly --
    try:
        import math

        if isinstance(value, float) and math.isnan(value):
            return ""
    except Exception:  # pragma: no cover -- defensive
        pass
    # -- explicit handling for pandas NaT --
    type_name = type(value).__name__
    if type_name == "NaTType":
        return ""

    if spec is None:
        if isinstance(value, (_dt.datetime, _dt.date)):
            # default ISO-ish rendering for date columns
            if isinstance(value, _dt.datetime):
                return value.strftime("%Y-%m-%d %H:%M:%S")
            return value.strftime("%Y-%m-%d")
        if isinstance(value, float):
            # avoid scientific notation surprises
            if value == int(value):
                return str(int(value))
            return repr(value)
        return str(value)

    if _looks_like_date_format(spec):
        if hasattr(value, "strftime"):
            return value.strftime(_date_spec_to_strftime(spec))
        # -- coerce ISO strings into datetime so the DSL still applies --
        if isinstance(value, str):
            try:
                parsed = _dt.datetime.fromisoformat(value)
            except ValueError:
                return value
            return parsed.strftime(_date_spec_to_strftime(spec))
        return str(value)

    return _format_number_spec(value, spec)


# ---------------------------------------------------------------------------
# Aggregators (total row)
# ---------------------------------------------------------------------------


def _aggregate(values: Sequence[Any], op: str) -> Any:
    """Reduce ``values`` according to ``op`` ∈ {sum, mean, count, none}.

    Non-numeric values are skipped for ``sum`` / ``mean``. ``count`` is
    the count of non-null values. ``none`` returns the empty string —
    convenient for caller-defined no-op columns in a total row.
    """
    if op == "none":
        return ""
    if op == "count":
        return sum(1 for v in values if v is not None and not _is_null(v))
    nums: List[float] = []
    for v in values:
        if v is None or _is_null(v):
            continue
        try:
            nums.append(float(v))
        except (TypeError, ValueError):
            continue
    if not nums:
        return ""
    if op == "sum":
        return sum(nums)
    if op == "mean":
        return sum(nums) / len(nums)
    raise ValueError(
        "unknown aggregator %r; expected one of sum/mean/count/none" % op
    )


def _is_null(value: Any) -> bool:
    if value is None:
        return True
    type_name = type(value).__name__
    if type_name == "NaTType":
        return True
    try:
        import math

        return isinstance(value, float) and math.isnan(value)
    except Exception:  # pragma: no cover
        return False


def _resolve_total_spec(
    show_total_row: TotalRowSpec,
    columns: Sequence[str],
    column_dtypes: Mapping[str, Any],
) -> Optional[Mapping[str, str]]:
    """Return ``{col -> aggregator}`` or |None| when no total row is requested.

    ``show_total_row`` may be:

    - |False| — no total row
    - |True| / ``"sum"`` — sum every numeric column, blank everything else
    - ``"mean"`` / ``"count"`` / ``"none"`` — apply across all numeric cols
    - a mapping ``{col_name: op}`` — explicit per-column override
    """
    if show_total_row is False or show_total_row is None:
        return None
    pd = None
    try:
        import pandas as _pd

        pd = _pd
    except ImportError:  # pragma: no cover
        pass

    def _is_numeric(col: str) -> bool:
        if pd is None:
            return False
        dtype = column_dtypes.get(col)
        return bool(dtype is not None and pd.api.types.is_numeric_dtype(dtype))

    if isinstance(show_total_row, Mapping):
        # explicit overrides — leave non-listed cols blank
        out = {col: "none" for col in columns}
        for col, op in show_total_row.items():
            if col not in out:
                raise ValueError(
                    "show_total_row references unknown column %r" % col
                )
            out[col] = op
        return out

    op = "sum" if show_total_row is True else show_total_row
    if op not in {"sum", "mean", "count", "none"}:
        raise ValueError(
            "show_total_row must be a bool, one of "
            "'sum'/'mean'/'count'/'none', or a mapping; got %r" % (show_total_row,)
        )
    return {col: (op if _is_numeric(col) else "none") for col in columns}


# ---------------------------------------------------------------------------
# Alignment helpers
# ---------------------------------------------------------------------------


_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "centre": WD_ALIGN_PARAGRAPH.CENTER,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _resolve_alignment(token: str) -> Any:
    try:
        return _ALIGN_MAP[token.lower()]
    except KeyError as exc:
        raise ValueError(
            "unknown alignment %r; expected left/right/center/justify" % token
        ) from exc


def _default_alignment(dtype: Any) -> Any:
    """Right-align numeric cols, left-align everything else."""
    try:
        import pandas as pd
    except ImportError:
        return WD_ALIGN_PARAGRAPH.LEFT
    if pd.api.types.is_numeric_dtype(dtype):
        return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT


# ---------------------------------------------------------------------------
# Cell-level styling primitives
# ---------------------------------------------------------------------------


def _set_cell_fill(cell: "_Cell", color: RGBColor) -> None:
    cell.shading.fill_color = color


def _stamp_run(
    cell: "_Cell",
    text: str,
    *,
    bold: bool = False,
    color: Optional[RGBColor] = None,
    monospace: bool = False,
    alignment: Optional[Any] = None,
) -> None:
    """Replace ``cell``'s contents with a single styled paragraph."""
    tc = cell._tc
    tc.clear_content()
    p = tc.add_p()
    if alignment is not None:
        from docx.text.paragraph import Paragraph

        proxy = Paragraph(p, cell)
        proxy.alignment = alignment
    r = p.add_r()
    r.text = text
    # -- run-level formatting via the existing Run proxy --
    from docx.text.run import Run

    run_proxy = Run(r, cell)
    if bold:
        run_proxy.bold = True
    if color is not None:
        run_proxy.font.color.rgb = color
    if monospace:
        run_proxy.font.name = "Consolas"


def _apply_borders(table: "Table", style: _StyleSpec) -> None:
    """Stamp the four border edges + insides according to ``style.border``."""
    from docx.enum.table import WD_BORDER_STYLE

    if style.border == "all":
        table.set_borders(
            top=True,
            bottom=True,
            left=True,
            right=True,
            inside_h=True,
            inside_v=True,
            style=WD_BORDER_STYLE.SINGLE,
        )
        return
    if style.border == "horizontal":
        table.set_borders(
            top=True,
            bottom=True,
            inside_h=True,
            style=WD_BORDER_STYLE.SINGLE,
        )
        return
    if style.border == "header_underline":
        # No global edges — the header underline is applied per-cell on the
        # header row downstream.
        table.set_borders(
            inside_h=False,
            inside_v=False,
            top=False,
            bottom=False,
            left=False,
            right=False,
        )
        return
    # "none"
    table.set_borders()


def _underline_header(cell: "_Cell", color: Optional[RGBColor] = None) -> None:
    """Draw a bottom border on the header cell only (minimal-style preset)."""
    from docx.enum.table import WD_BORDER_STYLE

    cell.borders.bottom.style = WD_BORDER_STYLE.SINGLE
    cell.borders.bottom.width = Pt(0.75)
    cell.borders.bottom.color = color or RGBColor(0x40, 0x40, 0x40)


def _top_border_total_row(cell: "_Cell") -> None:
    from docx.enum.table import WD_BORDER_STYLE

    cell.borders.top.style = WD_BORDER_STYLE.SINGLE
    cell.borders.top.width = Pt(0.75)
    cell.borders.top.color = RGBColor(0x40, 0x40, 0x40)


# ---------------------------------------------------------------------------
# Theme integration
# ---------------------------------------------------------------------------


def _theme_primary(document: "Document") -> Optional[RGBColor]:
    """Return the theme's primary accent (``accent1``), or |None|."""
    theme = document.theme
    if theme is None:
        return None
    return theme.colors.accent_1


def _theme_on_primary(document: "Document") -> Optional[RGBColor]:
    """Return a high-contrast text colour for the theme's primary fill."""
    theme = document.theme
    if theme is None:
        return None
    return theme.colors.light_1 or RGBColor(0xFF, 0xFF, 0xFF)


def _coerce_color(value: Any) -> Optional[RGBColor]:
    if value is None:
        return None
    if isinstance(value, RGBColor):
        return value
    if isinstance(value, str):
        return RGBColor.from_string(value.lstrip("#"))
    raise TypeError(
        "expected RGBColor, hex string, or None; got %r" % type(value).__name__
    )


# ---------------------------------------------------------------------------
# Top-level driver
# ---------------------------------------------------------------------------


def add_dataframe(
    document: "Document",
    df: Any,
    *,
    style: str = "executive",
    alternating_rows: Optional[bool] = None,
    header_color: Any = None,
    header_text_color: Any = None,
    autofit: bool = True,
    align: Optional[Mapping[str, str]] = None,
    number_format: Optional[Mapping[str, str]] = None,
    show_total_row: TotalRowSpec = False,
    table_style: Optional[str] = None,
) -> "Table":
    """Append a DataFrame to ``document`` as a styled Word table.

    See :meth:`docx.document.Document.add_dataframe` for the public
    contract; this helper is intentionally a free function so unit
    tests can exercise it without going through the proxy layer.
    """
    if not _is_dataframe(df):
        # Distinguish "pandas missing" from "wrong type"
        try:
            _require_pandas()
        except ImportError:
            raise
        raise TypeError(
            "add_dataframe(df) requires a pandas.DataFrame; got %r"
            % type(df).__name__
        )

    if style not in _BUILTIN_STYLES:
        raise ValueError(
            "unknown style %r; expected one of %s"
            % (style, ", ".join(_BUILTIN_STYLES))
        )

    spec = _resolve_style(style)

    # -- explicit caller overrides win over preset defaults --
    header_fill = _coerce_color(header_color) if header_color is not None else spec.header_fill
    header_txt = (
        _coerce_color(header_text_color)
        if header_text_color is not None
        else spec.header_text
    )

    columns = [str(c) for c in df.columns]
    column_dtypes = {str(c): df[c].dtype for c in df.columns}

    rows = df.shape[0]
    cols = len(columns)
    if cols == 0:
        raise ValueError("DataFrame must have at least one column")

    n_data_rows = rows
    total_spec = _resolve_total_spec(show_total_row, columns, column_dtypes)
    n_total_rows = 1 if total_spec is not None else 0
    total_rows = 1 + n_data_rows + n_total_rows  # header + data + total

    table = document.add_table(rows=total_rows, cols=cols, style=table_style)
    if autofit:
        table.autofit = True
    else:
        table.autofit = False

    _apply_borders(table, spec)

    # -- header row ---------------------------------------------------
    header_row = table.rows[0]
    header_row.is_header = True
    for col_idx, col_name in enumerate(columns):
        cell = header_row.cells[col_idx]
        if header_fill is not None:
            _set_cell_fill(cell, header_fill)
        align_token = align.get(col_name) if align else None
        alignment = (
            _resolve_alignment(align_token)
            if align_token is not None
            else _default_alignment(column_dtypes[col_name])
        )
        _stamp_run(
            cell,
            col_name,
            bold=True,
            color=header_txt,
            alignment=alignment,
        )
        if spec.header_underline_only:
            _underline_header(cell)

    # -- data rows -----------------------------------------------------
    # Default-on when the preset declares an alt-row tint; explicit
    # ``alternating_rows=True`` forces a tint even on presets that don't
    # ship one (we fall back to a neutral light grey).
    if alternating_rows is None:
        use_alt = spec.alt_row_fill is not None
    else:
        use_alt = bool(alternating_rows)
    if use_alt:
        alt_fill = spec.alt_row_fill or RGBColor(0xF2, 0xF2, 0xF2)
    else:
        alt_fill = None

    # Collect raw column values once (avoids repeated DataFrame indexing)
    raw_values: dict[str, list[Any]] = {
        col: df[col].tolist() for col in columns
    }

    for r_idx in range(n_data_rows):
        row = table.rows[1 + r_idx]
        row_is_alt = (r_idx % 2) == 1
        for col_idx, col_name in enumerate(columns):
            cell = row.cells[col_idx]
            value = raw_values[col_name][r_idx]
            spec_str = number_format.get(col_name) if number_format else None
            text = _format_value(value, spec_str)
            align_token = align.get(col_name) if align else None
            alignment = (
                _resolve_alignment(align_token)
                if align_token is not None
                else _default_alignment(column_dtypes[col_name])
            )
            if alt_fill is not None and row_is_alt:
                _set_cell_fill(cell, alt_fill)
            _stamp_run(
                cell,
                text,
                color=spec.row_text,
                monospace=spec.monospace_numbers,
                alignment=alignment,
            )

    # -- total row -----------------------------------------------------
    if total_spec is not None:
        total_row_idx = 1 + n_data_rows
        total_row = table.rows[total_row_idx]
        for col_idx, col_name in enumerate(columns):
            cell = total_row.cells[col_idx]
            op = total_spec[col_name]
            agg_value = _aggregate(raw_values[col_name], op)
            spec_str = number_format.get(col_name) if number_format else None
            if op in {"sum", "mean"} and agg_value != "":
                text = _format_value(agg_value, spec_str)
            elif op == "count" and agg_value != "":
                text = str(agg_value)
            else:
                text = ""
            align_token = align.get(col_name) if align else None
            alignment = (
                _resolve_alignment(align_token)
                if align_token is not None
                else _default_alignment(column_dtypes[col_name])
            )
            _stamp_run(
                cell,
                text,
                bold=True,
                color=spec.row_text,
                monospace=spec.monospace_numbers,
                alignment=alignment,
            )
            if spec.total_row_top_border:
                _top_border_total_row(cell)

    return table


__all__ = [
    "add_dataframe",
    "_BUILTIN_STYLES",
    "_aggregate",
    "_date_spec_to_strftime",
    "_format_value",
    "_is_dataframe",
    "_looks_like_date_format",
    "_resolve_alignment",
    "_resolve_style",
    "_resolve_total_spec",
]
