"""Styled-table authoring helpers — quick, opinionated tables in one call.

Closes #289.

This module exposes two helpers that build a styled |Table| at the end
of a |Document| in one call::

    from docx import Document
    from docx.kit import tables

    doc = Document()

    tables.styled_table(
        doc,
        headers=["Name", "Value"],
        rows=[["Alpha", 1], ["Beta", 2]],
        style="modern",
    )

    # pandas optional — degrades cleanly when not installed.
    import pandas as pd
    df = pd.DataFrame({"Name": ["X", "Y"], "Value": [10, 20]})
    tables.from_dataframe(doc, df, style="zebra", auto_format=True)

    doc.save("out.docx")

Both helpers compose only python-docx's public authoring API — no
``_element`` / ``oxml`` reach-down. They lean on
:meth:`Document.add_table`, :attr:`_Cell.shading`,
:meth:`_Row.apply_shading`, :meth:`Table.autofit`, and the |Run| /
|Paragraph| formatting properties to render a ready-to-ship table with
a styled header row and optional banded body rows.

Built-in styles
---------------

Four named styles are bundled. Each maps to a header fill colour, an
optional alternating-row "zebra" tint, and a header text colour. The
specs are deliberately conservative — colours sit well alongside Word's
default ``Calibri`` body font and survive printing in monochrome.

* ``"modern"`` — deep blue header, white text, no banding (default).
* ``"zebra"`` — medium-grey header, white text, alternating light-grey
  rows.
* ``"minimal"`` — no shading, single underline below the header,
  monospaced numeric columns where the renderer detects them.
* ``"corporate"`` — dark navy header with white text, light-blue zebra
  banding.

Callers override any single colour with the ``header_fill`` /
``alt_row_fill`` keyword arguments — passing |None| disables the band /
header fill. The header text colour follows the header fill's WCAG
contrast (white text on dark fills, automatic on light fills) without
the caller having to think about it.

Auto-formatting (``from_dataframe`` only)
-----------------------------------------

When ``auto_format=True`` (the default) :func:`from_dataframe`
inspects each column's dtype and applies a sensible alignment +
rendering rule:

* Integer / float columns right-align, render with ``"%g"`` to drop
  trailing zeros, and respect ``NaN`` (rendered as the empty string).
* Date / datetime columns render as ISO-8601 (``YYYY-MM-DD`` or
  ``YYYY-MM-DD HH:MM:SS``).
* Boolean columns render as ``"Yes"`` / ``"No"``.
* Everything else left-aligns and falls through to ``str(value)``.

Pandas is **optional**. When pandas is not installed the module imports
cleanly; only :func:`from_dataframe` raises (with an actionable
:class:`ImportError`) on call.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import datetime as _dt
from typing import (
    TYPE_CHECKING,
    Any,
    List,
    Optional,
    Sequence,
    Tuple,
    Union,
)

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls
    from docx.table import Table, _Cell, _Row
    from docx.text.paragraph import Paragraph


# -- Style descriptors ----------------------------------------------------


class _StyleSpec:
    """Inline style spec for one of the four built-in kit table styles.

    Captures four colour slots plus a header-text colour and a flag
    indicating whether to draw a single underline below the header row
    (the ``minimal`` style's signature).
    """

    __slots__ = (
        "name",
        "header_fill",
        "header_text",
        "alt_row_fill",
        "header_underline_only",
        "table_style",
    )

    def __init__(
        self,
        name: str,
        *,
        header_fill: Optional[RGBColor],
        header_text: Optional[RGBColor],
        alt_row_fill: Optional[RGBColor],
        header_underline_only: bool = False,
        table_style: Optional[str] = "Table Grid",
    ):
        self.name = name
        self.header_fill = header_fill
        self.header_text = header_text
        self.alt_row_fill = alt_row_fill
        self.header_underline_only = header_underline_only
        self.table_style = table_style


#: Built-in style names accepted by :func:`styled_table` and
#: :func:`from_dataframe`. Listed in the public docstring above.
BUILTIN_STYLES: Tuple[str, ...] = ("modern", "zebra", "minimal", "corporate")


def _resolve_style(name: str) -> _StyleSpec:
    """Return the style spec named ``name`` or raise :class:`ValueError`."""
    if name == "modern":
        return _StyleSpec(
            name="modern",
            header_fill=RGBColor(0x1F, 0x49, 0x7D),  # deep blue
            header_text=RGBColor(0xFF, 0xFF, 0xFF),
            alt_row_fill=None,
        )
    if name == "zebra":
        return _StyleSpec(
            name="zebra",
            header_fill=RGBColor(0x59, 0x59, 0x59),  # medium grey
            header_text=RGBColor(0xFF, 0xFF, 0xFF),
            alt_row_fill=RGBColor(0xF2, 0xF2, 0xF2),  # light grey
        )
    if name == "minimal":
        return _StyleSpec(
            name="minimal",
            header_fill=None,
            header_text=None,
            alt_row_fill=None,
            header_underline_only=True,
            table_style=None,  # no full grid
        )
    if name == "corporate":
        return _StyleSpec(
            name="corporate",
            header_fill=RGBColor(0x0B, 0x2D, 0x5C),  # dark navy
            header_text=RGBColor(0xFF, 0xFF, 0xFF),
            alt_row_fill=RGBColor(0xDC, 0xE6, 0xF2),  # light blue
        )
    raise ValueError(
        "unknown style %r; expected one of %s"
        % (name, ", ".join(BUILTIN_STYLES))
    )


# -- Colour coercion ------------------------------------------------------


def _coerce_color(
    value: Union[RGBColor, str, None],
) -> Optional[RGBColor]:
    """Return ``value`` as an :class:`RGBColor`, or |None|.

    Accepts an existing |RGBColor|, a 6-character hex string
    (``"FF0000"``), or |None|. Raises :class:`ValueError` on malformed
    input so caller mistakes surface immediately.
    """
    if value is None:
        return None
    if isinstance(value, RGBColor):
        return value
    if isinstance(value, str):
        return RGBColor.from_string(value)
    raise ValueError(
        "expected RGBColor, hex string, or None; got %r" % (value,)
    )


# -- Cell helpers ---------------------------------------------------------


def _set_cell_text(
    cell: "_Cell",
    text: str,
    *,
    bold: bool = False,
    text_color: Optional[RGBColor] = None,
    alignment: Optional[int] = None,
) -> None:
    """Write ``text`` into ``cell`` with optional formatting.

    Replaces the cell's content with a single paragraph containing one
    run, then applies ``bold`` / ``text_color`` to that run and
    ``alignment`` to the paragraph. Stays on the public python-docx
    API (no ``_element`` reach-down).
    """
    cell.text = text
    para = cell.paragraphs[0]
    if alignment is not None:
        para.alignment = alignment
    for run in para.runs:
        if bold:
            run.bold = True
        if text_color is not None:
            run.font.color.rgb = text_color


def _apply_table_style(table: "Table", style_name: Optional[str]) -> None:
    """Apply ``style_name`` to ``table``; silently fall back when missing."""
    if style_name is None:
        return
    try:
        table.style = style_name
    except KeyError:
        pass


def _apply_row_shading(row: "_Row", color: Optional[RGBColor]) -> None:
    """Shade every cell in ``row`` using the row-level helper."""
    if color is None:
        return
    row.apply_shading(color)


def _apply_header_underline(row: "_Row") -> None:
    """Bold the header row's runs and underline them.

    Used by the ``minimal`` style which forgoes a fill colour and
    instead leans on a typographic underline to delineate the header.
    """
    for cell in row.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
                run.underline = True


# -- Type-driven alignment + rendering -----------------------------------


def _is_numeric(value: Any) -> bool:
    """Return |True| when ``value`` is a real number (not bool)."""
    if isinstance(value, bool):
        return False
    return isinstance(value, (int, float))


def _is_date_like(value: Any) -> bool:
    """Return |True| when ``value`` is a :mod:`datetime` instance."""
    return isinstance(value, (_dt.datetime, _dt.date))


def _format_value(value: Any) -> str:
    """Render ``value`` as a string using sensible default formatters.

    Mirrors :mod:`docx.dataframe`'s default rendering but kept local
    to keep the kit module dependency-free of that internal surface:

    * |None| / pandas ``NaN`` / pandas ``NaT`` -> ``""``.
    * :class:`bool` -> ``"Yes"`` / ``"No"``.
    * :class:`int` -> decimal form.
    * :class:`float` -> ``%g`` (drops trailing zeros, no scientific).
    * :class:`datetime.datetime` -> ``YYYY-MM-DD HH:MM:SS``.
    * :class:`datetime.date` -> ``YYYY-MM-DD``.
    * Anything else -> ``str(value)``.
    """
    if value is None:
        return ""
    # -- pandas / numpy null sentinels --
    try:
        import math

        if isinstance(value, float) and math.isnan(value):
            return ""
    except Exception:  # pragma: no cover - defensive
        pass
    if type(value).__name__ == "NaTType":
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return f"{value:g}"
    if isinstance(value, _dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, _dt.date):
        return value.strftime("%Y-%m-%d")
    return str(value)


def _column_alignment(
    column_values: Sequence[Any],
) -> int:
    """Return the right alignment for a body column based on its values.

    Numbers and dates right-align (currency / dates read cleaner that
    way); everything else left-aligns. Empty / |None| values are
    ignored when sniffing — they don't tip a column's alignment.
    """
    has_value = False
    for v in column_values:
        if v is None or v == "":
            continue
        has_value = True
        if not (_is_numeric(v) or _is_date_like(v)):
            return WD_ALIGN_PARAGRAPH.LEFT
    return WD_ALIGN_PARAGRAPH.RIGHT if has_value else WD_ALIGN_PARAGRAPH.LEFT


# -- "use the style default" sentinel ------------------------------------


class _UseStyleDefault:
    """Marker used as the default for ``header_fill`` / ``alt_row_fill``.

    Distinguishes "caller didn't pass the keyword" (use the style's
    own value) from "caller explicitly passed None" (suppress the
    fill). The object is private; callers compare by identity via the
    public default values exposed below.
    """

    __slots__ = ()

    def __repr__(self) -> str:  # pragma: no cover - cosmetic
        return "<use-style-default>"


_USE_STYLE_DEFAULT = _UseStyleDefault()


# -- pandas import shim ---------------------------------------------------


def _require_pandas() -> Any:
    """Import and return the ``pandas`` module, or raise :class:`ImportError`.

    Matches the contract documented on :func:`from_dataframe` — pandas
    is an optional dependency, but the function cannot do anything
    useful without it.
    """
    try:
        import pandas as pd
    except ImportError as exc:
        raise ImportError(
            "docx.kit.tables.from_dataframe(...) requires the optional "
            "'pandas' dependency. Install it with `pip install pandas`."
        ) from exc
    return pd


# -- Public API: styled_table --------------------------------------------


def styled_table(
    document: "DocumentCls",
    *,
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
    style: str = "modern",
    header_fill: Any = _USE_STYLE_DEFAULT,
    alt_row_fill: Any = _USE_STYLE_DEFAULT,
    autofit: bool = True,
    column_alignments: Optional[Sequence[int]] = None,
) -> "Table":
    """Append a styled table to ``document`` and return it.

    Parameters
    ----------
    document
        Target |Document|. The table is appended at the end of the body.
    headers
        Column headers — one string per column. The header row is
        rendered with the resolved style's fill colour and white / dark
        text per the style spec.
    rows
        Sequence of body rows. Each row must be a sequence of the same
        length as ``headers``; mismatched lengths raise
        :class:`ValueError`. Cell values are rendered via
        :func:`_format_value` so |None|, ``NaN``, and ``NaT`` become
        the empty string.
    style
        Built-in style name — one of ``"modern"`` (default),
        ``"zebra"``, ``"minimal"``, ``"corporate"``. Unknown names
        raise :class:`ValueError`.
    header_fill
        Override the resolved style's header fill. Accepts an
        |RGBColor|, a 6-character hex string, or |None| (suppresses the
        fill). Defaults to the resolved style's value.
    alt_row_fill
        Override the resolved style's banded-row fill. Accepts an
        |RGBColor|, a hex string, or |None| (suppresses banding).
        Defaults to the resolved style's value.
    autofit
        When |True| (the default) call :meth:`Table.autofit` so Word
        sizes columns to their content. Pass |False| to leave column
        widths at the document defaults.
    column_alignments
        Optional per-column paragraph alignment overrides. One
        :class:`docx.enum.text.WD_ALIGN_PARAGRAPH` value per column.
        |None| entries fall back to the auto-detected alignment.

    Returns
    -------
    Table
        The freshly-appended |Table|.

    Raises
    ------
    ValueError
        On unknown ``style``, malformed colour overrides, or
        ``rows`` whose length disagrees with ``headers``.

    .. versionadded:: 2026.05.29
    """
    if not isinstance(style, str):
        raise ValueError("style must be a string; got %r" % (style,))
    spec = _resolve_style(style)

    # -- Resolve colour overrides. ``header_fill`` / ``alt_row_fill``
    # -- left at their default sentinel inherit from the style; an
    # -- explicit None suppresses the fill; an explicit RGBColor /
    # -- hex string overrides it.
    if isinstance(header_fill, _UseStyleDefault):
        header_color = spec.header_fill
    else:
        header_color = _coerce_color(header_fill)
    if isinstance(alt_row_fill, _UseStyleDefault):
        band_color = spec.alt_row_fill
    else:
        band_color = _coerce_color(alt_row_fill)

    headers_list = list(headers)
    if not headers_list:
        raise ValueError("headers must be a non-empty sequence")
    ncols = len(headers_list)

    # -- Validate every row up front so we don't half-emit a table.
    rows_list: List[List[Any]] = []
    for index, row in enumerate(rows):
        row_list = list(row)
        if len(row_list) != ncols:
            raise ValueError(
                "rows[%d] has %d cells; expected %d to match headers"
                % (index, len(row_list), ncols)
            )
        rows_list.append(row_list)

    # -- Per-column alignment based on the values in that column. --
    auto_aligns: List[int] = []
    for col in range(ncols):
        column_values = [row_list[col] for row_list in rows_list]
        auto_aligns.append(_column_alignment(column_values))

    if column_alignments is not None:
        overrides = list(column_alignments)
        if len(overrides) != ncols:
            raise ValueError(
                "column_alignments has %d entries; expected %d"
                % (len(overrides), ncols)
            )
        for i, ov in enumerate(overrides):
            if ov is not None:
                auto_aligns[i] = ov

    # -- Build the table. Start with one row (the header) and append
    # -- data rows so each row's cells start in a sensible state. --
    table = document.add_table(rows=1, cols=ncols)
    _apply_table_style(table, spec.table_style)
    if autofit:
        try:
            table.autofit = True
        except Exception:  # pragma: no cover - defensive
            pass

    # -- Header row --
    header_row = table.rows[0]
    for col, label in enumerate(headers_list):
        # -- Numeric / date columns right-align even in the header so
        # -- the column reads consistently top to bottom. The label
        # -- itself is always shown left-aligned for short headers and
        # -- right-aligned for numeric / date columns -> use auto. --
        _set_cell_text(
            header_row.cells[col],
            str(label),
            bold=True,
            text_color=spec.header_text,
            alignment=auto_aligns[col],
        )
    _apply_row_shading(header_row, header_color)
    if spec.header_underline_only:
        _apply_header_underline(header_row)

    # -- Body rows --
    for index, row_list in enumerate(rows_list):
        body_row = table.add_row()
        for col, value in enumerate(row_list):
            _set_cell_text(
                body_row.cells[col],
                _format_value(value),
                alignment=auto_aligns[col],
            )
        # -- Banded rows: every other body row gets the alt fill.
        # -- Index 0 here is the first body row; we shade odd indices
        # -- so the very first body row stays unshaded (matching
        # -- Word's default banded-row pattern). --
        if band_color is not None and (index % 2) == 1:
            _apply_row_shading(body_row, band_color)

    return table


# -- Public API: from_dataframe ------------------------------------------


def from_dataframe(
    document: "DocumentCls",
    df: Any,
    *,
    style: str = "modern",
    auto_format: bool = True,
    header_fill: Any = _USE_STYLE_DEFAULT,
    alt_row_fill: Any = _USE_STYLE_DEFAULT,
    autofit: bool = True,
    include_index: bool = False,
) -> "Table":
    """Append a styled table built from a pandas |DataFrame|.

    Parameters
    ----------
    document
        Target |Document|.
    df
        A :class:`pandas.DataFrame`. Pandas is an *optional* dependency
        — when not installed the call raises :class:`ImportError` with
        an actionable message.
    style
        Built-in style name. See :func:`styled_table` for the list.
    auto_format
        When |True| (the default), inspect each column's dtype and
        right-align numeric / date columns automatically; values are
        rendered with the same defaults as :func:`styled_table`. Pass
        |False| to render every value via ``str()`` and left-align
        every column.
    header_fill, alt_row_fill, autofit
        See :func:`styled_table`.
    include_index
        When |True| (default |False|), prepend the DataFrame's index
        as the first column. The index name (or ``""`` when unset) is
        used as that column's header label.

    Returns
    -------
    Table
        The freshly-appended |Table|.

    Raises
    ------
    ImportError
        When ``pandas`` is not installed.
    ValueError
        On unknown ``style``, malformed colour overrides, or when
        ``df`` does not quack like a DataFrame (e.g. a list passed by
        mistake).

    .. versionadded:: 2026.05.29
    """
    pd = _require_pandas()
    if not isinstance(df, pd.DataFrame):
        raise ValueError(
            "from_dataframe expects a pandas.DataFrame; got %s"
            % type(df).__name__
        )

    column_labels: List[str] = [str(c) for c in df.columns]
    headers: List[str]
    body_rows: List[List[Any]]

    if include_index:
        index_label = df.index.name if df.index.name is not None else ""
        headers = [str(index_label), *column_labels]
        body_rows = []
        for idx, row in zip(df.index, df.itertuples(index=False, name=None)):
            body_rows.append([idx, *list(row)])
    else:
        headers = column_labels
        body_rows = [
            list(row) for row in df.itertuples(index=False, name=None)
        ]

    if not auto_format:
        # -- Stringify every body cell so the styled-table sniffer
        # -- left-aligns every column. Headers pass through. --
        body_rows = [[str(v) for v in row] for row in body_rows]

    return styled_table(
        document,
        headers=headers,
        rows=body_rows,
        style=style,
        header_fill=header_fill,
        alt_row_fill=alt_row_fill,
        autofit=autofit,
    )


__all__ = [
    "BUILTIN_STYLES",
    "from_dataframe",
    "styled_table",
]
