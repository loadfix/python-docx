"""Comparison / pricing / rubric table builders.

Closes #290.

This module exposes three table-shaped recipe helpers that append a
fully-formed table to an existing |Document|::

    from docx import Document
    from docx.kit import tables_compare as tables

    doc = Document()

    # Comparison table — feature/option matrix
    tables.comparison(
        doc,
        options=["Plan A", "Plan B", "Plan C"],
        features={
            "Users":       ["1",     "10",    "Unlimited"],
            "Storage":     ["10 GB", "100 GB", "1 TB"],
            "Support":     ["Email", "Chat",  "Phone+SLA"],
            "Price/month": ["$9",    "$29",   "$99"],
        },
        recommended="Plan B",
    )

    # Pricing table — three-tier
    tables.pricing(doc, tiers=[
        {"name": "Starter",  "price": "$9/mo",  "bullets": ["1 user", "10 GB", "email support"]},
        {"name": "Pro",      "price": "$29/mo", "bullets": ["10 users", "100 GB", "chat support"], "highlighted": True},
        {"name": "Business", "price": "$99/mo", "bullets": ["unlimited", "1 TB", "phone+SLA"]},
    ])

    # Rubric — scoring grid
    tables.rubric(
        doc,
        criteria=["Clarity", "Accuracy", "Style"],
        levels=["Poor (1)", "OK (3)", "Excellent (5)"],
        cells=[
            ["unclear",   "mostly clear", "crystal clear"],
            ["3+ errors", "1-2 errors",   "no errors"],
            ["awkward",   "readable",     "polished"],
        ],
    )

    doc.save("out.docx")

Each helper returns the appended :class:`docx.table.Table` so callers
can post-process (resize columns, swap a cell style, etc.).

**Highlight semantics.** :func:`comparison` shades the *recommended*
column (matched by name against ``options``) and stamps a small
"Recommended" badge above the option header. :func:`pricing` shades
the column whose tier mapping carries ``"highlighted": True`` (multiple
highlighted tiers are tolerated — every flagged column is shaded). The
shade uses an accent fill colour (``#DCE6F1`` light-blue by default) —
pick a different fill via the per-call ``highlight_fill`` keyword.

**No XML reach-down** — every helper composes only public python-docx
API (``Document.add_paragraph``, ``Document.add_table``, ``_Cell.text``,
``_Cell.shading.fill_color``, ``Paragraph.alignment``, ``Run.bold``).

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import (
    TYPE_CHECKING,
    Any,
    List,
    Mapping,
    Optional,
    Sequence,
    Union,
)

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph


# -- Defaults -------------------------------------------------------------

#: Default highlight fill colour (light blue, Office "Accent 1, Lighter 80%").
DEFAULT_HIGHLIGHT_FILL: RGBColor = RGBColor(0xDC, 0xE6, 0xF1)

#: Recommended-column badge text, rendered above the matched option header.
RECOMMENDED_BADGE: str = "Recommended"


# -- Internal helpers -----------------------------------------------------


def _coerce_fill(value: Union[RGBColor, str, None]) -> Optional[RGBColor]:
    """Coerce ``value`` into an :class:`RGBColor` (or |None|)."""
    if value is None:
        return None
    if isinstance(value, RGBColor):
        return value
    text = str(value).strip().lstrip("#")
    if len(text) == 3:
        # -- short-form `#abc` -> `#aabbcc` --
        text = "".join(ch * 2 for ch in text)
    if len(text) != 6:
        raise ValueError(
            "highlight_fill must be a 6-digit hex string or RGBColor; got %r"
            % (value,)
        )
    try:
        return RGBColor.from_string(text.upper())
    except (ValueError, TypeError) as exc:
        raise ValueError(
            "highlight_fill must be a 6-digit hex string or RGBColor; got %r"
            % (value,)
        ) from exc


def _set_cell_text(
    cell: "_Cell",
    text: str,
    bold: bool = False,
    alignment: Optional[int] = None,
) -> None:
    """Write ``text`` into ``cell`` with optional bold + paragraph alignment."""
    cell.text = text
    para = cell.paragraphs[0]
    if alignment is not None:
        para.alignment = alignment
    if bold:
        for run in para.runs:
            run.bold = True


def _shade_cell(cell: "_Cell", fill: Optional[RGBColor]) -> None:
    """Apply ``fill`` background shading to ``cell`` (no-op when ``None``)."""
    if fill is None:
        return
    cell.shading.fill_color = fill


def _apply_table_grid(table: "Table") -> None:
    """Apply ``Table Grid`` style; silently fall back when missing."""
    try:
        table.style = "Table Grid"
    except KeyError:
        pass


def _add_heading(
    document: "DocumentCls", text: str, level: int = 2
) -> "Paragraph":
    """Append a heading; falls back to a bold paragraph when style absent."""
    try:
        return document.add_heading(text, level=level)
    except KeyError:
        para = document.add_paragraph()
        run = para.add_run(text)
        run.bold = True
        return para


def _require_non_empty_sequence(
    seq: Optional[Sequence[Any]], name: str
) -> List[Any]:
    """Return ``list(seq)`` after asserting it is non-empty."""
    if seq is None:
        raise ValueError(f"{name} is required and must be a non-empty sequence")
    items = list(seq)
    if not items:
        raise ValueError(f"{name} must contain at least one entry")
    return items


# -- Public: comparison ---------------------------------------------------


def comparison(
    document: "DocumentCls",
    *,
    options: Sequence[str],
    features: Mapping[str, Sequence[str]],
    recommended: Optional[str] = None,
    highlight_fill: Union[RGBColor, str, None] = None,
    title: Optional[str] = None,
) -> "Table":
    """Append a feature-by-option comparison table to ``document``.

    The first row is the option headers (one column per ``options``
    entry, plus a leading "Feature" label cell). Each subsequent row
    renders a feature name followed by that feature's value for each
    option. When ``recommended`` matches an entry in ``options`` (case
    sensitive), the matched column is shaded with ``highlight_fill``
    and the option header carries a small "Recommended" badge above
    the option name.

    Parameters
    ----------
    document
        The |Document| to append to.
    options
        Column headers — one per option (e.g. plan / vendor name).
        Required, must be non-empty.
    features
        Mapping of feature label -> per-option value sequence. Each
        value sequence must have ``len(options)`` entries (one per
        option, in column order).
    recommended
        Name of the option to flag as "Recommended". Must match one of
        ``options`` exactly. When omitted, no column is shaded.
    highlight_fill
        Override the recommended-column fill colour. Accepts an
        :class:`RGBColor`, a 3- or 6-digit hex string (with or without
        leading ``#``), or |None| (use the default).
    title
        Optional heading rendered above the table.

    Returns
    -------
    Table
        The freshly-appended :class:`docx.table.Table` (already in
        ``document.tables``).

    Raises
    ------
    ValueError
        When ``options`` is empty, when any feature row has the wrong
        number of values, when ``recommended`` doesn't match an
        option, or when ``highlight_fill`` is malformed.

    .. versionadded:: 2026.05.29
    """
    option_list = _require_non_empty_sequence(options, "options")
    if features is None or not features:
        raise ValueError("features is required and must contain at least one row")

    # -- Row-shape validation up front -- a clear error beats a partial table.
    for feature, values in features.items():
        if values is None:
            raise ValueError(
                "features[%r] must be a sequence of %d values; got None"
                % (feature, len(option_list))
            )
        if len(list(values)) != len(option_list):
            raise ValueError(
                "features[%r] must have %d values (one per option); got %d"
                % (feature, len(option_list), len(list(values)))
            )

    fill = (
        _coerce_fill(highlight_fill)
        if highlight_fill is not None
        else DEFAULT_HIGHLIGHT_FILL
    )

    recommended_index: Optional[int] = None
    if recommended is not None:
        if recommended not in option_list:
            raise ValueError(
                "recommended=%r does not match any option; valid options are %r"
                % (recommended, option_list)
            )
        recommended_index = option_list.index(recommended)

    if title:
        _add_heading(document, title, level=2)

    # -- Layout: 1 label column + N option columns. --
    cols = 1 + len(option_list)
    table = document.add_table(rows=1, cols=cols)
    _apply_table_grid(table)

    header_cells = table.rows[0].cells
    _set_cell_text(
        header_cells[0], "Feature", bold=True, alignment=WD_ALIGN_PARAGRAPH.LEFT
    )
    for col_idx, option_name in enumerate(option_list, start=1):
        cell_text = option_name
        if recommended_index is not None and col_idx == recommended_index + 1:
            cell_text = f"{RECOMMENDED_BADGE}\n{option_name}"
        _set_cell_text(
            header_cells[col_idx],
            cell_text,
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        if recommended_index is not None and col_idx == recommended_index + 1:
            _shade_cell(header_cells[col_idx], fill)

    for feature_label, values in features.items():
        row_cells = table.add_row().cells
        _set_cell_text(
            row_cells[0],
            str(feature_label),
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
        )
        for col_idx, value in enumerate(values, start=1):
            _set_cell_text(
                row_cells[col_idx],
                str(value),
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
            )
            if recommended_index is not None and col_idx == recommended_index + 1:
                _shade_cell(row_cells[col_idx], fill)

    return table


# -- Public: pricing ------------------------------------------------------


class _Tier:
    """Tiny normalised pricing-tier record."""

    __slots__ = ("name", "price", "bullets", "highlighted")

    def __init__(
        self,
        name: str,
        price: str,
        bullets: List[str],
        highlighted: bool,
    ) -> None:
        self.name = name
        self.price = price
        self.bullets = bullets
        self.highlighted = highlighted


def _normalise_tier(tier: Mapping[str, Any], index: int) -> _Tier:
    """Validate a single ``tiers`` entry."""
    if not isinstance(tier, Mapping):  # type: ignore[arg-type]
        raise ValueError(
            "tiers[%d] must be a mapping with at least 'name' and 'price'" % index
        )
    name = tier.get("name")
    if name is None or str(name).strip() == "":
        raise ValueError("tiers[%d] is missing a non-empty 'name'" % index)
    price = tier.get("price")
    if price is None or str(price).strip() == "":
        raise ValueError("tiers[%d] is missing a non-empty 'price'" % index)
    bullets_raw = tier.get("bullets") or []
    if isinstance(bullets_raw, str):
        # -- a bare string is a common typo; reject loudly --
        raise ValueError(
            "tiers[%d]: 'bullets' must be a sequence of strings, not a single string"
            % index
        )
    bullets = [str(b) for b in bullets_raw]
    highlighted = bool(tier.get("highlighted", False))
    return _Tier(
        name=str(name),
        price=str(price),
        bullets=bullets,
        highlighted=highlighted,
    )


def pricing(
    document: "DocumentCls",
    *,
    tiers: Sequence[Mapping[str, Any]],
    highlight_fill: Union[RGBColor, str, None] = None,
    title: Optional[str] = None,
) -> "Table":
    """Append a side-by-side pricing-tier table to ``document``.

    Each tier is a mapping with ``name``, ``price``, optional
    ``bullets`` (sequence of strings), and optional ``highlighted``
    (boolean — when truthy, the column is shaded with
    ``highlight_fill``). Three or more tiers is the conventional shape
    but the helper accepts any non-empty sequence.

    Layout (one column per tier): a name row, a price row, then one
    row per bullet position (padded to the longest tier with empty
    cells in shorter tiers).

    Parameters
    ----------
    document
        The |Document| to append to.
    tiers
        Sequence of tier mappings. Each must carry ``name`` and
        ``price``. ``bullets`` defaults to an empty list;
        ``highlighted`` defaults to |False|.
    highlight_fill
        Override the highlighted-column fill colour. Accepts the same
        forms as :func:`comparison`'s ``highlight_fill``.
    title
        Optional heading rendered above the table.

    Returns
    -------
    Table
        The freshly-appended :class:`docx.table.Table`.

    Raises
    ------
    ValueError
        When ``tiers`` is empty, when any tier is missing ``name`` /
        ``price``, when ``bullets`` is a bare string, or when
        ``highlight_fill`` is malformed.

    .. versionadded:: 2026.05.29
    """
    if not tiers:
        raise ValueError("tiers is required and must contain at least one tier")
    normalised: List[_Tier] = [
        _normalise_tier(tier, index) for index, tier in enumerate(tiers)
    ]
    fill = (
        _coerce_fill(highlight_fill)
        if highlight_fill is not None
        else DEFAULT_HIGHLIGHT_FILL
    )

    if title:
        _add_heading(document, title, level=2)

    cols = len(normalised)
    max_bullets = max((len(t.bullets) for t in normalised), default=0)
    # -- 2 fixed rows (name, price) + one row per bullet position. --
    rows = 2 + max_bullets

    table = document.add_table(rows=rows, cols=cols)
    _apply_table_grid(table)

    # -- Name row --
    name_cells = table.rows[0].cells
    for col_idx, tier in enumerate(normalised):
        _set_cell_text(
            name_cells[col_idx],
            tier.name,
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        if tier.highlighted:
            _shade_cell(name_cells[col_idx], fill)

    # -- Price row --
    price_cells = table.rows[1].cells
    for col_idx, tier in enumerate(normalised):
        _set_cell_text(
            price_cells[col_idx],
            tier.price,
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        if tier.highlighted:
            _shade_cell(price_cells[col_idx], fill)

    # -- Bullet rows --
    for bullet_idx in range(max_bullets):
        row_cells = table.rows[2 + bullet_idx].cells
        for col_idx, tier in enumerate(normalised):
            text = tier.bullets[bullet_idx] if bullet_idx < len(tier.bullets) else ""
            _set_cell_text(
                row_cells[col_idx],
                text,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
            )
            if tier.highlighted:
                _shade_cell(row_cells[col_idx], fill)

    return table


# -- Public: rubric -------------------------------------------------------


def rubric(
    document: "DocumentCls",
    *,
    criteria: Sequence[str],
    levels: Sequence[str],
    cells: Sequence[Sequence[str]],
    title: Optional[str] = None,
) -> "Table":
    """Append a scoring-rubric table to ``document``.

    The rendered grid has shape ``(1 + len(criteria)) x (1 + len(levels))``:
    the top-left cell is empty, the top row carries the level labels,
    the leftmost column carries the criteria labels, and ``cells[i][j]``
    fills the criterion-``i`` / level-``j`` body cell.

    Parameters
    ----------
    document
        The |Document| to append to.
    criteria
        Row labels — one per criterion (e.g. "Clarity", "Accuracy").
        Required, must be non-empty.
    levels
        Column labels — one per scoring level (e.g. "Poor (1)",
        "Excellent (5)"). Required, must be non-empty.
    cells
        ``len(criteria)`` x ``len(levels)`` grid of cell text. Row
        ``i`` is the body row for ``criteria[i]``; column ``j`` is the
        body column for ``levels[j]``.
    title
        Optional heading rendered above the table.

    Returns
    -------
    Table
        The freshly-appended :class:`docx.table.Table`.

    Raises
    ------
    ValueError
        When ``criteria`` / ``levels`` are empty, or when ``cells`` does
        not have shape ``len(criteria) x len(levels)``.

    .. versionadded:: 2026.05.29
    """
    criteria_list = _require_non_empty_sequence(criteria, "criteria")
    levels_list = _require_non_empty_sequence(levels, "levels")

    if cells is None:
        raise ValueError("cells is required and must be a 2-D sequence")
    cells_list = list(cells)
    if len(cells_list) != len(criteria_list):
        raise ValueError(
            "cells must have %d rows (one per criterion); got %d"
            % (len(criteria_list), len(cells_list))
        )
    for row_idx, row in enumerate(cells_list):
        if row is None:
            raise ValueError(
                "cells[%d] must be a sequence of %d values; got None"
                % (row_idx, len(levels_list))
            )
        row_values = list(row)
        if len(row_values) != len(levels_list):
            raise ValueError(
                "cells[%d] must have %d values (one per level); got %d"
                % (row_idx, len(levels_list), len(row_values))
            )

    if title:
        _add_heading(document, title, level=2)

    n_cols = 1 + len(levels_list)
    n_rows = 1 + len(criteria_list)
    table = document.add_table(rows=n_rows, cols=n_cols)
    _apply_table_grid(table)

    # -- Header row: blank corner + level labels. --
    header_cells = table.rows[0].cells
    _set_cell_text(header_cells[0], "")
    for col_idx, level_label in enumerate(levels_list, start=1):
        _set_cell_text(
            header_cells[col_idx],
            str(level_label),
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )

    # -- Body rows: criterion label + per-level cell. --
    for row_idx, criterion in enumerate(criteria_list, start=1):
        row_cells = table.rows[row_idx].cells
        _set_cell_text(
            row_cells[0],
            str(criterion),
            bold=True,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
        )
        for col_idx, value in enumerate(cells_list[row_idx - 1], start=1):
            _set_cell_text(
                row_cells[col_idx],
                str(value),
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
            )

    return table


__all__ = [
    "comparison",
    "pricing",
    "rubric",
    "DEFAULT_HIGHLIGHT_FILL",
    "RECOMMENDED_BADGE",
]
