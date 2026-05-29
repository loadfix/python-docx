"""Embed matplotlib figures as inline pictures with brand colours.

Closes #291.

Most "chart in a Word document" requirements arrive as a matplotlib
``Figure`` already wired up by the calling pipeline (notebooks, reports,
data-engineering scripts). The native :meth:`Document.add_picture` is
fine if you already have the chart on disk, but the typical authoring
shape — *take this Figure, make it look on-brand, drop it on page N
with alt-text* — wants three things bundled:

1. Render the figure to PNG **in-memory** (no temp files).
2. Apply a list of brand colours to the figure's lines / bars / pie
   wedges before render so corporate output isn't matplotlib teal.
3. Embed the result as an inline picture sized in inches, with a
   keyword-argument ``alt_text`` set on ``wp:docPr/@descr``.

Public API::

    from docx import Document
    from docx.kit import charts
    import matplotlib.pyplot as plt

    doc = Document()
    fig, ax = plt.subplots()
    ax.plot([1, 2, 3, 4], [10, 20, 15, 25])
    ax.set_title("Quarterly revenue")

    charts.add_chart(
        doc, fig,
        brand_colors=["#1f4e79", "#2e75b6", "#9dc3e6"],
        width_in=6.0,
        alt_text="Revenue chart",
    )

    # Convenience wrappers — no manual figure construction:
    charts.bar_chart(doc, x=["Q1", "Q2", "Q3", "Q4"], y=[10, 20, 15, 25],
                     title="Revenue")
    charts.line_chart(doc, x=[1, 2, 3, 4], y=[10, 20, 15, 25],
                      title="Trend")
    charts.pie_chart(doc, labels=["A", "B", "C"], values=[30, 40, 30],
                     title="Mix")

matplotlib is an **optional** dependency — the module is importable
without it, but every helper raises :class:`ImportError` with a
``pip install python-docx[matplotlib]`` hint when invoked. Tests gate
on ``pytest.importorskip("matplotlib")`` so a clean checkout without
matplotlib still runs the rest of the kit suite green.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import io
from typing import TYPE_CHECKING, Any, List, Optional, Sequence, Union

from docx.shared import Inches

if TYPE_CHECKING:  # pragma: no cover - import-time only
    from docx.document import Document
    from docx.shape import InlineShape


# -- Sentinel for the optional matplotlib import.  Each helper threads
# -- through ``_require_matplotlib()`` and re-raises with an actionable
# -- message rather than letting a bare ``ModuleNotFoundError`` bubble
# -- up from the import line.
_INSTALL_HINT = (
    "matplotlib is required for docx.kit.charts; install it via the "
    "optional extra: pip install 'python-docx[matplotlib]'"
)


def _require_matplotlib() -> Any:
    """Import and return the ``matplotlib`` module, or raise |ImportError|.

    Imports are deferred to call time so ``from docx.kit import charts``
    is import-safe in environments without matplotlib (the rest of the
    kit must keep working).
    """
    try:
        import matplotlib  # type: ignore[import-not-found]
    except ImportError as exc:  # pragma: no cover - exercised when missing
        raise ImportError(_INSTALL_HINT) from exc
    return matplotlib


def _require_pyplot() -> Any:
    """Import and return ``matplotlib.pyplot`` with the Agg backend forced.

    The convenience wrappers (``bar_chart`` / ``line_chart`` /
    ``pie_chart``) build a figure server-side; forcing the
    non-interactive ``Agg`` backend keeps them safe to invoke from a
    headless process (CI, web service, batch job) where no display is
    available.
    """
    matplotlib = _require_matplotlib()
    # -- ``use(..., force=False)`` so a caller who has already chosen a
    # -- backend (e.g. ``%matplotlib inline`` in a notebook) keeps theirs.
    try:
        matplotlib.use("Agg", force=False)
    except Exception:  # pragma: no cover - already-finalised backend
        pass
    import matplotlib.pyplot as plt  # type: ignore[import-not-found]

    return plt


def _apply_brand_colors(fig: Any, brand_colors: Sequence[str]) -> None:
    """Recolor lines / bars / pie wedges on `fig` using `brand_colors`.

    The colour list is cycled — supplying a shorter palette than the
    number of artists is fine, the helper wraps with modulo. matplotlib
    decomposes its visuals into a small number of artist categories;
    this helper covers the three the convenience wrappers emit (lines,
    bars, pie wedges) and any directly-equivalent artists the caller
    may have added before passing the figure in.
    """
    if not brand_colors:
        return

    n = len(brand_colors)

    for ax_index, ax in enumerate(fig.axes):
        # -- Lines (``ax.plot(...)`` returns Line2D artists). --
        for i, line in enumerate(ax.get_lines()):
            line.set_color(brand_colors[i % n])

        # -- Bars / patches (``ax.bar(...)`` populates ``ax.patches``;
        # -- pie wedges land in ``ax.patches`` too).  Each patch gets a
        # -- distinct colour so the chart reads as a categorical view.
        for i, patch in enumerate(ax.patches):
            patch.set_facecolor(brand_colors[i % n])

        # -- Re-draw the legend if one is showing so the legend swatches
        # -- pick up the new colours.  ``ax.get_legend()`` is None when
        # -- no legend was created.
        legend = ax.get_legend()
        if legend is not None:
            ax.legend()

        # -- Mark the axis to the brand cycler so any subsequent artist
        # -- the caller adds (between figure construction and ``add_chart``)
        # -- inherits the palette.  This is a no-op for most callers
        # -- (the figure is already finalised) but cheap when it isn't.
        try:
            from cycler import cycler  # type: ignore[import-not-found]

            ax.set_prop_cycle(cycler(color=list(brand_colors)))
        except ImportError:  # pragma: no cover - cycler ships with matplotlib
            pass

        # -- ``ax_index`` is unused; the loop variable is kept for
        # -- clarity when reading stack traces.
        del ax_index


def _render_to_png_bytes(fig: Any, dpi: int) -> bytes:
    """Render `fig` to PNG bytes via an in-memory buffer.

    ``bbox_inches="tight"`` trims the surrounding whitespace so the
    embedded picture sits flush against the figure content; this is the
    convention every "matplotlib in a report" pipeline uses.
    """
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
    buf.seek(0)
    return buf.getvalue()


def add_chart(
    document: "Document",
    figure: Any,
    *,
    brand_colors: Optional[Sequence[str]] = None,
    width_in: float = 6.0,
    alt_text: Optional[str] = None,
    dpi: int = 200,
) -> "InlineShape":
    """Render `figure` to PNG and embed it as an inline picture.

    Parameters
    ----------
    document
        The :class:`Document` to mutate. The picture is appended in its
        own paragraph at the end of the document body, matching
        :meth:`Document.add_picture`.
    figure
        A ``matplotlib.figure.Figure`` instance. Anything with a
        ``savefig(buf, format="png", dpi=...)`` method that produces
        PNG bytes is accepted, but the brand-colour pass assumes the
        matplotlib artist API (``axes`` / ``get_lines`` / ``patches``).
    brand_colors
        Optional list of CSS-style colour strings (``"#1f4e79"``,
        ``"red"``, ``"rgb(31,78,121)"``) cycled through the figure's
        artists before render. ``None`` (the default) preserves
        matplotlib's default palette.
    width_in
        Width of the embedded picture, in inches. Height scales
        proportionally to preserve the figure's aspect ratio. Defaults
        to ``6.0`` — the conventional "fits a letter / A4 page with
        1-inch margins" width.
    alt_text
        Accessibility description written to ``wp:inline/wp:docPr/@descr``
        on the inline shape. ``None`` (the default) leaves the attribute
        unset.
    dpi
        Resolution of the rendered PNG in dots-per-inch. Defaults to
        ``200`` — high enough to read crisply on a Word page at
        100% zoom but small enough to keep the embedded part under
        ~200 KiB for a typical chart.

    Returns
    -------
    InlineShape
        The freshly-appended inline shape wrapping the embedded PNG.
    """
    _require_matplotlib()

    if width_in <= 0:
        raise ValueError("width_in must be > 0; got %r" % (width_in,))
    if dpi <= 0:
        raise ValueError("dpi must be > 0; got %r" % (dpi,))

    if brand_colors:
        _apply_brand_colors(figure, brand_colors)

    png_bytes = _render_to_png_bytes(figure, dpi=dpi)

    # -- ``Document.add_picture`` accepts a binary file-like; wrap the
    # -- bytes in BytesIO so we don't need a temp file. --
    stream = io.BytesIO(png_bytes)
    inline_shape = document.add_picture(stream, width=Inches(width_in))

    if alt_text is not None:
        inline_shape.alt_text = alt_text

    return inline_shape


# -- Convenience wrappers.  Each builds a tiny matplotlib figure
# -- server-side, dispatches to ``add_chart``, and closes the figure to
# -- release the memory.  Closing is important when callers loop over
# -- many records — pyplot keeps a strong reference to every Figure it
# -- creates until ``plt.close(fig)``.


def _close(plt: Any, fig: Any) -> None:
    """Close `fig` via `plt.close`, ignoring any backend errors."""
    try:
        plt.close(fig)
    except Exception:  # pragma: no cover - defensive; close shouldn't raise
        pass


def bar_chart(
    document: "Document",
    *,
    x: Sequence[Union[str, float, int]],
    y: Sequence[Union[float, int]],
    title: Optional[str] = None,
    brand_colors: Optional[Sequence[str]] = None,
    width_in: float = 6.0,
    alt_text: Optional[str] = None,
    dpi: int = 200,
) -> "InlineShape":
    """Build a bar chart and embed it via :func:`add_chart`.

    Convenience wrapper for the common shape "vector of categories on
    the X axis, vector of values on the Y axis". The built figure is
    closed before return so the caller doesn't accumulate matplotlib
    state across many calls.
    """
    plt = _require_pyplot()

    if len(x) != len(y):
        raise ValueError(
            "x and y must have equal length; got len(x)=%d, len(y)=%d"
            % (len(x), len(y))
        )

    fig, ax = plt.subplots()
    try:
        ax.bar(list(x), list(y))
        if title is not None:
            ax.set_title(title)
        return add_chart(
            document,
            fig,
            brand_colors=brand_colors,
            width_in=width_in,
            alt_text=alt_text,
            dpi=dpi,
        )
    finally:
        _close(plt, fig)


def line_chart(
    document: "Document",
    *,
    x: Sequence[Union[float, int]],
    y: Sequence[Union[float, int]],
    title: Optional[str] = None,
    brand_colors: Optional[Sequence[str]] = None,
    width_in: float = 6.0,
    alt_text: Optional[str] = None,
    dpi: int = 200,
) -> "InlineShape":
    """Build a line chart and embed it via :func:`add_chart`."""
    plt = _require_pyplot()

    if len(x) != len(y):
        raise ValueError(
            "x and y must have equal length; got len(x)=%d, len(y)=%d"
            % (len(x), len(y))
        )

    fig, ax = plt.subplots()
    try:
        ax.plot(list(x), list(y))
        if title is not None:
            ax.set_title(title)
        return add_chart(
            document,
            fig,
            brand_colors=brand_colors,
            width_in=width_in,
            alt_text=alt_text,
            dpi=dpi,
        )
    finally:
        _close(plt, fig)


def pie_chart(
    document: "Document",
    *,
    labels: Sequence[str],
    values: Sequence[Union[float, int]],
    title: Optional[str] = None,
    brand_colors: Optional[Sequence[str]] = None,
    width_in: float = 6.0,
    alt_text: Optional[str] = None,
    dpi: int = 200,
) -> "InlineShape":
    """Build a pie chart and embed it via :func:`add_chart`."""
    plt = _require_pyplot()

    if len(labels) != len(values):
        raise ValueError(
            "labels and values must have equal length; "
            "got len(labels)=%d, len(values)=%d"
            % (len(labels), len(values))
        )

    fig, ax = plt.subplots()
    try:
        ax.pie(list(values), labels=list(labels))
        # -- A pie chart reads best on a square aspect ratio. --
        ax.set_aspect("equal")
        if title is not None:
            ax.set_title(title)
        return add_chart(
            document,
            fig,
            brand_colors=brand_colors,
            width_in=width_in,
            alt_text=alt_text,
            dpi=dpi,
        )
    finally:
        _close(plt, fig)


__all__: List[str] = [
    "add_chart",
    "bar_chart",
    "line_chart",
    "pie_chart",
]
