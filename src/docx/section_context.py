"""Context-manager helpers for ergonomic section authoring (issue #79).

The OOXML section model encodes per-region page setup (orientation,
margins, columns, headers/footers) by attaching a ``w:sectPr`` to the
*last paragraph* of the region. Authoring this by hand is error-prone:
the caller has to remember to take a snapshot of the previous section,
mutate the new ``sectPr``, and then add a fresh section break afterward
to revert. :func:`Document.section` exposes a context manager that
orchestrates that dance.

Usage::

    with doc.section(orientation='landscape', margins='narrow'):
        doc.add_paragraph(...)
        doc.add_table(...)
    # implicit section break here, returns to portrait

OOXML sections do not nest. Attempting to enter a section context while
another is already active raises :class:`NestedSectionError` with a
clear message.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Optional, Union

from docx.enum.section import WD_ORIENTATION, WD_SECTION
from docx.exceptions import NestedSectionError
from docx.shared import Inches, Length

if TYPE_CHECKING:
    from docx.document import Document
    from docx.section import Section


# -- Named margin presets (top, right, bottom, left) in inches.  Mirrors
# -- Word's UI "Narrow / Normal / Wide" preset list.
_MARGIN_PRESETS: dict[str, tuple[float, float, float, float]] = {
    "narrow": (0.5, 0.5, 0.5, 0.5),
    "normal": (1.0, 1.0, 1.0, 1.0),
    "moderate": (1.0, 0.75, 1.0, 0.75),
    "wide": (1.0, 2.0, 1.0, 2.0),
}


def _resolve_margins(
    spec: Union[str, dict, tuple, list, None],
) -> "Optional[dict[str, Length]]":
    """Translate a margin spec into a ``{edge: Length}`` mapping.

    Accepts a preset name (``"narrow"`` / ``"normal"`` / ``"moderate"`` /
    ``"wide"``), a 4-tuple of ``Length``-or-float-inches, or a dict with
    ``top``/``right``/``bottom``/``left`` keys (any subset). Returns
    |None| when ``spec`` is |None|.
    """
    if spec is None:
        return None
    if isinstance(spec, str):
        key = spec.lower()
        if key not in _MARGIN_PRESETS:
            raise ValueError(
                "unknown margins preset %r; expected one of %s or an explicit "
                "dict/tuple" % (spec, sorted(_MARGIN_PRESETS))
            )
        top, right, bottom, left = _MARGIN_PRESETS[key]
        return {
            "top": Inches(top),
            "right": Inches(right),
            "bottom": Inches(bottom),
            "left": Inches(left),
        }
    if isinstance(spec, (tuple, list)):
        if len(spec) != 4:
            raise ValueError(
                "margins tuple must have exactly 4 entries (top, right, "
                "bottom, left); got %d" % len(spec)
            )
        edges = ("top", "right", "bottom", "left")
        return {edge: _coerce_length(v) for edge, v in zip(edges, spec)}
    if isinstance(spec, dict):
        out: "dict[str, Length]" = {}
        for edge in ("top", "right", "bottom", "left"):
            if edge in spec and spec[edge] is not None:
                out[edge] = _coerce_length(spec[edge])
        return out
    raise TypeError(
        "margins must be a preset name, 4-tuple, dict, or None; got %r"
        % type(spec).__name__
    )


def _coerce_length(value: Any) -> Length:
    """Accept a |Length| or a numeric value interpreted as inches."""
    if isinstance(value, Length):
        return value
    if isinstance(value, (int, float)):
        return Inches(float(value))
    raise TypeError(
        "margin value must be a Length or a number-of-inches; got %r"
        % type(value).__name__
    )


def _resolve_page_size(
    spec: Union[str, tuple, list, dict, None],
) -> "Optional[tuple[Length, Length]]":
    """Translate a page-size spec into a ``(width, height)`` pair.

    Accepts a preset name (``"letter"`` / ``"legal"`` / ``"a4"`` /
    ``"a3"``), a 2-tuple ``(width, height)``, or a dict with ``width``
    and ``height`` keys. Returns |None| when ``spec`` is |None|.
    """
    if spec is None:
        return None
    presets: dict[str, tuple[float, float]] = {
        "letter": (8.5, 11.0),
        "legal": (8.5, 14.0),
        "a4": (8.27, 11.69),
        "a3": (11.69, 16.54),
        "tabloid": (11.0, 17.0),
    }
    if isinstance(spec, str):
        key = spec.lower()
        if key not in presets:
            raise ValueError(
                "unknown page_size preset %r; expected one of %s or an explicit "
                "(width, height)" % (spec, sorted(presets))
            )
        w, h = presets[key]
        return Inches(w), Inches(h)
    if isinstance(spec, (tuple, list)):
        if len(spec) != 2:
            raise ValueError(
                "page_size tuple must have exactly 2 entries (width, height); "
                "got %d" % len(spec)
            )
        return _coerce_length(spec[0]), _coerce_length(spec[1])
    if isinstance(spec, dict):
        if "width" not in spec or "height" not in spec:
            raise ValueError("page_size dict must have 'width' and 'height' keys")
        return _coerce_length(spec["width"]), _coerce_length(spec["height"])
    raise TypeError(
        "page_size must be a preset name, 2-tuple, dict, or None; got %r"
        % type(spec).__name__
    )


class _SectionContext:
    """Context manager returned by :meth:`Document.section`.

    On enter, appends a section break and applies the requested
    properties to the *new* section. On exit, appends a second section
    break that reverts to whatever section was active before — letting
    callers nest content under custom orientation/margins/columns
    without manually managing the section graph.

    OOXML does not allow nested sections, so this context tracks an
    activation flag on the parent document; entering while another
    context is active raises :class:`NestedSectionError`.
    """

    def __init__(
        self,
        document: "Document",
        *,
        orientation: Union[str, WD_ORIENTATION, None] = None,
        margins: Union[str, dict, tuple, list, None] = None,
        page_size: Union[str, tuple, list, dict, None] = None,
        page_numbering: Optional[dict] = None,
        header: Optional[str] = None,
        footer: Optional[str] = None,
        columns: Union[int, dict, None] = None,
        line_numbering: Union[bool, dict, None] = None,
    ) -> None:
        self._document = document
        self._orientation = orientation
        self._margins = _resolve_margins(margins)
        self._page_size = _resolve_page_size(page_size)
        self._page_numbering = page_numbering
        self._header = header
        self._footer = footer
        self._columns = columns
        self._line_numbering = line_numbering
        self._inner_section: "Optional[Section]" = None

    def __enter__(self) -> "Section":
        if getattr(self._document, "_in_section_context", False):
            raise NestedSectionError(
                "OOXML sections cannot nest — close the outer section "
                "context (`with doc.section(...)`) before opening another"
            )
        self._document._in_section_context = True  # type: ignore[attr-defined]

        # -- end the *prior* section here (turning the current sentinel
        # -- into a paragraph-anchored sectPr) and obtain a fresh sentinel
        # -- which now controls the inner region.
        section = self._document.add_section(WD_SECTION.CONTINUOUS)
        self._apply_properties(section)
        self._inner_section = section
        return section

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        try:
            # -- Close the inner region by anchoring its sectPr at the
            # -- last paragraph and starting a fresh section that reverts
            # -- to the prior layout. We look up the *second-last* section
            # -- (the one that existed before __enter__) and copy its
            # -- relevant page-setup attributes onto the new sentinel.
            sections = self._document.sections
            # -- before __enter__: N sections; after __enter__: N+1.
            # -- Now adding another break makes it N+2; the new sentinel
            # -- should match section N-1 (the one before our inner).
            prior = sections[-2] if len(sections) >= 2 else None
            new_section = self._document.add_section(WD_SECTION.CONTINUOUS)
            if prior is not None:
                self._copy_page_setup(prior, new_section)
        finally:
            self._document._in_section_context = False  # type: ignore[attr-defined]

    def _apply_properties(self, section: "Section") -> None:
        """Write the queued properties onto ``section``."""
        if self._page_size is not None:
            width, height = self._page_size
            section.page_width = width
            section.page_height = height
        if self._orientation is not None:
            orientation = self._orientation
            if isinstance(orientation, str):
                orientation = self._coerce_orientation(orientation)
            # -- Section.orientation setter already swaps page_width and
            # -- page_height when the orientation changes (matching Word's
            # -- own behaviour), so we only need to assign the value.
            section.orientation = orientation
        if self._margins is not None:
            edge_to_attr = {
                "top": "top_margin",
                "right": "right_margin",
                "bottom": "bottom_margin",
                "left": "left_margin",
            }
            for edge, length in self._margins.items():
                setattr(section, edge_to_attr[edge], length)
        if self._columns is not None:
            self._apply_columns(section)
        if self._page_numbering is not None:
            self._apply_page_numbering(section)
        if self._line_numbering is not None:
            self._apply_line_numbering(section)
        if self._header is not None:
            section.header.paragraphs[0].text = self._header
        if self._footer is not None:
            section.footer.paragraphs[0].text = self._footer

    @staticmethod
    def _coerce_orientation(value: str) -> WD_ORIENTATION:
        key = value.strip().lower()
        if key in ("landscape", "land"):
            return WD_ORIENTATION.LANDSCAPE
        if key in ("portrait", "port"):
            return WD_ORIENTATION.PORTRAIT
        raise ValueError(
            "orientation must be 'portrait' or 'landscape'; got %r" % value
        )

    def _apply_columns(self, section: "Section") -> None:
        spec = self._columns
        if isinstance(spec, int):
            section.set_columns(count=spec)
            return
        if isinstance(spec, dict):
            kwargs = {}
            for key in ("count", "space", "equal_width", "widths", "separator"):
                if key in spec:
                    kwargs[key] = spec[key]
            section.set_columns(**kwargs)
            return
        raise TypeError(
            "columns must be an int or dict; got %r" % type(spec).__name__
        )

    def _apply_page_numbering(self, section: "Section") -> None:
        spec = self._page_numbering or {}
        kwargs: dict = {}
        if "style" in spec and spec["style"] is not None:
            kwargs["fmt"] = spec["style"]
        if "fmt" in spec and spec["fmt"] is not None:
            kwargs["fmt"] = spec["fmt"]
        if "start" in spec and spec["start"] is not None:
            kwargs["start"] = spec["start"]
        if "restart" in spec and spec["restart"] is not None:
            # -- 'restart' may be a bool (True ⇒ start at 1) or an int.
            value = spec["restart"]
            if value is True:
                kwargs["start"] = kwargs.get("start", 1)
            elif isinstance(value, int) and value is not False:
                kwargs["start"] = value
        if kwargs:
            section.set_page_numbering(**kwargs)

    def _apply_line_numbering(self, section: "Section") -> None:
        spec = self._line_numbering
        if spec is False:
            # -- explicit disable: leave default (no w:lnNumType) --
            return
        if spec is True:
            section.set_line_numbering(count_by=1)
            return
        if isinstance(spec, dict):
            kwargs = {}
            for key in ("count_by", "start", "distance", "restart"):
                if key in spec:
                    kwargs[key] = spec[key]
            section.set_line_numbering(**kwargs)
            return
        raise TypeError(
            "line_numbering must be a bool or dict; got %r" % type(spec).__name__
        )

    @staticmethod
    def _copy_page_setup(src: "Section", dst: "Section") -> None:
        """Copy the page-setup attributes that the context tweaks.

        Used on exit to revert the trailing section to whatever was in
        effect before the context entered.
        """
        for attr in (
            "orientation",
            "page_width",
            "page_height",
            "top_margin",
            "right_margin",
            "bottom_margin",
            "left_margin",
        ):
            try:
                setattr(dst, attr, getattr(src, attr))
            except Exception:  # pragma: no cover - defensive
                pass


def open_section(
    document: "Document",
    *,
    orientation: Union[str, WD_ORIENTATION, None] = None,
    margins: Union[str, dict, tuple, list, None] = None,
    page_size: Union[str, tuple, list, dict, None] = None,
    page_numbering: Optional[dict] = None,
    header: Optional[str] = None,
    footer: Optional[str] = None,
    columns: Union[int, dict, None] = None,
    line_numbering: Union[bool, dict, None] = None,
) -> _SectionContext:
    """Build a :class:`_SectionContext` bound to ``document``.

    Thin factory used by :meth:`Document.section`. Kept module-level so
    the context-manager type stays importable for callers that want to
    type-annotate their own helpers.
    """
    return _SectionContext(
        document,
        orientation=orientation,
        margins=margins,
        page_size=page_size,
        page_numbering=page_numbering,
        header=header,
        footer=footer,
        columns=columns,
        line_numbering=line_numbering,
    )
