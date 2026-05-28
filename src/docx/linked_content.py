"""Cross-format linked-content support for python-docx.

A "linked target" is an external resource whose value is re-resolved at
document open time. Word natively models this via the ``INCLUDETEXT``
field — a complex field whose instruction names a file (and an optional
sub-target) and whose cached result is what readers without the linked
file see. python-docx exposes the pattern through three pieces:

* :meth:`docx.text.paragraph.Paragraph.link_to` — append an
  ``INCLUDETEXT`` field with a target URL.
* :meth:`docx.document.Document.linked_targets` — iterate every
  link record in the document body.
* :meth:`docx.document.Document.update_links` — best-effort re-resolve
  every link, replacing the cached field result with the current value
  fetched from the linked file.

Three target shapes are supported, all carried in the URL fragment:

* ``revenue.xlsx#RevenueQ1!B5`` — Excel cell reference (sheet name
  ``RevenueQ1``, cell ``B5``).
* ``revenue.xlsx#RevenueQ1[Total]`` — Excel structured-table column
  total (table named ``RevenueQ1``, column ``Total``).
* ``summary.pptx#slide-3`` — PowerPoint slide reference (1-indexed
  slide number).

Resolution of Excel cells is handled by the sibling ``xlsx`` package
when present; PowerPoint slides resolve to a placeholder string
(real rendering requires PowerPoint) and unknown shapes round-trip
unchanged. None of the resolution paths raise — a missing target,
parse error, or absent sibling library returns the cached
:attr:`docx.fields.Field.result_text` so the document is never
silently corrupted by a refresh.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import os
import re
from typing import TYPE_CHECKING, Iterator, Optional, Tuple, Union

if TYPE_CHECKING:
    from docx.document import Document
    from docx.fields import Field


# -- constants ----------------------------------------------------------

#: The Word field-type token used for cross-format linked content.
INCLUDETEXT = "INCLUDETEXT"

#: Sentinel inserted as the cached result when the link target cannot be
#: resolved at write time (e.g. ``link_to("missing.xlsx#Sheet1!A1")``
#: with no live workbook). Word displays this until the next refresh.
UNRESOLVED_PLACEHOLDER = "<<linked content not yet resolved>>"

#: Recognised target kinds, returned by :attr:`LinkedTarget.kind`.
LINK_KIND_XLSX_CELL = "xlsx-cell"
LINK_KIND_XLSX_TABLE_COLUMN = "xlsx-table-column"
LINK_KIND_PPTX_SLIDE = "pptx-slide"
LINK_KIND_UNKNOWN = "unknown"

# -- file extensions we recognise as Excel / PowerPoint sources --
_XLSX_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")
_PPTX_SUFFIXES = (".pptx", ".pptm", ".potx", ".potm")

# -- "RevenueQ1!B5" — sheet name (optionally quoted) plus a cell --
_CELL_FRAGMENT_RE = re.compile(
    r"^(?:'(?P<sheet_quoted>[^']+)'|(?P<sheet>[^!\[\]]+))!"
    r"(?P<cell>\$?[A-Za-z]{1,3}\$?\d{1,7})$"
)

# -- "RevenueQ1[Total]" — table name plus column name (column may be
#    bare or quoted; spaces and unicode allowed inside the brackets) --
_TABLE_COLUMN_FRAGMENT_RE = re.compile(
    r"^(?P<table>[^\[\]!]+)\[(?P<column>[^\[\]]+)\]$"
)

# -- "slide-3" — case-insensitive on the literal "slide" --
_SLIDE_FRAGMENT_RE = re.compile(r"^slide[-_]?(?P<index>\d+)$", re.IGNORECASE)


# -- target descriptors -------------------------------------------------


class ParsedLinkTarget:
    """Structured view of a parsed ``link_to(target_url)`` URL.

    Carries the four pieces a resolver needs:

    * :attr:`url` — the original ``target_url`` argument, verbatim.
    * :attr:`kind` — one of :data:`LINK_KIND_XLSX_CELL`,
      :data:`LINK_KIND_XLSX_TABLE_COLUMN`, :data:`LINK_KIND_PPTX_SLIDE`,
      or :data:`LINK_KIND_UNKNOWN`.
    * :attr:`path` — the file portion of the URL (everything before
      ``#``); empty string when the URL has no file portion.
    * :attr:`fragment` — the URL fragment (everything after the first
      ``#``), or the empty string.
    * :attr:`detail` — a kind-specific tuple of strings:

      * ``xlsx-cell`` → ``(sheet_name, cell_ref)``
      * ``xlsx-table-column`` → ``(table_name, column_name)``
      * ``pptx-slide`` → ``(slide_index_str,)``  (1-indexed)
      * ``unknown`` → ``()``

    .. versionadded:: 2026.05.13
    """

    __slots__ = ("url", "kind", "path", "fragment", "detail")

    def __init__(
        self,
        url: str,
        kind: str,
        path: str,
        fragment: str,
        detail: "Tuple[str, ...]",
    ):
        self.url = url
        self.kind = kind
        self.path = path
        self.fragment = fragment
        self.detail = detail

    def __eq__(self, other: object) -> bool:  # pragma: no cover - trivial
        if not isinstance(other, ParsedLinkTarget):
            return NotImplemented
        return (
            self.url == other.url
            and self.kind == other.kind
            and self.path == other.path
            and self.fragment == other.fragment
            and self.detail == other.detail
        )

    def __repr__(self) -> str:  # pragma: no cover - trivial
        return (
            f"ParsedLinkTarget(url={self.url!r}, kind={self.kind!r}, "
            f"path={self.path!r}, fragment={self.fragment!r}, "
            f"detail={self.detail!r})"
        )


def parse_link_target(target_url: str) -> ParsedLinkTarget:
    """Return a :class:`ParsedLinkTarget` describing `target_url`.

    The grammar is the URL-with-fragment form documented at the module
    docstring. Path and fragment are split on the first ``#``. The
    fragment is matched against three regexes in order — Excel cell,
    Excel table-column, PowerPoint slide; the first match wins. When
    nothing matches, :attr:`ParsedLinkTarget.kind` is set to
    :data:`LINK_KIND_UNKNOWN` and :attr:`detail` is an empty tuple, so
    callers can still round-trip the URL even when its shape is
    foreign.

    .. versionadded:: 2026.05.13
    """
    if not isinstance(target_url, str) or not target_url:
        return ParsedLinkTarget(
            url=target_url or "",
            kind=LINK_KIND_UNKNOWN,
            path="",
            fragment="",
            detail=(),
        )

    if "#" in target_url:
        path, _, fragment = target_url.partition("#")
    else:
        path, fragment = target_url, ""

    suffix = os.path.splitext(path)[1].lower()
    is_xlsx = suffix in _XLSX_SUFFIXES
    is_pptx = suffix in _PPTX_SUFFIXES

    # -- xlsx cell: sheet!cell --
    if is_xlsx and fragment:
        cell_match = _CELL_FRAGMENT_RE.match(fragment)
        if cell_match is not None:
            sheet = cell_match.group("sheet_quoted") or cell_match.group("sheet")
            cell = cell_match.group("cell")
            return ParsedLinkTarget(
                url=target_url,
                kind=LINK_KIND_XLSX_CELL,
                path=path,
                fragment=fragment,
                detail=(sheet, cell),
            )
        col_match = _TABLE_COLUMN_FRAGMENT_RE.match(fragment)
        if col_match is not None:
            return ParsedLinkTarget(
                url=target_url,
                kind=LINK_KIND_XLSX_TABLE_COLUMN,
                path=path,
                fragment=fragment,
                detail=(col_match.group("table"), col_match.group("column")),
            )

    # -- pptx slide: slide-N --
    if is_pptx and fragment:
        slide_match = _SLIDE_FRAGMENT_RE.match(fragment)
        if slide_match is not None:
            return ParsedLinkTarget(
                url=target_url,
                kind=LINK_KIND_PPTX_SLIDE,
                path=path,
                fragment=fragment,
                detail=(slide_match.group("index"),),
            )

    return ParsedLinkTarget(
        url=target_url,
        kind=LINK_KIND_UNKNOWN,
        path=path,
        fragment=fragment,
        detail=(),
    )


# -- field-instruction builder / parser ---------------------------------


def build_includetext_instruction(target_url: str) -> str:
    """Return the ``INCLUDETEXT`` field-instruction string for `target_url`.

    The emitted form is ``INCLUDETEXT "<target_url>"`` — the URL is
    always wrapped in double quotes, mirroring the shape Word writes
    when the user picks *Insert ▸ Quick Parts ▸ Field ▸ IncludeText*.
    Internal double quotes inside `target_url` are escaped by doubling,
    matching Word's escape rule for ``INCLUDETEXT`` arguments. This
    keeps round-trip safety for unusual URLs without losing the
    surrounding-quote disambiguation that Word's parser relies on.

    .. versionadded:: 2026.05.13
    """
    if not isinstance(target_url, str):
        raise TypeError(
            "target_url must be a string, got %r" % type(target_url).__name__
        )
    if not target_url:
        raise ValueError("target_url must be a non-empty string")
    escaped = target_url.replace('"', '""')
    return f'{INCLUDETEXT} "{escaped}"'


def parse_includetext_instruction(instruction: str) -> Optional[str]:
    """Return the URL argument from an ``INCLUDETEXT`` instruction.

    Strips the ``INCLUDETEXT`` token, unwraps the surrounding double
    quotes, undoes the ``""`` → ``"`` escape applied by
    :func:`build_includetext_instruction`. Returns |None| when the
    instruction is not an ``INCLUDETEXT`` field or has no URL argument.

    Switches (``\\* MERGEFORMAT``, ``\\!``, etc.) following the URL are
    ignored — only the URL is returned.

    .. versionadded:: 2026.05.13
    """
    stripped = instruction.strip()
    if not stripped.upper().startswith(INCLUDETEXT):
        return None
    remainder = stripped[len(INCLUDETEXT) :].strip()
    if not remainder:
        return None
    if remainder.startswith('"'):
        # -- find matching close quote, honouring "" as an escape --
        i = 1
        out: list[str] = []
        n = len(remainder)
        while i < n:
            ch = remainder[i]
            if ch == '"':
                # -- doubled quote? --
                if i + 1 < n and remainder[i + 1] == '"':
                    out.append('"')
                    i += 2
                    continue
                # -- end of quoted argument --
                return "".join(out)
            out.append(ch)
            i += 1
        # -- unterminated quote: return what we collected --
        return "".join(out)
    # -- bare token form: take everything up to the first whitespace or
    #    backslash switch --
    parts = remainder.split(None, 1)
    return parts[0] if parts else None


# -- proxy --------------------------------------------------------------


class LinkedTarget:
    """A linked external resource referenced from a paragraph.

    Returned by :meth:`docx.text.paragraph.Paragraph.link_to` and
    yielded by :attr:`docx.document.Document.linked_targets`. Wraps the
    underlying :class:`docx.fields.Field` proxy so callers can access
    the field's cached :attr:`~docx.fields.Field.result_text`,
    :attr:`~docx.fields.Field.is_dirty`, and the raw
    :attr:`~docx.fields.Field.instruction`. Adds linking-specific
    accessors:

    * :attr:`url` — the ``target_url`` originally passed to
      ``link_to`` (the unwrapped INCLUDETEXT argument).
    * :attr:`kind` — one of :data:`LINK_KIND_XLSX_CELL`,
      :data:`LINK_KIND_XLSX_TABLE_COLUMN`,
      :data:`LINK_KIND_PPTX_SLIDE`, or :data:`LINK_KIND_UNKNOWN`.
    * :attr:`parsed` — the :class:`ParsedLinkTarget` view of
      :attr:`url`.
    * :meth:`resolve` — best-effort fetch of the linked value, scoped
      to a `base_dir` for relative paths.
    * :meth:`refresh` — call :meth:`resolve` and write the result back
      into the field's cached result via
      :meth:`~docx.fields.Field.update_result_text`.

    The proxy is read-only with respect to the URL — to repoint a link
    to a new target, remove the field and re-call ``link_to``.

    .. versionadded:: 2026.05.13
    """

    def __init__(self, field: "Field"):
        self._field = field

    # -- public read-only views -------------------------------------------

    @property
    def field(self) -> "Field":
        """The underlying :class:`docx.fields.Field` for this link.

        Useful for reading the cached result, marking the field dirty,
        or inspecting the raw instruction.

        .. versionadded:: 2026.05.13
        """
        return self._field

    @property
    def url(self) -> str:
        """The link target URL as originally passed to ``link_to``.

        Returns the empty string when the underlying field is not an
        ``INCLUDETEXT`` field or the URL argument is missing.

        .. versionadded:: 2026.05.13
        """
        url = parse_includetext_instruction(self._field.instruction)
        return url or ""

    @property
    def parsed(self) -> ParsedLinkTarget:
        """A :class:`ParsedLinkTarget` view of :attr:`url`.

        .. versionadded:: 2026.05.13
        """
        return parse_link_target(self.url)

    @property
    def kind(self) -> str:
        """One of the :data:`LINK_KIND_*` constants.

        Shorthand for ``self.parsed.kind``.

        .. versionadded:: 2026.05.13
        """
        return self.parsed.kind

    @property
    def path(self) -> str:
        """The file portion of :attr:`url` (everything before ``#``).

        .. versionadded:: 2026.05.13
        """
        return self.parsed.path

    @property
    def fragment(self) -> str:
        """The fragment portion of :attr:`url` (everything after ``#``).

        .. versionadded:: 2026.05.13
        """
        return self.parsed.fragment

    @property
    def cached_text(self) -> str:
        """The cached result text Word displays for this link.

        Alias for ``self.field.result_text``. Empty string when no
        cached result has been written.

        .. versionadded:: 2026.05.13
        """
        return self._field.result_text

    # -- resolution -------------------------------------------------------

    def resolve(self, base_dir: Optional[str] = None) -> Optional[str]:
        """Best-effort fetch of the live value at :attr:`url`.

        For an Excel cell or table-column target, the sibling ``xlsx``
        package is loaded and the workbook at :attr:`path` is opened
        (relative paths are resolved against `base_dir`, falling back
        to the current working directory). The cell value is returned
        as a string. For a PowerPoint slide target, a placeholder of
        the form ``"[Slide N]"`` is returned (real rendering requires
        PowerPoint or LibreOffice and is intentionally out of scope).

        Returns |None| in any of these cases — the call site can
        distinguish them from a successfully-resolved empty string by
        checking for |None|:

        * :attr:`kind` is :data:`LINK_KIND_UNKNOWN`
        * The target file does not exist on disk
        * The sibling resolver library is not installed
        * Any unexpected error from the resolver

        Never raises — failed resolution always returns |None| so
        :meth:`refresh` (and hence :meth:`Document.update_links`) can
        run without aborting on the first broken link.

        .. versionadded:: 2026.05.13
        """
        parsed = self.parsed
        if parsed.kind == LINK_KIND_UNKNOWN:
            return None
        path = _resolve_link_path(parsed.path, base_dir)
        if path is None:
            return None

        if parsed.kind == LINK_KIND_XLSX_CELL:
            return _resolve_xlsx_cell(path, parsed.detail)
        if parsed.kind == LINK_KIND_XLSX_TABLE_COLUMN:
            return _resolve_xlsx_table_column(path, parsed.detail)
        if parsed.kind == LINK_KIND_PPTX_SLIDE:
            return _resolve_pptx_slide(path, parsed.detail)
        return None  # pragma: no cover - defensive

    def refresh(self, base_dir: Optional[str] = None) -> Optional[str]:
        """Resolve the link and write the result back into the field cache.

        Returns the resolved value (the new cached text) when
        resolution succeeded, or |None| when it didn't — in which case
        the existing :attr:`cached_text` is left untouched. Mark-dirty
        state is preserved so Word still re-evaluates the link the
        next time it's opened.

        .. versionadded:: 2026.05.13
        """
        resolved = self.resolve(base_dir=base_dir)
        if resolved is None:
            return None
        self._field.update_result_text(resolved)
        return resolved


# -- resolution helpers (intentionally lazy on the sibling import) ------


def _resolve_link_path(path: str, base_dir: Optional[str]) -> Optional[str]:
    """Return an existing absolute path for `path`, or |None`.

    Relative paths are resolved against `base_dir` first, then against
    the current working directory as a fallback. Returns |None| when
    no candidate exists on disk — the caller treats that as
    "unresolvable" and leaves the cached field result alone.
    """
    if not path:
        return None
    candidates: list[str] = []
    if os.path.isabs(path):
        candidates.append(path)
    else:
        if base_dir:
            candidates.append(os.path.join(base_dir, path))
        candidates.append(os.path.abspath(path))
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return None


def _resolve_xlsx_cell(
    path: str, detail: "Tuple[str, ...]"
) -> Optional[str]:
    """Open the workbook at `path` and return ``str(sheet[cell].value)``.

    `detail` is ``(sheet_name, cell_ref)``. Returns |None| on any
    failure — missing sheet, missing cell, sibling ``xlsx`` package not
    installed, parse errors, or any unexpected exception.
    """
    if len(detail) != 2:
        return None
    sheet_name, cell_ref = detail
    try:
        from xlsx import load_workbook  # type: ignore[import-not-found]
    except Exception:
        return None
    try:
        workbook = load_workbook(path, data_only=True)
    except Exception:
        return None
    try:
        try:
            sheet = workbook[sheet_name]
        except Exception:
            return None
        try:
            cell = sheet[cell_ref]
        except Exception:
            return None
        value = getattr(cell, "value", None)
        if value is None:
            return ""
        return str(value)
    finally:
        # -- xlsx workbooks expose `close()` to release any underlying
        #    file handles; best-effort cleanup so resource warnings
        #    don't leak into pytest --
        close = getattr(workbook, "close", None)
        if callable(close):
            try:
                close()
            except Exception:
                pass


def _resolve_xlsx_table_column(
    path: str, detail: "Tuple[str, ...]"
) -> Optional[str]:
    """Return the totals-row value of `column` in table `table`, as a string.

    Walks every worksheet in the workbook at `path`, looking for a
    table with the given name, then locates the named column on the
    table. Returns the totals-row cell value when present; otherwise
    falls back to the last data row's value in that column. Returns
    |None| on any failure.
    """
    if len(detail) != 2:
        return None
    table_name, column_name = detail
    try:
        from xlsx import load_workbook  # type: ignore[import-not-found]
        from xlsx.utils.cell import (  # type: ignore[import-not-found]
            range_boundaries,
            get_column_letter,
        )
    except Exception:
        return None
    try:
        workbook = load_workbook(path, data_only=True)
    except Exception:
        return None
    try:
        for ws in workbook.worksheets:
            tables = getattr(ws, "tables", {})
            try:
                table = tables.get(table_name)
            except AttributeError:
                # -- older xlsx versions: tables is a list-like
                table = None
                for candidate in tables:
                    if getattr(candidate, "name", None) == table_name:
                        table = candidate
                        break
            if table is None:
                continue
            columns = getattr(table, "tableColumns", None)
            if not columns:
                continue
            col_index: Optional[int] = None
            for i, tc in enumerate(columns):
                if getattr(tc, "name", None) == column_name:
                    col_index = i
                    break
            if col_index is None:
                return None
            ref = getattr(table, "ref", None)
            if not ref:
                return None
            try:
                min_col, min_row, max_col, max_row = range_boundaries(ref)
            except Exception:
                return None
            target_col_letter = get_column_letter(min_col + col_index)
            # -- totals row is the bottom row when totalsRowCount > 0 --
            totals_row_count = getattr(table, "totalsRowCount", 0) or 0
            if totals_row_count > 0:
                target_row = max_row
            else:
                # -- last data row (header is row 0) --
                target_row = max_row
            try:
                cell = ws[f"{target_col_letter}{target_row}"]
            except Exception:
                return None
            value = getattr(cell, "value", None)
            if value is None:
                return ""
            return str(value)
        return None
    finally:
        close = getattr(workbook, "close", None)
        if callable(close):
            try:
                close()
            except Exception:
                pass


def _resolve_pptx_slide(
    path: str, detail: "Tuple[str, ...]"
) -> Optional[str]:
    """Return a placeholder string for a PowerPoint slide reference.

    python-docx has no rendering pipeline for PowerPoint slides, so
    this resolver intentionally returns a short ``"[Slide N: <title>]"``
    summary instead of the rendered bitmap Word would normally embed.
    The intent is twofold:

    * The cached result is human-readable so a Word reader without
      PowerPoint sees something more useful than the URL.
    * Word still re-resolves the link on open — when the user has
      PowerPoint installed, Word fetches a real preview; the
      placeholder is overwritten on the first refresh.

    Returns |None| when the sibling ``pptx`` package is not installed
    *and* the slide index can't be parsed, so the caller can leave the
    cached text alone.
    """
    if not detail:
        return None
    index_str = detail[0]
    try:
        slide_index = int(index_str)
    except (TypeError, ValueError):
        return None
    # -- best-effort: try to read the slide title via sibling pptx --
    title_text: Optional[str] = None
    try:
        from pptx import Presentation  # type: ignore[import-not-found]
    except Exception:
        return f"[Slide {slide_index}]"
    try:
        prs = Presentation(path)
    except Exception:
        return f"[Slide {slide_index}]"
    try:
        slides = list(prs.slides)
        # -- 1-indexed in the URL, 0-indexed in pptx --
        if 1 <= slide_index <= len(slides):
            slide = slides[slide_index - 1]
            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue
                tf = getattr(shape, "text_frame", None)
                text = getattr(tf, "text", "") if tf is not None else ""
                if text and text.strip():
                    title_text = text.strip().splitlines()[0]
                    break
    except Exception:
        title_text = None
    if title_text:
        return f"[Slide {slide_index}: {title_text}]"
    return f"[Slide {slide_index}]"


# -- iteration over a document ------------------------------------------


def iter_linked_targets(document: "Document") -> Iterator[LinkedTarget]:
    """Yield every :class:`LinkedTarget` in `document`'s body.

    Walks :attr:`docx.document.Document.fields` and yields one
    :class:`LinkedTarget` per ``INCLUDETEXT`` field. The order matches
    document order. Used internally by
    :attr:`Document.linked_targets`; exposed at module scope for
    callers that already hold a list of fields.

    .. versionadded:: 2026.05.13
    """
    for field in document.fields:
        if field.type.upper() == INCLUDETEXT:
            yield LinkedTarget(field)


def update_document_links(
    document: "Document",
    base_dir: Optional[str] = None,
) -> int:
    """Refresh every link in `document`. Returns the count of updates.

    Calls :meth:`LinkedTarget.refresh` on every record yielded by
    :func:`iter_linked_targets`. Returns the number of fields whose
    cached text was actually rewritten — failed resolutions are
    silently skipped (the field's existing cached text is left in
    place, which is the correct round-trip behaviour for an
    unreachable target).

    `base_dir` is forwarded to :meth:`LinkedTarget.refresh` to scope
    relative paths in the link URLs. When |None|, relative paths fall
    back to the process working directory.

    .. versionadded:: 2026.05.13
    """
    count = 0
    for link in iter_linked_targets(document):
        result = link.refresh(base_dir=base_dir)
        if result is not None:
            count += 1
    return count


# -- public re-exports --------------------------------------------------

__all__ = [
    "INCLUDETEXT",
    "LINK_KIND_PPTX_SLIDE",
    "LINK_KIND_UNKNOWN",
    "LINK_KIND_XLSX_CELL",
    "LINK_KIND_XLSX_TABLE_COLUMN",
    "LinkedTarget",
    "ParsedLinkTarget",
    "UNRESOLVED_PLACEHOLDER",
    "build_includetext_instruction",
    "iter_linked_targets",
    "parse_includetext_instruction",
    "parse_link_target",
    "update_document_links",
]
