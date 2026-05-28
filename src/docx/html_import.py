# pyright: reportPrivateUsage=false

"""Minimal HTML → docx importer.

Stdlib-only HTML parser (``html.parser``) that walks an HTML5 fragment
or document and emits the closest equivalent WordprocessingML
constructs. Intended as the read-side companion to
:mod:`docx.html_export` — preview-grade fidelity, not a full HTML
rendering engine.

Element mapping:

* ``<h1>`` … ``<h6>``     → ``Heading 1`` … ``Heading 6`` paragraphs.
* ``<p>``                 → body paragraph.
* ``<strong>`` / ``<b>``  → bold runs.
* ``<em>`` / ``<i>``      → italic runs.
* ``<u>``                 → underlined runs.
* ``<a href="…">``        → hyperlink (``http`` / ``https`` /
                            ``mailto`` only).
* ``<ul>`` / ``<ol>``     → ``List Bullet`` / ``List Number`` items.
* ``<table>`` / ``<tr>``  → ``Document.add_table`` with one row per
                            ``<tr>`` and one cell per ``<td>``.
* ``<img src="…">``       → embedded picture (``data:`` URLs decoded
                            and embedded; remote URLs degrade to
                            alt-text — they are **not** fetched).
* ``<blockquote>``        → ``Quote`` style paragraph.
* ``<code>`` / ``<pre>``  → monospace runs / paragraphs (``Courier
                            New`` font; ``<pre>`` also preserves
                            whitespace).

When ``clean=True`` (the default) the parser strips ``<script>``,
``<style>``, and HTML comments, drops ``class`` / ``id`` attributes,
and ignores ``style`` attributes other than ``color: <hex>`` (which
is preserved on the run as a colour). With ``clean=False`` the same
elements are still skipped (they would never be useful) but
``class`` / ``id`` are passed through untouched (they are inert in
the docx mapping today regardless).

LaTeX import is **not** in scope for this module — see
:meth:`docx.api._from_html` for the deferred-roadmap entry.

.. versionadded:: 2026.05.14
"""

from __future__ import annotations

import base64
import io
import os
import re
from html.parser import HTMLParser
from typing import IO, TYPE_CHECKING, List, Optional, Union

from docx.shared import RGBColor

if TYPE_CHECKING:
    from docx.document import Document as _DocumentObject
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------


#: Tags whose entire subtree is skipped (open / close pair, content discarded).
_SKIP_TAGS = frozenset({"script", "style", "head", "title"})

#: Void tags (no end tag) that we silently swallow without altering the
#: parser's skip-depth counter — ``<meta>`` and ``<link>`` are common in
#: HTML <head> regions.
_VOID_SKIP_TAGS = frozenset({"meta", "link"})

#: URL schemes accepted on ``<a href>`` and ``<img src>``. Anything
#: else (``javascript:``, ``vbscript:``, ``file:``, ...) is dropped to
#: avoid carrying an XSS payload through into a Word document.
_ALLOWED_URL_SCHEMES = frozenset({"http", "https", "mailto"})


# ---------------------------------------------------------------------------
# Public entry points
# ---------------------------------------------------------------------------


def from_html(
    source: "Union[str, os.PathLike[str], IO[bytes], IO[str]]",
    clean: bool = True,
) -> "_DocumentObject":
    """Build a |Document| from an HTML file or stream.

    `source` may be a path (``str`` / :class:`os.PathLike`) or any
    binary or text file-like object. Encoding for binary input is
    sniffed from a leading ``<meta charset>`` declaration when
    present, falling back to UTF-8.

    See :func:`from_html_string` for the in-memory equivalent and
    :meth:`docx.api._from_html` for the public-API contract attached
    to ``Document.from_html``.

    .. versionadded:: 2026.05.14
    """
    return from_html_string(_read_html_source(source), clean=clean)


def from_html_string(html_text: str, clean: bool = True) -> "_DocumentObject":
    """Build a |Document| from an in-memory HTML string.

    `clean=True` (the default) strips ``<script>`` / ``<style>`` /
    comments and drops ``class`` / ``id`` attributes. ``style``
    attributes are honoured only for ``color`` (best-effort).

    .. versionadded:: 2026.05.14
    """
    # -- imported here to dodge the api → document → html_import cycle. --
    from docx.api import Document as _Document

    document = _Document()
    _strip_initial_empty_paragraph(document)

    builder = _Builder(document)
    parser = _HtmlImportParser(builder, clean=clean)
    parser.feed(html_text)
    parser.close()
    return document


# ---------------------------------------------------------------------------
# Source-resolution helpers
# ---------------------------------------------------------------------------


_CHARSET_PATTERN = re.compile(rb"<meta[^>]+charset\s*=\s*[\"']?([\w-]+)", re.I)


def _read_html_source(
    source: "Union[str, os.PathLike[str], IO[bytes], IO[str]]",
) -> str:
    """Return the HTML text payload at `source`."""
    if isinstance(source, (str, os.PathLike)) and not hasattr(source, "read"):
        with open(os.fspath(source), "rb") as fp:
            return _decode_html_bytes(fp.read())

    raw = source.read()  # type: ignore[union-attr]
    if isinstance(raw, bytes):
        return _decode_html_bytes(raw)
    return raw


def _decode_html_bytes(data: bytes) -> str:
    """Decode `data` honouring a ``<meta charset>`` hint, falling back to UTF-8."""
    match = _CHARSET_PATTERN.search(data[:2048])
    if match:
        try:
            return data.decode(match.group(1).decode("ascii", "ignore"))
        except (LookupError, UnicodeDecodeError):
            pass
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("utf-8", errors="replace")


def _strip_initial_empty_paragraph(document: "_DocumentObject") -> None:
    """Remove the empty body paragraph the default template ships with."""
    paragraphs = list(document.paragraphs)
    if len(paragraphs) == 1 and not paragraphs[0].text:
        body = document._body._body  # pyright: ignore[reportPrivateUsage]
        body.remove(paragraphs[0]._p)  # pyright: ignore[reportPrivateUsage]


# ---------------------------------------------------------------------------
# Run-formatting state tracked while walking inline content
# ---------------------------------------------------------------------------


class _RunFormat:
    """Mutable bag of run-level toggles maintained while parsing inline HTML."""

    __slots__ = ("bold", "italic", "underline", "monospace", "color")

    def __init__(self) -> None:
        self.bold = 0
        self.italic = 0
        self.underline = 0
        self.monospace = 0
        self.color: Optional[RGBColor] = None

    def apply(self, run: "Run") -> None:
        if self.bold:
            run.bold = True
        if self.italic:
            run.italic = True
        if self.underline:
            run.underline = True
        if self.monospace:
            run.font.name = "Courier New"
        if self.color is not None:
            run.font.color.rgb = self.color


# ---------------------------------------------------------------------------
# Pending-table buffer — collects <tr>/<td> events until the table closes
# ---------------------------------------------------------------------------


class _PendingTable:
    """Buffer text content per cell until the closing ``</table>`` arrives.

    ``rows`` is a list of rows; each row is a list of cells; each cell
    is a list of accumulated text fragments. Inline formatting inside
    a cell is *not* preserved by this minimal importer — fidelity
    upgrades belong to a later PR.
    """

    def __init__(self) -> None:
        self.rows: List[List[List[str]]] = []
        self._open_row: Optional[List[List[str]]] = None
        self._open_cell: Optional[List[str]] = None

    def start_row(self) -> None:
        self.end_row()
        self._open_row = []

    def end_row(self) -> None:
        if self._open_row is None:
            return
        self.end_cell()
        self.rows.append(self._open_row)
        self._open_row = None

    def start_cell(self) -> None:
        self.end_cell()
        self._open_cell = []
        if self._open_row is None:
            self._open_row = []
        self._open_row.append(self._open_cell)

    def end_cell(self) -> None:
        self._open_cell = None

    def is_collecting(self) -> bool:
        return self._open_cell is not None

    def append_text(self, text: str) -> None:
        if self._open_cell is None:
            return
        self._open_cell.append(text)


# ---------------------------------------------------------------------------
# Builder — drives the docx side of the import
# ---------------------------------------------------------------------------


class _Builder:
    """Stateful sink fed by :class:`_HtmlImportParser`."""

    def __init__(self, document: "_DocumentObject") -> None:
        self._document = document
        self._format = _RunFormat()

        # -- the paragraph currently accepting runs, or |None| when no
        # -- block is open (next character data starts a fresh ``<p>``). --
        self._current_paragraph: Optional[Paragraph] = None
        # -- when > 0, character data is appended verbatim (preserving
        # -- whitespace) — used inside <pre>. --
        self._pre_depth: int = 0

        # -- list nesting: stack of "ul" | "ol" tokens. --
        self._list_stack: List[str] = []

        # -- hyperlink state. --
        self._link_url: Optional[str] = None
        self._link_text_buf: Optional[List[str]] = None

        # -- pending table (single, top-level — nested tables collapse). --
        self._table: Optional[_PendingTable] = None

    # -- start / end tag handling -----------------------------------------

    def start(self, tag: str, attrs: dict) -> None:
        # -- tables ----------------------------------------------------
        if tag == "table":
            if self._table is None:
                # -- only top-level tables are buffered; inner tables
                # -- collapse into surrounding cell text. --
                self._close_paragraph()
                self._table = _PendingTable()
            return
        if self._table is not None:
            if tag == "tr":
                self._table.start_row()
            elif tag in ("td", "th"):
                self._table.start_cell()
            return  # -- ignore inline tags inside cells; we capture text only --

        # -- inline-formatting -----------------------------------------
        if tag in ("strong", "b"):
            self._format.bold += 1
            return
        if tag in ("em", "i"):
            self._format.italic += 1
            return
        if tag == "u":
            self._format.underline += 1
            return
        if tag == "code":
            self._format.monospace += 1
            return

        # -- block-level -----------------------------------------------
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._close_paragraph()
            self._current_paragraph = self._add_paragraph(
                style=f"Heading {tag[1]}"
            )
            return
        if tag == "p":
            self._close_paragraph()
            self._current_paragraph = self._add_paragraph()
            return
        if tag == "blockquote":
            self._close_paragraph()
            self._current_paragraph = self._add_paragraph(style="Quote")
            return
        if tag == "pre":
            self._close_paragraph()
            self._pre_depth += 1
            self._current_paragraph = self._add_paragraph()
            return
        if tag == "br":
            paragraph = self._ensure_paragraph()
            run = paragraph.add_run()
            self._format.apply(run)
            run.add_break()
            return
        if tag == "hr":
            self._close_paragraph()
            self._add_paragraph()  # -- visual separator, blank line --
            return

        # -- lists -----------------------------------------------------
        if tag in ("ul", "ol"):
            self._close_paragraph()
            self._list_stack.append(tag)
            return
        if tag == "li":
            self._close_paragraph()
            kind = self._list_stack[-1] if self._list_stack else "ul"
            style = "List Bullet" if kind == "ul" else "List Number"
            self._current_paragraph = self._add_paragraph(style=style)
            return

        # -- hyperlinks ------------------------------------------------
        if tag == "a":
            href = (attrs.get("href") or "").strip()
            self._link_url = _sanitize_url(href) if href else None
            self._link_text_buf = []
            return

        # -- inline image ----------------------------------------------
        if tag == "img":
            self._handle_img(attrs)
            return

        # -- inline ``style="color:#..."`` pickup ----------------------
        color = _parse_color_from_style(attrs.get("style"))
        if color is not None:
            self._format.color = color

    def end(self, tag: str) -> None:
        # -- tables ----------------------------------------------------
        if tag == "table":
            self._close_table()
            return
        if self._table is not None:
            if tag == "tr":
                self._table.end_row()
            elif tag in ("td", "th"):
                self._table.end_cell()
            return

        # -- inline-formatting -----------------------------------------
        if tag in ("strong", "b"):
            self._format.bold = max(0, self._format.bold - 1)
            return
        if tag in ("em", "i"):
            self._format.italic = max(0, self._format.italic - 1)
            return
        if tag == "u":
            self._format.underline = max(0, self._format.underline - 1)
            return
        if tag == "code":
            self._format.monospace = max(0, self._format.monospace - 1)
            return

        # -- block-level -----------------------------------------------
        if tag in (
            "h1", "h2", "h3", "h4", "h5", "h6",
            "p", "blockquote", "li", "div",
        ):
            self._close_paragraph()
            return
        if tag == "pre":
            self._close_paragraph()
            self._pre_depth = max(0, self._pre_depth - 1)
            return
        if tag in ("ul", "ol"):
            if self._list_stack:
                self._list_stack.pop()
            return
        if tag == "a":
            self._flush_hyperlink()
            return

    # -- character data ----------------------------------------------------

    def data(self, text: str) -> None:
        if not text:
            return
        if self._table is not None:
            if self._table.is_collecting():
                # -- collapse whitespace in a cell, like the body path --
                if not text.strip():
                    self._table.append_text(" ")
                else:
                    self._table.append_text(re.sub(r"\s+", " ", text))
            return  # -- inter-row / inter-cell whitespace is dropped --

        if self._link_text_buf is not None:
            self._link_text_buf.append(text)
            return

        if self._pre_depth == 0:
            collapsed = re.sub(r"\s+", " ", text)
            if not collapsed.strip() and self._current_paragraph is None:
                # -- pure whitespace between blocks: drop --
                return
            text = collapsed

        paragraph = self._ensure_paragraph()
        run = paragraph.add_run(text)
        self._format.apply(run)

    # -- helpers -----------------------------------------------------------

    def _add_paragraph(self, style: Optional[str] = None) -> "Paragraph":
        try:
            return self._document.add_paragraph(style=style) if style else self._document.add_paragraph()
        except Exception:
            # -- style not registered (Quote / List Bullet on a stripped
            # -- template) — fall back to a plain paragraph. --
            return self._document.add_paragraph()

    def _ensure_paragraph(self) -> "Paragraph":
        if self._current_paragraph is None:
            self._current_paragraph = self._add_paragraph()
        return self._current_paragraph

    def _close_paragraph(self) -> None:
        self._current_paragraph = None

    def _flush_hyperlink(self) -> None:
        url = self._link_url
        text = "".join(self._link_text_buf or [])
        self._link_url = None
        self._link_text_buf = None
        if not text:
            return
        if not url:
            # -- bare anchor / unsafe scheme: degrade to plain run --
            paragraph = self._ensure_paragraph()
            run = paragraph.add_run(text)
            self._format.apply(run)
            return
        paragraph = self._ensure_paragraph()
        # -- pass `style=None` so an unregistered ``Hyperlink`` character
        # -- style does not fail the call. The default-template
        # -- stylesheet does not register a Hyperlink character style. --
        try:
            paragraph.add_hyperlink(url=url, text=text, style=None)
        except Exception:
            run = paragraph.add_run(f"{text} ({url})")
            self._format.apply(run)

    # -- images ------------------------------------------------------------

    def _handle_img(self, attrs: dict) -> None:
        src = (attrs.get("src") or "").strip()
        alt = attrs.get("alt") or ""
        if not src:
            return

        paragraph = self._ensure_paragraph()
        if src.startswith("data:"):
            stream = _decode_data_url(src)
            if stream is None:
                if alt:
                    paragraph.add_run(alt)
                return
            try:
                paragraph.add_run().add_picture(stream)
            except Exception:
                if alt:
                    paragraph.add_run(alt)
            return

        # -- non-data URLs are not fetched; we degrade to alt-text plus
        # -- the URL so the document is still self-describing. --
        paragraph.add_run(alt or src)

    # -- tables ------------------------------------------------------------

    def _close_table(self) -> None:
        if self._table is None:
            return
        pending = self._table
        self._table = None
        if not pending.rows or not any(pending.rows):
            return

        cols = max(len(row) for row in pending.rows)
        if cols == 0:
            return

        try:
            table = self._document.add_table(rows=len(pending.rows), cols=cols)
        except Exception:
            return

        for r_idx, row in enumerate(pending.rows):
            for c_idx, cell_text_parts in enumerate(row):
                text = "".join(cell_text_parts).strip()
                if not text:
                    continue
                cell = table.rows[r_idx].cells[c_idx]
                cell.paragraphs[0].add_run(text)


# ---------------------------------------------------------------------------
# HTML parser — turns SAX-ish events into builder calls
# ---------------------------------------------------------------------------


class _HtmlImportParser(HTMLParser):
    """Stdlib :class:`html.parser.HTMLParser` driving a :class:`_Builder`."""

    def __init__(self, builder: _Builder, clean: bool = True) -> None:
        super().__init__(convert_charrefs=True)
        self._builder = builder
        self._clean = clean
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs):  # type: ignore[override]
        tag = tag.lower()
        if tag in _VOID_SKIP_TAGS:
            return
        if tag in _SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        self._builder.start(tag, self._normalise_attrs(attrs))

    def handle_startendtag(self, tag: str, attrs):  # type: ignore[override]
        tag = tag.lower()
        if tag in _VOID_SKIP_TAGS or tag in _SKIP_TAGS or self._skip_depth:
            return
        attr_dict = self._normalise_attrs(attrs)
        self._builder.start(tag, attr_dict)
        self._builder.end(tag)

    def handle_endtag(self, tag: str):  # type: ignore[override]
        tag = tag.lower()
        if tag in _VOID_SKIP_TAGS:
            return
        if tag in _SKIP_TAGS:
            if self._skip_depth:
                self._skip_depth -= 1
            return
        if self._skip_depth:
            return
        self._builder.end(tag)

    def handle_data(self, data: str):  # type: ignore[override]
        if self._skip_depth:
            return
        self._builder.data(data)

    def handle_comment(self, data: str):  # type: ignore[override]
        # -- always strip comments. --
        return

    def _normalise_attrs(self, attrs) -> dict:
        out: dict = {}
        for name, value in attrs:
            if name is None:
                continue
            lname = name.lower()
            if self._clean and lname in ("class", "id"):
                continue
            out[lname] = "" if value is None else value
        return out


# ---------------------------------------------------------------------------
# Inline-style + URL helpers
# ---------------------------------------------------------------------------


_DATA_URL_PATTERN = re.compile(
    r"^data:(?P<mime>[^;,]+)(?:;(?P<enc>base64))?,(?P<payload>.*)$",
    re.DOTALL,
)


def _decode_data_url(url: str) -> "Optional[io.BytesIO]":
    """Decode a ``data:`` URL into a binary stream, or |None| on failure."""
    match = _DATA_URL_PATTERN.match(url)
    if not match:
        return None
    payload = match.group("payload")
    encoding = match.group("enc")
    try:
        if encoding == "base64":
            blob = base64.b64decode(payload, validate=False)
        else:
            blob = payload.encode("latin-1", errors="replace")
    except Exception:
        return None
    return io.BytesIO(blob)


_HEX_COLOR_PATTERN = re.compile(r"#?([0-9a-fA-F]{6})")


def _parse_color_from_style(style: Optional[str]) -> Optional[RGBColor]:
    """Pull a ``color: #aabbcc`` value out of an HTML inline style."""
    if not style:
        return None
    for prop in style.split(";"):
        prop = prop.strip()
        if not prop.lower().startswith("color"):
            continue
        _, _, value = prop.partition(":")
        match = _HEX_COLOR_PATTERN.search(value)
        if not match:
            continue
        hex6 = match.group(1)
        try:
            return RGBColor(
                int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
            )
        except Exception:
            return None
    return None


def _sanitize_url(url: str) -> Optional[str]:
    """Return `url` if it's safe to embed in a hyperlink, else |None|.

    Mirrors the allow-list used by :func:`docx.html_export._sanitize_href`
    so a round-trip ``from_html`` → ``to_html`` does not promote an
    attacker-controlled scheme into the output.
    """
    if not url:
        return None
    if url.startswith("#"):
        return url
    colon = url.find(":")
    if colon == -1:
        return url
    slash = url.find("/")
    qmark = url.find("?")
    hmark = url.find("#")
    for sep in (slash, qmark, hmark):
        if 0 <= sep < colon:
            return url
    if url[:colon].lower() in _ALLOWED_URL_SCHEMES:
        return url
    return None
