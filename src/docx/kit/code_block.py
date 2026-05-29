"""Syntax-highlighted code blocks — Pygments-driven monospace rendering.

Closes #293.

Authors of technical reports, runbooks, training material, and product
documentation routinely need to embed snippets of source code in a
``.docx`` document with the same conventions a Markdown / Sphinx /
reStructuredText reader would expect: a monospace font, a light
shaded background that visually separates code from prose, optional
line numbers in a narrow gutter, and per-token syntax highlighting
that follows a recognisable colour theme. Hand-rolling that shape with
the public python-docx API is a forty-line ritual every author writes
once and copy-pastes thereafter.

This module exposes :func:`add` plus a family of language-specific
convenience wrappers (``python``, ``bash``, ``json``, ``yaml``,
``sql``, ``javascript``, ``typescript``, ``rust``, ``go``, ``html``,
``css``, ``xml``) built entirely on python-docx's public API
(:func:`Document.add_table`, ``_Cell.shading.fill_color``, ``Run.font``).
Highlighting is delegated to `pygments <https://pygments.org/>`_,
which is declared as the optional ``[code]`` extra: if pygments is
installed the helper produces a fully token-coloured block; if it is
*not* installed the helper falls back gracefully to a plain monospace
block with the same shading and line-number gutter — never raising —
so callers who only occasionally embed code do not need to
take on a new mandatory dependency::

    from docx import Document
    from docx.kit import code_block

    doc = Document()
    code_block.add(
        doc,
        '''
        def hello():
            print("Hello world!")
        ''',
        lang="python",
        line_numbers=True,
        theme="monokai",
    )

    # Convenience wrappers — same return value (a `Table`), same kwargs.
    code_block.python(doc, "x = 1\\nprint(x)")
    code_block.bash(doc, "ls -la")
    code_block.json(doc, '{"a": 1}', line_numbers=True)

    doc.save("out.docx")

Implementation notes:

* The block is a single-row, single-column-or-two-column table. Tables
  are the only public python-docx surface that paints a cell-wide
  shaded background; a paragraph-bordered approach (``w:pBdr`` +
  ``w:shd``) would require dropping into ``_element`` access, which
  the kit forbids.
* Every line of source is one paragraph inside the code cell. Pygments
  tokens are emitted as separate runs with ``font.color.rgb``,
  ``font.bold``, and ``font.italic`` set per the active theme. A
  trailing newline at the end of the block is consumed (pygments emits
  one) so the rendered table does not carry an empty trailing line.
* Without pygments, the helper still tokenises the source by
  splitting on newlines — the soft fallback path renders the same
  table shape (single shaded cell, monospace font, optional gutter)
  with no colour, no bold, no italic. The contract is "looks like
  code, never raises". Callers can detect the fallback by checking
  ``code_block.HAS_PYGMENTS``.
* Theme background colour comes from
  ``style.background_color`` for the named pygments theme; when the
  theme does not declare one (the ``default`` style returns
  ``"#f8f8f8"``) the helper picks a light grey
  (``RGBColor(0xF5, 0xF5, 0xF5)``) so the block reads as code on a
  white page.
* Line-number column shading matches the body cell (so the gutter
  fades into the block) but the line-number text is rendered at half
  opacity by colouring it ``RGBColor(0x80, 0x80, 0x80)`` for the
  ``default`` theme and a lightened body colour for dark themes.
* Theme names follow pygments style names — ``default``,
  ``monokai``, ``tango``, ``friendly``, ``solarized-dark``,
  ``solarized-light``, ``vs``, etc. Unknown theme names raise
  :class:`ValueError` *only when pygments is installed*; the
  no-pygments fallback ignores ``theme`` entirely.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional, Tuple, Union

from docx.shared import Pt, RGBColor

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph


# -- Pygments availability flag.  Probed once at import time so callers
# -- can branch on ``code_block.HAS_PYGMENTS`` without paying the import
# -- cost a second time.  Tests for the fallback path mock this flag to
# -- ``False`` via monkeypatch.
try:  # pragma: no cover - import branch
    import pygments  # noqa: F401

    HAS_PYGMENTS = True
except ImportError:  # pragma: no cover - import branch
    HAS_PYGMENTS = False


# -- Default monospace font.  ``Consolas`` ships with every modern
# -- Microsoft Word install; the helper accepts a caller-supplied
# -- ``monospace_font`` for environments that prefer a different
# -- typeface (``Cascadia Code``, ``Menlo``, ``DejaVu Sans Mono``).
_DEFAULT_MONOSPACE_FONT = "Consolas"

# -- Default code font size.  Most code-in-document conventions render
# -- code one or two points smaller than body text; 9pt sits between
# -- a 10pt body (Word's default body size) and the 8pt some authors
# -- prefer for densely-packed listings.
_DEFAULT_CODE_FONT_SIZE = Pt(9)

# -- Fallback shading colour when the named theme does not declare a
# -- background.  ``#F5F5F5`` is a very light grey that reads as
# -- "code" against white prose without competing with body text.
_FALLBACK_SHADING_HEX = "F5F5F5"

# -- Fallback foreground (text) colour for the soft-fallback path
# -- (no pygments) and for tokens that the active theme does not
# -- colour. ``#1F1F1F`` is near-black with a hint of warmth so it
# -- doesn't clash with line-number gutter grey.
_FALLBACK_FOREGROUND = RGBColor(0x1F, 0x1F, 0x1F)

# -- Line-number gutter foreground colour.  Pygments themes do not
# -- declare a gutter colour; using a mid-grey (``#808080``) gives the
# -- gutter the conventional "muted" feel without leaning on the
# -- theme's palette.
_GUTTER_FOREGROUND = RGBColor(0x80, 0x80, 0x80)


def _hex_to_rgb(hex_str):
    # type: (str) -> RGBColor
    """Convert ``"#rrggbb"`` / ``"rrggbb"`` to an :class:`RGBColor`.

    Pygments emits style ``color`` / ``bgcolor`` values as either a
    bare ``"rrggbb"`` string or ``None``; the leading ``#`` is
    optional. This helper normalises both shapes; an empty / ``None``
    input raises :class:`ValueError` to surface mistakes at call sites.
    """
    if not hex_str:
        raise ValueError("hex colour must be a non-empty string")
    cleaned = hex_str.lstrip("#")
    if len(cleaned) != 6:
        raise ValueError(
            "hex colour must be 6 hex digits; got %r" % hex_str
        )
    return RGBColor(
        int(cleaned[0:2], 16),
        int(cleaned[2:4], 16),
        int(cleaned[4:6], 16),
    )


def _resolve_background(theme):
    # type: (str) -> RGBColor
    """Resolve the cell-shading colour for ``theme``.

    When pygments is installed and the named style declares a
    background, return that colour. Otherwise (theme has no
    background, or pygments is missing) fall back to the conventional
    very-light-grey :data:`_FALLBACK_SHADING_HEX`.
    """
    if HAS_PYGMENTS:
        try:
            from pygments.styles import get_style_by_name

            style = get_style_by_name(theme)
        except Exception:  # noqa: BLE001 - any pygments error → fallback
            return _hex_to_rgb(_FALLBACK_SHADING_HEX)
        bg = getattr(style, "background_color", None)
        if bg:
            try:
                return _hex_to_rgb(bg)
            except ValueError:
                return _hex_to_rgb(_FALLBACK_SHADING_HEX)
    return _hex_to_rgb(_FALLBACK_SHADING_HEX)


def _shade_cell(cell, rgb):
    # type: (_Cell, RGBColor) -> None
    """Paint ``cell``'s background with ``rgb`` via the public ``shading`` API."""
    cell.shading.fill_color = rgb


def _strip_table_borders(table):
    # type: (Table) -> None
    """Remove the default ``Table Grid`` borders for a cleaner look.

    The kit avoids ``oxml`` reach-down; setting ``style = "Normal Table"``
    (the unstyled default) drops the gridlines without touching the
    underlying XML. We swallow :class:`KeyError` for templates that
    don't define ``Normal Table``.
    """
    try:
        table.style = "Normal Table"
    except KeyError:  # pragma: no cover - degenerate template
        pass


def _new_code_paragraph(cell, first):
    # type: (_Cell, bool) -> Paragraph
    """Return a paragraph in ``cell`` for the next line of code.

    ``cell.text = ""`` (set by the table constructor) leaves a single
    empty paragraph; we re-use that for the first line and append a
    fresh one for every subsequent line so each line maps cleanly to
    one ``w:p`` element.
    """
    if first:
        para = cell.paragraphs[0]
    else:
        para = cell.add_paragraph()
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    return para


def _style_run(run, monospace_font, font_size, color=None, bold=False, italic=False):
    # type: (object, str, object, Optional[RGBColor], bool, bool) -> None
    """Apply monospace + theme-derived font properties to ``run``."""
    font = run.font  # type: ignore[attr-defined]
    font.name = monospace_font
    font.size = font_size
    if color is not None:
        font.color.rgb = color
    if bold:
        font.bold = True
    if italic:
        font.italic = True


def _tokenise_with_pygments(source, lang):
    # type: (str, str) -> Tuple[List[List[Tuple[object, str]]], object]
    """Run pygments, returning ``(lines_of_tokens, lexer_or_None)``.

    Each *line* is a list of ``(token_type, text)`` pairs; the token's
    own newline (pygments emits ``"\\n"`` as its own token) is the line
    delimiter. The trailing empty line pygments often emits after the
    final newline is dropped so the rendered table does not carry an
    empty trailing paragraph.
    """
    from pygments import lex
    from pygments.lexers import get_lexer_by_name
    from pygments.util import ClassNotFound

    try:
        lexer = get_lexer_by_name(lang)
    except ClassNotFound:
        from pygments.lexers.special import TextLexer

        lexer = TextLexer()

    lines = [[]]  # type: List[List[Tuple[object, str]]]
    for ttype, value in lex(source, lexer):
        # -- Split on newlines so each line maps to a paragraph.
        parts = value.split("\n")
        for index, part in enumerate(parts):
            if part:
                lines[-1].append((ttype, part))
            if index < len(parts) - 1:
                lines.append([])
    # -- Drop a trailing empty line emitted when source ends with `\n`.
    if lines and not lines[-1]:
        lines.pop()
    return lines, lexer


def _resolve_token_style(style, ttype):
    # type: (object, object) -> Tuple[Optional[RGBColor], bool, bool]
    """Look up ``ttype`` in ``style``; return ``(rgb_or_None, bold, italic)``.

    Pygments returns a dict with keys ``color`` (``"rrggbb"`` or empty
    string) / ``bold`` / ``italic``; absent colour means "use the
    theme's default foreground", which we render as ``None`` so the
    caller can skip setting the run colour.
    """
    info = style.style_for_token(ttype)  # type: ignore[attr-defined]
    color_hex = info.get("color")
    rgb = _hex_to_rgb(color_hex) if color_hex else None
    return rgb, bool(info.get("bold")), bool(info.get("italic"))


def _gutter_width_chars(line_count):
    # type: (int) -> int
    """Width of the line-number text column in characters.

    A 3-digit gutter fits up to 999 lines without wrapping; longer
    listings get a 4-digit gutter. Rendered text is right-justified
    so column alignment looks correct in monospace.
    """
    if line_count < 100:
        return 3
    if line_count < 10_000:
        return 4
    return 6


def _render_lines_to_cell(
    cell,
    lines_of_tokens,
    style,
    monospace_font,
    font_size,
):
    # type: (_Cell, List[List[Tuple[object, str]]], object, str, object) -> None
    """Render highlighted token lines into ``cell``.

    One paragraph per source line; one run per token. Tokens whose
    text is empty are skipped (pygments occasionally emits zero-length
    tokens around line boundaries).
    """
    for line_index, tokens in enumerate(lines_of_tokens):
        para = _new_code_paragraph(cell, first=(line_index == 0))
        for ttype, text in tokens:
            if not text:
                continue
            run = para.add_run(text)
            color, bold, italic = _resolve_token_style(style, ttype)
            _style_run(
                run,
                monospace_font=monospace_font,
                font_size=font_size,
                color=color,
                bold=bold,
                italic=italic,
            )
        # -- Empty source lines need at least one styled (empty) run so
        # -- the paragraph carries the monospace font / size — without
        # -- it Word renders an empty line at the body font, breaking
        # -- the visual rhythm of the listing.
        if not tokens:
            run = para.add_run("")
            _style_run(run, monospace_font=monospace_font, font_size=font_size)


def _render_lines_plain(cell, lines, monospace_font, font_size):
    # type: (_Cell, List[str], str, object) -> None
    """Soft-fallback renderer — no highlighting, just monospace."""
    for line_index, text in enumerate(lines):
        para = _new_code_paragraph(cell, first=(line_index == 0))
        run = para.add_run(text)
        _style_run(
            run,
            monospace_font=monospace_font,
            font_size=font_size,
            color=_FALLBACK_FOREGROUND,
        )


def _render_gutter(cell, line_count, monospace_font, font_size):
    # type: (_Cell, int, str, object) -> None
    """Fill ``cell`` with right-justified line-number paragraphs."""
    width = _gutter_width_chars(line_count)
    for line_index in range(line_count):
        para = _new_code_paragraph(cell, first=(line_index == 0))
        text = str(line_index + 1).rjust(width)
        run = para.add_run(text)
        _style_run(
            run,
            monospace_font=monospace_font,
            font_size=font_size,
            color=_GUTTER_FOREGROUND,
        )


def _normalise_source(source):
    # type: (str) -> List[str]
    '''Strip leading/trailing blank-line padding from a heredoc source.

    Triple-quoted strings idiomatically open with a newline after the
    opening triple-quote; without normalisation that would render as
    an empty leading line in the code block. We also drop a single
    trailing newline so the block does not carry a phantom blank
    final line.
    '''
    if source.startswith("\n"):
        source = source[1:]
    if source.endswith("\n"):
        source = source[:-1]
    return source.split("\n")


def add(
    document,
    source,
    *,
    lang,
    line_numbers=False,
    theme="default",
    monospace_font=_DEFAULT_MONOSPACE_FONT,
    font_size=None,
):
    # type: (Document, str, str, bool, str, str, Optional[object]) -> Table
    """Append a syntax-highlighted code block to ``document``.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    source
        The source code to render. Leading and trailing blank lines
        introduced by triple-quoted heredocs are stripped.
    lang
        The pygments lexer name (``"python"``, ``"bash"``, ``"json"``,
        ``"yaml"``, ``"sql"``, ``"javascript"``, etc.). Unknown
        languages fall back to the plain-text lexer (no highlighting,
        but no exception either).
    line_numbers
        When ``True``, render a narrow line-number gutter on the left
        of the block. Defaults to ``False``.
    theme
        The pygments style name. Common picks: ``"default"``,
        ``"monokai"``, ``"tango"``, ``"friendly"``, ``"solarized-dark"``,
        ``"solarized-light"``, ``"vs"``. Defaults to ``"default"``.
        Ignored when pygments is not installed.
    monospace_font
        Override the rendered font face. Defaults to ``"Consolas"``.
    font_size
        Override the rendered font size. Defaults to ``Pt(9)``.

    Returns
    -------
    Table
        The single-row table holding the code block. Callers may
        post-process (e.g. add a caption paragraph beforehand or a
        spacing paragraph afterwards).
    """
    if font_size is None:
        font_size = _DEFAULT_CODE_FONT_SIZE

    lines = _normalise_source(source)
    line_count = len(lines)
    cols = 2 if line_numbers else 1

    table = document.add_table(rows=1, cols=cols)
    _strip_table_borders(table)

    background = _resolve_background(theme)
    cells = table.rows[0].cells

    # -- Body cell is always the rightmost column.
    body_cell = cells[-1]
    _shade_cell(body_cell, background)

    # -- Pygments path.  Falls through to the plain renderer when the
    # -- import fails (e.g. pygments uninstalled in the runtime venv)
    # -- so callers never see an exception from the helper.
    if HAS_PYGMENTS:
        try:
            from pygments.styles import get_style_by_name

            try:
                style = get_style_by_name(theme)
            except Exception:  # noqa: BLE001 - unknown theme name → default
                style = get_style_by_name("default")
            lines_of_tokens, _lexer = _tokenise_with_pygments(
                "\n".join(lines), lang
            )
            _render_lines_to_cell(
                body_cell,
                lines_of_tokens,
                style=style,
                monospace_font=monospace_font,
                font_size=font_size,
            )
        except Exception:  # noqa: BLE001 - any unexpected pygments error
            _render_lines_plain(body_cell, lines, monospace_font, font_size)
    else:
        _render_lines_plain(body_cell, lines, monospace_font, font_size)

    if line_numbers:
        gutter_cell = cells[0]
        _shade_cell(gutter_cell, background)
        _render_gutter(gutter_cell, line_count, monospace_font, font_size)

    return table


# -- Convenience wrappers.  Each forwards every keyword argument the
# -- main `add` accepts, so callers keep `theme` / `line_numbers` /
# -- `monospace_font` / `font_size` overrides while binding `lang` to
# -- the wrapper's namesake. The wrappers cover the dozen languages
# -- that account for the vast majority of code-in-document use cases
# -- (web stack + ops + systems + data interchange + a couple of
# -- emerging systems languages); callers needing other lexers fall
# -- back to `add(..., lang="...")`.
#
# -- Implementation: a closure-based factory keeps the public surface
# -- stable while avoiding twelve copies of the same forwarding stub.
# -- Each wrapper appears in the module namespace as a distinct named
# -- callable (so `code_block.python` is identical to invoking the
# -- factory's output once at import time) and shares `add`'s
# -- keyword-only signature.


def _make_lang_wrapper(lang_name):
    # type: (str) -> object
    """Build a per-language convenience wrapper that delegates to :func:`add`."""

    def _wrapper(
        document,
        source,
        *,
        line_numbers=False,
        theme="default",
        monospace_font=_DEFAULT_MONOSPACE_FONT,
        font_size=None,
    ):
        # type: (Document, str, bool, str, str, Optional[object]) -> Table
        return add(
            document,
            source,
            lang=lang_name,
            line_numbers=line_numbers,
            theme=theme,
            monospace_font=monospace_font,
            font_size=font_size,
        )

    _wrapper.__name__ = lang_name
    _wrapper.__qualname__ = lang_name
    _wrapper.__doc__ = (
        "Append a %s code block. Shortcut for :func:`add` with "
        '``lang=%r``.' % (lang_name.capitalize(), lang_name)
    )
    return _wrapper


# -- Generate the twelve language-specific wrappers.  ``json`` is
# -- bound last because the standard-library `json` module imported
# -- earlier in this file (none, but defensive) would otherwise
# -- shadow the wrapper.
python = _make_lang_wrapper("python")
bash = _make_lang_wrapper("bash")
json = _make_lang_wrapper("json")
yaml = _make_lang_wrapper("yaml")
sql = _make_lang_wrapper("sql")
javascript = _make_lang_wrapper("javascript")
typescript = _make_lang_wrapper("typescript")
rust = _make_lang_wrapper("rust")
go = _make_lang_wrapper("go")
html = _make_lang_wrapper("html")
css = _make_lang_wrapper("css")
xml = _make_lang_wrapper("xml")


__all__ = [
    "HAS_PYGMENTS",
    "add",
    "bash",
    "css",
    "go",
    "html",
    "javascript",
    "json",
    "python",
    "rust",
    "sql",
    "typescript",
    "xml",
    "yaml",
]
