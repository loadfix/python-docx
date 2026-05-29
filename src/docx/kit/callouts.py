"""Callout / admonition boxes — Note / Warning / Tip / Caution / Important / Example.

Closes #287.

Technical writing routinely uses *callout boxes* (sometimes called
*admonitions*, *call-outs*, or *aside boxes*) to set short pieces of
auxiliary content visually apart from the body — a tip, a hazard
warning, a worked example, or a "you must read this" note. Each
callout is a coloured box with a short label / icon and a body.
Convention varies between style guides (DITA, AsciiDoc, MkDocs,
Microsoft Manual of Style, GitHub Flavoured Markdown) but the six
labels covered here — *Note*, *Warning*, *Tip*, *Caution*,
*Important*, *Example* — are the union that appears in nearly every
guide.

This module exposes seven composition helpers built entirely on
python-docx's public API (``Document.add_table``, ``_Cell.shading``,
``_Cell.paragraphs[0].add_run``, ``Run.font``)::

    from docx import Document
    from docx.kit import callouts

    doc = Document()
    callouts.note(doc, "This is informational.")
    callouts.warning(doc, "Be careful here.")
    callouts.tip(doc, "Pro tip: use the kit.")
    callouts.caution(doc, "May cause data loss.")
    callouts.important(doc, "You must read this.")
    callouts.example(doc, "Here is an example.")

    # Custom callout — pick your own style + icon + title.
    callouts.box(doc, "Custom callout", style="info", icon="ℹ")

Each helper appends a *single-cell* one-row table to the document and
returns the resulting :class:`~docx.table.Table`. The cell is
shaded with a style-specific pastel fill, the body text is preceded
by a Unicode icon (emoji or symbol) prefix, and the optional title
(``"Note"``, ``"Warning"``, …) is rendered bold on its own line above
the body. Multi-paragraph bodies are supported by passing a
``list[str]`` rather than a single ``str``.

Style → fill colour:

================  =======================  ==========================
Style             Fill (pastel)            Title-bar emoji
================  =======================  ==========================
``note``          light blue ``DEEBF7``    ``\U0001F4DD`` MEMO
``warning``       amber     ``FFE699``     ``⚠`` WARNING SIGN
``caution``       light red ``F8CBAD``     ``⛔`` NO ENTRY
``tip``           light green ``E2EFDA``   ``\U0001F4A1`` LIGHT BULB
``important``     lavender  ``E4DFEC``     ``❗`` HEAVY EXCLAMATION
``example``       neutral   ``F2F2F2``     ``\U0001F4D8`` BLUE BOOK
``info``          light cyan ``DEF1F5``    ``ℹ`` INFORMATION SOURCE
================  =======================  ==========================

The final ``info`` row is the conventional default for the underlying
:func:`box` helper when callers pass ``style="info"`` (or any
unrecognised style — :func:`box` falls back to ``info`` rather than
raising so an unfamiliar style name doesn't crash a document build).

Implementation notes:

* The single-cell table approach matches what Word's built-in
  *Insert > Quick Parts > Building Block Gallery > Sidebars / Pull
  Quotes* gallery emits, and is the only shape that survives a
  Word ``.docx`` round-trip with both fill and inline-flow intact.
  An alternative ``w:pBdr`` + ``w:shd`` paragraph-level approach
  loses the side margins.
* Pastel fills are picked so black 11pt body text remains readable
  (every fill is luminance > 0.85). Office's Word 2013-themed
  *Subtle Reference* and *Intense Reference* styles use a similar
  palette.
* The icon glyph is prepended *to the body run* (not stamped as a
  separate run) so callers iterating ``cell.paragraphs[0].runs``
  see one run per logical chunk. The title, when present, is a
  separate paragraph above the body so it can carry bold styling
  independently.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional, Sequence, Tuple, Union

from docx.shared import Pt, RGBColor

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table


# -- Style → (fill colour, default icon glyph, default title).
# -- Fills are pastel hex triples; icons are single Unicode glyphs
# -- chosen from the BMP (or astral plane for the four emoji) so they
# -- render in Word's default Calibri / Segoe UI Emoji glyph fall-back.
_StyleSpec = Tuple[RGBColor, str, str]

_STYLES: "dict[str, _StyleSpec]" = {
    "note": (RGBColor(0xDE, 0xEB, 0xF7), "\U0001F4DD", "Note"),
    "warning": (RGBColor(0xFF, 0xE6, 0x99), "⚠", "Warning"),
    "caution": (RGBColor(0xF8, 0xCB, 0xAD), "⛔", "Caution"),
    "tip": (RGBColor(0xE2, 0xEF, 0xDA), "\U0001F4A1", "Tip"),
    "important": (RGBColor(0xE4, 0xDF, 0xEC), "❗", "Important"),
    "example": (RGBColor(0xF2, 0xF2, 0xF2), "\U0001F4D8", "Example"),
    "info": (RGBColor(0xDE, 0xF1, 0xF5), "ℹ", "Info"),
}

# -- Recognised aliases for the six convenience helpers.  Used by
# -- :func:`box` to look up the fill / icon / title triple.  Order
# -- matters only for error messages — Python ``dict`` preserves
# -- insertion order, and the convenience helpers below dispatch by
# -- name so the iteration order of ``_STYLES`` is also the public
# -- listing order in error messages.
_KNOWN_STYLES = tuple(_STYLES.keys())


def _coerce_body(body):
    # type: (Union[str, Sequence[str]]) -> List[str]
    """Normalise ``body`` to a non-empty list of paragraph strings.

    Accepts a single ``str`` or any iterable of ``str``. Empty
    strings inside an iterable are preserved so the caller can emit
    a deliberate blank-line separator between body paragraphs.
    """
    if isinstance(body, str):
        return [body]
    paragraphs = [str(item) for item in body]
    if not paragraphs:
        raise ValueError("body must be a non-empty string or iterable of strings")
    return paragraphs


def box(
    document,
    body,
    style="info",
    icon=None,
    title=None,
):
    # type: (Document, Union[str, Sequence[str]], str, Optional[str], Optional[str]) -> Table
    """Append a single-cell coloured callout box to ``document``.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    body
        The callout body. Either a single ``str`` (rendered as one
        paragraph, the icon glyph prepended) or an iterable of
        ``str`` (rendered as N paragraphs; the icon prepends only
        the first).
    style
        One of ``"note"`` / ``"warning"`` / ``"caution"`` / ``"tip"``
        / ``"important"`` / ``"example"`` / ``"info"``. Unrecognised
        values fall back to ``"info"`` rather than raise so a
        caller's typo doesn't crash a long document build. The fall-
        back behaviour is deliberate — callouts are visual decoration,
        not load-bearing structure.
    icon
        Optional override for the icon glyph (or ``""`` to suppress
        the default). When |None| the style's default icon is used.
    title
        Optional bold title rendered on its own line above the body
        (e.g. ``"Note"``, ``"Pro tip"``). When |None| no title line
        is emitted; the icon and body share a single paragraph.

    Returns
    -------
    Table
        The newly-appended single-cell, single-row :class:`Table`
        carrying the callout. Callers may further mutate the table
        (resize, restyle, append rows) via the public table API.
    """
    # -- Look up the style triple; unrecognised styles fall back to
    # -- "info" rather than raise.  See the docstring for rationale.
    spec = _STYLES.get(style, _STYLES["info"])
    fill, default_icon, _default_title = spec

    paragraphs = _coerce_body(body)

    # -- Resolve the icon.  ``None`` = use style default; ``""`` =
    # -- suppress; any other string = use as-is.
    if icon is None:
        glyph = default_icon
    else:
        glyph = icon

    table = document.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    cell.shading.fill_color = fill

    # -- Replace the auto-emitted empty paragraph with our content.
    # -- _Cell.text=... clears existing paragraphs to a single blank
    # -- paragraph; we then append paragraphs as needed.
    first_para = cell.paragraphs[0]
    # -- Clear any default text the table inserted.
    for run in list(first_para.runs):
        run.text = ""

    if title is not None:
        # -- Title line: bold, optionally prefixed by the icon glyph.
        title_text = title
        if glyph:
            title_text = "%s %s" % (glyph, title)
        title_run = first_para.add_run(title_text)
        title_run.bold = True
        # -- Body paragraphs follow the title line.
        for body_text in paragraphs:
            body_para = cell.add_paragraph()
            body_para.add_run(body_text)
    else:
        # -- No title: prepend the icon to the first body paragraph
        # -- so the call-out reads ``<icon> <body...>``.
        first_text = paragraphs[0]
        if glyph:
            first_text = "%s %s" % (glyph, first_text)
        first_para.add_run(first_text)
        for body_text in paragraphs[1:]:
            body_para = cell.add_paragraph()
            body_para.add_run(body_text)

    return table


def note(document, body, *, title="Note"):
    # type: (Document, Union[str, Sequence[str]], Optional[str]) -> Table
    """Append a blue-shaded *Note* callout to ``document``.

    Convenience wrapper over :func:`box` with ``style="note"`` and a
    default ``title`` of ``"Note"``. Pass ``title=None`` to suppress
    the title line.
    """
    return box(document, body, style="note", title=title)


def warning(document, body, *, title="Warning"):
    # type: (Document, Union[str, Sequence[str]], Optional[str]) -> Table
    """Append an amber-shaded *Warning* callout to ``document``.

    Convenience wrapper over :func:`box` with ``style="warning"`` and
    a default ``title`` of ``"Warning"``. Pass ``title=None`` to
    suppress the title line.
    """
    return box(document, body, style="warning", title=title)


def tip(document, body, *, title="Tip"):
    # type: (Document, Union[str, Sequence[str]], Optional[str]) -> Table
    """Append a green-shaded *Tip* callout to ``document``.

    Convenience wrapper over :func:`box` with ``style="tip"`` and a
    default ``title`` of ``"Tip"``. Pass ``title=None`` to suppress
    the title line.
    """
    return box(document, body, style="tip", title=title)


def caution(document, body, *, title="Caution"):
    # type: (Document, Union[str, Sequence[str]], Optional[str]) -> Table
    """Append a red-shaded *Caution* callout to ``document``.

    Convenience wrapper over :func:`box` with ``style="caution"`` and
    a default ``title`` of ``"Caution"``. Pass ``title=None`` to
    suppress the title line.
    """
    return box(document, body, style="caution", title=title)


def important(document, body, *, title="Important"):
    # type: (Document, Union[str, Sequence[str]], Optional[str]) -> Table
    """Append a purple-shaded *Important* callout to ``document``.

    Convenience wrapper over :func:`box` with ``style="important"``
    and a default ``title`` of ``"Important"``. Pass ``title=None``
    to suppress the title line.
    """
    return box(document, body, style="important", title=title)


def example(document, body, *, title="Example"):
    # type: (Document, Union[str, Sequence[str]], Optional[str]) -> Table
    """Append a neutral-shaded *Example* callout to ``document``.

    Convenience wrapper over :func:`box` with ``style="example"`` and
    a default ``title`` of ``"Example"``. Pass ``title=None`` to
    suppress the title line.
    """
    return box(document, body, style="example", title=title)


# -- Deliberately silence pyflakes on the ``Pt`` import; future
# -- callers may want to set body-paragraph spacing via Pt(...) on
# -- the returned table.  Listed in the public surface below so it
# -- is re-exported alongside the rest of the helpers when callers
# -- ``from docx.kit.callouts import *``.
_ = Pt


__all__ = [
    "box",
    "note",
    "warning",
    "tip",
    "caution",
    "important",
    "example",
]
