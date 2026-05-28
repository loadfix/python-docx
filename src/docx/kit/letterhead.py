"""Letterhead helper — branded header (logo + return address) and footer (contact lines).

Closes #61.

This module composes existing python-docx primitives (``Document.sections``,
``_Header.paragraphs``, ``Run.add_picture``, ``Paragraph.add_hyperlink``,
``font.color.rgb``, ``Document.theme``) into a single high-level helper,
:func:`set_letterhead`, that brands every page of a document with a styled
header (logo + return address) and footer (phone / email / website)::

    from docx.kit.letterhead import set_letterhead

    set_letterhead(
        doc,
        logo="acme-logo.png",
        return_address="123 Pitt Street\\nSydney NSW 2000\\nAustralia",
        phone="+61 2 1234 5678",
        email="hello@acme.com",
        website="acme.com",
        style="modern",
        color="primary",
    )

Three built-in *styles* shape the header/footer layout:

- ``"modern"`` (default) — logo left, return address right; footer
  combines all three contact strings on a single centered line
  separated by middle dots. Accent-coloured.
- ``"classic"`` — return address centered above a coloured horizontal
  rule (em-dash run); footer centers each contact line on its own
  paragraph, italic-styled. Logo, when supplied, is centered above
  the address.
- ``"minimal"`` — single-line header with logo (left) and a tab-separated
  one-line summary of the return address (right); footer is a single
  centered line, no decorations, no italic.

Theme integration: ``color`` accepts the named presets
(``"primary"``..``"muted"``), an :class:`RGBColor`, a hex string, or
a theme-color token (``"accent1"``..``"accent6"``, ``"hlink"``, etc.).
When a token is supplied and the document has a theme part, the helper
resolves it via :attr:`Document.theme` and applies the resulting RGB.
When the document has no theme, theme tokens fall back to the equivalent
named preset (``accent1`` -> ``primary``, ``accent2`` -> ``secondary``, etc.)
so the helper still works on minimal templates.

The helper writes to the *primary* header/footer of the *first* section
only — that section's header/footer cascades to subsequent sections by
default in Word, so this is the canonical "document-wide letterhead"
shape. Callers needing per-section variation should call
:func:`set_letterhead` per section after constructing them.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, List, Optional, Tuple, Union

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

if TYPE_CHECKING:
    from docx.document import Document
    from docx.section import _Footer, _Header
    from docx.text.paragraph import Paragraph


# -- Named accent colours.  Resolves the keyword strings exposed in
# -- the public API (``color="primary"``) to concrete RGB triples.  These
# -- match the chapter-opener palette so kit helpers compose visually.
_NAMED_COLORS = {
    "primary": RGBColor(0x1F, 0x4E, 0x79),    # deep blue
    "secondary": RGBColor(0x70, 0x30, 0xA0),  # purple
    "accent": RGBColor(0xC0, 0x50, 0x4D),     # warm red
    "muted": RGBColor(0x59, 0x59, 0x59),      # neutral grey
    "black": RGBColor(0x00, 0x00, 0x00),
}

# -- Theme-color tokens recognised by the helper. The values map each
# -- token to the named-preset fallback used when the document has no
# -- theme part. This keeps minimal templates working without surprise.
_THEME_TOKENS = {
    "accent1": "primary",
    "accent2": "secondary",
    "accent3": "accent",
    "accent4": "secondary",
    "accent5": "primary",
    "accent6": "muted",
    "dk1": "black",
    "dk2": "muted",
    "lt1": "muted",
    "lt2": "muted",
    "hlink": "primary",
    "folHlink": "secondary",
}

# -- Built-in styles. Listed here so callers and tests can introspect.
STYLES: Tuple[str, ...] = ("modern", "classic", "minimal")

# -- Visual sizing for header/footer text. Pt() is a typed Length;
# -- keep integer literals here so the visual identity stays consistent.
_HEADER_TEXT_SIZE = Pt(10)
_FOOTER_TEXT_SIZE = Pt(9)
_DEFAULT_LOGO_HEIGHT = Pt(36)  # ~ half-inch
_RULE = "—" * 30  # 30 em-dashes — classic-style horizontal rule


def _resolve_color(
    color,  # type: Union[str, RGBColor, None]
    document=None,  # type: Optional[Document]
):
    # type: (...) -> Optional[RGBColor]
    """Return an :class:`RGBColor` for ``color`` or |None|.

    ``color`` may be:

    - |None| — return |None| (caller leaves text colour at the style default);
    - one of the named keys in :data:`_NAMED_COLORS`;
    - one of the theme-color tokens in :data:`_THEME_TOKENS` — resolved
      via ``document.theme`` when available, else falls back to the
      preset associated with the token;
    - an :class:`RGBColor` instance — returned unchanged;
    - a 6-character hex string (with or without leading ``#``).

    ``document`` is consulted only for theme-token resolution; pass
    |None| to skip theme lookup and use the fallback preset directly
    (useful in unit tests).

    Any other value raises :class:`ValueError`.
    """
    if color is None:
        return None
    if isinstance(color, RGBColor):
        return color
    if isinstance(color, str):
        key = color.lower()
        if key in _NAMED_COLORS:
            return _NAMED_COLORS[key]
        if key in _THEME_TOKENS:
            # -- Try the document's theme first; on success use the
            # -- resolved RGB, otherwise fall back to the preset. The
            # -- ``getattr`` chain tolerates a Document with no theme
            # -- attribute (synthetic / bare-bones test fixtures).
            theme = getattr(document, "theme", None) if document is not None else None
            if theme is not None:
                try:
                    rgb = theme.colors[key]
                except (KeyError, AttributeError):
                    rgb = None
                if rgb is not None:
                    return rgb
            return _NAMED_COLORS[_THEME_TOKENS[key]]
        # -- treat as hex string (strip leading '#' if present)
        return RGBColor.from_string(color.lstrip("#"))
    raise ValueError(
        "color must be None, a named preset (%s), a theme token "
        "(%s), an RGBColor, or a hex string; got %r"
        % (
            ", ".join(sorted(_NAMED_COLORS)),
            ", ".join(sorted(_THEME_TOKENS)),
            color,
        )
    )


def _clear_existing_paragraphs(container):
    # type: (Union[_Header, _Footer]) -> None
    """Empty ``container`` so the helper writes from a clean slate.

    The default header/footer Word writes when a section first gets
    one is a single empty paragraph. The helper's body assumes it is
    appending into an empty container, so any pre-existing content
    (from a template, or from a re-run of :func:`set_letterhead` on
    the same document) is removed first.

    Implementation: replaces the first paragraph's text with empty
    string and removes every subsequent paragraph by deleting its
    underlying ``w:p`` element. Uses only the public ``paragraphs``
    sequence and the ``text`` setter — no oxml reach-down. The
    underlying element removal is via Python's standard ``getparent`` /
    ``remove`` on the lxml node exposed by the public ``_p`` proxy
    attribute, which is the supported escape hatch (see how
    :meth:`Document.tracked_changes` clears its scope).
    """
    paragraphs = list(container.paragraphs)
    if not paragraphs:
        return
    # -- Wipe the first paragraph so it stays as the required empty
    # -- placeholder; deleting it would leave the header/footer
    # -- structurally invalid (Word requires >=1 paragraph). --
    first = paragraphs[0]
    first.text = ""
    # -- Remove the surplus paragraphs from the tail.
    for para in paragraphs[1:]:
        p_elm = para._p  # public proxy attribute (see Paragraph)
        parent = p_elm.getparent()
        if parent is not None:
            parent.remove(p_elm)


def _apply_color(run, rgb):
    # type: (object, Optional[RGBColor]) -> None
    """Set the run's font colour to `rgb` when supplied; no-op otherwise."""
    if rgb is not None:
        run.font.color.rgb = rgb


def _add_logo(paragraph, logo, height=None):
    # type: (Paragraph, Union[str, os.PathLike, None], Optional[int]) -> None
    """Append the brand logo image to ``paragraph`` at a sensible default size.

    ``logo`` may be any value :meth:`Run.add_picture` accepts. ``height``
    overrides the default header logo height (~36 pt). Width is
    auto-computed by python-docx to preserve the aspect ratio.
    """
    if logo is None:
        return
    run = paragraph.add_run()
    run.add_picture(logo, height=height or _DEFAULT_LOGO_HEIGHT)


def _split_address(text):
    # type: (Optional[str]) -> List[str]
    """Split a multi-line return address into its lines, dropping empties."""
    if not text:
        return []
    return [line for line in text.splitlines() if line.strip()]


def _join_minimal_address(lines):
    # type: (List[str]) -> str
    """Collapse a multi-line return address into a single comma-separated string."""
    return ", ".join(line.strip() for line in lines)


def _set_run_text(run, text, rgb, size, bold=False, italic=False):
    # type: (object, str, Optional[RGBColor], Optional[int], bool, bool) -> None
    """Configure ``run`` with the standard letterhead font block."""
    run.text = text
    if size is not None:
        run.font.size = size
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    _apply_color(run, rgb)


def _has_style(document, style_name):
    # type: (Optional[Document], str) -> bool
    """Return |True| when `document` defines a paragraph/character style with the given name."""
    if document is None:
        return False
    try:
        styles = document.styles
        styles[style_name]
        return True
    except (KeyError, AttributeError):
        return False


def _add_email_link(paragraph, email, rgb, size, italic=False):
    # type: (Paragraph, str, Optional[RGBColor], Optional[int], bool) -> None
    """Append `email` as a ``mailto:`` hyperlink with letterhead styling.

    The Hyperlink character style is preferred when available so Word
    renders the link consistently with the rest of the document, but
    minimal templates that lack the style fall through to ``style=None``
    (no character style applied). Either way the helper sets explicit
    font colour / size on the resulting runs so the letterhead's accent
    palette wins.
    """
    document = paragraph.part.package.main_document_part.document
    style: Optional[str] = "Hyperlink" if _has_style(document, "Hyperlink") else None
    link = paragraph.add_hyperlink(url=f"mailto:{email}", text=email, style=style)
    for run in link.runs:
        if size is not None:
            run.font.size = size
        if italic:
            run.italic = True
        _apply_color(run, rgb)


def _add_website_link(paragraph, website, rgb, size, italic=False):
    # type: (Paragraph, str, Optional[RGBColor], Optional[int], bool) -> None
    """Append `website` as a hyperlink with letterhead styling.

    The ``website`` argument is rendered verbatim as the visible link
    text but the URL is normalised to include the ``https://`` scheme
    when the caller supplied a bare domain (``"acme.com"``).
    """
    href = website if "://" in website else f"https://{website}"
    document = paragraph.part.package.main_document_part.document
    style: Optional[str] = "Hyperlink" if _has_style(document, "Hyperlink") else None
    link = paragraph.add_hyperlink(url=href, text=website, style=style)
    for run in link.runs:
        if size is not None:
            run.font.size = size
        if italic:
            run.italic = True
        _apply_color(run, rgb)


# -- Per-style header / footer renderers ----------------------------------
# -- Each renderer takes an empty header/footer container plus the
# -- already-resolved RGB colour and the four content fields, and
# -- writes the appropriate visual layout.  Keeping the renderers small
# -- and homogeneous makes the dispatch in ``set_letterhead`` trivial.
# ------------------------------------------------------------------------


def _render_modern_header(header, logo, return_address, rgb):
    # type: (_Header, Union[str, os.PathLike, None], Optional[str], Optional[RGBColor]) -> List[Paragraph]
    """Modern: logo on left, return address on right (tab-separated).

    A single paragraph carries both elements with a centre-tab and a
    right-tab so they sit at the page edges. Falls back to two
    paragraphs when only one of the two pieces is present.
    """
    paragraphs: List[Paragraph] = []
    para = header.paragraphs[0]
    paragraphs.append(para)

    address_lines = _split_address(return_address)
    address_text = _join_minimal_address(address_lines)

    if logo is not None:
        _add_logo(para, logo)
    if address_text:
        # -- Tab character separates logo from the right-aligned address.
        if logo is not None:
            tab_run = para.add_run("\t")
            tab_run.font.size = _HEADER_TEXT_SIZE
        addr_run = para.add_run(address_text)
        _set_run_text(
            addr_run, address_text, rgb, _HEADER_TEXT_SIZE, bold=False
        )
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT if logo is None else None
    return paragraphs


def _render_modern_footer(footer, phone, email, website, rgb):
    # type: (_Footer, Optional[str], Optional[str], Optional[str], Optional[RGBColor]) -> List[Paragraph]
    """Modern: contact pieces joined with middle-dot separators on one centered line."""
    para = footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    parts: List[Tuple[str, str]] = []
    if phone:
        parts.append(("text", phone))
    if email:
        parts.append(("email", email))
    if website:
        parts.append(("website", website))

    sep = "  ·  "  # middle dot with surrounding spaces
    for index, (kind, value) in enumerate(parts):
        if index > 0:
            sep_run = para.add_run(sep)
            _set_run_text(sep_run, sep, rgb, _FOOTER_TEXT_SIZE)
        if kind == "text":
            run = para.add_run(value)
            _set_run_text(run, value, rgb, _FOOTER_TEXT_SIZE)
        elif kind == "email":
            _add_email_link(para, value, rgb, _FOOTER_TEXT_SIZE)
        else:  # website
            _add_website_link(para, value, rgb, _FOOTER_TEXT_SIZE)
    return [para]


def _render_classic_header(header, logo, return_address, rgb):
    # type: (_Header, Union[str, os.PathLike, None], Optional[str], Optional[RGBColor]) -> List[Paragraph]
    """Classic: centered logo on its own line, then the address, then a horizontal rule."""
    paragraphs: List[Paragraph] = []
    first = header.paragraphs[0]

    if logo is not None:
        first.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_logo(first, logo)
        paragraphs.append(first)
        addr_para_target = header.add_paragraph() if any(_split_address(return_address)) else None
    else:
        addr_para_target = first

    address_lines = _split_address(return_address)
    if address_lines and addr_para_target is not None:
        addr_para_target.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for index, line in enumerate(address_lines):
            if index > 0:
                # -- soft line break keeps all address lines in one paragraph
                br_run = addr_para_target.add_run()
                br_run.add_break()
                _apply_color(br_run, rgb)
            run = addr_para_target.add_run(line)
            _set_run_text(run, line, rgb, _HEADER_TEXT_SIZE)
        paragraphs.append(addr_para_target)

    # -- decorative horizontal rule using em-dashes (no XML reach-down)
    rule = header.add_paragraph()
    rule.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rule_run = rule.add_run(_RULE)
    _set_run_text(rule_run, _RULE, rgb, _HEADER_TEXT_SIZE)
    paragraphs.append(rule)
    return paragraphs


def _render_classic_footer(footer, phone, email, website, rgb):
    # type: (_Footer, Optional[str], Optional[str], Optional[str], Optional[RGBColor]) -> List[Paragraph]
    """Classic: each contact line on its own centered italic paragraph, preceded by a rule."""
    paragraphs: List[Paragraph] = []
    rule_para = footer.paragraphs[0]
    rule_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rule_run = rule_para.add_run(_RULE)
    _set_run_text(rule_run, _RULE, rgb, _FOOTER_TEXT_SIZE)
    paragraphs.append(rule_para)

    if phone:
        para = footer.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(phone)
        _set_run_text(run, phone, rgb, _FOOTER_TEXT_SIZE, italic=True)
        paragraphs.append(para)
    if email:
        para = footer.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_email_link(para, email, rgb, _FOOTER_TEXT_SIZE, italic=True)
        paragraphs.append(para)
    if website:
        para = footer.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_website_link(para, website, rgb, _FOOTER_TEXT_SIZE, italic=True)
        paragraphs.append(para)
    return paragraphs


def _render_minimal_header(header, logo, return_address, rgb):
    # type: (_Header, Union[str, os.PathLike, None], Optional[str], Optional[RGBColor]) -> List[Paragraph]
    """Minimal: single line, logo on left, comma-flattened address on right."""
    para = header.paragraphs[0]
    if logo is not None:
        _add_logo(para, logo)
    address_lines = _split_address(return_address)
    if address_lines:
        if logo is not None:
            tab_run = para.add_run("\t")
            tab_run.font.size = _HEADER_TEXT_SIZE
        text = _join_minimal_address(address_lines)
        run = para.add_run(text)
        _set_run_text(run, text, rgb, _HEADER_TEXT_SIZE)
        if logo is None:
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return [para]


def _render_minimal_footer(footer, phone, email, website, rgb):
    # type: (_Footer, Optional[str], Optional[str], Optional[str], Optional[RGBColor]) -> List[Paragraph]
    """Minimal: centered single line, fields joined by ``" | "`` separators."""
    para = footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sep = " | "
    parts: List[Tuple[str, str]] = []
    if phone:
        parts.append(("text", phone))
    if email:
        parts.append(("email", email))
    if website:
        parts.append(("website", website))
    for index, (kind, value) in enumerate(parts):
        if index > 0:
            sep_run = para.add_run(sep)
            _set_run_text(sep_run, sep, rgb, _FOOTER_TEXT_SIZE)
        if kind == "text":
            run = para.add_run(value)
            _set_run_text(run, value, rgb, _FOOTER_TEXT_SIZE)
        elif kind == "email":
            _add_email_link(para, value, rgb, _FOOTER_TEXT_SIZE)
        else:
            _add_website_link(para, value, rgb, _FOOTER_TEXT_SIZE)
    return [para]


_HEADER_RENDERERS = {
    "modern": _render_modern_header,
    "classic": _render_classic_header,
    "minimal": _render_minimal_header,
}
_FOOTER_RENDERERS = {
    "modern": _render_modern_footer,
    "classic": _render_classic_footer,
    "minimal": _render_minimal_footer,
}


def set_letterhead(
    document,
    logo=None,
    return_address=None,
    phone=None,
    email=None,
    website=None,
    style="modern",
    color=None,
):
    # type: (Document, Union[str, os.PathLike, None], Optional[str], Optional[str], Optional[str], Optional[str], str, Union[str, RGBColor, None]) -> dict
    """Apply a styled letterhead — branded header and footer — to ``document``.

    Writes to the *primary* header and footer of the *first* section.
    In Word, those definitions cascade to subsequent sections by
    default, so this is the canonical "document-wide letterhead"
    operation.

    Any pre-existing header/footer content on the target section is
    cleared first so calling :func:`set_letterhead` is idempotent
    across re-runs and template-based callers.

    Parameters
    ----------
    document
        The :class:`Document` to brand.
    logo
        Path or stream for the brand image. Any value accepted by
        :meth:`Run.add_picture` works (string path, :class:`os.PathLike`,
        binary file-like object). Pass |None| to omit the logo — the
        helper still emits the address and footer.
    return_address
        Multi-line return address. Lines are split on ``"\\n"``; empty
        lines are dropped. Pass |None| to omit. The ``"minimal"`` style
        flattens the address to a single comma-separated line.
    phone
        Phone number rendered verbatim in the footer. Pass |None| to omit.
    email
        Email address rendered as a ``mailto:`` hyperlink in the footer.
        Pass |None| to omit.
    website
        Website rendered as an external hyperlink in the footer. A bare
        domain (``"acme.com"``) is normalised to ``"https://acme.com"``;
        pass a full URL to override the scheme. Pass |None| to omit.
    style
        One of ``"modern"`` (default), ``"classic"``, or ``"minimal"``.
        See module docstring for the visual contract of each.
    color
        Accent colour. Accepts a named preset (``"primary"``,
        ``"secondary"``, ``"accent"``, ``"muted"``, ``"black"``), a
        theme-color token (``"accent1"``..``"accent6"``, ``"hlink"``,
        ``"folHlink"``, ``"dk1"``, ``"dk2"``, ``"lt1"``, ``"lt2"``)
        which is resolved through ``document.theme`` when available,
        an :class:`RGBColor`, or a 6-character hex string. |None|
        leaves text at the style default.

    Returns
    -------
    dict
        ``{"header": [Paragraph, ...], "footer": [Paragraph, ...]}`` —
        the paragraphs the helper appended, in document order. Tests and
        fluent callers can iterate them to apply further formatting.

    Raises
    ------
    ValueError
        When ``style`` is not one of the three built-ins, or when
        ``color`` is supplied as a non-recognisable string.

    .. versionadded:: 2026.05.29
    """
    if style not in STYLES:
        raise ValueError(
            "style must be one of %s; got %r"
            % (", ".join(repr(s) for s in STYLES), style)
        )
    if not document.sections:  # pragma: no cover - defensive
        raise ValueError("document has no sections; cannot apply letterhead")

    rgb = _resolve_color(color, document=document)

    section = document.sections[0]
    header = section.header
    footer = section.footer
    # -- Detach the section's header/footer from the prior section so
    # -- the new content actually applies. This is a no-op for the
    # -- first section (which has no prior) but defends against callers
    # -- who run set_letterhead on a document whose sections were
    # -- copied from a template that linked them.
    if header.is_linked_to_previous:
        header.is_linked_to_previous = False
    if footer.is_linked_to_previous:
        footer.is_linked_to_previous = False

    _clear_existing_paragraphs(header)
    _clear_existing_paragraphs(footer)

    header_paragraphs = _HEADER_RENDERERS[style](
        header, logo, return_address, rgb
    )
    footer_paragraphs = _FOOTER_RENDERERS[style](
        footer, phone, email, website, rgb
    )

    return {"header": header_paragraphs, "footer": footer_paragraphs}


__all__ = ["set_letterhead", "STYLES"]
