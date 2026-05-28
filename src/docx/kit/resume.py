"""Resume / CV template family — chronological / functional / technical.

Closes #63.

Three template factories that build a fully-styled |Document| from
plain-Python keyword arguments:

- :func:`resume_chronological` — reverse-chronological work history.
- :func:`resume_functional` — skills / focus-area first, condensed
  history below.
- :func:`resume_technical` — projects + tech-stack first, suited to
  engineers.

Three visual *styles* shape typography and section headings: ``modern``
(large coloured name, uppercased coloured section headings),
``classic`` (centered name, italic subtitle, conservative
``Heading 1`` sections), and ``minimal`` (template defaults only — no
colour overrides). Each factory accepts ``style="modern"`` /
``"classic"`` / ``"minimal"`` (default ``"modern"``).

Every helper composes only python-docx's *public* API — no
``_element`` / ``oxml`` reach-down. Recognised contact links
(``email``, ``linkedin``, ``github``, ``website``) are emitted as
:class:`Hyperlink` runs styled with Word's ``Hyperlink`` character
style when available.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, List, Mapping, Optional, Sequence, Tuple, Union

from docx import Document as _new_document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# -- Built-in style names this kit reaches for; falls back to ``Normal``
# -- when a custom-loaded template lacks one.
_STYLE_TITLE = "Title"
_STYLE_SUBTITLE = "Subtitle"
_STYLE_HEADING_1 = "Heading 1"
_STYLE_HEADING_2 = "Heading 2"
_STYLE_NORMAL = "Normal"
_STYLE_LIST_BULLET = "List Bullet"
_STYLE_HYPERLINK = "Hyperlink"

#: Built-in visual styles.
STYLES: Tuple[str, ...] = ("modern", "classic", "minimal")
#: Built-in template factory names.
TEMPLATES: Tuple[str, ...] = ("chronological", "functional", "technical")

# -- Visual identity. Modern's accent matches the letterhead/chapter
# -- palettes so kit helpers compose visually when chained.
_MODERN_ACCENT = RGBColor(0x1F, 0x4E, 0x79)  # deep blue
_NAME_SIZE = Pt(28)
_TITLE_SIZE = Pt(14)
_CONTACT_SIZE = Pt(10)
_SECTION_HEADING_SIZE = Pt(13)
_BODY_SIZE = Pt(11)
_DATE_RANGE_SIZE = Pt(10)

# -- Recognised contact-link kinds rendered in this fixed order.
_CONTACT_ORDER = ("email", "phone", "linkedin", "github", "website", "location")


# -- Style helpers --------------------------------------------------------


def _has_style(document, style_name):
    # type: (Document, str) -> bool
    """Return |True| when ``document`` defines a style named ``style_name``."""
    try:
        styles = document.styles
    except Exception:  # pragma: no cover - defensive
        return False
    try:
        styles[style_name]
        return True
    except KeyError:
        return False


def _resolve_style(document, preferred):
    # type: (Document, str) -> str
    """Return ``preferred`` when it exists on ``document``, else ``"Normal"``."""
    return preferred if _has_style(document, preferred) else _STYLE_NORMAL


def _apply_color(run, rgb):
    # type: (Any, Optional[RGBColor]) -> None
    """Set the run's font colour to ``rgb`` when supplied; no-op otherwise."""
    if rgb is not None:
        run.font.color.rgb = rgb


def _accent_for(style):
    # type: (str) -> Optional[RGBColor]
    """Return the accent colour to apply for ``style``, or |None| for plain styles."""
    return _MODERN_ACCENT if style == "modern" else None


def _validate_style(style):
    # type: (str) -> None
    """Raise :class:`ValueError` when ``style`` is not one of the three built-ins."""
    if style not in STYLES:
        raise ValueError(
            "style must be one of %s; got %r"
            % (", ".join(repr(s) for s in STYLES), style)
        )


# -- Contact-line URL formatters ------------------------------------------


def _format_email(value):
    # type: (str) -> Tuple[str, str]
    return f"mailto:{value}", value


def _format_linkedin(value):
    # type: (str) -> Tuple[str, str]
    """Bare handles like ``"in/x"`` or ``"x"`` -> ``https://linkedin.com/in/x``."""
    if "://" in value:
        return value, value
    handle = value.lstrip("/")
    if not handle.startswith("in/"):
        handle = f"in/{handle}"
    return f"https://linkedin.com/{handle}", value


def _format_github(value):
    # type: (str) -> Tuple[str, str]
    if "://" in value:
        return value, value
    return f"https://github.com/{value.lstrip('/')}", value


def _format_website(value):
    # type: (str) -> Tuple[str, str]
    """Bare domains -> ``https://...``; full URLs pass through unchanged."""
    href = value if "://" in value else f"https://{value}"
    return href, value


_CONTACT_FORMATTERS = {
    "email": _format_email,
    "linkedin": _format_linkedin,
    "github": _format_github,
    "website": _format_website,
}


# -- Run / paragraph helpers ----------------------------------------------


def _add_hyperlink_run(paragraph, href, text, rgb, size, document):
    # type: (Paragraph, str, str, Optional[RGBColor], Optional[int], Document) -> None
    """Append ``text`` as a hyperlink to ``paragraph`` with letterhead-style font."""
    style = _STYLE_HYPERLINK if _has_style(document, _STYLE_HYPERLINK) else None
    link = paragraph.add_hyperlink(url=href, text=text, style=style)
    for run in link.runs:
        if size is not None:
            run.font.size = size
        _apply_color(run, rgb)


def _styled_run(paragraph, text, rgb=None, size=None, bold=False, italic=False):
    # type: (Paragraph, str, Optional[RGBColor], Optional[int], bool, bool) -> Any
    """Append a run carrying ``text`` with the standard formatting block."""
    run = paragraph.add_run(text)
    if size is not None:
        run.font.size = size
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    _apply_color(run, rgb)
    return run


def _bullet_paragraphs(document, items):
    # type: (Document, Sequence[Any]) -> List[Paragraph]
    """Append one bullet paragraph per non-empty item; returns the paragraphs."""
    if not items:
        return []
    bullet_style = (
        _STYLE_LIST_BULLET if _has_style(document, _STYLE_LIST_BULLET) else None
    )
    out: List[Paragraph] = []
    for item in items:
        if not item:
            continue
        text = str(item)
        if bullet_style is not None:
            out.append(document.add_paragraph(text, style=bullet_style))
        else:
            out.append(document.add_paragraph(f"• {text}"))
    return out


# -- Header (name + title + contact line) --------------------------------


def _add_contact_line(document, contact, rgb, size):
    # type: (Document, Mapping[str, str], Optional[RGBColor], Optional[int]) -> Optional[Paragraph]
    """Append a centered contact-line paragraph; returns it (or |None| when empty)."""
    if not contact:
        return None
    para = document.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sep = "  ·  "
    # -- Recognised keys first, in fixed order; then any unknown keys
    # -- in caller-insertion order so the helper stays extensible.
    pieces = [(k, contact[k]) for k in _CONTACT_ORDER if contact.get(k)]
    for key, value in contact.items():
        if key not in _CONTACT_ORDER and value:
            pieces.append((key, value))

    for index, (kind, value) in enumerate(pieces):
        if index > 0:
            _styled_run(para, sep, rgb=rgb, size=size)
        formatter = _CONTACT_FORMATTERS.get(kind)
        if formatter is not None:
            href, display = formatter(value)
            _add_hyperlink_run(para, href, display, rgb, size, document)
        else:
            _styled_run(para, value, rgb=rgb, size=size)
    return para


def _add_name_block(document, name, title, style):
    # type: (Document, str, Optional[str], str) -> List[Paragraph]
    """Append the name + title block; layout depends on ``style``.

    - ``modern`` — large left-aligned accent-coloured name, smaller
      subtitle line.
    - ``classic`` — centered ``Title`` paragraph, italic centered
      ``Subtitle`` paragraph.
    - ``minimal`` — bold name, plain subtitle; no colour overrides.
    """
    paragraphs: List[Paragraph] = []
    rgb = _accent_for(style)

    if style == "modern":
        name_para = document.add_paragraph(style=_resolve_style(document, _STYLE_TITLE))
        _styled_run(name_para, name, rgb=rgb, size=_NAME_SIZE, bold=True)
        paragraphs.append(name_para)
        if title:
            sub_para = document.add_paragraph(
                style=_resolve_style(document, _STYLE_SUBTITLE)
            )
            _styled_run(sub_para, title, rgb=rgb, size=_TITLE_SIZE)
            paragraphs.append(sub_para)
    elif style == "classic":
        name_para = document.add_paragraph(
            name, style=_resolve_style(document, _STYLE_TITLE)
        )
        name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraphs.append(name_para)
        if title:
            sub_para = document.add_paragraph(
                title, style=_resolve_style(document, _STYLE_SUBTITLE)
            )
            sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in sub_para.runs:
                run.italic = True
            paragraphs.append(sub_para)
    else:  # minimal
        name_para = document.add_paragraph()
        _styled_run(name_para, name, size=_NAME_SIZE, bold=True)
        paragraphs.append(name_para)
        if title:
            paragraphs.append(document.add_paragraph(title))
    return paragraphs


# -- Section heading + reusable section helpers --------------------------


def _add_section_heading(document, text, style):
    # type: (Document, str, str) -> Paragraph
    """Append a section heading whose look matches ``style``."""
    rgb = _accent_for(style)
    if style == "modern":
        para = document.add_paragraph(style=_resolve_style(document, _STYLE_HEADING_2))
        _styled_run(para, text.upper(), rgb=rgb, size=_SECTION_HEADING_SIZE, bold=True)
        return para
    if style == "classic":
        para = document.add_paragraph(
            text, style=_resolve_style(document, _STYLE_HEADING_1)
        )
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return para
    return document.add_paragraph(text, style=_resolve_style(document, _STYLE_HEADING_2))


def _add_summary(document, summary, style):
    # type: (Document, Optional[str], str) -> List[Paragraph]
    """Append the optional summary block (heading + paragraph)."""
    if not summary:
        return []
    return [
        _add_section_heading(document, "Summary", style),
        document.add_paragraph(summary),
    ]


def _format_date_range(start, end):
    # type: (Optional[str], Optional[str]) -> str
    """Render ``start``–``end`` as ``"start – end"`` (en-dash)."""
    parts: List[str] = []
    if start:
        parts.append(str(start))
    if end:
        parts.append(str(end))
    return " – ".join(parts)


def _add_role_header(document, primary, secondary, trailing, style, trailing_kind="text"):
    # type: (Document, str, str, str, str, str) -> Paragraph
    """Append a one-line "primary · secondary <tab> trailing" header.

    Used for both experience entries (``role · company <tab> dates``)
    and project entries (``name · role <tab> url``). ``trailing_kind``
    is ``"text"`` (italic date range / plain trailing) or ``"hyperlink"``
    (rendered as a clickable URL).
    """
    rgb = _accent_for(style)
    para = document.add_paragraph()
    if primary:
        _styled_run(para, primary, rgb=rgb, size=_BODY_SIZE, bold=True)
    if secondary:
        if primary:
            _styled_run(para, " · ", size=_BODY_SIZE)
        _styled_run(para, secondary, size=_BODY_SIZE, italic=True)
    if trailing:
        if primary or secondary:
            _styled_run(para, "\t", size=_BODY_SIZE)
        if trailing_kind == "hyperlink":
            href, display = _format_website(trailing)
            _add_hyperlink_run(para, href, display, rgb, _DATE_RANGE_SIZE, document)
        else:
            _styled_run(para, trailing, size=_DATE_RANGE_SIZE, italic=True)
    return para


def _add_experience(document, experience, style, heading="Experience"):
    # type: (Document, Sequence[Mapping[str, Any]], str, str) -> List[Paragraph]
    """Append the experience block (heading + one entry per item)."""
    if not experience:
        return []
    paragraphs: List[Paragraph] = [_add_section_heading(document, heading, style)]
    for entry in experience:
        company = str(entry.get("company") or "")
        role = str(entry.get("title") or entry.get("role") or "")
        date_range = _format_date_range(
            entry.get("start") or entry.get("from"),
            entry.get("end") or entry.get("to"),
        )
        paragraphs.append(_add_role_header(document, role, company, date_range, style))
        paragraphs.extend(
            _bullet_paragraphs(
                document, entry.get("bullets") or entry.get("achievements") or []
            )
        )
    return paragraphs


def _add_education(document, education, style):
    # type: (Document, Sequence[Mapping[str, Any]], str) -> List[Paragraph]
    """Append the education block (heading + one entry per item).

    Each entry: ``degree · school <tab> year``; an optional ``details``
    key becomes a follow-up plain paragraph.
    """
    if not education:
        return []
    paragraphs: List[Paragraph] = [_add_section_heading(document, "Education", style)]
    for entry in education:
        year = entry.get("year")
        paragraphs.append(
            _add_role_header(
                document,
                str(entry.get("degree") or ""),
                str(entry.get("school") or ""),
                "" if year is None else str(year),
                style,
            )
        )
        details = entry.get("details")
        if details:
            paragraphs.append(document.add_paragraph(str(details)))
    return paragraphs


def _add_categorised_or_flat(document, items, heading, style):
    # type: (Document, Union[Sequence[str], Mapping[str, Sequence[str]], None], str, str) -> List[Paragraph]
    """Render ``items`` under ``heading`` — flat list or category map.

    Used by the ``skills`` and ``tech_stack`` sections. A flat sequence
    becomes a single comma-joined paragraph; a mapping becomes one
    bolded "Category: items" paragraph per entry.
    """
    if not items:
        return []
    paragraphs: List[Paragraph] = [_add_section_heading(document, heading, style)]
    rgb = _accent_for(style)
    if isinstance(items, Mapping):
        for category, values in items.items():
            if not values:
                continue
            para = document.add_paragraph()
            _styled_run(para, f"{category}: ", rgb=rgb, size=_BODY_SIZE, bold=True)
            _styled_run(para, ", ".join(str(v) for v in values), size=_BODY_SIZE)
            paragraphs.append(para)
    else:
        paragraphs.append(
            document.add_paragraph(", ".join(str(item) for item in items))
        )
    return paragraphs


def _add_focus_areas(document, focus_areas, style):
    # type: (Document, Sequence[str], str) -> List[Paragraph]
    """Append the focus-areas block (heading + bullet per area)."""
    if not focus_areas:
        return []
    paragraphs: List[Paragraph] = [
        _add_section_heading(document, "Areas of Expertise", style)
    ]
    paragraphs.extend(_bullet_paragraphs(document, focus_areas))
    return paragraphs


def _add_projects(document, projects, style):
    # type: (Document, Sequence[Mapping[str, Any]], str) -> List[Paragraph]
    """Append the projects block.

    Per-entry recognised keys: ``name``, ``role``, ``url``, ``tech``,
    ``bullets`` / ``achievements``. ``url`` becomes a hyperlink in the
    header line; ``tech`` (string or sequence) renders as a comma-joined
    italic line below the header.
    """
    if not projects:
        return []
    paragraphs: List[Paragraph] = [_add_section_heading(document, "Projects", style)]
    for entry in projects:
        url = entry.get("url")
        paragraphs.append(
            _add_role_header(
                document,
                str(entry.get("name") or ""),
                str(entry.get("role") or ""),
                str(url) if url else "",
                style,
                trailing_kind="hyperlink" if url else "text",
            )
        )
        tech = entry.get("tech")
        if tech:
            tech_text = tech if isinstance(tech, str) else ", ".join(str(t) for t in tech)
            tech_para = document.add_paragraph()
            _styled_run(tech_para, tech_text, size=_DATE_RANGE_SIZE, italic=True)
            paragraphs.append(tech_para)
        paragraphs.extend(
            _bullet_paragraphs(
                document, entry.get("bullets") or entry.get("achievements") or []
            )
        )
    return paragraphs


def _build_header(document, name, title, contact, style):
    # type: (Document, str, Optional[str], Optional[Mapping[str, str]], str) -> List[Paragraph]
    """Append the resume header (name + title + contact line) shared by all factories."""
    paragraphs = _add_name_block(document, name, title, style)
    contact_para = _add_contact_line(
        document, contact or {}, _accent_for(style), _CONTACT_SIZE
    )
    if contact_para is not None:
        paragraphs.append(contact_para)
    return paragraphs


# -- Public template factories -------------------------------------------


def resume_chronological(
    name,
    title=None,
    contact=None,
    summary=None,
    experience=None,
    education=None,
    skills=None,
    style="modern",
):
    # type: (str, Optional[str], Optional[Mapping[str, str]], Optional[str], Optional[Sequence[Mapping[str, Any]]], Optional[Sequence[Mapping[str, Any]]], Union[Sequence[str], Mapping[str, Sequence[str]], None], str) -> Document
    """Build and return a reverse-chronological resume |Document|.

    Section order: name + title + contact, summary, experience,
    education, skills. Pass ``experience`` entries in
    reverse-chronological order; the helper does *not* re-sort.

    ``contact`` recognises the keys ``email``, ``phone``, ``linkedin``,
    ``github``, ``website``, ``location``; recognised link kinds become
    hyperlinks. Unrecognised keys render verbatim.

    Each ``experience`` entry recognises ``company``, ``title`` /
    ``role``, ``start`` / ``from``, ``end`` / ``to``, and
    ``bullets`` / ``achievements``.

    Each ``education`` entry recognises ``school``, ``degree``,
    ``year``, ``details``.

    ``skills`` accepts a flat sequence (rendered as one comma-joined
    line) or a category mapping (one bolded line per category).

    ``style`` is one of ``"modern"`` (default), ``"classic"``, or
    ``"minimal"``.

    Raises :class:`ValueError` when ``name`` is empty or ``style`` is
    not one of the three built-ins.

    .. versionadded:: 2026.05.29
    """
    if not name:
        raise ValueError("name must be a non-empty string")
    _validate_style(style)

    document = _new_document()
    _build_header(document, name, title, contact, style)
    _add_summary(document, summary, style)
    _add_experience(document, experience or [], style)
    _add_education(document, education or [], style)
    _add_categorised_or_flat(document, skills, "Skills", style)
    return document


def resume_functional(
    name,
    title=None,
    contact=None,
    summary=None,
    focus_areas=None,
    experience=None,
    education=None,
    skills=None,
    style="modern",
):
    # type: (str, Optional[str], Optional[Mapping[str, str]], Optional[str], Optional[Sequence[str]], Optional[Sequence[Mapping[str, Any]]], Optional[Sequence[Mapping[str, Any]]], Union[Sequence[str], Mapping[str, Sequence[str]], None], str) -> Document
    """Build and return a functional resume |Document|.

    Functional resumes lead with skills and focus areas, then condense
    work history below — appropriate for career-changers and
    consultants whose chronology is less compelling than their
    capabilities.

    Section order: name + title + contact, summary, areas of expertise
    (``focus_areas`` rendered as bullets), skills, experience,
    education.

    See :func:`resume_chronological` for the shape of ``contact``,
    ``experience``, ``education``, and ``skills``.

    ``style`` is one of ``"modern"`` (default), ``"classic"``, or
    ``"minimal"``.

    Raises :class:`ValueError` when ``name`` is empty or ``style`` is
    not one of the three built-ins.

    .. versionadded:: 2026.05.29
    """
    if not name:
        raise ValueError("name must be a non-empty string")
    _validate_style(style)

    document = _new_document()
    _build_header(document, name, title, contact, style)
    _add_summary(document, summary, style)
    _add_focus_areas(document, focus_areas or [], style)
    _add_categorised_or_flat(document, skills, "Skills", style)
    _add_experience(document, experience or [], style)
    _add_education(document, education or [], style)
    return document


def resume_technical(
    name,
    title=None,
    contact=None,
    summary=None,
    projects=None,
    tech_stack=None,
    experience=None,
    education=None,
    skills=None,
    style="modern",
):
    # type: (str, Optional[str], Optional[Mapping[str, str]], Optional[str], Optional[Sequence[Mapping[str, Any]]], Union[Sequence[str], Mapping[str, Sequence[str]], None], Optional[Sequence[Mapping[str, Any]]], Optional[Sequence[Mapping[str, Any]]], Union[Sequence[str], Mapping[str, Sequence[str]], None], str) -> Document
    """Build and return a technical resume |Document|.

    Technical resumes lead with notable projects and a tech-stack
    matrix — appropriate for engineers whose work product is best
    illustrated by what they shipped.

    Section order: name + title + contact, summary, projects, tech
    stack (heading "Technical Skills"), experience, education, skills.

    Each ``projects`` entry recognises ``name``, ``role``, ``url``,
    ``tech`` (string or sequence), and ``bullets`` / ``achievements``.

    ``tech_stack`` accepts a flat sequence or a category mapping
    (rendered separately from ``skills`` so a technical resume can
    carry both a hard-skills matrix and a free-form skills line).

    See :func:`resume_chronological` for the shape of ``contact``,
    ``experience``, ``education``, and ``skills``.

    ``style`` is one of ``"modern"`` (default), ``"classic"``, or
    ``"minimal"``.

    Raises :class:`ValueError` when ``name`` is empty or ``style`` is
    not one of the three built-ins.

    .. versionadded:: 2026.05.29
    """
    if not name:
        raise ValueError("name must be a non-empty string")
    _validate_style(style)

    document = _new_document()
    _build_header(document, name, title, contact, style)
    _add_summary(document, summary, style)
    _add_projects(document, projects or [], style)
    _add_categorised_or_flat(document, tech_stack, "Technical Skills", style)
    _add_experience(document, experience or [], style)
    _add_education(document, education or [], style)
    _add_categorised_or_flat(document, skills, "Skills", style)
    return document


__all__ = [
    "STYLES",
    "TEMPLATES",
    "resume_chronological",
    "resume_functional",
    "resume_technical",
]
