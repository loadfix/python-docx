"""Document-lint framework + heading-hierarchy rules (issue #57).

The original framework (issue #57) shipped paragraph-scope rules whose
signature is ``(paragraphs) -> Iterable[LintFinding]``. Issue #15
(Wave-10) extends the registry with accessibility-scope rules that need
visibility into the whole :class:`docx.document.Document` (core
properties, inline shapes, tables, document defaults). Those rules are
marked with ``rule._NEEDS_DOCUMENT = True``; the dispatcher passes
``document`` instead of ``paragraphs`` for any rule that carries the
marker, so the public ``rules=[...]`` surface and the existing
paragraph-scope rule contract stay backward-compatible.

.. versionadded:: 2026.05.13
.. versionchanged:: 2026.05.dev0
   Accessibility rules (``image-no-alt-text``, ``table-no-caption``,
   ``no-language-tag``, ``low-contrast``, ``no-document-title``) added
   for issue #15.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Callable, Iterable, List, Optional, Sequence

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


__all__ = [
    "LintFinding",
    "Severity",
    "lint_document",
    "DEFAULT_RULES",
    "ALL_RULES",
    "ACCESSIBILITY_RULES",
    "rule_heading_skip",
    "rule_heading_multiple_h1",
    "rule_heading_no_h1",
    "rule_heading_direct_formatting",
    "rule_heading_empty",
    "rule_heading_too_long",
    "rule_image_no_alt_text",
    "rule_table_no_caption",
    "rule_no_language_tag",
    "rule_low_contrast",
    "rule_no_document_title",
]


class Severity(str):
    """Severity tokens used by :class:`LintFinding`."""

    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


@dataclass(frozen=True)
class LintFinding:
    """Structured lint result.

    ``severity`` is ``"error"`` / ``"warning"`` / ``"info"``; ``rule_id``
    is the machine-readable kebab-case identifier; ``paragraph_index``
    is |None| for document-level findings.
    """

    severity: str
    paragraph_index: Optional[int]
    rule_id: str
    message: str


def _heading_level(paragraph: "Paragraph") -> Optional[int]:
    """Return the 1-9 heading level for ``paragraph`` or |None|."""
    style = paragraph.style
    if style is None:
        return None
    name = (style.name or "").strip()
    if not name:
        return None
    if name.lower() == "title":
        return 1
    if name.lower().startswith("heading "):
        suffix = name[len("Heading ") :].strip()
        if suffix.isdigit():
            level = int(suffix)
            if 1 <= level <= 9:
                return level
    return None


def _looks_like_heading(paragraph: "Paragraph") -> bool:
    """Heuristic: short body-styled paragraph that is bold or has large font."""
    if _heading_level(paragraph) is not None:
        return False
    text = (paragraph.text or "").strip()
    if not text:
        return False
    if len(text) > 80:
        return False
    if text.endswith("."):
        return False
    if "\n" in text:
        return False
    runs = list(paragraph.runs)
    if not runs:
        return False
    bold_runs = [r for r in runs if r.bold is True]
    has_large_font = False
    for r in runs:
        try:
            size = r.font.size
        except Exception:  # pragma: no cover -- defensive
            size = None
        if size is not None and int(size) > 14 * 12700:  # 14 pt in EMU
            has_large_font = True
            break
    if not bold_runs and not has_large_font:
        return False
    # -- if every text-bearing run is bold or font is bumped, treat
    # -- as a probable heading.
    return len(bold_runs) == len(runs) or has_large_font


# ---------------------------------------------------------------------------
# Built-in rules
# ---------------------------------------------------------------------------


def rule_heading_skip(paragraphs: Sequence["Paragraph"]) -> Iterable[LintFinding]:
    """Flag a heading whose level jumps by more than 1 (e.g. H1 -> H3)."""
    last_level: Optional[int] = None
    for idx, paragraph in enumerate(paragraphs):
        level = _heading_level(paragraph)
        if level is None:
            continue
        if last_level is not None and level > last_level + 1:
            yield LintFinding(
                severity=Severity.ERROR,
                paragraph_index=idx,
                rule_id="heading-skip",
                message=(
                    "Heading level %d follows heading level %d (skipped %d)"
                    % (level, last_level, level - last_level - 1)
                ),
            )
        last_level = level


def rule_heading_multiple_h1(
    paragraphs: Sequence["Paragraph"],
) -> Iterable[LintFinding]:
    """Flag every Heading 1 after the first."""
    seen_first = False
    for idx, paragraph in enumerate(paragraphs):
        if _heading_level(paragraph) == 1:
            if seen_first:
                yield LintFinding(
                    severity=Severity.WARNING,
                    paragraph_index=idx,
                    rule_id="heading-multiple-h1",
                    message="Document contains more than one Heading 1",
                )
            else:
                seen_first = True


def rule_heading_no_h1(paragraphs: Sequence["Paragraph"]) -> Iterable[LintFinding]:
    """Emit an info finding when no Heading 1 is present."""
    for paragraph in paragraphs:
        if _heading_level(paragraph) == 1:
            return
    yield LintFinding(
        severity=Severity.INFO,
        paragraph_index=None,
        rule_id="heading-no-h1",
        message="Document has no Heading 1",
    )


def rule_heading_direct_formatting(
    paragraphs: Sequence["Paragraph"],
) -> Iterable[LintFinding]:
    """Flag body paragraphs that look like headings via direct formatting."""
    for idx, paragraph in enumerate(paragraphs):
        if _looks_like_heading(paragraph):
            yield LintFinding(
                severity=Severity.WARNING,
                paragraph_index=idx,
                rule_id="heading-direct-formatting",
                message=(
                    "Paragraph appears to be a heading (bold/large) but uses "
                    "the body style; apply a Heading style instead"
                ),
            )


def rule_heading_empty(paragraphs: Sequence["Paragraph"]) -> Iterable[LintFinding]:
    """Flag heading paragraphs whose text is empty or whitespace-only."""
    for idx, paragraph in enumerate(paragraphs):
        if _heading_level(paragraph) is None:
            continue
        text = paragraph.text or ""
        if not text.strip():
            yield LintFinding(
                severity=Severity.ERROR,
                paragraph_index=idx,
                rule_id="heading-empty",
                message="Heading paragraph contains no visible text",
            )


def rule_heading_too_long(
    paragraphs: Sequence["Paragraph"],
    *,
    max_chars: int = 120,
) -> Iterable[LintFinding]:
    """Flag heading paragraphs longer than ``max_chars`` characters."""
    for idx, paragraph in enumerate(paragraphs):
        if _heading_level(paragraph) is None:
            continue
        text = (paragraph.text or "").strip()
        if len(text) > max_chars:
            yield LintFinding(
                severity=Severity.WARNING,
                paragraph_index=idx,
                rule_id="heading-too-long",
                message=(
                    "Heading is %d characters long (>%d); consider shortening"
                    % (len(text), max_chars)
                ),
            )


# ---------------------------------------------------------------------------
# Accessibility rules (issue #15) — operate on the Document object so
# they can reach inline shapes, tables, core properties, and the
# document defaults / styles tree. Marked with ``_NEEDS_DOCUMENT`` so
# :func:`lint_document` knows to pass ``document`` instead of the
# paragraph sequence.
# ---------------------------------------------------------------------------


def _needs_document(func: Callable) -> Callable:
    """Mark `func` as a Document-scope rule (sees the whole document)."""
    func._NEEDS_DOCUMENT = True  # type: ignore[attr-defined]
    return func


@_needs_document
def rule_image_no_alt_text(document: "Document") -> Iterable[LintFinding]:
    """Flag every inline image whose ``wp:docPr/@descr`` is missing or empty.

    A picture is considered "described" when its ``alt_text`` is
    non-empty after stripping. Decorative-role pictures (``[decorative]``
    prefix) are exempt — the role declaration is itself an a11y signal.
    """
    try:
        shapes = list(document.inline_shapes)
    except Exception:  # pragma: no cover -- defensive; older fixtures
        return
    for idx, shape in enumerate(shapes):
        try:
            role = getattr(shape, "a11y_role", None)
        except Exception:
            role = None
        if role == "decorative":
            continue
        try:
            alt = shape.alt_text
        except Exception:
            alt = None
        if alt is None or not alt.strip():
            yield LintFinding(
                severity=Severity.ERROR,
                paragraph_index=None,
                rule_id="image-no-alt-text",
                message=(
                    "Inline image #%d has no alt text (wp:docPr/@descr "
                    "missing or empty); add alt text or mark the image "
                    "as decorative" % idx
                ),
            )


@_needs_document
def rule_table_no_caption(document: "Document") -> Iterable[LintFinding]:
    """Flag every top-level table whose ``w:tblCaption/@w:val`` is missing or empty.

    Reads :attr:`docx.table.Table.alt_text` (the OOXML
    ``w:tblPr/w:tblCaption`` element — Word labels it "Title" in the
    Alt Text dialog and refers to it as the table's accessibility
    caption).
    """
    try:
        tables = list(document.tables)
    except Exception:  # pragma: no cover -- defensive
        return
    for idx, table in enumerate(tables):
        caption = None
        try:
            caption = table.alt_text
        except Exception:
            caption = None
        if caption is None or not str(caption).strip():
            yield LintFinding(
                severity=Severity.WARNING,
                paragraph_index=None,
                rule_id="table-no-caption",
                message=(
                    "Table #%d has no caption (w:tblCaption missing or "
                    "empty); set Table.alt_text to a short title for "
                    "screen-reader users" % idx
                ),
            )


def _document_has_lang(document: "Document") -> bool:
    """Return True when *any* ``w:lang/@w:val`` is present at doc/style/run scope."""
    try:
        body = document._element.body  # CT_Body
    except Exception:
        return False
    # -- look anywhere in the document tree for a w:lang carrying a w:val --
    try:
        for el in body.iter():
            tag = getattr(el, "tag", "")
            if isinstance(tag, str) and tag.endswith("}lang"):
                val = el.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                )
                if val:
                    return True
    except Exception:
        pass
    # -- fall through: also scan the styles part (docDefaults / styles) --
    try:
        styles_element = document.styles._element
    except Exception:
        return False
    try:
        for el in styles_element.iter():
            tag = getattr(el, "tag", "")
            if isinstance(tag, str) and tag.endswith("}lang"):
                val = el.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                )
                if val:
                    return True
    except Exception:
        pass
    return False


@_needs_document
def rule_no_language_tag(document: "Document") -> Iterable[LintFinding]:
    """Flag a document with no ``w:lang/@w:val`` anywhere (body, styles, defaults)."""
    if _document_has_lang(document):
        return
    yield LintFinding(
        severity=Severity.WARNING,
        paragraph_index=None,
        rule_id="no-language-tag",
        message=(
            "Document declares no language (no w:lang/@w:val on any "
            "run, paragraph, or document-defaults); set "
            "document.styles['Normal'].font.lang or per-run lang for "
            "assistive-tech support"
        ),
    )


# -- WCAG 2.x relative-luminance approximation. The proper formula
# -- requires sRGB linearisation; we use the cheap approximation
# -- (0.299*r + 0.587*g + 0.114*b) for the heuristic — perceived
# -- luminance is good enough to flag the egregious "yellow on white"
# -- and "light-grey on white" mistakes the rule targets. The 4.5:1
# -- WCAG AA threshold maps to a ~0.18 normalised-luminance gap on
# -- this approximation; we use 0.20 to leave a small safety margin.
def _relative_luminance(rgb: Any) -> float:
    """Return a 0.0..1.0 perceptual-luminance value for an RGBColor-ish triple."""
    try:
        r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    except Exception:
        return 1.0
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


@_needs_document
def rule_low_contrast(document: "Document") -> Iterable[LintFinding]:
    """Flag runs whose explicit text colour has low contrast against white.

    A pure heuristic — python-docx has no layout engine and no theme
    resolver, so the rule only fires on runs with an *explicit*
    ``font.color.rgb``. The check assumes a white page background
    (the Word default) and the WCAG AA 4.5:1 minimum, approximated
    as a normalised-luminance gap of 0.20. Theme-resolved colours are
    skipped to avoid false positives.
    """
    paragraphs = list(document.paragraphs)
    for p_idx, paragraph in enumerate(paragraphs):
        for run in paragraph.runs:
            try:
                rgb = run.font.color.rgb
            except Exception:
                rgb = None
            if rgb is None:
                continue
            text_lum = _relative_luminance(rgb)
            # -- assume white page background ⇒ luminance 1.0 --
            if (1.0 - text_lum) < 0.20:
                yield LintFinding(
                    severity=Severity.INFO,
                    paragraph_index=p_idx,
                    rule_id="low-contrast",
                    message=(
                        "Run text colour #%02X%02X%02X has insufficient "
                        "contrast against a white page background "
                        "(luminance gap %.2f < 0.20); aim for WCAG AA "
                        "(4.5:1)" % (rgb[0], rgb[1], rgb[2], 1.0 - text_lum)
                    ),
                )
                # -- one finding per paragraph is plenty; the next
                # -- offending run in the same paragraph is almost
                # -- always the same colour. --
                break


@_needs_document
def rule_no_document_title(document: "Document") -> Iterable[LintFinding]:
    """Flag a document whose ``cp:coreProperties/dc:title`` is missing or empty."""
    try:
        title = document.core_properties.title
    except Exception:
        title = None
    if title is None or not str(title).strip():
        yield LintFinding(
            severity=Severity.WARNING,
            paragraph_index=None,
            rule_id="no-document-title",
            message=(
                "Document core property 'title' is empty; assistive "
                "technology and search engines use it as the document "
                "name — set document.core_properties.title"
            ),
        )


# Registry --------------------------------------------------------------

#: Heading-hierarchy rules enabled by default when callers pass
#: ``rules=None`` to :meth:`Document.lint`.
DEFAULT_RULES: tuple = (
    rule_heading_skip,
    rule_heading_multiple_h1,
    rule_heading_no_h1,
    rule_heading_direct_formatting,
    rule_heading_empty,
    rule_heading_too_long,
)


#: Accessibility rules (issue #15). Off by default — opt in via
#: ``Document.lint(rules=ACCESSIBILITY_RULES)`` or pass the rule ids.
ACCESSIBILITY_RULES: tuple = (
    rule_image_no_alt_text,
    rule_table_no_caption,
    rule_no_language_tag,
    rule_low_contrast,
    rule_no_document_title,
)


#: Convenience alias for callers that want every shipped rule.
ALL_RULES: tuple = DEFAULT_RULES + ACCESSIBILITY_RULES


_RULE_BY_ID = {
    "heading-skip": rule_heading_skip,
    "heading-multiple-h1": rule_heading_multiple_h1,
    "heading-no-h1": rule_heading_no_h1,
    "heading-direct-formatting": rule_heading_direct_formatting,
    "heading-empty": rule_heading_empty,
    "heading-too-long": rule_heading_too_long,
    "image-no-alt-text": rule_image_no_alt_text,
    "table-no-caption": rule_table_no_caption,
    "no-language-tag": rule_no_language_tag,
    "low-contrast": rule_low_contrast,
    "no-document-title": rule_no_document_title,
}


def _resolve_rule(spec) -> Callable:
    """Translate a string id into a rule callable; passthrough callables."""
    if callable(spec):
        return spec
    if isinstance(spec, str):
        try:
            return _RULE_BY_ID[spec]
        except KeyError:
            raise ValueError(
                "unknown lint rule id %r; expected one of %s"
                % (spec, sorted(_RULE_BY_ID))
            ) from None
    raise TypeError(
        "rule must be a callable or a rule id string; got %r"
        % type(spec).__name__
    )


def lint_document(
    document: "Document",
    rules: Optional[Sequence] = None,
) -> List[LintFinding]:
    """Run ``rules`` against ``document``'s body and return findings.

    ``rules`` may be |None| (defaults), a sequence of callables, or a
    sequence of rule-id strings. Findings sort by
    ``(paragraph_index, rule_id)``, doc-level findings last.

    Each rule callable receives either ``paragraphs`` (the body
    paragraph sequence) or ``document`` itself; the dispatcher checks
    for a ``rule._NEEDS_DOCUMENT`` attribute (set by :func:`_needs_document`)
    and routes accordingly. User-registered rules default to the
    paragraph-scope shape — set ``rule._NEEDS_DOCUMENT = True`` to opt
    into the document-scope shape.
    """
    paragraphs = list(document.paragraphs)
    selected = DEFAULT_RULES if rules is None else [_resolve_rule(r) for r in rules]
    findings: List[LintFinding] = []
    for rule in selected:
        if getattr(rule, "_NEEDS_DOCUMENT", False):
            iterable = rule(document)
        else:
            iterable = rule(paragraphs)
        for finding in iterable:
            findings.append(finding)
    findings.sort(
        key=lambda f: (
            f.paragraph_index if f.paragraph_index is not None else 10**9,
            f.rule_id,
        )
    )
    return findings
