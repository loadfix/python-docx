"""Style audit + consolidation helpers (issue #59).

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterable, List, Optional, Sequence

from docx.enum.style import WD_STYLE_TYPE

if TYPE_CHECKING:
    from docx.document import Document
    from docx.styles.style import BaseStyle
    from docx.text.paragraph import Paragraph


__all__ = [
    "StyleAudit",
    "StyleIssue",
    "audit_styles",
]


@dataclass(frozen=True)
class StyleIssue:
    """One finding in a :class:`StyleAudit`.

    ``severity`` is ``"error"`` / ``"warning"`` / ``"info"``;
    ``rule_id`` is the kebab-case identifier; ``paragraph_index`` is
    |None| for document-level findings.
    """

    severity: str
    rule_id: str
    message: str
    paragraph_index: Optional[int] = None
    style_names: tuple = ()


def _normalise_style_name(name: Optional[str]) -> str:
    """Lower-case + strip + collapse whitespace for fuzzy comparison."""
    if name is None:
        return ""
    cleaned = re.sub(r"\s+", " ", name.strip().lower())
    # -- treat short alias forms (H1, H2 ...) as Heading N for matching --
    m = re.fullmatch(r"h(\d)", cleaned)
    if m:
        return "heading %s" % m.group(1)
    return cleaned


def _font_signature(style: "BaseStyle") -> tuple:
    """Return a hashable signature of a style's visual font properties."""
    font = getattr(style, "font", None)
    if font is None:
        return ()
    try:
        size = font.size
    except Exception:  # pragma: no cover - defensive
        size = None
    try:
        color = font.color.rgb if font.color is not None else None
    except Exception:  # pragma: no cover - defensive
        color = None
    return (
        getattr(font, "name", None),
        int(size) if size is not None else None,
        font.bold,
        font.italic,
        getattr(font, "underline", None),
        str(color) if color is not None else None,
    )


def _paragraph_font_signature(paragraph: "Paragraph") -> Optional[tuple]:
    """Aggregate run-level direct formatting into a signature; |None| when none."""
    runs = list(paragraph.runs)
    if not runs:
        return None
    sigs = []
    for run in runs:
        try:
            color = run.font.color.rgb if run.font.color is not None else None
        except Exception:  # pragma: no cover - defensive
            color = None
        try:
            size = run.font.size
        except Exception:  # pragma: no cover - defensive
            size = None
        sigs.append(
            (
                run.font.name,
                int(size) if size is not None else None,
                run.font.bold,
                run.font.italic,
                run.font.underline,
                str(color) if color is not None else None,
            )
        )
    # -- collapse: keep value if every run agrees, else None --
    if all(s == sigs[0] for s in sigs):
        result = sigs[0]
    else:
        result = tuple(
            s[i] if all(t[i] == s[i] for t in sigs) else None
            for i, s in enumerate(sigs[0:1] * len(sigs[0]))
        )
    if all(v is None for v in result):
        return None
    return result


@dataclass
class StyleAudit:
    """Audit result returned by :meth:`Document.audit_styles`."""

    issues: List[StyleIssue] = field(default_factory=list)
    summary: dict = field(default_factory=dict)
    document: "Document | None" = None

    def __iter__(self):
        return iter(self.issues)

    def __len__(self) -> int:
        return len(self.issues)

    def by_rule(self, rule_id: str) -> List[StyleIssue]:
        """Return every issue whose ``rule_id`` matches ``rule_id``."""
        return [i for i in self.issues if i.rule_id == rule_id]

    def consolidate_styles(
        self,
        canonical: str,
        drop: Sequence[str] = (),
    ) -> int:
        """Rewrite body paragraphs using a style in ``drop`` to use ``canonical``.

        Then delete the dropped styles. Returns the number of paragraphs
        rewritten. Raises :class:`KeyError` when ``canonical`` is not
        defined. Round-trip safe.
        """
        if self.document is None:
            raise RuntimeError(
                "consolidate_styles requires the audit to be bound to a "
                "Document; this audit was constructed without one"
            )
        document = self.document
        styles = document.styles
        if canonical not in styles:
            raise KeyError(
                "canonical style %r is not defined in this document" % canonical
            )
        drop_set = {_normalise_style_name(d) for d in drop}
        canonical_norm = _normalise_style_name(canonical)
        drop_set.discard(canonical_norm)

        rewritten = 0
        for paragraph in document.paragraphs:
            style = paragraph.style
            if style is None:
                continue
            if _normalise_style_name(style.name) in drop_set:
                paragraph.style = canonical
                rewritten += 1

        # -- delete styles --
        for name in drop:
            if name == canonical:
                continue
            try:
                style = styles[name]
            except KeyError:
                continue
            try:
                style.delete()
            except Exception:  # pragma: no cover - defensive
                pass
        return rewritten


# ---------------------------------------------------------------------------
# Detection passes
# ---------------------------------------------------------------------------


def _detect_duplicate_styles(document: "Document") -> Iterable[StyleIssue]:
    """Group styles by normalised name + font signature and flag clusters."""
    groups: dict = {}
    for style in document.styles:
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        norm = _normalise_style_name(style.name)
        sig = _font_signature(style)
        key = (norm, sig)
        groups.setdefault(key, []).append(style.name)
    for (norm, _sig), names in groups.items():
        if len(names) > 1:
            yield StyleIssue(
                severity="info",
                rule_id="duplicate-styles",
                message=(
                    "Styles share visual properties and a similar name "
                    "(%s); consolidate via audit.consolidate_styles" % ", ".join(names)
                ),
                style_names=tuple(names),
            )


def _detect_direct_formatting(
    document: "Document", paragraphs: Sequence["Paragraph"]
) -> Iterable[StyleIssue]:
    """Flag paragraphs whose direct formatting matches an existing style."""
    style_sigs: dict = {}
    for style in document.styles:
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        sig = _font_signature(style)
        if sig and any(v is not None for v in sig):
            style_sigs[sig] = style.name
    for idx, paragraph in enumerate(paragraphs):
        sig = _paragraph_font_signature(paragraph)
        if sig is None:
            continue
        match = style_sigs.get(sig)
        if match is None:
            continue
        # -- only flag when the paragraph's current style differs --
        current = paragraph.style.name if paragraph.style is not None else None
        if current == match:
            continue
        yield StyleIssue(
            severity="info",
            rule_id="direct-formatting",
            message=(
                "Paragraph uses direct formatting matching style %r; consider "
                "applying that style instead" % match
            ),
            paragraph_index=idx,
            style_names=(match,),
        )


def _detect_mixed_fonts(paragraphs: Sequence["Paragraph"]) -> Iterable[StyleIssue]:
    """Flag paragraphs with runs in ≥ 2 different font families."""
    for idx, paragraph in enumerate(paragraphs):
        names = set()
        for run in paragraph.runs:
            name = run.font.name
            if name is not None:
                names.add(name)
        if len(names) >= 2:
            yield StyleIssue(
                severity="warning",
                rule_id="mixed-fonts",
                message=(
                    "Paragraph mixes %d font families (%s)"
                    % (len(names), ", ".join(sorted(names)))
                ),
                paragraph_index=idx,
            )


def _detect_unstyled_paragraphs(
    paragraphs: Sequence["Paragraph"],
) -> Iterable[StyleIssue]:
    """Flag paragraphs that fall back to the default Normal style."""
    for idx, paragraph in enumerate(paragraphs):
        text = (paragraph.text or "").strip()
        if not text:
            continue
        # -- explicit style id of the paragraph; None ⇒ inherits "Normal"
        explicit = paragraph._p.style  # pyright: ignore[reportPrivateUsage]
        if explicit is None:
            yield StyleIssue(
                severity="info",
                rule_id="unstyled-paragraph",
                message=(
                    "Paragraph uses the default Normal style; consider an "
                    "explicit body / heading style"
                ),
                paragraph_index=idx,
            )


def _detect_heading_without_style(
    paragraphs: Sequence["Paragraph"],
) -> Iterable[StyleIssue]:
    """Body paragraph that visually looks like a heading."""
    from docx.lint import _heading_level, _looks_like_heading

    for idx, paragraph in enumerate(paragraphs):
        if _heading_level(paragraph) is not None:
            continue
        if _looks_like_heading(paragraph):
            yield StyleIssue(
                severity="error",
                rule_id="heading-without-style",
                message=(
                    "Paragraph appears to be a heading but does not use a "
                    "Heading style; apply 'Heading N' so it joins the outline"
                ),
                paragraph_index=idx,
            )


def _detect_orphan_styles(
    document: "Document", paragraphs: Sequence["Paragraph"]
) -> Iterable[StyleIssue]:
    """Paragraph styles defined but never used by any body paragraph."""
    used: set = set()
    for paragraph in paragraphs:
        if paragraph.style is not None:
            used.add(paragraph.style.style_id)
    for style in document.styles:
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        if style.builtin:
            continue
        if style.style_id in used:
            continue
        yield StyleIssue(
            severity="info",
            rule_id="orphan-style",
            message="Custom style %r is defined but unused" % style.name,
            style_names=(style.name,),
        )


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def audit_styles(document: "Document") -> StyleAudit:
    """Run every audit pass against ``document`` and return a :class:`StyleAudit`."""
    paragraphs = list(document.paragraphs)
    issues: List[StyleIssue] = []
    issues.extend(_detect_duplicate_styles(document))
    issues.extend(_detect_direct_formatting(document, paragraphs))
    issues.extend(_detect_mixed_fonts(paragraphs))
    issues.extend(_detect_unstyled_paragraphs(paragraphs))
    issues.extend(_detect_heading_without_style(paragraphs))
    issues.extend(_detect_orphan_styles(document, paragraphs))

    summary: dict = {}
    for issue in issues:
        summary[issue.rule_id] = summary.get(issue.rule_id, 0) + 1
    summary.setdefault("total", 0)
    summary["total"] = len(issues)

    issues.sort(
        key=lambda i: (
            i.paragraph_index if i.paragraph_index is not None else 10**9,
            i.rule_id,
        )
    )
    return StyleAudit(issues=issues, summary=summary, document=document)
