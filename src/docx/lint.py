"""Document-lint framework + heading-hierarchy rules (issue #57).

Provides :class:`LintFinding` (dataclass with ``severity``,
``paragraph_index``, ``rule_id``, and ``message``) plus :func:`lint`
which walks a |Document| body and yields findings according to a list
of rule callables. The heading-hierarchy ruleset ships preconfigured;
callers can pass their own list to extend or restrict the checks.

This is the minimal viable lint surface — meant to give authors and
review tooling a structured way to flag accessibility / readability
issues that would otherwise live in human review notes. New rules are
ordinary callables ``(paragraphs) -> Iterable[LintFinding]`` so adding
one is as easy as a function.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Callable, Iterable, List, Optional, Sequence

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


__all__ = [
    "LintFinding",
    "Severity",
    "lint_document",
    "DEFAULT_RULES",
    "ALL_RULES",
    "rule_heading_skip",
    "rule_heading_multiple_h1",
    "rule_heading_no_h1",
    "rule_heading_direct_formatting",
    "rule_heading_empty",
    "rule_heading_too_long",
]


class Severity(str):
    """String subclass exposing the supported severity tokens.

    Stored as plain strings (``"error"`` / ``"warning"`` / ``"info"``)
    so JSON-style consumers don't need to translate an enum.
    """

    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


@dataclass(frozen=True)
class LintFinding:
    """Structured lint result.

    ``severity`` is one of ``"error"`` / ``"warning"`` / ``"info"``.
    ``paragraph_index`` is the zero-based index into ``document.paragraphs``
    (top-level body paragraphs only); |None| when the finding is
    document-level (e.g. "no Heading 1 anywhere"). ``rule_id`` is the
    machine-readable identifier (kebab-case) and ``message`` is the
    human-readable description.
    """

    severity: str
    paragraph_index: Optional[int]
    rule_id: str
    message: str


def _heading_level(paragraph: "Paragraph") -> Optional[int]:
    """Return the 1-9 heading level for ``paragraph`` or |None|.

    ``"Title"`` and aliases (``"Heading"`` without trailing digit) are
    treated as level 1 for hierarchy purposes; styles whose names don't
    match the ``"Heading N"`` shape return |None|.
    """
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
    """Heuristic: paragraph is short, bold, and not styled as a heading.

    Used by ``heading-direct-formatting`` and ``heading-without-style``
    rules. A "looks like" heading is a body-styled paragraph whose
    visible runs are entirely bold *or* whose font size is larger than
    14 pt, and whose text is a single short line (≤ 80 chars, no
    trailing period).
    """
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
    """Flag a heading whose level jumps by more than 1 from the prior.

    For example, an ``H1`` followed by an ``H3`` is reported because
    screen-readers rely on contiguous heading levels.
    """
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


#: Convenience alias used by callers that want every shipped rule.
ALL_RULES: tuple = DEFAULT_RULES


_RULE_BY_ID = {
    "heading-skip": rule_heading_skip,
    "heading-multiple-h1": rule_heading_multiple_h1,
    "heading-no-h1": rule_heading_no_h1,
    "heading-direct-formatting": rule_heading_direct_formatting,
    "heading-empty": rule_heading_empty,
    "heading-too-long": rule_heading_too_long,
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
    """Run ``rules`` against the body of ``document`` and return findings.

    ``rules`` may be |None| (uses :data:`DEFAULT_RULES`), a sequence of
    rule callables, or a sequence of rule-id strings. Callable rules
    receive the list of top-level body paragraphs and return an
    iterable of :class:`LintFinding` instances. Findings appear in
    ``(paragraph_index, rule_id)`` order, with document-level findings
    (no paragraph index) sorted last.
    """
    paragraphs = list(document.paragraphs)
    selected = DEFAULT_RULES if rules is None else [_resolve_rule(r) for r in rules]
    findings: List[LintFinding] = []
    for rule in selected:
        for finding in rule(paragraphs):
            findings.append(finding)
    findings.sort(
        key=lambda f: (
            f.paragraph_index if f.paragraph_index is not None else 10**9,
            f.rule_id,
        )
    )
    return findings
