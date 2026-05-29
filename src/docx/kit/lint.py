"""Visual / structural lint rules with auto-fix suggestions for ``.docx`` documents.

Closes #304.

This module ships a small, opinionated linter for the kind of rough-edge
defects that survive a Word author's last pass and slip into the printed
output: stray double spaces, lonely tabs, blank-line drifts, missing alt
text, leftover ``[TBD]`` placeholders, mixed straight / smart quotes,
and so on. The lint surface is intentionally distinct from the
structural-accessibility ``docx.lint`` module — that one focuses on
heading hierarchy and a11y; this one is the *visual / micro-typography*
counterpart.

Two-stage workflow::

    from docx import Document
    from docx.kit import lint

    doc = Document("draft.docx")
    report = lint.lint(doc)

    for finding in report.findings:
        print(finding.rule, finding.severity, finding.message)
        if finding.autofix_available:
            print("  fix:", finding.autofix_description)

    # Apply every available autofix in one call.
    report.autofix()
    doc.save("clean.docx")

Stage one (``lint``) is *read-only* — every check inspects the document
without mutation and returns a :class:`Finding`. Stage two
(``LintReport.autofix``) is the only mutation step; callers can opt out
per-rule via the ``rules=[...]`` filter.

Built-in rules (eleven total):

* ``multiple-spaces`` (warning, autofix) — two or more consecutive
  spaces inside a run; collapses to a single space.
* ``trailing-whitespace`` (warning, autofix) — paragraph ends with a
  whitespace character; trims the trailing whitespace.
* ``tab-instead-of-indent`` (warning, autofix) — paragraph starts with
  a literal ``\\t`` character; removes the leading tab.
* ``mixed-quotes`` (info, no-fix) — paragraph mixes "smart" and
  "straight" quote characters (manual review — auto-converting can
  destroy intentional code samples).
* ``empty-paragraph`` (info, autofix) — consecutive empty / whitespace-
  only paragraphs; keeps the first, removes the rest.
* ``inconsistent-heading-levels`` (warning, no-fix) — heading skips a
  level (e.g. H1 then H3). Manual fix only — automatically renumbering
  changes the table of contents.
* ``missing-alt-text`` (warning, no-fix) — inline image without an alt
  text attribute. Manual fix only — alt text is meaning-bearing and
  should be authored, not generated.
* ``mixed-fonts`` (info, no-fix) — paragraph runs use multiple font
  families. Manual review only — sometimes intentional (code spans,
  emphasis runs).
* ``missing-document-title`` (info, autofix-from-filename) — core
  property ``title`` is empty; autofix sets it to the document
  filename's stem when one is known.
* ``over-long-paragraph`` (info, no-fix) — paragraph longer than 1000
  characters. Manual review only — splitting may break list / TOC
  numbering.
* ``placeholder-text`` (warning, no-fix) — paragraph still contains
  ``[PLACEHOLDER]`` / ``[TBD]`` / ``Lorem ipsum`` sentinels. Manual
  fix — autoreplace cannot guess the intended replacement.

Custom rules plug in via :func:`register_rule`::

    def check_no_emoji(doc):
        for i, p in enumerate(doc.paragraphs):
            if any(ord(c) > 0x1F000 for c in p.text):
                yield Finding(
                    rule="no-emoji", severity="info",
                    message="paragraph contains emoji",
                    paragraph_index=i, autofix_available=False,
                )

    lint.register_rule("no-emoji", check_no_emoji)

The custom rule's ``check_callback`` may be a generator yielding
:class:`Finding` instances or a function returning a
:class:`~typing.Sequence` of them. The optional ``autofix_callback``
takes ``(document, finding)`` and returns ``True`` when it applied the
fix and ``False`` otherwise; ``register_rule`` does not require the
autofix callback when the rule is read-only.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import os
import re
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import (
    TYPE_CHECKING,
    Any,
    Callable,
    Dict,
    Iterable,
    Iterator,
    List,
    Optional,
    Sequence,
    Tuple,
    Union,
)

if TYPE_CHECKING:  # pragma: no cover - import-time hints only
    from docx.document import Document
    from docx.shape import InlineShape
    from docx.text.paragraph import Paragraph


__all__ = [
    "Finding",
    "LintReport",
    "Rule",
    "lint",
    "register_rule",
    "unregister_rule",
    "registered_rules",
    "BUILTIN_RULES",
    "SEVERITIES",
]


# ---------------------------------------------------------------------------
# Public dataclasses
# ---------------------------------------------------------------------------


SEVERITIES: Tuple[str, ...] = ("error", "warning", "info")
"""The three severity tiers, ordered most-urgent first."""


@dataclass(frozen=True)
class Finding:
    """A single lint finding surfaced by a rule.

    Attributes
    ----------
    rule
        The rule identifier (e.g. ``"multiple-spaces"``).
    severity
        One of ``"error"`` / ``"warning"`` / ``"info"``.
    message
        Human-readable explanation of the issue, suitable for printing.
    paragraph_index
        Index of the paragraph the finding applies to, or ``None`` when
        the finding is document-level (e.g. ``missing-document-title``)
        or table-scoped (in which case ``location`` carries the locator).
    autofix_available
        ``True`` when the rule registered an autofix callback *and* the
        callback can apply for this finding.
    autofix_description
        Short human-readable description of what the autofix would do,
        or ``None`` when no autofix is available.
    location
        Optional human-readable locator (e.g. ``"table 2 row 3 cell 1"``)
        used when ``paragraph_index`` is not the right scope.
    """

    rule: str
    severity: str
    message: str
    paragraph_index: Optional[int] = None
    autofix_available: bool = False
    autofix_description: Optional[str] = None
    location: Optional[str] = None


@dataclass
class Rule:
    """A registered lint rule.

    Internal — instances are created by :func:`register_rule` and the
    built-in rule loader. The dataclass is exposed so callers can
    introspect ``registered_rules()`` output.
    """

    name: str
    check: Callable[["Document"], Iterable[Finding]]
    autofix: Optional[Callable[["Document", Finding], bool]] = None


# ---------------------------------------------------------------------------
# Rule registry
# ---------------------------------------------------------------------------


_REGISTRY: "OrderedDict[str, Rule]" = OrderedDict()


def register_rule(
    name: str,
    check_callback: Callable[["Document"], Iterable[Finding]],
    autofix_callback: Optional[Callable[["Document", Finding], bool]] = None,
) -> Rule:
    """Register a custom rule under *name*.

    *check_callback* receives the :class:`~docx.document.Document` and
    returns (or yields) :class:`Finding` instances. *autofix_callback*,
    when supplied, takes ``(document, finding)`` and returns ``True``
    when it applied the fix.

    Re-registering an existing rule replaces the previous entry. Returns
    the new :class:`Rule` so callers can chain.
    """

    if not isinstance(name, str) or not name:
        raise ValueError("rule name must be a non-empty string")
    if not callable(check_callback):
        raise TypeError("check_callback must be callable")
    if autofix_callback is not None and not callable(autofix_callback):
        raise TypeError("autofix_callback must be callable when provided")
    rule = Rule(name=name, check=check_callback, autofix=autofix_callback)
    _REGISTRY[name] = rule
    return rule


def unregister_rule(name: str) -> bool:
    """Remove the rule named *name* from the registry.

    Returns ``True`` when a rule was removed, ``False`` when no rule of
    that name was registered. Built-in rules can be unregistered; call
    :func:`_install_builtin_rules` (or import this module fresh) to
    restore them.
    """

    return _REGISTRY.pop(name, None) is not None


def registered_rules() -> Tuple[str, ...]:
    """Return the names of every currently-registered rule, in registration order."""

    return tuple(_REGISTRY.keys())


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def lint(document: "Document") -> "LintReport":
    """Run every registered rule against *document* and return a :class:`LintReport`.

    The returned report carries the ``findings`` list (in document
    order, then rule order) and exposes an :meth:`LintReport.autofix`
    method that mutates the document in place.

    .. versionadded:: 2026.05.29
    """

    findings: List[Finding] = []
    finding_to_rule: Dict[int, str] = {}
    for rule in _REGISTRY.values():
        emitted = rule.check(document)
        if emitted is None:
            continue
        for f in emitted:
            if not isinstance(f, Finding):  # pragma: no cover - defensive
                raise TypeError(
                    f"rule {rule.name!r} yielded a non-Finding object: {f!r}"
                )
            findings.append(f)
            finding_to_rule[id(f)] = rule.name
    findings.sort(key=_finding_sort_key)
    return LintReport(document=document, findings=findings)


def _finding_sort_key(f: Finding) -> Tuple[int, int, str]:
    """Sort findings document-first, then by rule name."""
    # ``None`` paragraph_index sorts last (document-level findings).
    idx = f.paragraph_index if f.paragraph_index is not None else 10**9
    severity_order = {"error": 0, "warning": 1, "info": 2}.get(f.severity, 3)
    return (idx, severity_order, f.rule)


# ---------------------------------------------------------------------------
# LintReport
# ---------------------------------------------------------------------------


@dataclass
class LintReport:
    """Result of running :func:`lint` against a document.

    The report is *bound* to the document — calling :meth:`autofix`
    mutates the same document instance the report was generated from.
    Callers may inspect :attr:`findings` freely (read-only is the
    intended use); rerun :func:`lint` after mutations to refresh.
    """

    document: "Document"
    findings: List[Finding] = field(default_factory=list)

    # -- Aggregations -----------------------------------------------------

    def summary(self) -> str:
        """Return a multi-line summary of findings, one row per rule.

        Format::

            multiple-spaces      warning  3
            trailing-whitespace  warning  1
            mixed-quotes         info     2
            ---
            6 findings (1 error, 4 warnings, 2 infos)

        Always includes the totals line at the bottom even when no
        findings are present (in which case the body is empty and the
        totals line reads ``0 findings``).
        """

        per_rule: "OrderedDict[str, Tuple[str, int]]" = OrderedDict()
        errors = warnings = infos = 0
        for f in self.findings:
            sev = f.severity
            if sev == "error":
                errors += 1
            elif sev == "warning":
                warnings += 1
            else:
                infos += 1
            existing = per_rule.get(f.rule)
            if existing is None:
                per_rule[f.rule] = (sev, 1)
            else:
                # When the same rule has mixed severities, surface the
                # most urgent one in the summary line.
                merged_sev = _max_severity(existing[0], sev)
                per_rule[f.rule] = (merged_sev, existing[1] + 1)
        rule_width = max((len(name) for name in per_rule), default=0)
        sev_width = max(
            (len(sev) for sev, _ in per_rule.values()), default=len("warning")
        )
        body_lines: List[str] = []
        for rule_name, (sev, count) in per_rule.items():
            body_lines.append(
                f"{rule_name.ljust(rule_width)}  {sev.ljust(sev_width)}  {count}"
            )
        total = len(self.findings)
        totals_line = (
            f"{total} findings ({errors} error{'s' if errors != 1 else ''}, "
            f"{warnings} warning{'s' if warnings != 1 else ''}, "
            f"{infos} info{'s' if infos != 1 else ''})"
        )
        if not body_lines:
            return totals_line
        return "\n".join(body_lines + ["---", totals_line])

    # -- Mutation --------------------------------------------------------

    def autofix(
        self,
        rules: Optional[Sequence[str]] = None,
    ) -> int:
        """Apply every available autofix and return the number applied.

        When *rules* is ``None`` every finding whose ``autofix_available``
        flag is ``True`` is fixed. When a sequence is supplied, only
        findings whose ``rule`` is in the sequence are fixed.

        Returns the count of fixes that the underlying autofix callbacks
        reported as successful.
        """

        if rules is not None:
            wanted = set(rules)
        else:
            wanted = None
        applied = 0
        # Fix in document-reverse order so paragraph_index changes
        # caused by removal don't invalidate later indices.
        ordered = sorted(
            self.findings,
            key=lambda f: (
                f.paragraph_index if f.paragraph_index is not None else -1
            ),
            reverse=True,
        )
        for f in ordered:
            if not f.autofix_available:
                continue
            if wanted is not None and f.rule not in wanted:
                continue
            rule = _REGISTRY.get(f.rule)
            if rule is None or rule.autofix is None:
                continue
            try:
                ok = rule.autofix(self.document, f)
            except Exception:  # pragma: no cover - defensive
                ok = False
            if ok:
                applied += 1
        return applied

    # -- Convenience -----------------------------------------------------

    def __iter__(self) -> Iterator[Finding]:
        return iter(self.findings)

    def __len__(self) -> int:
        return len(self.findings)

    def __bool__(self) -> bool:
        return bool(self.findings)


def _max_severity(a: str, b: str) -> str:
    order = {"error": 0, "warning": 1, "info": 2}
    return a if order.get(a, 3) <= order.get(b, 3) else b


# ---------------------------------------------------------------------------
# Built-in rules: helpers
# ---------------------------------------------------------------------------


_MULTI_SPACE_RE = re.compile(r"  +")  # two or more spaces

_HEADING_STYLE_RE = re.compile(r"^Heading\s+(\d+)$")

_PLACEHOLDER_PATTERNS: Tuple[re.Pattern[str], ...] = (
    re.compile(r"\[PLACEHOLDER\]", re.IGNORECASE),
    re.compile(r"\[TBD\]", re.IGNORECASE),
    re.compile(r"\bLorem\s+ipsum\b", re.IGNORECASE),
)

_SMART_QUOTES = "“”‘’"  # “ ” ‘ ’
_STRAIGHT_QUOTES = "\"'"

_OVER_LONG_THRESHOLD = 1000


def _heading_level(paragraph: "Paragraph") -> Optional[int]:
    """Return the heading level (1-9) of *paragraph* or |None|."""
    style = paragraph.style
    if style is None:
        return None
    name = getattr(style, "name", None) or ""
    match = _HEADING_STYLE_RE.match(name)
    if match is None:
        if name in ("Title", "Subtitle"):
            return 0  # treat Title/Subtitle as the document root level
        return None
    return int(match.group(1))


def _document_inline_shapes(document: "Document") -> List["InlineShape"]:
    try:
        return list(document.inline_shapes)
    except Exception:  # pragma: no cover - defensive
        return []


# ---------------------------------------------------------------------------
# Built-in rules: check + autofix callbacks
# ---------------------------------------------------------------------------


def _check_multiple_spaces(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        for run_index, run in enumerate(paragraph.runs):
            text = run.text
            if not text:
                continue
            match = _MULTI_SPACE_RE.search(text)
            if match is None:
                continue
            yield Finding(
                rule="multiple-spaces",
                severity="warning",
                message=(
                    f"paragraph {index} run {run_index} contains "
                    f"{len(match.group(0))} consecutive spaces"
                ),
                paragraph_index=index,
                autofix_available=True,
                autofix_description="collapse runs of spaces to a single space",
                location=f"paragraph {index} run {run_index}",
            )
            break  # one finding per run is enough; autofix collapses all


def _autofix_multiple_spaces(document: "Document", finding: Finding) -> bool:
    if finding.paragraph_index is None:
        return False
    try:
        paragraph = document.paragraphs[finding.paragraph_index]
    except IndexError:
        return False
    fixed_any = False
    for run in paragraph.runs:
        new_text = _MULTI_SPACE_RE.sub(" ", run.text)
        if new_text != run.text:
            run.text = new_text
            fixed_any = True
    return fixed_any


def _check_trailing_whitespace(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        if not text:
            continue
        if text != text.rstrip():
            yield Finding(
                rule="trailing-whitespace",
                severity="warning",
                message=f"paragraph {index} ends with whitespace",
                paragraph_index=index,
                autofix_available=True,
                autofix_description="trim trailing whitespace",
                location=f"paragraph {index}",
            )


def _autofix_trailing_whitespace(document: "Document", finding: Finding) -> bool:
    if finding.paragraph_index is None:
        return False
    try:
        paragraph = document.paragraphs[finding.paragraph_index]
    except IndexError:
        return False
    runs = paragraph.runs
    if not runs:
        return False
    fixed_any = False
    # Walk runs in reverse, peeling trailing whitespace off the last
    # non-empty run. Empty runs are skipped (they may carry formatting
    # state we shouldn't touch).
    for run in reversed(runs):
        if run.text == "":
            continue
        new_text = run.text.rstrip()
        if new_text != run.text:
            run.text = new_text
            fixed_any = True
        if new_text:
            break  # the last visible character is now non-whitespace
    return fixed_any


def _check_tab_instead_of_indent(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        # Use the first run's text rather than paragraph.text — Paragraph.text
        # maps `<w:tab/>` elements to `\t`, which would mis-fire on
        # legitimate field separators. Run.text only contains the literal
        # characters the author typed.
        runs = paragraph.runs
        if not runs:
            continue
        first_text = runs[0].text
        if first_text.startswith("\t"):
            yield Finding(
                rule="tab-instead-of-indent",
                severity="warning",
                message=f"paragraph {index} starts with a literal tab character",
                paragraph_index=index,
                autofix_available=True,
                autofix_description="remove leading tab character",
                location=f"paragraph {index}",
            )


def _autofix_tab_instead_of_indent(document: "Document", finding: Finding) -> bool:
    if finding.paragraph_index is None:
        return False
    try:
        paragraph = document.paragraphs[finding.paragraph_index]
    except IndexError:
        return False
    runs = paragraph.runs
    if not runs:
        return False
    first_run = runs[0]
    new_text = first_run.text.lstrip("\t")
    if new_text != first_run.text:
        first_run.text = new_text
        return True
    return False


def _check_mixed_quotes(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        has_smart = any(ch in _SMART_QUOTES for ch in text)
        has_straight = any(ch in _STRAIGHT_QUOTES for ch in text)
        if has_smart and has_straight:
            yield Finding(
                rule="mixed-quotes",
                severity="info",
                message=(
                    f"paragraph {index} mixes smart (curly) and straight quotes"
                ),
                paragraph_index=index,
                autofix_available=False,
                autofix_description=None,
                location=f"paragraph {index}",
            )


def _check_empty_paragraph(document: "Document") -> Iterable[Finding]:
    paragraphs = document.paragraphs
    in_run = False
    run_start: Optional[int] = None
    for index, paragraph in enumerate(paragraphs):
        is_empty = not paragraph.text.strip()
        if is_empty:
            if not in_run:
                in_run = True
                run_start = index
            else:
                yield Finding(
                    rule="empty-paragraph",
                    severity="info",
                    message=(
                        f"paragraph {index} is a consecutive empty "
                        f"paragraph (run started at {run_start})"
                    ),
                    paragraph_index=index,
                    autofix_available=True,
                    autofix_description="remove this consecutive empty paragraph",
                    location=f"paragraph {index}",
                )
        else:
            in_run = False
            run_start = None


def _autofix_empty_paragraph(document: "Document", finding: Finding) -> bool:
    if finding.paragraph_index is None:
        return False
    try:
        paragraph = document.paragraphs[finding.paragraph_index]
    except IndexError:
        return False
    if paragraph.text.strip():
        return False
    try:
        paragraph.delete()
    except Exception:  # pragma: no cover - defensive
        return False
    return True


def _check_inconsistent_heading_levels(document: "Document") -> Iterable[Finding]:
    previous_level: Optional[int] = None
    for index, paragraph in enumerate(document.paragraphs):
        level = _heading_level(paragraph)
        if level is None or level == 0:
            continue
        if previous_level is not None and level > previous_level + 1:
            yield Finding(
                rule="inconsistent-heading-levels",
                severity="warning",
                message=(
                    f"heading at paragraph {index} jumps from level "
                    f"{previous_level} to level {level}"
                ),
                paragraph_index=index,
                autofix_available=False,
                autofix_description=None,
                location=f"paragraph {index}",
            )
        previous_level = level


def _check_missing_alt_text(document: "Document") -> Iterable[Finding]:
    for shape_index, shape in enumerate(_document_inline_shapes(document)):
        alt = getattr(shape, "alt_text", None)
        title = getattr(shape, "title", None)
        # Treat a non-empty alt OR title as sufficient — Word's own UI
        # accepts either as a screen-reader hint.
        if alt and alt.strip():
            continue
        if title and title.strip():
            continue
        yield Finding(
            rule="missing-alt-text",
            severity="warning",
            message=f"inline image {shape_index} has no alt text",
            paragraph_index=None,
            autofix_available=False,
            autofix_description=None,
            location=f"inline image {shape_index}",
        )


def _check_mixed_fonts(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        names = {run.font.name for run in paragraph.runs if run.font.name}
        if len(names) > 1:
            yield Finding(
                rule="mixed-fonts",
                severity="info",
                message=(
                    f"paragraph {index} uses multiple font families: "
                    + ", ".join(sorted(names))
                ),
                paragraph_index=index,
                autofix_available=False,
                autofix_description=None,
                location=f"paragraph {index}",
            )


def _document_filename_stem(document: "Document") -> Optional[str]:
    """Best-effort guess at the filename stem the document was loaded from.

    python-docx's ``Document()`` factory does not retain the load path
    on the document object, so the linter looks for a side-channel hint
    set by the caller as ``document._lint_filename = "..."``. That keeps
    the lint module a strict consumer of the public API while still
    giving callers a clean way to opt into the filename-based autofix::

        doc = Document("draft.docx")
        doc._lint_filename = "draft.docx"
        report = lint.lint(doc)
        report.autofix(rules=["missing-document-title"])

    Falls back to scanning the package / part for a stored path
    attribute should one ever be added to the core API.
    """

    hint = getattr(document, "_lint_filename", None)
    if isinstance(hint, str) and hint:
        stem = os.path.splitext(os.path.basename(hint))[0]
        if stem:
            return stem
    part = getattr(document, "part", None)
    package = getattr(part, "package", None) if part is not None else None
    for source in (package, part, document):
        path = getattr(source, "_pkg_filename", None) or getattr(
            source, "_filename", None
        )
        if isinstance(path, str) and path:
            stem = os.path.splitext(os.path.basename(path))[0]
            if stem:
                return stem
    return None


def _check_missing_document_title(document: "Document") -> Iterable[Finding]:
    try:
        title = document.core_properties.title
    except Exception:  # pragma: no cover - defensive
        title = None
    if title and title.strip():
        return
    stem = _document_filename_stem(document)
    autofix = stem is not None
    yield Finding(
        rule="missing-document-title",
        severity="info",
        message="document core property 'title' is empty",
        paragraph_index=None,
        autofix_available=autofix,
        autofix_description=(
            f"set core property 'title' to {stem!r}" if autofix else None
        ),
        location="core properties",
    )


def _autofix_missing_document_title(
    document: "Document", finding: Finding
) -> bool:
    stem = _document_filename_stem(document)
    if stem is None:
        return False
    try:
        document.core_properties.title = stem
    except Exception:  # pragma: no cover - defensive
        return False
    return True


def _check_over_long_paragraph(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        if len(text) > _OVER_LONG_THRESHOLD:
            # Skip headings — the rule targets body prose, not titles.
            if _heading_level(paragraph) is not None:
                continue
            yield Finding(
                rule="over-long-paragraph",
                severity="info",
                message=(
                    f"paragraph {index} is {len(text)} characters long "
                    f"(threshold {_OVER_LONG_THRESHOLD})"
                ),
                paragraph_index=index,
                autofix_available=False,
                autofix_description=None,
                location=f"paragraph {index}",
            )


def _check_placeholder_text(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        for pattern in _PLACEHOLDER_PATTERNS:
            match = pattern.search(text)
            if match is None:
                continue
            yield Finding(
                rule="placeholder-text",
                severity="warning",
                message=(
                    f"paragraph {index} contains placeholder "
                    f"{match.group(0)!r}"
                ),
                paragraph_index=index,
                autofix_available=False,
                autofix_description=None,
                location=f"paragraph {index}",
            )
            break  # one finding per paragraph regardless of how many


# ---------------------------------------------------------------------------
# Built-in rule registration
# ---------------------------------------------------------------------------


BUILTIN_RULES: Tuple[str, ...] = (
    "multiple-spaces",
    "trailing-whitespace",
    "tab-instead-of-indent",
    "mixed-quotes",
    "empty-paragraph",
    "inconsistent-heading-levels",
    "missing-alt-text",
    "mixed-fonts",
    "missing-document-title",
    "over-long-paragraph",
    "placeholder-text",
)
"""The eleven built-in rule identifiers, in registration order."""


def _install_builtin_rules() -> None:
    """(Re-)install the built-in rules into the registry.

    Idempotent — calling this after :func:`unregister_rule` restores
    any built-ins that were removed without disturbing custom rules
    registered alongside.
    """

    register_rule(
        "multiple-spaces", _check_multiple_spaces, _autofix_multiple_spaces
    )
    register_rule(
        "trailing-whitespace",
        _check_trailing_whitespace,
        _autofix_trailing_whitespace,
    )
    register_rule(
        "tab-instead-of-indent",
        _check_tab_instead_of_indent,
        _autofix_tab_instead_of_indent,
    )
    register_rule("mixed-quotes", _check_mixed_quotes)
    register_rule(
        "empty-paragraph", _check_empty_paragraph, _autofix_empty_paragraph
    )
    register_rule(
        "inconsistent-heading-levels", _check_inconsistent_heading_levels
    )
    register_rule("missing-alt-text", _check_missing_alt_text)
    register_rule("mixed-fonts", _check_mixed_fonts)
    register_rule(
        "missing-document-title",
        _check_missing_document_title,
        _autofix_missing_document_title,
    )
    register_rule("over-long-paragraph", _check_over_long_paragraph)
    register_rule("placeholder-text", _check_placeholder_text)


_install_builtin_rules()


# ---------------------------------------------------------------------------
# Re-exports (purely for clearer ``help(docx.kit.lint)`` output)
# ---------------------------------------------------------------------------


def _typing_aliases() -> Tuple[Any, ...]:  # pragma: no cover - documentation aid
    return (Union[str, os.PathLike], Sequence[str])
