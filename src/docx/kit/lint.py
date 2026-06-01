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

Built-in rules (twelve total):

* ``multiple-spaces`` (warning, autofix) — three or more consecutive
  spaces *inside* a run (interior, not leading); collapses to a single
  space. Paragraphs styled as ``List *``, ``Body Text Indent`` or
  ``Quote`` are skipped because their hanging indents legitimately use
  multi-space prefixes. Heading paragraphs whose double-space sits
  immediately after a leading numbering token (``4.1  Title``) are also
  skipped — that gap is a deliberate template convention. Threshold is
  configurable via the module-level :data:`MULTIPLE_SPACES_MIN_RUN`.
* ``trailing-whitespace`` (warning, autofix) — paragraph ends with a
  whitespace character; trims the trailing whitespace.
* ``tab-instead-of-indent`` (warning, autofix) — body paragraph starts
  with one or more literal ``\\t`` characters; replaces the leading
  tab(s) with a real ``paragraph_format.left_indent`` so the visual
  indent survives. Skips heading and list paragraphs, where a leading
  tab is typically a list/numbering leader, not an author-typed
  indent.
* ``leading-spaces-instead-of-indent`` (info, autofix) — body
  paragraph starts with four-or-more literal space characters
  (configurable via :data:`LEADING_SPACES_MIN_RUN`); replaces the
  leading space-run with a real ``paragraph_format.left_indent`` so
  the visual indent survives. Same heading / list / hanging-indent
  skip-list as ``tab-instead-of-indent``. Sibling of that rule for
  authors who fake an indent with the spacebar (common from web /
  markdown copy-paste). Closes #676.
* ``mixed-quotes`` (info, no-fix) — paragraph mixes "smart" and
  "straight" quote characters (manual review — auto-converting can
  destroy intentional code samples).
* ``empty-paragraph`` (info, autofix) — consecutive empty / whitespace-
  only paragraphs; keeps the first, removes the rest. Paragraphs whose
  XML carries layout / annotation intent (page / column / line break,
  tab, drawing, picture, embedded object, bookmark anchor, comment-
  range marker, SDT, section properties, ink annotation, complex /
  simple field, hyperlink) are never reported and never auto-fixed,
  even when their rendered text is empty (issue #656). The autofix
  also honours :attr:`Finding.safe_to_delete` — a caller-built
  ``Finding`` with ``safe_to_delete=False`` is skipped and a one-line
  preservation note is appended to
  :attr:`LintReport.preservation_notes`.
* ``trailing-empty-paragraph`` (info, autofix) — empty paragraphs at
  the very end of the document; deletes them. Closes the gap left by
  ``empty-paragraph`` (which only catches the second-and-subsequent
  in a consecutive run).
* ``inconsistent-heading-levels`` (warning, no-fix) — heading skips a
  level (e.g. H1 then H3). Manual fix only — automatically renumbering
  changes the table of contents.
* ``trailing-heading`` (info, no-fix) — heading paragraph at the end
  of the document with no body content beneath it (every following
  block is an empty paragraph, or there are no following blocks).
  Catches sections promised but never delivered, plus the common
  authoring bug where ``Heading`` style is auto-applied to a final
  pasted line. A trailing table counts as content, even when its
  cells are empty. Manual fix only — the rule cannot guess what
  content the author intended; the right move is to either delete
  the heading or add the missing section body.
* ``missing-alt-text`` (info or warning, no-fix) — inline image without
  an alt text attribute. Default severity is ``info``; escalates to
  ``warning`` when the document already declares accessibility intent
  (a non-empty ``core_properties.title`` *and* at least one inline image
  that already carries alt text). Decorative images — those flagged via
  python-docx's :attr:`~docx.shape.InlineShape.a11y_role` of
  ``"decorative"`` or carrying Office 365's
  ``<a16:decorative val="1"/>`` extension marker — are skipped. Repeat
  insertions of the same image binary are collapsed to a single finding
  per unique image. Manual fix only — alt text is meaning-bearing and
  should be authored, not generated.
* ``mixed-fonts`` (info, no-fix) — paragraph runs use multiple font
  families. Manual review only — sometimes intentional (code spans,
  emphasis runs).
* ``missing-document-title`` (info, autofix-from-filename) — core
  property ``title`` is empty; autofix sets it to the document
  filename's stem when one is known. The ``Document(path)`` factory
  records the load path automatically, so the autofix is available
  out-of-the-box for documents loaded from disk. Pass an explicit
  ``source_path=...`` to :func:`lint` for documents loaded from
  in-memory streams when a filename is known by other means. When
  *no* filename is available the finding is suppressed entirely —
  there is nothing the caller can do about a missing title without
  context, so the rule stays silent rather than emitting permanent
  ``info`` noise.
* ``over-long-paragraph`` (info, no-fix) — paragraph longer than the
  configured threshold (default ``1000`` characters). Manual review
  only — splitting may break list / TOC numbering. List / caption /
  footnote / quote styles are exempt by default; tune via
  :class:`LintConfig`.
* ``placeholder-text`` (warning, no-fix) — paragraph still contains a
  known placeholder sentinel. The bundled patterns cover the
  bracket-token forms (``[PLACEHOLDER]``, ``[TBD]``, ``[FILL IN]``),
  the Latin filler (``Lorem ipsum``), the author-marker conventions
  (``TODO:``, ``FIXME``, ``XXX``, ``TKTK``), and angle-bracket
  sentinels (``<replace me>``, ``<your text here>``, ``<insert
  name>``). Manual fix — autoreplace cannot guess the intended
  replacement.
* ``table-without-header-row`` (warning, autofix) — the first row of
  a table is not flagged as a header (``<w:trPr>/<w:tblHeader/>``
  absent). Word will not repeat the row when the table breaks across
  pages and screen readers will not announce it as a header — a WCAG
  1.3.1 (Info & Relationships) failure. Autofix sets
  ``rows[0].is_header = True`` on the affected table; opt out via
  ``report.autofix(rules=[...])`` when the first row is genuinely a
  data row rather than headings.
* ``bare-url`` (info, no-fix) — paragraph contains a raw URL string
  (``https://...``, ``http://...``, ``www....``) that is not wrapped in
  a ``<w:hyperlink>`` element. Manual fix only — choosing the visible
  link text and the relationship target is meaning-bearing.
* ``excessive-font-size-variation`` (info, no-fix) — body runs use
  more than four distinct explicit font sizes across the document
  (heading paragraphs are skipped). Manual review only — collapsing
  sizes is a meaning-bearing decision the author must make.

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
from types import MappingProxyType
from typing import (
    TYPE_CHECKING,
    Any,
    Callable,
    Dict,
    Iterable,
    Iterator,
    List,
    Mapping,
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
    "LintConfig",
    "LintReport",
    "Rule",
    "lint",
    "register_rule",
    "unregister_rule",
    "registered_rules",
    "BUILTIN_RULES",
    "DEFAULT_STYLE_EXEMPTIONS",
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
    details
        Optional read-only mapping carrying rule-specific structured
        data (e.g. ``inconsistent-heading-levels`` exposes
        ``{"level": 3, "previous_level": 1, "skipped": 1}``). Empty
        mapping when the rule has no structured payload. Tools that
        need to reason about a finding programmatically should prefer
        ``details[...]`` over regex-parsing :attr:`message`. Closes
        #678.
    safe_to_delete
        ``True`` (the default) when the autofix may delete the
        underlying paragraph / element without losing structural
        intent. Set to ``False`` by rules that detected an "empty"
        paragraph carrying load-bearing XML (page break, bookmark,
        section properties, comment anchor, SDT, field, hyperlink,
        ink, embedded object, etc.).
        :meth:`LintReport.autofix` skips findings whose
        ``safe_to_delete`` is ``False`` and records a one-line
        preservation note in :attr:`LintReport.preservation_notes`.
        Closes #656.
    """

    rule: str
    severity: str
    message: str
    paragraph_index: Optional[int] = None
    autofix_available: bool = False
    autofix_description: Optional[str] = None
    location: Optional[str] = None
    details: Mapping[str, Any] = field(
        default_factory=lambda: MappingProxyType({})
    )
    safe_to_delete: bool = True


DEFAULT_STYLE_EXEMPTIONS: Tuple[str, ...] = (
    "List Bullet",
    "List Number",
    "List Paragraph",
    "Caption",
    "Footnote Text",
    "Quote",
)
"""Paragraph style names exempted by default from prose-length heuristics.

These styles legitimately carry long compound content (bulleted
explanations, captions, quoted blocks, footnote bodies) whose length is
bounded by editorial intent rather than reading-line ergonomics, so the
``over-long-paragraph`` rule skips them by default. Override via
:class:`LintConfig`.
"""


@dataclass(frozen=True)
class LintConfig:
    """Tunable thresholds and exemptions for the built-in rules.

    Pass an instance to :func:`lint` to override any of the heuristics
    without monkey-patching module-level constants. Every field has a
    sane default chosen to match the historical behavior::

        from docx.kit.lint import lint, LintConfig

        report = lint(doc, config=LintConfig(over_long_threshold=2000))

    Attributes
    ----------
    over_long_threshold
        Maximum paragraph character length before
        ``over-long-paragraph`` fires. Defaults to ``1000``.
    multi_space_minimum
        Minimum run of consecutive ``ASCII space`` characters that
        triggers ``multiple-spaces``. Must be ``>= 2``. Defaults to
        ``2``.
    style_exemptions
        Paragraph style names exempted from ``over-long-paragraph``.
        The default covers list / caption / footnote / quote families
        whose long bodies are usually intentional. Pass an explicit
        empty ``frozenset()`` to disable exemptions entirely.

    .. versionadded:: 2026.05.31
    """

    over_long_threshold: int = 1000
    multi_space_minimum: int = 2
    style_exemptions: frozenset = field(
        default_factory=lambda: frozenset(DEFAULT_STYLE_EXEMPTIONS)
    )

    def __post_init__(self) -> None:
        if self.over_long_threshold < 1:
            raise ValueError(
                "over_long_threshold must be a positive integer; "
                f"got {self.over_long_threshold!r}"
            )
        if self.multi_space_minimum < 2:
            raise ValueError(
                "multi_space_minimum must be at least 2 to detect 'runs' "
                f"of spaces; got {self.multi_space_minimum!r}"
            )
        # Coerce mutable iterables (set, list, tuple) to frozenset for
        # immutability — the dataclass is frozen=True at the top level
        # so callers can pass any iterable here without surprise.
        if not isinstance(self.style_exemptions, frozenset):
            object.__setattr__(
                self, "style_exemptions", frozenset(self.style_exemptions)
            )


_DEFAULT_CONFIG = LintConfig()
_ACTIVE_CONFIG: LintConfig = _DEFAULT_CONFIG


def _current_config() -> LintConfig:
    """Return the config the running ``lint()`` call should consult."""
    return _ACTIVE_CONFIG


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


def lint(
    document: "Document",
    *,
    source_path: "Optional[Union[str, os.PathLike[str]]]" = None,
    config: Optional["LintConfig"] = None,
) -> "LintReport":
    """Run every registered rule against *document* and return a :class:`LintReport`.

    *config*, when supplied, overrides the built-in thresholds and
    style exemptions consulted by the bundled rules. Pass an instance
    of :class:`LintConfig`; pass |None| (the default) to use the
    defaults documented on that class.

    The returned report carries the ``findings`` list (in document
    order, then rule order) and exposes an :meth:`LintReport.autofix`
    method that mutates the document in place.

    *source_path* is an optional explicit filename hint used by
    filename-based rules (today: ``missing-document-title``). When
    omitted, the rule falls back to any path captured automatically
    by :func:`docx.Document` when the document was loaded from disk.
    Pass *source_path* explicitly when the document was loaded from
    an in-memory stream but the original filename is known.

    .. versionadded:: 2026.05.29
    .. versionchanged:: 2026.05.31
       Added *source_path* keyword.
    .. versionchanged:: 2026.06.01
       Added the *config* parameter.
    """

    global _ACTIVE_CONFIG
    if config is not None and not isinstance(config, LintConfig):
        raise TypeError(
            "config must be a LintConfig instance or None; got "
            f"{type(config).__name__}"
        )
    previous_config = _ACTIVE_CONFIG
    _ACTIVE_CONFIG = config if config is not None else _DEFAULT_CONFIG
    try:
        if source_path is not None:
            path_str = os.fspath(source_path)
            prev_attr_set = hasattr(document, "_lint_filename")
            prev_value = (
                getattr(document, "_lint_filename", None)
                if prev_attr_set
                else None
            )
            try:
                document._lint_filename = path_str  # type: ignore[attr-defined]
            except Exception:  # pragma: no cover - defensive
                return _run_rules(document, source_path=None)
            try:
                return _run_rules(document, source_path=path_str)
            finally:
                try:
                    if prev_attr_set:
                        document._lint_filename = prev_value  # type: ignore[attr-defined]
                    else:
                        try:
                            del document._lint_filename
                        except AttributeError:  # pragma: no cover
                            pass
                except Exception:  # pragma: no cover - defensive
                    pass
        return _run_rules(document, source_path=None)
    finally:
        _ACTIVE_CONFIG = previous_config


def _run_rules(
    document: "Document", source_path: Optional[str] = None
) -> "LintReport":
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
    return LintReport(
        document=document,
        findings=findings,
        source_path=source_path,
        config=_ACTIVE_CONFIG,
    )


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

    The :attr:`config` attribute records the :class:`LintConfig` the
    findings were produced with so callers can introspect which
    thresholds were in effect.
    """

    document: "Document"
    findings: List[Finding] = field(default_factory=list)
    source_path: Optional[str] = None
    """Filename hint passed to :func:`lint` (or |None|).

    Re-applied to the bound document for the duration of
    :meth:`autofix` so filename-based autofixes (today:
    ``missing-document-title``) succeed even though the rule's
    side-channel attribute was scrubbed when :func:`lint` returned.
    """
    config: "LintConfig" = field(default_factory=lambda: _DEFAULT_CONFIG)
    preservation_notes: List[str] = field(default_factory=list)
    """One-line messages describing findings whose autofix was skipped
    because ``Finding.safe_to_delete`` was ``False`` — e.g. an empty-
    paragraph autofix that would have destroyed a page break or
    section break.

    Populated by :meth:`autofix` / :meth:`autofix_breakdown` on every
    invocation (cleared at the start of each call). Closes #656.
    """

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

        return sum(self.autofix_breakdown(rules=rules).values())

    def autofix_breakdown(
        self,
        rules: Optional[Sequence[str]] = None,
    ) -> Dict[str, int]:
        """Apply autofixes and return a per-rule breakdown of successes.

        Same selection semantics as :meth:`autofix` (``rules=None``
        applies every available autofix; a sequence restricts to those
        rule names) but the return value is a ``{rule_name: count}``
        mapping instead of a single aggregate. Rules with zero successful
        fixes are omitted from the mapping.

        Useful for tooling that wants to display "fixed: multiple-spaces
        x3" without re-running :func:`lint` before and after to diff the
        rule counts.

        Closes #679.
        """

        if rules is not None:
            wanted = set(rules)
        else:
            wanted = None
        breakdown: Dict[str, int] = {}
        # Reset preservation notes — every autofix invocation publishes
        # its own list. Callers reading the previous run's notes should
        # snapshot the list themselves.
        self.preservation_notes = []
        # Fix in document-reverse order so paragraph_index changes
        # caused by removal don't invalidate later indices.
        ordered = sorted(
            self.findings,
            key=lambda f: (
                f.paragraph_index if f.paragraph_index is not None else -1
            ),
            reverse=True,
        )
        # Re-apply the source-path hint that ``lint(..., source_path=...)``
        # captured, so filename-based autofixes can re-derive the stem.
        # Restore the document's prior state on exit.
        path_hint = self.source_path
        prev_attr_set = hasattr(self.document, "_lint_filename")
        prev_value = (
            getattr(self.document, "_lint_filename", None)
            if prev_attr_set
            else None
        )
        if path_hint is not None:
            try:
                self.document._lint_filename = path_hint  # type: ignore[attr-defined]
            except Exception:  # pragma: no cover - defensive
                path_hint = None
        try:
            for f in ordered:
                if not f.autofix_available:
                    continue
                if wanted is not None and f.rule not in wanted:
                    continue
                rule = _REGISTRY.get(f.rule)
                if rule is None or rule.autofix is None:
                    continue
                # Issue #656: never delete a paragraph the rule itself
                # flagged as load-bearing. Record a one-line note so
                # callers can show *why* the autofix was a no-op.
                if not f.safe_to_delete:
                    locator = f.location or (
                        f"paragraph {f.paragraph_index}"
                        if f.paragraph_index is not None
                        else f.rule
                    )
                    self.preservation_notes.append(
                        f"preserved {locator}: {f.rule} autofix skipped "
                        f"because the paragraph carries load-bearing "
                        f"content (page/section break, bookmark, "
                        f"comment anchor, SDT, field, hyperlink, ink, "
                        f"or embedded object)"
                    )
                    continue
                try:
                    ok = rule.autofix(self.document, f)
                except Exception:  # pragma: no cover - defensive
                    ok = False
                if ok:
                    breakdown[f.rule] = breakdown.get(f.rule, 0) + 1
        finally:
            if path_hint is not None:
                try:
                    if prev_attr_set:
                        self.document._lint_filename = prev_value  # type: ignore[attr-defined]
                    else:
                        try:
                            del self.document._lint_filename
                        except AttributeError:  # pragma: no cover
                            pass
                except Exception:  # pragma: no cover - defensive
                    pass
        return breakdown

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


MULTIPLE_SPACES_MIN_RUN = 3
"""Minimum number of consecutive spaces required to trigger ``multiple-spaces``.

Intentional formatting (e.g. spacing after a `4.1` heading-number prefix or
list-bullet hanging indents) routinely uses *exactly* two spaces, so the
default threshold is three. Callers can tighten or loosen by reassigning
this module-level constant before invoking :func:`lint`. Values below 2 are
clamped up to 2 — a single space is never multi-space.
"""

# Pattern used to flag interior runs of spaces inside a run's text. Built
# lazily so callers can override ``MULTIPLE_SPACES_MIN_RUN`` between calls.
def _multi_space_re() -> "re.Pattern[str]":
    n = max(2, int(MULTIPLE_SPACES_MIN_RUN))
    # ``(?<=\S)`` ensures we don't match leading-whitespace runs (list
    # hanging indents, intentional pre-bullet padding); we want *interior*
    # runs only.
    return re.compile(rf"(?<=\S) {{{n},}}")


# Legacy module-level pattern preserved for back-compat with any external
# callers that imported ``_MULTI_SPACE_RE`` directly. The autofix uses a
# greedier collapse pattern (any run of two-or-more spaces) so that once a
# finding is emitted, every space-run in the paragraph is normalised.
_MULTI_SPACE_RE = re.compile(r"  +")

# Heading paragraphs whose multi-space gap sits immediately after a leading
# numeric token (``4.1  Three-LZA``) are intentional — skip them.
_HEADING_NUMBERING_GAP_RE = re.compile(r"^\s*\d+(?:\.\d+)*\s{2,}\S")

# Styles whose paragraphs commonly carry intentional leading whitespace
# (hanging indents). Membership is tested case-insensitively against the
# style name; "List Bullet 2", "List Number", etc. all match the "list"
# token.
_INDENTED_STYLE_TOKENS: Tuple[str, ...] = (
    "list",  # List Bullet, List Number, List Paragraph, ...
    "body text indent",
    "quote",
)

_HEADING_STYLE_RE = re.compile(r"^Heading\s+(\d+)$")

def _multi_space_pattern(minimum: int) -> "re.Pattern[str]":
    """Return a compiled regex matching *minimum*+ consecutive spaces."""
    if minimum <= 2:
        return _MULTI_SPACE_RE
    return re.compile(r" {%d,}" % minimum)


# Placeholder-text patterns paired with a stable ``category`` tag that
# rule consumers (autofixers, authoring UIs) can group findings by
# without regex-parsing the matched substring. The categories are part
# of the public Finding.details contract once #681 lands.
_PLACEHOLDER_PATTERNS: Tuple[Tuple[re.Pattern[str], str], ...] = (
    (re.compile(r"\[PLACEHOLDER\]", re.IGNORECASE), "bracket-token"),
    (re.compile(r"\[TBD\]", re.IGNORECASE), "bracket-token"),
    (re.compile(r"\bLorem\s+ipsum\b", re.IGNORECASE), "lorem-ipsum"),
    # ``TODO`` is the most common author-marker placeholder in real
    # drafts. Require a delimiter (``:``, ``-``, whitespace) after the
    # word so we don't false-match ``TODOLIST`` or product names that
    # incidentally contain the substring.
    (re.compile(r"\bTODO\b[\s:\-–]", re.IGNORECASE), "todo-marker"),
    # ``FIXME`` and ``XXX`` are the programmer-style placeholder
    # conventions that bleed into prose drafts via copy-paste from code.
    # ``XXX`` stays case-sensitive — lower-case ``xxx`` is too noisy
    # (it appears in URLs, anonymisation tokens, etc.).
    (re.compile(r"\bFIXME\b", re.IGNORECASE), "todo-marker"),
    (re.compile(r"\bXXX\b"), "todo-marker"),
    # ``TKTK`` is the journalism convention for "to come". The four-
    # letter form is unambiguous enough to flag case-insensitively;
    # we deliberately do not match the two-letter ``TK`` because it
    # collides with too many product / acronym uses.
    (re.compile(r"\bTKTK\b", re.IGNORECASE), "to-come"),
    # Generic angle-bracket sentinels: ``<replace me>``, ``<your
    # text here>``, ``<insert name>``. The pattern requires the
    # opening ``<`` and matching ``>`` so legitimate prose using the
    # word "replace" stays untouched.
    (
        re.compile(
            r"<\s*(?:replace[\s_-]*me|your[\s_-]*text[\s_-]*here|"
            r"insert[\s_-]*\w+)\s*>",
            re.IGNORECASE,
        ),
        "angle-bracket",
    ),
    # ``[FILL IN]`` / ``[FILL ME]`` mirror the existing ``[TBD]``
    # bracket-token convention.
    (re.compile(r"\[\s*FILL\s*(?:IN|ME)\s*\]", re.IGNORECASE), "bracket-token"),
)

_SMART_QUOTES = "“”‘’"  # “ ” ‘ ’
_STRAIGHT_QUOTES = "\"'"

_OVER_LONG_THRESHOLD = 1000

# Match http/https/www URLs. Trailing punctuation (``.,;:)]}>"'``) is stripped
# below so a sentence-ending period or closing parenthesis is not treated as
# part of the URL — that would produce noisy / inaccurate findings.
_BARE_URL_RE = re.compile(r"\b(?:https?://|www\.)\S+")
_URL_TRAILING_PUNCT = ".,;:!?)]}>\"'"

_EXCESSIVE_FONT_SIZE_THRESHOLD = 4


def _paragraph_style_name(paragraph: "Paragraph") -> str:
    """Return the paragraph's style name as a lowercased string (or ``""``).

    Lowercased so existing case-insensitive callers (e.g.
    :func:`_is_indented_style`) can match style families without
    re-normalising. Callers that need the case-preserved form (e.g.
    matching against :data:`DEFAULT_STYLE_EXEMPTIONS`) should read
    ``paragraph.style.name`` directly via :func:`_paragraph_style_name_raw`.
    """
    style = paragraph.style
    if style is None:
        return ""
    name = getattr(style, "name", None) or ""
    return name.lower()


def _paragraph_style_name_raw(paragraph: "Paragraph") -> str:
    """Return the paragraph style name with case preserved (or empty)."""
    style = paragraph.style
    if style is None:
        return ""
    return getattr(style, "name", None) or ""


def _heading_level(paragraph: "Paragraph") -> Optional[int]:
    """Return the heading level (1-9) of *paragraph* or |None|."""
    name = _paragraph_style_name_raw(paragraph)
    if not name:
        return None
    match = _HEADING_STYLE_RE.match(name)
    if match is None:
        if name in ("Title", "Subtitle"):
            return 0  # treat Title/Subtitle as the document root level
        return None
    return int(match.group(1))


def _is_indented_style(paragraph: "Paragraph") -> bool:
    """``True`` for paragraphs whose style typically uses a hanging indent.

    Hanging-indent styles (``List Bullet``, ``List Number``, ``List
    Paragraph``, ``Body Text Indent``, ``Quote``) routinely begin with
    multi-space prefixes in the authored text — those should not trigger
    ``multiple-spaces``.
    """
    name = _paragraph_style_name(paragraph)
    if not name:
        return False
    return any(token in name for token in _INDENTED_STYLE_TOKENS)


def _document_inline_shapes(document: "Document") -> List["InlineShape"]:
    try:
        return list(document.inline_shapes)
    except Exception:  # pragma: no cover - defensive
        return []


# -- Office 365 "Mark as decorative" marker. The flag lives on
# ``wp:docPr/a:extLst/a:ext/a16:decorative/@val``. ``a16`` is the Office
# 2018 drawing namespace; the URI is stable per Microsoft's published
# extension catalog. We do not register the prefix in ``docx.oxml.ns`` —
# the lookup here is read-only and uses the fully-qualified URI directly
# so we don't perturb the global namespace map. --
_A16_DECORATIVE_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative"
_A16_DECORATIVE_QN = f"{{{_A16_DECORATIVE_NS}}}decorative"


def _shape_is_decorative(shape: "InlineShape") -> bool:
    """Return ``True`` when *shape* is flagged decorative.

    Two paths are recognised:

    * python-docx's :attr:`~docx.shape.InlineShape.a11y_role` returns
      ``"decorative"`` when the descr carries a ``[decorative]`` prefix.
    * Office 365 writes ``<a16:decorative val="1"/>`` inside
      ``wp:docPr/a:extLst/a:ext`` when the user ticks the *Mark as
      decorative* checkbox. We resolve the extension element by Clark
      name so an unregistered ``a16`` prefix doesn't trip us up.
    """
    role = getattr(shape, "a11y_role", None)
    if role == "decorative":
        return True
    inline = getattr(shape, "_inline", None)
    if inline is None:
        return False
    docPr = getattr(inline, "docPr", None)
    if docPr is None:
        return False
    for elem in docPr.iter(_A16_DECORATIVE_QN):
        val = elem.get("val")
        # Per the Office 365 schema the attribute is a boolean — values
        # of "1" or "true" mean decorative. Anything else (including a
        # missing attribute, which spec-wise defaults to false) is
        # ignored to stay on the conservative side.
        if val in ("1", "true"):
            return True
    return False


def _shape_identity(shape: "InlineShape") -> Optional[str]:
    """Return a stable identity for *shape*'s underlying image, or |None|.

    Prefers the SHA-1 of the image binary so that two pictures inserted
    from the same file collapse to one finding even when Word stores
    them as separate parts. Falls back to the part-name (a string like
    ``"/word/media/image1.png"``) so chart / SmartArt shapes that don't
    expose a blob still dedupe correctly. Returns |None| when neither
    can be resolved — those shapes always emit a finding.
    """
    try:
        image = shape.image  # type: ignore[attr-defined]
    except Exception:
        image = None
    sha1 = getattr(image, "sha1", None)
    if isinstance(sha1, str) and sha1:
        return f"sha1:{sha1}"
    inline = getattr(shape, "_inline", None)
    part = getattr(shape, "_part", None)
    if inline is not None and part is not None:
        try:
            blip = shape._blip()  # type: ignore[attr-defined]
        except Exception:
            blip = None
        if blip is not None:
            rId = getattr(blip, "embed", None) or getattr(blip, "link", None)
            if rId:
                related = getattr(part, "related_parts", None) or {}
                related_part = related.get(rId)
                partname = getattr(related_part, "partname", None)
                if partname:
                    return f"partname:{partname}"
    return None


def _document_has_a11y_intent(
    document: "Document", shapes: Sequence["InlineShape"]
) -> bool:
    """Heuristic: does the author show signs of authoring for accessibility?

    Two things have to be true:

    * the document carries a non-empty ``core_properties.title``, and
    * at least one inline shape already has alt text or a title
      attribute set.

    The combination means the author is paying attention to a11y
    metadata in general; a missing alt text in *that* document is much
    more likely to be a real defect than a decorative leftover.
    """
    try:
        title = document.core_properties.title
    except Exception:  # pragma: no cover - defensive
        title = None
    if not (title and title.strip()):
        return False
    for shape in shapes:
        alt = getattr(shape, "alt_text", None)
        if alt and alt.strip():
            return True
        ttl = getattr(shape, "title", None)
        if ttl and ttl.strip():
            return True
    return False


# ---------------------------------------------------------------------------
# Built-in rules: check + autofix callbacks
# ---------------------------------------------------------------------------


def _joined_runs_text_with_offsets(
    paragraph: "Paragraph",
) -> Tuple[str, List[Tuple[int, int, int]]]:
    """Return the joined text of *paragraph*'s top-level runs and a span map.

    The span map is a list of ``(run_index, start, end)`` tuples giving
    the half-open ``[start, end)`` slice of the joined string occupied
    by each run. ``end - start`` equals ``len(run.text)``. Empty runs
    contribute an empty span at the boundary so callers can still
    locate insertion points consistently.

    This is the coordinate system used by ``multiple-spaces`` for
    cross-run detection (issue #657): the joined string lets the
    pattern see double-spaces that straddle a run boundary, and the
    span map lets the autofix translate matches back into per-run
    edits.
    """

    parts: List[str] = []
    spans: List[Tuple[int, int, int]] = []
    cursor = 0
    for run_index, run in enumerate(paragraph.runs):
        text = run.text
        spans.append((run_index, cursor, cursor + len(text)))
        parts.append(text)
        cursor += len(text)
    return "".join(parts), spans


def _runs_for_match(
    spans: Sequence[Tuple[int, int, int]],
    match_start: int,
    match_end: int,
) -> List[int]:
    """Return the indices of runs whose span overlaps ``[match_start, match_end)``.

    A zero-length run sitting exactly on the boundary is excluded — it
    has no characters to edit. The result is in ascending run-index
    order.
    """

    hits: List[int] = []
    for run_index, start, end in spans:
        if start >= match_end:
            break
        if end <= match_start:
            continue
        if start == end:  # empty run, no characters to mutate
            continue
        hits.append(run_index)
    return hits


def _collapse_cross_run_spaces(
    paragraph: "Paragraph",
    match_start: int,
    match_end: int,
) -> bool:
    """Collapse a run of spaces spanning ``[match_start, match_end)`` to one space.

    The match positions are in the joined-runs coordinate space (see
    :func:`_joined_runs_text_with_offsets`). The surviving single space
    lands in the *first* run that contributed at least one space — i.e.
    trailing-space-in-A wins, leading-space(s)-in-B drop. This rule is
    deterministic and preserves the formatting of the run that already
    "owned" the gap.

    Returns ``True`` when at least one run's text was rewritten.
    """

    runs = paragraph.runs
    if not runs:
        return False
    _, spans = _joined_runs_text_with_offsets(paragraph)
    affected = _runs_for_match(spans, match_start, match_end)
    if not affected:
        return False

    # Find the first affected run that actually contributes a space to
    # the matched region. This is where the surviving single space will
    # live; every other affected run drops its spaces in the matched
    # region entirely.
    survivor: Optional[int] = None
    for run_index in affected:
        _, start, end = spans[run_index]
        seg_start = max(start, match_start) - start
        seg_end = min(end, match_end) - start
        segment = runs[run_index].text[seg_start:seg_end]
        if any(ch == " " for ch in segment):
            survivor = run_index
            break
    if survivor is None:  # pragma: no cover - defensive
        return False

    fixed_any = False
    for run_index in affected:
        _, start, end = spans[run_index]
        seg_start = max(start, match_start) - start
        seg_end = min(end, match_end) - start
        run = runs[run_index]
        original = run.text
        before = original[:seg_start]
        segment = original[seg_start:seg_end]
        after = original[seg_end:]
        if run_index == survivor:
            replacement = " "
        else:
            replacement = ""
        # Only rewrite when this run actually carried spaces in the
        # matched region — leaves runs whose overlap is non-space text
        # untouched (defensive; the matcher is space-only so this is
        # belt-and-braces).
        if not any(ch == " " for ch in segment) and replacement == segment:
            continue
        new_text = before + replacement + after
        if new_text != original:
            run.text = new_text
            fixed_any = True
    return fixed_any


def _is_intentional_multiple_spaces(
    paragraph: "Paragraph", match: "re.Match[str]"
) -> bool:
    """Return |True| when *match* is an intentional formatting convention.

    The detection regex is shared by both the check and the autofix,
    but some matches in real-world Word documents are deliberate:

    * Heading-styled paragraphs (``Heading 1`` … ``Heading 9``,
      ``Title``, ``Subtitle``) routinely use a multi-space gap between
      a leading numeric prefix and the title — ``4.1  Three-LZA
      topology`` is a template convention, not a defect.

    * List-styled paragraphs (``List Bullet``, ``List Number``, ``List
      Paragraph``, ``List Continue``, ``Body Text Indent``, ``Quote``)
      commonly start with a multi-space hanging indent before a bullet
      glyph (``    - bullet text``).

    Returning |True| for a given match exempts that match from both
    flagging and the autofix; mid-sentence defects in the same
    paragraph remain in scope (issue #645).
    """

    match_start = match.start()

    # Heading paragraphs whose multi-space match sits immediately after
    # a leading ``\d+(\.\d+)*`` numeric prefix are using the gap as a
    # deliberate number-to-title separator.
    if _heading_level(paragraph) is not None:
        # ``paragraph.text`` covers hyperlink content the joined-runs
        # text omits; we want the predicate to see the full prefix.
        prefix = _HEADING_NUMBERING_GAP_RE.match(paragraph.text)
        if prefix is not None:
            # The numbering gap ends one character before the match's
            # final ``\S`` (which is the first character of the title)
            # — the gap *is* the run of spaces ending at ``prefix.end()
            # - 1``. A match whose end aligns with that position (and
            # whose start sits inside the spaces of the prefix) is the
            # intentional gap.
            gap_end = prefix.end() - 1
            if match.end() == gap_end and match_start < gap_end:
                return True

    # List- / hanging-indent-styled paragraphs whose match starts at
    # the very beginning of the paragraph carry the hanging-indent
    # padding the author typed before the bullet glyph.
    if _is_indented_style(paragraph) and match_start == 0:
        return True

    return False


def _check_multiple_spaces(document: "Document") -> Iterable[Finding]:
    config = _current_config()
    # Prefer the legacy module-level constant when the caller has tuned
    # it past the default; otherwise honour the LintConfig setting. This
    # keeps the existing ``MULTIPLE_SPACES_MIN_RUN`` override path
    # working for callers who never adopt ``LintConfig``. The legacy
    # pattern is interior-only (``(?<=\S)``) — applied to the joined
    # runs string so cross-run double-spaces become detectable (#657).
    legacy_n = max(2, int(MULTIPLE_SPACES_MIN_RUN))
    if legacy_n != 2:
        pattern = _multi_space_re()
    else:
        pattern = _multi_space_pattern(config.multi_space_minimum)
    for index, paragraph in enumerate(document.paragraphs):
        joined, spans = _joined_runs_text_with_offsets(paragraph)
        if not joined:
            continue
        # Walk every match and emit the first non-exempt one. The
        # per-match predicate (issue #645) lets an intentional heading
        # numbering gap coexist with a real mid-sentence defect in the
        # same paragraph — only the latter fires.
        match: "Optional[re.Match[str]]" = None
        for candidate in pattern.finditer(joined):
            if _is_intentional_multiple_spaces(paragraph, candidate):
                continue
            match = candidate
            break
        if match is None:
            continue
        affected = _runs_for_match(spans, match.start(), match.end())
        first_run = affected[0] if affected else 0
        location = (
            f"paragraph {index} run {first_run}"
            if len(affected) <= 1
            else (
                f"paragraph {index} runs {affected[0]}-{affected[-1]}"
            )
        )
        if len(affected) <= 1:
            message = (
                f"paragraph {index} run {first_run} contains "
                f"{len(match.group(0))} consecutive spaces"
            )
        else:
            message = (
                f"paragraph {index} contains "
                f"{len(match.group(0))} consecutive spaces "
                f"spanning runs {affected[0]}-{affected[-1]}"
            )
        yield Finding(
            rule="multiple-spaces",
            severity="warning",
            message=message,
            paragraph_index=index,
            autofix_available=True,
            autofix_description="collapse runs of spaces to a single space",
            location=location,
            details=MappingProxyType(
                {
                    "run_index": first_run,
                    "space_count": len(match.group(0)),
                    "match_start": match.start(),
                    "match_end": match.end(),
                    "run_indices": tuple(affected),
                }
            ),
        )


def _autofix_multiple_spaces(document: "Document", finding: Finding) -> bool:
    if finding.paragraph_index is None:
        return False
    try:
        paragraph = document.paragraphs[finding.paragraph_index]
    except IndexError:
        return False
    # Collapse every non-exempt multi-space run in the paragraph
    # (greedy ``  +`` pattern), operating on the joined-runs string so
    # cross-run double-spaces are caught alongside in-run ones. The
    # per-match exemption predicate (issue #645) leaves intentional
    # heading numbering gaps and list hanging-indent padding alone.
    # Iterating in right-to-left order keeps span offsets stable
    # across edits.
    fixed_any = False
    while True:
        joined, _spans = _joined_runs_text_with_offsets(paragraph)
        matches = [
            m
            for m in _MULTI_SPACE_RE.finditer(joined)
            if not _is_intentional_multiple_spaces(paragraph, m)
        ]
        if not matches:
            break
        # Process the right-most match first; left-side offsets are
        # unaffected by a later edit.
        match = matches[-1]
        if not _collapse_cross_run_spaces(
            paragraph, match.start(), match.end()
        ):
            break
        fixed_any = True
    return fixed_any


def _has_trailing_authored_whitespace(paragraph: "Paragraph") -> bool:
    """Return |True| when the paragraph ends with author-typed whitespace.

    Walks the runs in reverse, examining ``Run.text`` (which decodes
    structural ``<w:tab/>`` and ``<w:br/>`` elements to ``\\t`` and
    ``\\n`` respectively). The last visible character is examined only
    after stripping trailing ``\\t`` and ``\\n`` characters — those
    represent structural elements rather than literal whitespace the
    author typed, so they should not poison the check. Empty runs
    (formatting-only) are skipped.
    """

    for run in reversed(paragraph.runs):
        text = run.text
        if text == "":
            continue
        # Strip structural-element decode artifacts from the right —
        # `<w:tab/>` → '\t' and `<w:br/>` → '\n'. Those don't represent
        # literal whitespace authored at the end of a w:t element, so we
        # ignore them when deciding whether the paragraph trails space.
        cleaned = text.rstrip("\t\n")
        if not cleaned:
            # The run is entirely structural (tabs / breaks); keep
            # walking left for a run carrying actual text characters.
            continue
        return cleaned != cleaned.rstrip()
    return False


def _check_trailing_whitespace(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        if not paragraph.runs:
            continue
        if not _has_trailing_authored_whitespace(paragraph):
            continue
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


def _is_list_paragraph(paragraph: "Paragraph") -> bool:
    """Return ``True`` when *paragraph* is part of a numbered/bulleted list.

    The check uses the public ``Paragraph.list_level`` accessor, which
    returns ``None`` when the paragraph has no ``w:numPr`` (i.e. is not
    a list item). The accessor was introduced for list-handling support
    and is the documented public way to detect list membership.
    """
    try:
        return paragraph.list_level is not None
    except Exception:  # pragma: no cover - defensive
        return False


def _is_heading_or_toc(paragraph: "Paragraph") -> bool:
    """Return ``True`` for heading/title/TOC-style paragraphs.

    A leading tab on these paragraphs is almost always a rendered
    leader between the number and the title (e.g. ``"1.\\tIntroduction"``)
    and stripping it is destructive. ``_heading_level`` handles the
    ``Heading N`` / ``Title`` / ``Subtitle`` styles; we additionally
    skip ``TOC N`` and ``Table of Contents``-style names.
    """
    if _heading_level(paragraph) is not None:
        return True
    style = paragraph.style
    name = getattr(style, "name", None) or ""
    if name.startswith("TOC ") or name == "TOC Heading":
        return True
    if name == "Table of Contents":
        return True
    return False


# Word's default tab-stop is 36 points (≈ 0.5 inch). Each stripped tab
# becomes one tab-stop's worth of left_indent so the visual position
# survives the substitution.
_TAB_INDENT_PT = 36


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
        if not first_text.startswith("\t"):
            continue
        # Skip paragraphs where a leading tab is structural rather
        # than an author-typed indent: heading/TOC paragraphs render a
        # tab-leader between number and title, and list paragraphs
        # carry their indent through ``w:numPr`` already.
        if _is_heading_or_toc(paragraph) or _is_list_paragraph(paragraph):
            continue
        # Count consecutive leading tabs so the autofix can compensate
        # with the right indent multiple and the message can report it.
        tab_count = len(first_text) - len(first_text.lstrip("\t"))
        plural = "s" if tab_count != 1 else ""
        yield Finding(
            rule="tab-instead-of-indent",
            severity="warning",
            message=(
                f"paragraph {index} starts with {tab_count} literal "
                f"tab character{plural}"
            ),
            paragraph_index=index,
            autofix_available=True,
            autofix_description=(
                "replace leading tab(s) with paragraph left-indent"
            ),
            location=f"paragraph {index}",
        )


def _autofix_tab_instead_of_indent(document: "Document", finding: Finding) -> bool:
    # Local import — keeps ``docx.shared`` out of module import time
    # and avoids cycles for callers using the lint module standalone.
    from docx.shared import Emu, Pt

    if finding.paragraph_index is None:
        return False
    try:
        paragraph = document.paragraphs[finding.paragraph_index]
    except IndexError:
        return False
    runs = paragraph.runs
    if not runs:
        return False
    # Re-check the skip conditions in case the document was edited
    # between :func:`lint` and :meth:`LintReport.autofix`. This keeps
    # the autofix idempotent on heading/list paragraphs even if a
    # caller fabricated a Finding by hand.
    if _is_heading_or_toc(paragraph) or _is_list_paragraph(paragraph):
        return False
    first_run = runs[0]
    original = first_run.text
    stripped = original.lstrip("\t")
    tab_count = len(original) - len(stripped)
    if tab_count == 0:
        return False
    first_run.text = stripped
    # Add to any existing direct left_indent so we don't clobber an
    # author-set value; treat ``None`` (inherited) as zero baseline.
    pf = paragraph.paragraph_format
    existing = pf.left_indent
    addition = Pt(_TAB_INDENT_PT * tab_count)
    # ``Length`` is an ``int`` subclass; arithmetic returns plain int
    # so wrap the sum back into ``Emu`` for the setter to keep the
    # value typed.
    pf.left_indent = Emu(int(existing or 0) + int(addition))
    return True


# Threshold for ``leading-spaces-instead-of-indent``. Authors who fake an
# indent with the spacebar typically tap it four times (the standard
# tab-equivalent in web / markdown source); two-or-three spaces are
# common in body prose (quoted dialogue, continuation lines) and
# routinely intentional. Four-or-more is the unambiguous "fake tab"
# threshold the issue (#676) settles on.
LEADING_SPACES_MIN_RUN = 4
"""Minimum consecutive leading-space count before
``leading-spaces-instead-of-indent`` fires.

Default is 4 — the common "fake tab" width on web / markdown input.
Lower the threshold to 2 to flag every double-space leader; raise it
to 8 to only catch the deepest fakes.
"""


def _leading_space_count(text: str) -> int:
    """Return the number of leading ASCII-space characters in *text*."""
    count = 0
    for ch in text:
        if ch == " ":
            count += 1
        else:
            break
    return count


def _check_leading_spaces_instead_of_indent(
    document: "Document",
) -> Iterable[Finding]:
    threshold = max(2, int(LEADING_SPACES_MIN_RUN))
    for index, paragraph in enumerate(document.paragraphs):
        # Use the first run's text — ``paragraph.text`` flattens
        # ``<w:tab/>`` to ``\t`` and would mis-attribute leading
        # whitespace on a paragraph whose first run starts with a tab.
        runs = paragraph.runs
        if not runs:
            continue
        first_text = runs[0].text
        if not first_text:
            continue
        space_count = _leading_space_count(first_text)
        if space_count < threshold:
            continue
        # Skip the same structural carriers as ``tab-instead-of-indent``:
        # heading / TOC paragraphs (where leading whitespace is part of
        # the rendered numbering leader) and list paragraphs (whose
        # indent is controlled by ``w:numPr``).
        if _is_heading_or_toc(paragraph) or _is_list_paragraph(paragraph):
            continue
        # Hanging-indent body styles (``List Paragraph``, ``Body Text
        # Indent``, ``Quote``) intentionally lead with multi-space
        # padding — defer to ``multiple-spaces``'s skip-list so the
        # two rules stay in sync.
        if _is_indented_style(paragraph):
            continue
        plural = "s" if space_count != 1 else ""
        yield Finding(
            rule="leading-spaces-instead-of-indent",
            severity="info",
            message=(
                f"paragraph {index} starts with {space_count} leading "
                f"space{plural}"
            ),
            paragraph_index=index,
            autofix_available=True,
            autofix_description=(
                "replace leading space-run with paragraph left-indent"
            ),
            location=f"paragraph {index}",
            details=MappingProxyType(
                {
                    "space_count": space_count,
                    "threshold": threshold,
                }
            ),
        )


def _autofix_leading_spaces_instead_of_indent(
    document: "Document", finding: Finding
) -> bool:
    # Local import — mirrors ``_autofix_tab_instead_of_indent``; keeps
    # ``docx.shared`` out of module import time.
    from docx.shared import Emu, Pt

    if finding.paragraph_index is None:
        return False
    try:
        paragraph = document.paragraphs[finding.paragraph_index]
    except IndexError:
        return False
    runs = paragraph.runs
    if not runs:
        return False
    # Defence in depth: re-check the skip conditions in case the
    # document was edited between ``lint()`` and ``autofix()``, or a
    # caller fabricated a Finding by hand.
    if _is_heading_or_toc(paragraph) or _is_list_paragraph(paragraph):
        return False
    if _is_indented_style(paragraph):
        return False
    threshold = max(2, int(LEADING_SPACES_MIN_RUN))
    first_run = runs[0]
    original = first_run.text
    space_count = _leading_space_count(original)
    if space_count < threshold:
        return False
    stripped = original[space_count:]
    first_run.text = stripped
    # Each ``threshold``-wide block of leading spaces converts to one
    # tab-stop's worth of left_indent — i.e. the same 0.5 inch (36 pt)
    # multiple ``tab-instead-of-indent`` uses. A run shorter than the
    # threshold is dropped (the rule only fires above threshold), but
    # a partial block (e.g. 6 spaces with threshold 4) only counts as
    # one block — the leftover two spaces would re-enter the body text
    # if we credited them as a stop. Use floor division so the indent
    # never overshoots what the author appears to have intended.
    blocks = space_count // threshold
    if blocks <= 0:
        return False
    pf = paragraph.paragraph_format
    existing = pf.left_indent
    addition = Pt(_TAB_INDENT_PT * blocks)
    pf.left_indent = Emu(int(existing or 0) + int(addition))
    return True


def _check_mixed_quotes(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        smart_count = sum(1 for ch in text if ch in _SMART_QUOTES)
        straight_count = sum(1 for ch in text if ch in _STRAIGHT_QUOTES)
        if smart_count and straight_count:
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
                details=MappingProxyType(
                    {
                        "smart_count": smart_count,
                        "straight_count": straight_count,
                    }
                ),
            )


# XML element local names whose presence in a paragraph means the paragraph
# carries load-bearing layout / annotation intent even when its plain text is
# empty. The empty-paragraph rule must skip such paragraphs — auto-deleting
# them silently destroys page breaks, bookmarks, comment anchors, etc.
#
# Issue #656: this list is the ground truth for "what protects a paragraph
# from the empty-paragraph autofix". Add to it; do not remove from it.
_STRUCTURAL_EMPTY_BLOCKERS: Tuple[str, ...] = (
    "br",  # <w:br> — page, column, textWrapping or line breaks
    "tab",  # <w:tab/>
    "drawing",  # <w:drawing> — inline / floating images, charts, ink
    "pict",  # <w:pict> — legacy VML drawings
    "object",  # <w:object> — embedded OLE
    "bookmarkStart",  # <w:bookmarkStart> / End — anchor targets
    "bookmarkEnd",
    "commentRangeStart",  # comment-range markers
    "commentRangeEnd",
    "commentReference",
    "sdt",  # <w:sdt> — structured-document-tag (content controls)
    "sdtContent",  # <w:sdtContent> — content-control body
    "contentPart",  # <w:contentPart> — ink annotations
    "fldChar",  # <w:fldChar> — complex field begin/separate/end
    "fldSimple",  # <w:fldSimple> — simple field
    "hyperlink",  # <w:hyperlink> — anchor or external link
)


def _paragraph_has_structural_content(paragraph: "Paragraph") -> bool:
    """Return ``True`` when *paragraph* carries XML with layout / annotation
    intent that must not be discarded as "empty drift".

    A paragraph whose plain text is empty may still carry a page break, a
    bookmark anchor, a comment-range marker, an SDT, ink, a field, a
    hyperlink, etc. The ``empty-paragraph`` rule must skip such
    paragraphs — both at finding time (so the autofix is never offered)
    and inside the autofix callback (defence in depth, since callers can
    build :class:`Finding` instances directly via :func:`register_rule`).

    Implementation note: the checks read the underlying ``_p`` element
    via its ``xpath`` helper. That is the same pattern the rest of
    ``docx.text.paragraph`` uses to expose ``has_page_break``,
    ``has_section_break``, ``drawings``, ``ink_annotations``, etc., so
    the linter is not introducing a new private-XML coupling.
    """

    p = getattr(paragraph, "_p", None)
    if p is None:  # pragma: no cover - defensive, every Paragraph has _p
        return False
    # Section-property carrier inside <w:pPr> — a section break.
    pPr = getattr(p, "pPr", None)
    if pPr is not None and getattr(pPr, "sectPr", None) is not None:
        return True
    for local_name in _STRUCTURAL_EMPTY_BLOCKERS:
        if p.xpath(f".//w:{local_name}"):
            return True
    return False


def _paragraph_is_truly_empty(paragraph: "Paragraph") -> bool:
    """Return ``True`` when *paragraph* is structurally empty — i.e. safe to
    delete as "blank-line drift".

    A paragraph is truly empty only when *both* of the following hold:

    1. ``paragraph.text.strip() == ""`` — its rendered plain text is
       empty (the original loose check).
    2. The underlying ``<w:p>`` element carries no
       :data:`_STRUCTURAL_EMPTY_BLOCKERS` and no ``<w:pPr>/<w:sectPr>``
       — i.e. no page / column / line break, tab, drawing, picture,
       embedded object, bookmark / comment anchor, SDT, ink, field
       (simple or complex), hyperlink, or section break.

    Closes #656 — historically, the loose ``text.strip() == ""`` check
    silently classified paragraphs whose only content was a
    ``<w:br w:type="page"/>`` (or any other structural sibling) as
    "empty", and the autofix deleted them, destroying load-bearing
    page breaks, section breaks, bookmark anchors, etc.
    """

    text = paragraph.text
    if text and text.strip():
        return False
    if _paragraph_has_structural_content(paragraph):
        return False
    return True


def _check_empty_paragraph(document: "Document") -> Iterable[Finding]:
    paragraphs = document.paragraphs
    in_run = False
    run_start: Optional[int] = None
    for index, paragraph in enumerate(paragraphs):
        # Use the tightened predicate — a paragraph is "empty" only when
        # both its rendered text is blank *and* it carries no
        # load-bearing XML (page break, bookmark, comment anchor, SDT,
        # section properties, field, hyperlink, ink, embedded object).
        # See ``_paragraph_is_truly_empty`` and #656.
        if _paragraph_is_truly_empty(paragraph):
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
            # Either the paragraph has visible text (breaks the run) or
            # it carries load-bearing XML — in both cases the run of
            # consecutive empties stops here. A paragraph carrying
            # structural content is never a finding and never counts
            # toward the "consecutive empties" run.
            in_run = False
            run_start = None


def _autofix_empty_paragraph(document: "Document", finding: Finding) -> bool:
    if finding.paragraph_index is None:
        return False
    try:
        paragraph = document.paragraphs[finding.paragraph_index]
    except IndexError:
        return False
    # Defence in depth: even when a Finding was hand-built by a caller,
    # never delete a paragraph that carries layout / annotation intent.
    # The tightened predicate in ``_paragraph_is_truly_empty`` catches
    # text-bearing paragraphs *and* paragraphs whose only content is a
    # break, bookmark, comment anchor, SDT, sectPr, field, hyperlink,
    # ink annotation, or embedded object.
    if not _paragraph_is_truly_empty(paragraph):
        return False
    try:
        paragraph.delete()
    except Exception:  # pragma: no cover - defensive
        return False
    return True


def _check_inconsistent_heading_levels(document: "Document") -> Iterable[Finding]:
    # Treat the implicit pre-document state as "level 0" — same value
    # ``Title`` / ``Subtitle`` already report from ``_heading_level`` —
    # so the very first heading is required to be ``Heading 1`` (or a
    # ``Title``-equivalent). A document that starts with ``Heading 2``
    # or deeper is a level-skip from the document root and is flagged.
    previous_level: int = 0
    for index, paragraph in enumerate(document.paragraphs):
        level = _heading_level(paragraph)
        if level is None or level == 0:
            continue
        if level > previous_level + 1:
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
                details=MappingProxyType(
                    {
                        "level": level,
                        "previous_level": previous_level,
                        "skipped": level - previous_level - 1,
                    }
                ),
            )
        previous_level = level


def _check_trailing_heading(document: "Document") -> Iterable[Finding]:
    """Yield a finding for every heading whose document body trails into nothing.

    A heading paragraph is "trailing" iff every block (paragraph or
    table) that follows it in document order is empty — i.e. the
    section title promises content the document never delivers. This
    catches two real-world authoring bugs:

    * the document ends mid-thought with a section title (e.g. the
      last paragraph is ``Heading 1: '11. Glossary'`` and there is
      no glossary body beneath it);
    * the very last pasted line is auto-styled as ``Heading`` by
      Word and the author never noticed.

    The walk uses :meth:`Document.iter_inner_content` so a heading
    followed by a table counts as having body content (a table is
    content, even when its cells are empty — the author has at
    minimum sketched the structure). A heading is flagged only when
    every following block is a paragraph whose ``text.strip()`` is
    empty.

    Read-only — the rule cannot guess what content the author
    intended, so no autofix is offered. The right move is for the
    author to either delete the heading or add the missing section
    content.

    Closes #644.
    """

    try:
        blocks = list(document.iter_inner_content())
    except Exception:  # pragma: no cover - defensive
        return
    # Map each heading paragraph to its position in ``document.paragraphs``
    # so the finding can carry the right ``paragraph_index`` (the public
    # locator the rest of the lint surface uses). ``iter_inner_content``
    # yields paragraphs and tables interleaved; ``document.paragraphs``
    # is paragraphs-only, so we can't use the inner-content index. The
    # mapping keys on the underlying ``<w:p>`` element rather than on
    # ``id(paragraph)`` because :class:`Paragraph` proxies are created
    # fresh on each iteration and would not collide on identity.
    paragraph_index_by_p: Dict[int, int] = {
        id(p._p): i for i, p in enumerate(document.paragraphs)
    }
    # Local import — keeps the table type out of module import time.
    from docx.table import Table as _Table

    for i, block in enumerate(blocks):
        if isinstance(block, _Table):
            continue
        # Block is a paragraph; check whether it is a heading.
        level = _heading_level(block)
        if level is None or level == 0:
            # Skip non-headings *and* Title / Subtitle (level 0) — a
            # ``Title`` at end-of-document with no body is a different
            # animal (a one-line cover page), not the "promised section
            # never delivered" pattern this rule targets.
            continue
        # Examine every block after this one. If any is a non-empty
        # *body* paragraph or any table, the heading has body content
        # and is not trailing. A subsequent *heading* doesn't count as
        # body content — it's the next section title, not the missing
        # content for this one. Two adjacent trailing headings at
        # end-of-document are both flagged.
        has_following_content = False
        for follower in blocks[i + 1:]:
            if isinstance(follower, _Table):
                has_following_content = True
                break
            text = follower.text
            if not (text and text.strip()):
                continue
            if _heading_level(follower) is not None:
                # Another heading — body content for *that* heading,
                # not for this one. Keep walking; the next non-empty
                # *non-heading* paragraph is what counts.
                continue
            has_following_content = True
            break
        if has_following_content:
            continue
        heading_text = block.text.strip()
        paragraph_index = paragraph_index_by_p.get(id(block._p))
        location = (
            f"paragraph {paragraph_index}"
            if paragraph_index is not None
            else "document body"
        )
        yield Finding(
            rule="trailing-heading",
            severity="info",
            message=(
                f"heading {heading_text!r} at {location} has no body "
                f"content beneath it"
            ),
            paragraph_index=paragraph_index,
            autofix_available=False,
            autofix_description=None,
            location=location,
            details=MappingProxyType(
                {
                    "heading_level": level,
                    "heading_text": heading_text,
                }
            ),
        )


def _check_missing_alt_text(document: "Document") -> Iterable[Finding]:
    shapes = _document_inline_shapes(document)
    severity = (
        "warning" if _document_has_a11y_intent(document, shapes) else "info"
    )
    # Track each unique image identity so duplicate insertions of the
    # same blob produce one finding instead of N. Shapes with no
    # resolvable identity (charts, SmartArt, malformed blips) always
    # emit — they are the cases most likely to need human attention.
    # ``seen`` maps identity -> ordered list of every occurrence index;
    # the first occurrence becomes the canonical location and the rest
    # surface via ``Finding.details["additional_locations"]``.
    seen: Dict[str, List[int]] = {}
    deferred: List[Tuple[int, str, str]] = []
    for shape_index, shape in enumerate(shapes):
        alt = getattr(shape, "alt_text", None)
        title = getattr(shape, "title", None)
        # Treat a non-empty alt OR title as sufficient — Word's own UI
        # accepts either as a screen-reader hint.
        if alt and alt.strip():
            continue
        if title and title.strip():
            continue
        if _shape_is_decorative(shape):
            continue
        identity = _shape_identity(shape)
        if identity is not None:
            existing = seen.get(identity)
            if existing is None:
                seen[identity] = [shape_index]
                deferred.append((shape_index, identity, severity))
            else:
                existing.append(shape_index)
            continue
        # Unkeyed shape — emit immediately, can't dedupe.
        yield Finding(
            rule="missing-alt-text",
            severity=severity,
            message=f"inline image {shape_index} has no alt text",
            paragraph_index=None,
            autofix_available=False,
            autofix_description=None,
            location=f"inline image {shape_index}",
            details=MappingProxyType(
                {
                    "occurrence_count": 1,
                    "additional_locations": (),
                }
            ),
        )
    for first_index, identity, sev in deferred:
        # The dedupe loop above may have grown the index list; pull the
        # final occurrence list from ``seen`` so the message and details
        # reflect every duplicate.
        occurrences = seen[identity]
        count = len(occurrences)
        additional = tuple(
            f"inline image {idx}" for idx in occurrences[1:]
        )
        if count > 1:
            message = (
                f"inline image {first_index} has no alt text "
                f"(repeated on {count} shapes; same image binary)"
            )
        else:
            message = f"inline image {first_index} has no alt text"
        yield Finding(
            rule="missing-alt-text",
            severity=sev,
            message=message,
            paragraph_index=None,
            autofix_available=False,
            autofix_description=None,
            location=f"inline image {first_index}",
            details=MappingProxyType(
                {
                    "occurrence_count": count,
                    "additional_locations": additional,
                }
            ),
        )
# Conservative serif / sans-serif font sets used by the mixed-fonts
# rule to grade severity. The lists cover the Word/Office defaults that
# show up in real documents; an unknown font name falls back to ``info``
# severity (the prior behaviour). The sets are deliberately small —
# adding a stray entry costs nothing if it's wrong, but a bigger list
# would invite false-positive severity escalations.
_SERIF_FONTS: frozenset[str] = frozenset(
    {
        "Times New Roman",
        "Times",
        "Cambria",
        "Georgia",
        "Garamond",
        "Book Antiqua",
        "Palatino",
        "Palatino Linotype",
        "Constantia",
        "Sitka",
    }
)
_SANS_FONTS: frozenset[str] = frozenset(
    {
        "Calibri",
        "Calibri Light",
        "Arial",
        "Arial Black",
        "Helvetica",
        "Verdana",
        "Tahoma",
        "Aptos",
        "Aptos Display",
        "Segoe UI",
        "Trebuchet MS",
        "Lucida Sans Unicode",
        "Corbel",
    }
)


def _font_clash_straddles_serif_sans(font_names: Iterable[str]) -> bool:
    """Return ``True`` when *font_names* contains both a serif and a
    sans-serif family (the visually loudest mixed-fonts case)."""

    fonts = set(font_names)
    return bool(fonts & _SERIF_FONTS) and bool(fonts & _SANS_FONTS)


def _check_mixed_fonts(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        names = {run.font.name for run in paragraph.runs if run.font.name}
        if len(names) > 1:
            font_names = tuple(sorted(names))
            straddles = _font_clash_straddles_serif_sans(font_names)
            yield Finding(
                rule="mixed-fonts",
                # Issue #680: a serif + sans clash is a much more
                # visible defect than two sans-serif fonts that
                # happen to differ — escalate to ``warning`` for the
                # straddling case so an LLM-author lint pass treats
                # it as actionable rather than noise.
                severity="warning" if straddles else "info",
                message=(
                    f"paragraph {index} uses multiple font families: "
                    + ", ".join(font_names)
                ),
                paragraph_index=index,
                autofix_available=False,
                autofix_description=None,
                location=f"paragraph {index}",
                # Issue #680: callers should not have to regex-parse
                # the message to recover the offending font names.
                details=MappingProxyType(
                    {
                        "font_names": font_names,
                        "count": len(font_names),
                        "straddles_serif_sans": straddles,
                    }
                ),
            )


def _document_filename_stem(document: "Document") -> Optional[str]:
    """Best-effort guess at the filename stem the document was loaded from.

    The :func:`docx.Document` factory automatically records the load path
    on the document as the public-ish ``_loaded_from_path`` attribute
    when called with a ``str`` / :class:`os.PathLike` argument, so this
    works out-of-the-box for documents loaded from disk::

        from docx import Document
        from docx.kit import lint

        doc = Document("draft.docx")            # _loaded_from_path auto-set
        report = lint.lint(doc)
        report.autofix(rules=["missing-document-title"])

    Callers loading from an in-memory stream can pass the filename
    explicitly via :func:`lint`'s ``source_path`` keyword (which sets
    the side-channel attribute for the duration of the lint pass) or
    by assigning ``document._loaded_from_path`` directly.

    For back-compat the legacy ``_lint_filename`` attribute is also
    consulted; ``_loaded_from_path`` wins when both are present.

    Falls back to scanning the package / part for a stored path
    attribute should one ever be added to the core API.
    """

    # Prefer the public-ish ``_loaded_from_path`` attribute, fall back
    # to the legacy private ``_lint_filename`` for code that still
    # writes to the older name (issue #648).
    for attr in ("_loaded_from_path", "_lint_filename"):
        hint = getattr(document, attr, None)
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
    """``missing-document-title`` (info, autofix-from-filename).

    Fires when ``document.core_properties.title`` is empty *and* a
    filename hint is available. The hint comes from (in priority
    order): ``document._loaded_from_path`` (set automatically by the
    :func:`docx.Document` factory when loaded from disk),
    ``document._lint_filename`` (legacy back-compat name),
    :func:`lint`'s ``source_path=`` kwarg, or a path attribute on the
    package / part. When no hint is available the finding is
    suppressed (issue #648 — emitting a permanent info finding the
    caller can't act on is just noise).
    """
    try:
        title = document.core_properties.title
    except Exception:  # pragma: no cover - defensive
        title = None
    if title and title.strip():
        return
    stem = _document_filename_stem(document)
    if stem is None:
        # No filename hint available — there's no autofix path and no
        # actionable signal to the caller, so stay silent rather than
        # emitting a permanent ``info`` finding the user can't address.
        # When a hint becomes available (loaded via ``Document(path)``,
        # passed via ``lint(..., source_path=...)``, or set directly on
        # ``document._loaded_from_path``) the finding fires with the
        # autofix attached.
        return
    yield Finding(
        rule="missing-document-title",
        severity="info",
        message="document core property 'title' is empty",
        paragraph_index=None,
        autofix_available=True,
        autofix_description=f"set core property 'title' to {stem!r}",
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
    config = _current_config()
    threshold = config.over_long_threshold
    exemptions = config.style_exemptions
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        if len(text) <= threshold:
            continue
        # Skip headings — the rule targets body prose, not titles.
        if _heading_level(paragraph) is not None:
            continue
        # Skip styles whose long bodies are bounded by editorial intent
        # rather than reading-line ergonomics (lists, captions,
        # footnotes, quotes, etc.).
        if exemptions and _paragraph_style_name_raw(paragraph) in exemptions:
            continue
        yield Finding(
            rule="over-long-paragraph",
            severity="info",
            message=(
                f"paragraph {index} is {len(text)} characters long "
                f"(threshold {threshold})"
            ),
            paragraph_index=index,
            autofix_available=False,
            autofix_description=None,
            location=f"paragraph {index}",
            details=MappingProxyType(
                {
                    "char_count": len(text),
                    "threshold": threshold,
                }
            ),
        )


def _check_table_without_header_row(document: "Document") -> Iterable[Finding]:
    """Yield a finding for every table whose first row is not flagged as a header.

    A WCAG 1.3.1 (Info & Relationships) accessibility check. Word
    represents the header-row marker as ``<w:trPr>/<w:tblHeader/>`` on
    the row's XML; python-docx exposes it as the public
    :attr:`docx.table._Row.is_header` boolean. When the flag is absent
    Word will not repeat the row when the table breaks across pages and
    screen readers will not announce it as a header.

    Autofix — sets ``rows[0].is_header = True`` on the affected table,
    which adds ``<w:tblHeader/>`` to the row's ``<w:trPr>``. The
    finding's ``details`` mapping carries the table index so the
    autofix can target the same table even after intervening edits
    shift positions.
    """
    try:
        tables = list(document.tables)
    except Exception:  # pragma: no cover - defensive
        return
    for table_index, table in enumerate(tables):
        try:
            rows = list(table.rows)
        except Exception:  # pragma: no cover - defensive
            continue
        if not rows:
            continue
        first_row = rows[0]
        try:
            if first_row.is_header:
                continue
        except Exception:  # pragma: no cover - defensive
            continue
        yield Finding(
            rule="table-without-header-row",
            severity="warning",
            message=(
                f"table {table_index} first row is not flagged as a "
                f"header (w:tblHeader missing); Word will not repeat "
                f"the row across pages and screen readers will not "
                f"announce it as a header"
            ),
            paragraph_index=None,
            autofix_available=True,
            autofix_description=(
                "set rows[0].is_header = True on the affected table "
                "(adds <w:tblHeader/> to the row's <w:trPr>)"
            ),
            location=f"table {table_index}",
            details=MappingProxyType({"table_index": table_index}),
        )


def _autofix_table_without_header_row(
    document: "Document", finding: Finding
) -> bool:
    table_index = finding.details.get("table_index") if finding.details else None
    if table_index is None:
        # Fall back to parsing the location string for hand-built findings.
        loc = finding.location or ""
        prefix = "table "
        if loc.startswith(prefix):
            try:
                table_index = int(loc[len(prefix):])
            except ValueError:
                return False
        else:
            return False
    try:
        tables = list(document.tables)
    except Exception:  # pragma: no cover - defensive
        return False
    if not isinstance(table_index, int) or not (0 <= table_index < len(tables)):
        return False
    table = tables[table_index]
    try:
        rows = list(table.rows)
    except Exception:  # pragma: no cover - defensive
        return False
    if not rows:
        return False
    first_row = rows[0]
    try:
        if first_row.is_header:
            return False
        first_row.is_header = True
    except Exception:  # pragma: no cover - defensive
        return False
    return True


def _check_trailing_empty_paragraph(
    document: "Document",
) -> Iterable[Finding]:
    """Flag trailing empty paragraphs at the very end of the document.

    Closes #677. The standard ``empty-paragraph`` rule only catches the
    second-and-subsequent paragraph in a *consecutive* run, so a single
    trailing empty paragraph at end-of-document (or two — the first is
    silent on the existing rule) is silently shipped.

    This rule complements ``empty-paragraph`` by surfacing every empty
    paragraph in the trailing run, including the first. Word users
    routinely leave a phantom empty paragraph at the bottom; LLM
    authors emit them even more reliably. The autofix removes them
    one-by-one in reverse order.
    """

    paragraphs = document.paragraphs
    if not paragraphs:
        return
    # Walk backwards from the end, collecting trailing empties.
    trailing_indices: List[int] = []
    for i in range(len(paragraphs) - 1, -1, -1):
        if paragraphs[i].text.strip():
            break
        trailing_indices.append(i)
    if not trailing_indices:
        return
    # Word almost always carries a single empty paragraph at the end of
    # body content as a section-properties anchor; flagging that one is
    # noisy. Only surface a finding when the trailing run is two or
    # more, OR when the document is genuinely tiny (<= 3 paragraphs and
    # the last is empty — that's clearly authoring residue).
    if len(trailing_indices) < 2 and len(paragraphs) > 3:
        return
    # Emit findings in document order, oldest-first, so the autofix
    # ordering (reverse-paragraph-index in LintReport.autofix) deletes
    # from the bottom up.
    for idx in sorted(trailing_indices):
        yield Finding(
            rule="trailing-empty-paragraph",
            severity="info",
            message=(
                f"paragraph {idx} is a trailing empty paragraph "
                f"({len(trailing_indices)} trailing empties total)"
            ),
            paragraph_index=idx,
            autofix_available=True,
            autofix_description="remove trailing empty paragraph",
            location=f"paragraph {idx}",
        )


def _autofix_trailing_empty_paragraph(
    document: "Document", finding: Finding
) -> bool:
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



def _check_bare_url(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        if not text:
            continue
        # Pre-filter cheaply before invoking the regex.
        if "http" not in text and "www." not in text:
            continue
        # Collect the visible text of any hyperlinks already present in
        # the paragraph; URL strings appearing inside that visible text
        # are wrapped, so they should not be flagged.
        try:
            hyperlink_texts = [hl.text for hl in paragraph.hyperlinks]
        except Exception:  # pragma: no cover - defensive
            hyperlink_texts = []
        for match in _BARE_URL_RE.finditer(text):
            url = match.group(0).rstrip(_URL_TRAILING_PUNCT)
            if not url:
                continue
            wrapped = any(url in ht for ht in hyperlink_texts)
            if wrapped:
                continue
            yield Finding(
                rule="bare-url",
                severity="info",
                message=(
                    f"paragraph {index} contains bare URL {url!r} "
                    f"not wrapped in a hyperlink"
                ),
                paragraph_index=index,
                autofix_available=False,
                autofix_description=None,
                location=f"paragraph {index}",
            )


def _check_placeholder_text(document: "Document") -> Iterable[Finding]:
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text
        for pattern, category in _PLACEHOLDER_PATTERNS:
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
                # Issue #681: surface the matched placeholder and a
                # stable category tag so consumers can group findings
                # without regex-parsing the message.
                details=MappingProxyType(
                    {
                        "placeholder": match.group(0),
                        "category": category,
                    }
                ),
            )
            break  # one finding per paragraph regardless of how many


def _check_excessive_font_size_variation(
    document: "Document",
) -> Iterable[Finding]:
    # Aggregate every explicit run-level font size on non-heading
    # paragraphs across every story (body, table cells, headers /
    # footers, footnotes / endnotes, comments — see #673). ``run.font.size``
    # is ``None`` when the run inherits from its paragraph / character
    # style — those cases are *not* drift, so we skip them. Sizes are
    # stored as ``Length`` (EMU) but compare and render naturally as
    # point values via ``.pt``.
    sizes: "OrderedDict[int, None]" = OrderedDict()
    iter_paragraphs = getattr(document, "iter_all_paragraphs", None)
    if callable(iter_paragraphs):
        paragraph_iter: Iterable["Paragraph"] = (
            paragraph for paragraph, _location in iter_paragraphs()
        )
    else:  # pragma: no cover - defensive (older Document without #662)
        paragraph_iter = document.paragraphs
    for paragraph in paragraph_iter:
        if _heading_level(paragraph) is not None:
            # Headings are intentionally larger / smaller than body
            # prose; including them would false-positive every styled
            # document.
            continue
        for run in paragraph.runs:
            try:
                size = run.font.size
            except Exception:  # pragma: no cover - defensive
                size = None
            if size is None:
                continue
            try:
                pt_value = int(round(float(size.pt)))
            except Exception:  # pragma: no cover - defensive
                continue
            sizes.setdefault(pt_value, None)
    if len(sizes) <= _EXCESSIVE_FONT_SIZE_THRESHOLD:
        return
    sorted_sizes = sorted(sizes.keys())
    pretty = ", ".join(str(s) for s in sorted_sizes)
    yield Finding(
        rule="excessive-font-size-variation",
        severity="info",
        message=(
            f"document body uses {len(sorted_sizes)} distinct explicit "
            f"font sizes ({pretty} pt); consider consolidating via styles"
        ),
        paragraph_index=None,
        autofix_available=False,
        autofix_description=None,
        location="document body",
    )


# ---------------------------------------------------------------------------
# Built-in rule registration
# ---------------------------------------------------------------------------


BUILTIN_RULES: Tuple[str, ...] = (
    "multiple-spaces",
    "trailing-whitespace",
    "tab-instead-of-indent",
    "leading-spaces-instead-of-indent",
    "mixed-quotes",
    "empty-paragraph",
    "trailing-empty-paragraph",
    "inconsistent-heading-levels",
    "trailing-heading",
    "missing-alt-text",
    "mixed-fonts",
    "missing-document-title",
    "over-long-paragraph",
    "placeholder-text",
    "table-without-header-row",
    "bare-url",
    "excessive-font-size-variation",
)
"""The built-in rule identifiers, in registration order."""


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
    register_rule(
        "leading-spaces-instead-of-indent",
        _check_leading_spaces_instead_of_indent,
        _autofix_leading_spaces_instead_of_indent,
    )
    register_rule("mixed-quotes", _check_mixed_quotes)
    register_rule(
        "empty-paragraph", _check_empty_paragraph, _autofix_empty_paragraph
    )
    register_rule(
        "trailing-empty-paragraph",
        _check_trailing_empty_paragraph,
        _autofix_trailing_empty_paragraph,
    )
    register_rule(
        "inconsistent-heading-levels", _check_inconsistent_heading_levels
    )
    register_rule("trailing-heading", _check_trailing_heading)
    register_rule("missing-alt-text", _check_missing_alt_text)
    register_rule("mixed-fonts", _check_mixed_fonts)
    register_rule(
        "missing-document-title",
        _check_missing_document_title,
        _autofix_missing_document_title,
    )
    register_rule("over-long-paragraph", _check_over_long_paragraph)
    register_rule("placeholder-text", _check_placeholder_text)
    register_rule(
        "table-without-header-row",
        _check_table_without_header_row,
        _autofix_table_without_header_row,
    )
    register_rule("bare-url", _check_bare_url)
    register_rule(
        "excessive-font-size-variation",
        _check_excessive_font_size_variation,
    )


_install_builtin_rules()


# ---------------------------------------------------------------------------
# Re-exports (purely for clearer ``help(docx.kit.lint)`` output)
# ---------------------------------------------------------------------------


def _typing_aliases() -> Tuple[Any, ...]:  # pragma: no cover - documentation aid
    return (Union[str, os.PathLike], Sequence[str])
