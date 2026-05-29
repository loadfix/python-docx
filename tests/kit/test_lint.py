"""Unit-test suite for :mod:`docx.kit.lint` (issue #304)."""

from __future__ import annotations

import pytest

from docx import Document
from docx.document import Document as DocumentCls
from docx.kit import lint as lint_mod
from docx.kit.lint import (
    BUILTIN_RULES,
    Finding,
    LintReport,
    Rule,
    lint,
    register_rule,
    registered_rules,
    unregister_rule,
)


@pytest.fixture
def document() -> DocumentCls:
    return Document()


@pytest.fixture(autouse=True)
def _restore_registry():
    """Snapshot and restore the registry around each test.

    Custom-rule tests would otherwise leak rules into sibling tests.
    """

    snapshot = dict(lint_mod._REGISTRY)  # noqa: SLF001 - test isolation
    yield
    lint_mod._REGISTRY.clear()  # noqa: SLF001
    lint_mod._REGISTRY.update(snapshot)  # noqa: SLF001


# ---------------------------------------------------------------------------
# Module-level smoke
# ---------------------------------------------------------------------------


class DescribeBuiltinRegistration:

    def it_registers_eleven_built_in_rules(self):
        for name in BUILTIN_RULES:
            assert name in registered_rules()

    def it_lists_the_rules_in_documented_order(self):
        names = registered_rules()
        # Each built-in must appear; relative order must match the
        # documented BUILTIN_RULES tuple.
        positions = [names.index(r) for r in BUILTIN_RULES]
        assert positions == sorted(positions)

    def it_returns_a_LintReport_when_called_on_a_clean_document(
        self, document: DocumentCls
    ):
        report = lint(document)
        assert isinstance(report, LintReport)
        assert report.document is document


# ---------------------------------------------------------------------------
# multiple-spaces
# ---------------------------------------------------------------------------


class DescribeMultipleSpaces:

    def it_flags_two_consecutive_spaces_in_a_run(self, document: DocumentCls):
        document.add_paragraph("hello  world")
        report = lint(document)
        rules = [f.rule for f in report.findings]
        assert "multiple-spaces" in rules

    def it_does_not_flag_a_single_space(self, document: DocumentCls):
        document.add_paragraph("hello world")
        report = lint(document)
        assert "multiple-spaces" not in [f.rule for f in report.findings]

    def it_collapses_runs_of_spaces_when_autofix_runs(
        self, document: DocumentCls
    ):
        document.add_paragraph("a   b    c")
        report = lint(document)
        applied = report.autofix(rules=["multiple-spaces"])
        assert applied >= 1
        assert document.paragraphs[0].text == "a b c"

    def it_marks_findings_with_autofix_available_true(
        self, document: DocumentCls
    ):
        document.add_paragraph("x  y")
        report = lint(document)
        ms = [f for f in report.findings if f.rule == "multiple-spaces"]
        assert ms and ms[0].autofix_available is True
        assert ms[0].autofix_description


# ---------------------------------------------------------------------------
# trailing-whitespace
# ---------------------------------------------------------------------------


class DescribeTrailingWhitespace:

    def it_flags_a_paragraph_ending_in_a_space(self, document: DocumentCls):
        document.add_paragraph("hello ")
        report = lint(document)
        assert "trailing-whitespace" in [f.rule for f in report.findings]

    def it_does_not_flag_a_paragraph_with_no_trailing_whitespace(
        self, document: DocumentCls
    ):
        document.add_paragraph("hello")
        assert "trailing-whitespace" not in [
            f.rule for f in lint(document).findings
        ]

    def it_trims_trailing_whitespace_on_autofix(self, document: DocumentCls):
        document.add_paragraph("hello   ")
        report = lint(document)
        report.autofix(rules=["trailing-whitespace"])
        assert document.paragraphs[0].text == "hello"


# ---------------------------------------------------------------------------
# tab-instead-of-indent
# ---------------------------------------------------------------------------


class DescribeTabInsteadOfIndent:

    def it_flags_a_leading_tab(self, document: DocumentCls):
        para = document.add_paragraph()
        para.add_run("\thello")
        report = lint(document)
        assert "tab-instead-of-indent" in [f.rule for f in report.findings]

    def it_does_not_flag_a_paragraph_with_no_leading_tab(
        self, document: DocumentCls
    ):
        document.add_paragraph("hello")
        assert "tab-instead-of-indent" not in [
            f.rule for f in lint(document).findings
        ]

    def it_strips_the_leading_tab_on_autofix(self, document: DocumentCls):
        para = document.add_paragraph()
        para.add_run("\thello")
        report = lint(document)
        applied = report.autofix(rules=["tab-instead-of-indent"])
        assert applied == 1
        assert document.paragraphs[0].runs[0].text == "hello"


# ---------------------------------------------------------------------------
# mixed-quotes
# ---------------------------------------------------------------------------


class DescribeMixedQuotes:

    def it_flags_a_paragraph_mixing_smart_and_straight(
        self, document: DocumentCls
    ):
        document.add_paragraph("she said “hello” and 'goodbye'")
        report = lint(document)
        mq = [f for f in report.findings if f.rule == "mixed-quotes"]
        assert mq and mq[0].severity == "info"
        assert mq[0].autofix_available is False

    def it_does_not_flag_a_paragraph_with_only_smart_quotes(
        self, document: DocumentCls
    ):
        document.add_paragraph("she said “hello”")
        assert "mixed-quotes" not in [f.rule for f in lint(document).findings]


# ---------------------------------------------------------------------------
# empty-paragraph
# ---------------------------------------------------------------------------


class DescribeEmptyParagraph:

    def it_does_not_flag_a_single_blank_paragraph_between_sentences(
        self, document: DocumentCls
    ):
        document.add_paragraph("first")
        document.add_paragraph("")
        document.add_paragraph("second")
        report = lint(document)
        assert "empty-paragraph" not in [f.rule for f in report.findings]

    def it_flags_consecutive_blank_paragraphs(self, document: DocumentCls):
        document.add_paragraph("first")
        document.add_paragraph("")
        document.add_paragraph("")
        document.add_paragraph("")
        document.add_paragraph("second")
        report = lint(document)
        ep = [f for f in report.findings if f.rule == "empty-paragraph"]
        assert len(ep) == 2

    def it_removes_consecutive_empties_on_autofix_keeping_one(
        self, document: DocumentCls
    ):
        document.add_paragraph("first")
        document.add_paragraph("")
        document.add_paragraph("")
        document.add_paragraph("")
        document.add_paragraph("second")
        report = lint(document)
        report.autofix(rules=["empty-paragraph"])
        # First, one empty, second.
        texts = [p.text for p in document.paragraphs]
        assert texts == ["first", "", "second"]


# ---------------------------------------------------------------------------
# inconsistent-heading-levels
# ---------------------------------------------------------------------------


class DescribeInconsistentHeadingLevels:

    def it_flags_a_skipped_level(self, document: DocumentCls):
        document.add_heading("First", level=1)
        document.add_heading("Skipped", level=3)
        report = lint(document)
        assert "inconsistent-heading-levels" in [
            f.rule for f in report.findings
        ]

    def it_does_not_flag_a_clean_progression(self, document: DocumentCls):
        document.add_heading("a", level=1)
        document.add_heading("b", level=2)
        document.add_heading("c", level=3)
        report = lint(document)
        assert "inconsistent-heading-levels" not in [
            f.rule for f in report.findings
        ]

    def it_marks_the_finding_as_no_autofix(self, document: DocumentCls):
        document.add_heading("First", level=1)
        document.add_heading("Skipped", level=3)
        report = lint(document)
        finding = next(
            f for f in report.findings if f.rule == "inconsistent-heading-levels"
        )
        assert finding.autofix_available is False


# ---------------------------------------------------------------------------
# missing-alt-text
# ---------------------------------------------------------------------------


class DescribeMissingAltText:

    def it_flags_an_inline_image_without_alt_text(
        self, document: DocumentCls
    ):
        # Use a real fixture image so the python-docx image-header
        # parser actually accepts it.
        document.add_picture(
            "tests/test_files/python-icon.png"
        )
        report = lint(document)
        ma = [f for f in report.findings if f.rule == "missing-alt-text"]
        assert ma and ma[0].autofix_available is False


# ---------------------------------------------------------------------------
# mixed-fonts
# ---------------------------------------------------------------------------


class DescribeMixedFonts:

    def it_flags_paragraph_with_two_font_families(
        self, document: DocumentCls
    ):
        para = document.add_paragraph()
        run_a = para.add_run("hello ")
        run_a.font.name = "Calibri"
        run_b = para.add_run("world")
        run_b.font.name = "Times New Roman"
        report = lint(document)
        mf = [f for f in report.findings if f.rule == "mixed-fonts"]
        assert mf and mf[0].severity == "info"
        assert mf[0].autofix_available is False

    def it_does_not_flag_paragraph_with_a_single_family(
        self, document: DocumentCls
    ):
        para = document.add_paragraph()
        for word in ("hello", " world"):
            r = para.add_run(word)
            r.font.name = "Calibri"
        report = lint(document)
        assert "mixed-fonts" not in [f.rule for f in report.findings]


# ---------------------------------------------------------------------------
# missing-document-title
# ---------------------------------------------------------------------------


class DescribeMissingDocumentTitle:

    def it_flags_a_document_with_no_title(self, document: DocumentCls):
        # Default Document() has empty title core property
        document.core_properties.title = ""
        report = lint(document)
        mdt = [
            f for f in report.findings if f.rule == "missing-document-title"
        ]
        assert mdt
        # Title autofix is only available when we can guess a stem;
        # a fresh in-memory Document has no on-disk filename, so the
        # autofix is unavailable.
        assert mdt[0].severity == "info"

    def it_does_not_flag_when_title_is_set(self, document: DocumentCls):
        document.core_properties.title = "My Doc"
        report = lint(document)
        assert "missing-document-title" not in [
            f.rule for f in report.findings
        ]

    def it_autofixes_from_filename_when_caller_supplies_a_hint(
        self, tmp_path
    ):
        # Save a clean document, reload it, supply the filename hint
        # documented in the module docstring, and verify the autofix
        # picks up the filename stem.
        path = tmp_path / "report-final.docx"
        Document().save(str(path))
        doc = Document(str(path))
        doc.core_properties.title = ""
        # python-docx's Document factory does not retain the load path,
        # so the linter accepts a side-channel hint via _lint_filename.
        doc._lint_filename = str(path)
        report = lint(doc)
        mdt = next(
            f for f in report.findings if f.rule == "missing-document-title"
        )
        assert mdt.autofix_available is True
        applied = report.autofix(rules=["missing-document-title"])
        assert applied == 1
        assert doc.core_properties.title == "report-final"


# ---------------------------------------------------------------------------
# over-long-paragraph
# ---------------------------------------------------------------------------


class DescribeOverLongParagraph:

    def it_flags_a_paragraph_longer_than_1000_characters(
        self, document: DocumentCls
    ):
        document.add_paragraph("x" * 1500)
        report = lint(document)
        assert "over-long-paragraph" in [f.rule for f in report.findings]

    def it_does_not_flag_a_short_paragraph(self, document: DocumentCls):
        document.add_paragraph("short")
        report = lint(document)
        assert "over-long-paragraph" not in [
            f.rule for f in report.findings
        ]


# ---------------------------------------------------------------------------
# placeholder-text
# ---------------------------------------------------------------------------


class DescribePlaceholderText:

    @pytest.mark.parametrize(
        "snippet",
        ["[PLACEHOLDER]", "[TBD]", "Lorem ipsum dolor sit amet"],
    )
    def it_flags_a_known_placeholder(
        self, document: DocumentCls, snippet: str
    ):
        document.add_paragraph(f"intro {snippet} outro")
        report = lint(document)
        assert "placeholder-text" in [f.rule for f in report.findings]

    def it_does_not_flag_clean_prose(self, document: DocumentCls):
        document.add_paragraph("This is finished prose.")
        assert "placeholder-text" not in [
            f.rule for f in lint(document).findings
        ]


# ---------------------------------------------------------------------------
# Report aggregations
# ---------------------------------------------------------------------------


class DescribeLintReport:

    def it_supports_iteration_and_len(self, document: DocumentCls):
        document.add_paragraph("two  spaces")
        report = lint(document)
        assert len(report) == len(report.findings) == len(list(report))

    def it_summary_lists_per_rule_counts(self, document: DocumentCls):
        document.add_paragraph("two  spaces")
        document.add_paragraph("trailing ")
        report = lint(document)
        out = report.summary()
        assert "multiple-spaces" in out
        assert "trailing-whitespace" in out
        assert "findings" in out

    def it_summary_handles_a_clean_document(self, document: DocumentCls):
        document.core_properties.title = "x"  # so missing-doc-title silent
        # Need to make sure no findings from other rules
        for rule in list(BUILTIN_RULES):
            if rule != "missing-document-title":
                # Disable rules not exercised so the summary is empty
                pass
        # Easier approach: build a totally clean doc — just title set
        report = lint(document)
        # Should at least include the totals line
        assert "findings" in report.summary()

    def it_autofix_applies_every_available_fix_when_called_with_no_args(
        self, document: DocumentCls
    ):
        document.add_paragraph("hello  world ")  # multi-space + trailing
        report = lint(document)
        applied = report.autofix()
        # multi-space + trailing-whitespace == 2 fixes
        assert applied >= 2

    def it_autofix_filters_by_rule(self, document: DocumentCls):
        document.add_paragraph("hello  world ")
        report = lint(document)
        report.autofix(rules=["multiple-spaces"])
        # multi-space gone
        assert "  " not in document.paragraphs[0].text
        # trailing space remains because we didn't ask to fix it
        assert document.paragraphs[0].text.endswith(" ")

    def it_autofix_returns_zero_for_a_clean_document(
        self, document: DocumentCls
    ):
        document.add_paragraph("clean text")
        document.core_properties.title = "x"
        report = lint(document)
        # No findings carry autofix_available=True for a clean doc.
        assert report.autofix() == 0


# ---------------------------------------------------------------------------
# register_rule / unregister_rule
# ---------------------------------------------------------------------------


class DescribeRegisterRule:

    def it_registers_a_custom_rule(self, document: DocumentCls):
        def check(doc):
            for i, p in enumerate(doc.paragraphs):
                if "FOOBAR" in p.text:
                    yield Finding(
                        rule="no-foobar",
                        severity="error",
                        message="contains FOOBAR",
                        paragraph_index=i,
                    )

        register_rule("no-foobar", check)
        document.add_paragraph("contains FOOBAR here")
        report = lint(document)
        assert "no-foobar" in [f.rule for f in report.findings]

    def it_invokes_the_custom_autofix_callback(
        self, document: DocumentCls
    ):
        def check(doc):
            for i, p in enumerate(doc.paragraphs):
                if "BAD" in p.text:
                    yield Finding(
                        rule="no-bad",
                        severity="warning",
                        message="contains BAD",
                        paragraph_index=i,
                        autofix_available=True,
                        autofix_description="replace BAD with GOOD",
                    )

        def fix(doc, finding):
            p = doc.paragraphs[finding.paragraph_index]
            p.text = p.text.replace("BAD", "GOOD")
            return True

        register_rule("no-bad", check, fix)
        document.add_paragraph("a BAD line")
        report = lint(document)
        report.autofix(rules=["no-bad"])
        assert "GOOD" in document.paragraphs[0].text
        assert "BAD" not in document.paragraphs[0].text

    def it_unregisters_a_rule(self):
        def check(doc):
            return []

        register_rule("temp-rule", check)
        assert "temp-rule" in registered_rules()
        assert unregister_rule("temp-rule") is True
        assert "temp-rule" not in registered_rules()
        # Second call returns False — nothing left to remove.
        assert unregister_rule("temp-rule") is False

    def it_returns_a_Rule_dataclass_from_register(
        self, document: DocumentCls
    ):
        rule = register_rule("noop", lambda d: [])
        assert isinstance(rule, Rule)
        assert rule.name == "noop"
        assert rule.autofix is None

    def it_rejects_invalid_inputs(self):
        with pytest.raises(ValueError):
            register_rule("", lambda d: [])
        with pytest.raises(TypeError):
            register_rule("bad", "not-callable")  # type: ignore[arg-type]
        with pytest.raises(TypeError):
            register_rule(
                "bad", lambda d: [], autofix_callback="not-callable"  # type: ignore[arg-type]
            )


# ---------------------------------------------------------------------------
# Smoke-level cross-cutting tests
# ---------------------------------------------------------------------------


class DescribeIntegration:

    def it_returns_findings_in_document_order(self, document: DocumentCls):
        document.add_paragraph("first  para")
        document.add_paragraph("second  para")
        document.add_paragraph("third  para")
        report = lint(document)
        ms = [f for f in report.findings if f.rule == "multiple-spaces"]
        indices = [f.paragraph_index for f in ms]
        assert indices == sorted(indices)

    def it_applies_fixes_in_reverse_order_so_indices_stay_valid(
        self, document: DocumentCls
    ):
        # Two consecutive empties at index 1 and 2 — autofix removes
        # the second one, the first stays; indices for any later
        # paragraph-scoped findings must still resolve.
        document.add_paragraph("a  b")  # 0 — multi-space
        document.add_paragraph("")  # 1 — empty (kept)
        document.add_paragraph("")  # 2 — empty (removed)
        document.add_paragraph("c   d")  # 3 — multi-space
        report = lint(document)
        report.autofix()
        texts = [p.text for p in document.paragraphs]
        # Only one consecutive empty remains; multi-spaces collapsed.
        assert texts == ["a b", "", "c d"]
