"""Tests for `Document.lint()` (issue #57)."""

from __future__ import annotations

import pytest

from docx import Document
from docx.lint import LintFinding, Severity, lint_document
from docx.shared import Pt


class DescribeDocumentLint:
    def it_returns_an_empty_list_for_a_well_formed_outline(self):
        document = Document()
        document.add_heading("Title", level=1)
        document.add_heading("Subsection", level=2)
        document.add_paragraph("Body text.")
        findings = document.lint(rules=["heading-skip"])
        assert findings == []

    def it_flags_heading_skip(self):
        document = Document()
        document.add_heading("Top", level=1)
        document.add_heading("Skipped", level=3)
        findings = document.lint(rules=["heading-skip"])
        assert len(findings) == 1
        assert findings[0].rule_id == "heading-skip"
        assert findings[0].severity == Severity.ERROR

    def it_flags_multiple_h1(self):
        document = Document()
        document.add_heading("First", level=1)
        document.add_heading("Second", level=1)
        findings = document.lint(rules=["heading-multiple-h1"])
        assert len(findings) == 1
        assert findings[0].rule_id == "heading-multiple-h1"
        assert findings[0].severity == Severity.WARNING

    def it_emits_info_when_no_h1_present(self):
        document = Document()
        document.add_heading("Sub", level=2)
        findings = document.lint(rules=["heading-no-h1"])
        assert len(findings) == 1
        assert findings[0].rule_id == "heading-no-h1"
        assert findings[0].severity == Severity.INFO
        assert findings[0].paragraph_index is None

    def it_flags_an_empty_heading(self):
        document = Document()
        document.add_heading("   ", level=1)
        findings = document.lint(rules=["heading-empty"])
        assert len(findings) == 1
        assert findings[0].rule_id == "heading-empty"
        assert findings[0].severity == Severity.ERROR

    def it_flags_an_overly_long_heading(self):
        document = Document()
        document.add_heading("X" * 200, level=1)
        findings = document.lint(rules=["heading-too-long"])
        assert len(findings) == 1
        assert findings[0].rule_id == "heading-too-long"
        assert findings[0].severity == Severity.WARNING

    def it_flags_body_text_that_looks_like_a_heading(self):
        document = Document()
        paragraph = document.add_paragraph("My Section")
        run = paragraph.runs[0]
        run.bold = True
        findings = document.lint(rules=["heading-direct-formatting"])
        assert len(findings) == 1
        assert findings[0].rule_id == "heading-direct-formatting"

    def it_flags_large_font_body_text_as_heading_like(self):
        document = Document()
        paragraph = document.add_paragraph("Big Title")
        run = paragraph.runs[0]
        run.font.size = Pt(20)
        findings = document.lint(rules=["heading-direct-formatting"])
        assert len(findings) == 1

    def it_uses_default_rules_when_none_specified(self):
        document = Document()
        document.add_heading("Top", level=1)
        document.add_heading("Skip", level=3)
        findings = document.lint()
        assert any(f.rule_id == "heading-skip" for f in findings)

    def it_accepts_a_callable_rule(self):
        document = Document()
        document.add_paragraph("body")

        def custom_rule(paragraphs):
            yield LintFinding(
                severity="info",
                paragraph_index=0,
                rule_id="custom",
                message="hello",
            )

        findings = document.lint(rules=[custom_rule])
        assert len(findings) == 1
        assert findings[0].rule_id == "custom"

    def it_rejects_an_unknown_rule_id(self):
        document = Document()
        with pytest.raises(ValueError):
            document.lint(rules=["no-such-rule"])

    def it_sorts_findings_by_paragraph_index_then_id(self):
        document = Document()
        document.add_heading("Top", level=1)
        document.add_heading("Skip", level=3)
        document.add_heading("X" * 200, level=4)
        findings = lint_document(document)
        # -- paragraph_index ordered, with None last
        indices = [f.paragraph_index for f in findings]
        assert indices == sorted(
            indices, key=lambda x: x if x is not None else 10**9
        )
