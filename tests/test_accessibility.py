# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.accessibility` module."""

from __future__ import annotations

import pytest

from docx.accessibility import (
    EMPTY_HEADING,
    MULTIPLE_H1,
    NO_H1,
    SKIPPED_LEVEL,
    HeadingIssue,
    _heading_level,
    validate_heading_structure,
)

from .unitutil.mock import Mock


def _fake_paragraph(style_name: str | None, text: str = "lorem"):
    """Return a Mock paragraph whose ``.style.name`` and ``.text`` match the args.

    When `style_name` is |None|, the paragraph has no style attached.
    """
    paragraph = Mock(name="Paragraph")
    paragraph.text = text
    if style_name is None:
        paragraph.style = None
    else:
        style = Mock(name="ParagraphStyle")
        style.name = style_name
        paragraph.style = style
    return paragraph


class DescribeHeadingIssue:
    """Unit-test suite for `docx.accessibility.HeadingIssue`."""

    def it_exposes_its_paragraph_kind_and_message(self):
        paragraph_ = Mock(name="Paragraph")
        issue = HeadingIssue(
            paragraph=paragraph_, kind=SKIPPED_LEVEL, message="Heading 3 follows Heading 1"
        )
        assert issue.paragraph is paragraph_
        assert issue.kind == SKIPPED_LEVEL
        assert issue.message == "Heading 3 follows Heading 1"

    def it_is_immutable(self):
        paragraph_ = Mock(name="Paragraph")
        issue = HeadingIssue(paragraph=paragraph_, kind=SKIPPED_LEVEL, message="msg")
        with pytest.raises(Exception):
            issue.kind = MULTIPLE_H1  # type: ignore[misc]


class Describe_heading_level:
    """Unit-test suite for `docx.accessibility._heading_level`."""

    @pytest.mark.parametrize(
        ("style_name", "expected"),
        [
            ("Heading 1", 1),
            ("Heading 2", 2),
            ("Heading 9", 9),
            ("heading 3", 3),
            ("HEADING 4", 4),
            ("  Heading 2  ", 2),
            ("Normal", None),
            ("Title", None),
            ("Heading 10", None),
            ("Heading", None),
            ("Heading1", None),
            ("", None),
            (None, None),
        ],
    )
    def it_returns_the_level_for_a_heading_style(self, style_name, expected):
        paragraph = _fake_paragraph(style_name)
        assert _heading_level(paragraph) == expected

    def it_returns_None_when_paragraph_has_no_style(self):
        paragraph = _fake_paragraph(None)
        assert _heading_level(paragraph) is None


class DescribeValidateHeadingStructure:
    """Unit-test suite for `docx.accessibility.validate_heading_structure`."""

    def it_returns_empty_list_for_no_paragraphs(self):
        assert validate_heading_structure([]) == []

    def it_returns_empty_list_for_document_with_no_headings(self):
        paragraphs = [
            _fake_paragraph("Normal"),
            _fake_paragraph("Body Text"),
            _fake_paragraph(None),
        ]
        assert validate_heading_structure(paragraphs) == []

    def it_returns_empty_list_for_well_formed_heading_structure(self):
        paragraphs = [
            _fake_paragraph("Heading 1"),
            _fake_paragraph("Normal"),
            _fake_paragraph("Heading 2"),
            _fake_paragraph("Normal"),
            _fake_paragraph("Heading 3"),
            _fake_paragraph("Heading 2"),
            _fake_paragraph("Heading 3"),
        ]
        assert validate_heading_structure(paragraphs) == []

    def it_reports_skipped_levels(self):
        h1 = _fake_paragraph("Heading 1")
        h3 = _fake_paragraph("Heading 3")
        paragraphs = [h1, _fake_paragraph("Normal"), h3]

        issues = validate_heading_structure(paragraphs)

        assert len(issues) == 1
        assert issues[0].paragraph is h3
        assert issues[0].kind == SKIPPED_LEVEL
        assert "Heading 3 follows Heading 1" in issues[0].message
        assert "Heading 2 is missing" in issues[0].message

    def it_reports_skipped_levels_spanning_multiple_levels(self):
        h1 = _fake_paragraph("Heading 1")
        h4 = _fake_paragraph("Heading 4")
        issues = validate_heading_structure([h1, h4])

        assert len(issues) == 1
        assert issues[0].kind == SKIPPED_LEVEL
        # -- when the jump spans multiple levels, we name the first missing level --
        assert "Heading 2 is missing" in issues[0].message

    def it_reports_multiple_h1_paragraphs(self):
        h1a = _fake_paragraph("Heading 1")
        h1b = _fake_paragraph("Heading 1")
        h1c = _fake_paragraph("Heading 1")

        issues = validate_heading_structure([h1a, h1b, h1c])

        # -- only the second and subsequent H1s are flagged --
        multi = [i for i in issues if i.kind == MULTIPLE_H1]
        assert len(multi) == 2
        assert multi[0].paragraph is h1b
        assert multi[1].paragraph is h1c

    def it_reports_empty_heading_paragraphs(self):
        empty = _fake_paragraph("Heading 1", text="")
        whitespace = _fake_paragraph("Heading 2", text="   \t\n  ")

        issues = validate_heading_structure([empty, whitespace])

        empties = [i for i in issues if i.kind == EMPTY_HEADING]
        assert len(empties) == 2
        assert empties[0].paragraph is empty
        assert empties[1].paragraph is whitespace
        assert "empty" in empties[0].message.lower()

    def it_does_not_report_empty_heading_for_nonheading_paragraphs(self):
        paragraphs = [
            _fake_paragraph("Normal", text=""),
            _fake_paragraph("Heading 1"),
        ]
        issues = validate_heading_structure(paragraphs)
        assert [i for i in issues if i.kind == EMPTY_HEADING] == []

    def it_reports_no_h1_when_first_heading_is_below_H1(self):
        h2 = _fake_paragraph("Heading 2")
        h3 = _fake_paragraph("Heading 3")

        issues = validate_heading_structure([_fake_paragraph("Normal"), h2, h3])

        no_h1 = [i for i in issues if i.kind == NO_H1]
        assert len(no_h1) == 1
        assert no_h1[0].paragraph is h2
        assert "Heading 1" in no_h1[0].message

    def it_does_not_report_no_h1_when_first_heading_is_H1(self):
        paragraphs = [_fake_paragraph("Heading 1"), _fake_paragraph("Heading 2")]
        issues = validate_heading_structure(paragraphs)
        assert [i for i in issues if i.kind == NO_H1] == []

    def it_ignores_non_heading_paragraphs_between_headings(self):
        h1 = _fake_paragraph("Heading 1")
        h2 = _fake_paragraph("Heading 2")
        paragraphs = [
            h1,
            _fake_paragraph("Normal"),
            _fake_paragraph("Body Text"),
            _fake_paragraph(None),
            h2,
        ]
        assert validate_heading_structure(paragraphs) == []

    def it_reports_multiple_issues_in_document_order(self):
        h2 = _fake_paragraph("Heading 2")  # triggers NO_H1
        h4 = _fake_paragraph("Heading 4")  # triggers SKIPPED_LEVEL
        h1 = _fake_paragraph("Heading 1", text="")  # triggers EMPTY_HEADING

        issues = validate_heading_structure([h2, h4, h1])

        # -- h2 produces NO_H1; h4 produces SKIPPED_LEVEL; h1 produces EMPTY_HEADING --
        kinds = [i.kind for i in issues]
        assert NO_H1 in kinds
        assert SKIPPED_LEVEL in kinds
        assert EMPTY_HEADING in kinds

        # -- issues appear in document order --
        paragraph_order = [h2, h4, h1]
        assert [
            paragraph_order.index(i.paragraph) for i in issues
        ] == sorted([paragraph_order.index(i.paragraph) for i in issues])


class DescribeDocument_validate_heading_structure:
    """Integration test: `Document.validate_heading_structure` delegates to the helper."""

    def it_calls_the_module_function_with_document_paragraphs(self):
        from typing import cast
        from unittest.mock import patch

        from docx.document import Document
        from docx.oxml.document import CT_Document

        from .unitutil.cxml import element

        document_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"just text"'),
        )
        doc = Document(document_elm, Mock(name="DocumentPart"))

        with patch(
            "docx.accessibility.validate_heading_structure",
            return_value=[Mock(name="HeadingIssue")],
        ) as validate_mock:
            result = doc.validate_heading_structure()

        # -- the helper is called exactly once, with a list of Paragraph objects --
        validate_mock.assert_called_once()
        (call_paragraphs,) = validate_mock.call_args.args
        from docx.text.paragraph import Paragraph

        assert isinstance(call_paragraphs, list)
        assert len(call_paragraphs) == 1
        assert isinstance(call_paragraphs[0], Paragraph)
        assert result == validate_mock.return_value
