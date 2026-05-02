"""Step implementations for accessibility (heading-structure) features."""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.accessibility import HeadingIssue

from helpers import test_docx

_OUTLINE_FIXTURES = {
    "a valid": "acc-valid-headings",
    "a missing-H2": "acc-missing-h2",
    "a skipped-level": "acc-skipped-level",
}


# given ====================================================


@given("a document with {outline} heading outline")
def given_a_document_with_outline_heading_outline(context: Context, outline: str):
    fixture_name = _OUTLINE_FIXTURES[outline]
    context.document = Document(test_docx(fixture_name))


# when =====================================================


@when("I call document.validate_heading_structure()")
def when_I_call_document_validate_heading_structure(context: Context):
    context.issues = context.document.validate_heading_structure()


# then =====================================================


@then("the result is a list of {count} HeadingIssue objects")
def then_result_is_a_list_of_count_heading_issues(context: Context, count: str):
    expected = int(count)
    issues = context.issues
    assert isinstance(issues, list), f"expected a list, got {type(issues)}"
    assert len(issues) == expected, (
        f"expected {expected} HeadingIssue objects, got {len(issues)}"
    )
    for issue in issues:
        assert isinstance(issue, HeadingIssue), (
            f"expected HeadingIssue, got {type(issue)}"
        )


@then('the first reported issue has kind ""')
def then_first_reported_issue_has_no_kind(context: Context):
    # -- an empty kind cell means no issues expected --
    assert context.issues == [], (
        f"expected no issues, got {[i.kind for i in context.issues]}"
    )


@then('the first reported issue has kind "{kind}"')
def then_first_reported_issue_has_kind(context: Context, kind: str):
    assert context.issues, "expected at least one issue, got none"
    actual = context.issues[0].kind
    assert actual == kind, f"expected first issue kind '{kind}', got '{actual}'"
