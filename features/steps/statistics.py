"""Step implementations for Document.statistics features."""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.statistics import DocumentStatistics

from helpers import test_docx


# given ====================================================


@given("a document with known body text")
def given_a_document_with_known_body_text(context: Context):
    # -- par-known-paragraphs.docx has two non-empty paragraphs:
    # --   "Heading Paragraph" (Heading 1) and "Normal paragraph" (Normal).
    # -- That's 2 paragraphs, 4 whitespace-delimited words, 33 characters
    # -- (including the single space in each of the two paragraph texts),
    # -- and 31 non-space characters.
    context.document = Document(test_docx("par-known-paragraphs"))


# when =====================================================


@when("I access document.statistics")
def when_I_access_document_statistics(context: Context):
    context.statistics = context.document.statistics


# then =====================================================


@then("statistics is a DocumentStatistics object")
def then_statistics_is_a_DocumentStatistics_object(context: Context):
    stats = context.statistics
    assert isinstance(stats, DocumentStatistics), (
        f"expected a DocumentStatistics, got {type(stats)}"
    )


@then("statistics.{field} == {expected}")
def then_statistics_field_eq_expected(context: Context, field: str, expected: str):
    stats = context.statistics
    actual = getattr(stats, field)
    assert actual == int(expected), (
        f"expected statistics.{field} == {expected}, got {actual}"
    )
