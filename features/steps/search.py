"""Step implementations for search/replace features."""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.search import SearchMatch
from docx.text.paragraph import Paragraph

from helpers import test_docx

# given ====================================================


@given("a Document loaded from srh-target-text.docx")
def given_a_document_loaded_from_srh_target_text(context: Context):
    context.document = Document(test_docx("srh-target-text"))


# when =====================================================


@when('I call document.search("{text}")')
def when_I_call_document_search(context: Context, text: str):
    context.result = context.document.search(text)


@when('I call document.search("{text}", case_sensitive=False)')
def when_I_call_document_search_case_insensitive(context: Context, text: str):
    context.result = context.document.search(text, case_sensitive=False)


@when("I call document.search_regex(r\"{pattern}\")")
def when_I_call_document_search_regex(context: Context, pattern: str):
    context.result = context.document.search_regex(pattern)


@when('I call document.search_all("{text}")')
def when_I_call_document_search_all(context: Context, text: str):
    context.result = context.document.search_all(text)


@when('I call document.replace("{old}", "{new}")')
def when_I_call_document_replace(context: Context, old: str, new: str):
    context.returned_count = context.document.replace(old, new)


@when('I call document.replace_all("{old}", "{new}")')
def when_I_call_document_replace_all(context: Context, old: str, new: str):
    context.returned_count = context.document.replace_all(old, new)


@when("I call document.replace_regex(r\"{pattern}\", r\"{replacement}\")")
def when_I_call_document_replace_regex(
    context: Context, pattern: str, replacement: str
):
    context.returned_count = context.document.replace_regex(pattern, replacement)


# then =====================================================


@then("the result is a list of {count:d} SearchMatch objects")
def then_the_result_is_a_list_of_count_searchmatch_objects(
    context: Context, count: int
):
    result = context.result
    assert isinstance(result, list), f"expected a list, got {type(result).__name__}"
    assert len(result) == count, (
        f"expected {count} matches, got {len(result)}: {result!r}"
    )
    for match in result:
        assert isinstance(match, SearchMatch), (
            f"expected SearchMatch, got {type(match).__name__}"
        )


@then("every match.location is None")
def then_every_match_location_is_none(context: Context):
    for match in context.result:
        assert match.location is None, (
            f"expected location None, got {match.location!r}"
        )


@then("every match.paragraph is a Paragraph")
def then_every_match_paragraph_is_a_paragraph(context: Context):
    for match in context.result:
        assert isinstance(match.paragraph, Paragraph), (
            f"expected Paragraph, got {type(match.paragraph).__name__}"
        )


@then('every match.text equals "{text}"')
def then_every_match_text_equals(context: Context, text: str):
    for match in context.result:
        actual = match.paragraph.text[match.start : match.end]
        assert actual == text, f"expected {text!r}, got {actual!r}"


@then("match_texts == [{items}]")
def then_match_texts_eq(context: Context, items: str):
    expected = [s.strip().strip('"') for s in items.split(",")]
    actual = [m.paragraph.text[m.start : m.end] for m in context.result]
    assert actual == expected, f"expected {expected!r}, got {actual!r}"


@then("the returned count is {count:d}")
def then_the_returned_count_is(context: Context, count: int):
    assert context.returned_count == count, (
        f"expected count {count}, got {context.returned_count}"
    )


@then('document.search("{text}") returns {count:d} matches')
def then_document_search_returns_count_matches(
    context: Context, text: str, count: int
):
    matches = context.document.search(text)
    assert len(matches) == count, (
        f"expected {count} matches for {text!r}, got {len(matches)}"
    )


@then('document.search_all("{text}") returns {count:d} matches')
def then_document_search_all_returns_count_matches(
    context: Context, text: str, count: int
):
    matches = context.document.search_all(text)
    assert len(matches) == count, (
        f"expected {count} matches for {text!r}, got {len(matches)}"
    )


@then("at least one match spans multiple runs")
def then_at_least_one_match_spans_multiple_runs(context: Context):
    multi = [m for m in context.result if len(m.run_indices) >= 2]
    assert multi, (
        f"expected at least one multi-run match in {context.result!r}, got none"
    )


@then('match.location starts with "{prefix}"')
def then_match_location_starts_with(context: Context, prefix: str):
    assert len(context.result) >= 1, "expected at least one match"
    match = context.result[0]
    assert match.location is not None and match.location.startswith(prefix), (
        f"expected location to start with {prefix!r}, got {match.location!r}"
    )
