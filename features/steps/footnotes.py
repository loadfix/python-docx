"""Step implementations for footnote-related features."""

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.footnotes import Footnote, Footnotes


# given ====================================================


@given("a new document")
def given_a_new_document(context: Context):
    context.document = Document()


@given("a new document with a paragraph")
def given_a_new_document_with_a_paragraph(context: Context):
    document = Document()
    document.add_paragraph("Test paragraph.")
    context.document = document


@given("a new document with two footnotes")
def given_a_new_document_with_two_footnotes(context: Context):
    document = Document()
    para = document.add_paragraph("Test paragraph.")
    run = para.runs[0]
    document.footnotes.add(run, text="First footnote.")
    run2 = para.add_run(" More text.")
    document.footnotes.add(run2, text="Second footnote.")
    context.document = document


@given("a new document with a footnote containing text")
def given_a_new_document_with_a_footnote_containing_text(context: Context):
    document = Document()
    para = document.add_paragraph("Test paragraph.")
    run = para.runs[0]
    footnote = document.footnotes.add(run, text="Original text.")
    context.document = document
    context.footnote = footnote


@given("a new document with a footnote")
def given_a_new_document_with_a_footnote(context: Context):
    document = Document()
    para = document.add_paragraph("Test paragraph.")
    run = para.runs[0]
    document.footnotes.add(run, text="A footnote.")
    context.document = document


# when =====================================================


@when("I add a footnote to the first run")
def when_I_add_a_footnote_to_the_first_run(context: Context):
    run = context.document.paragraphs[-1].runs[0]
    context.footnote = context.document.footnotes.add(run)


@when('I add a footnote with text "{text}" to the first run')
def when_I_add_a_footnote_with_text(context: Context, text: str):
    run = context.document.paragraphs[-1].runs[0]
    context.footnote = context.document.footnotes.add(run, text=text)


@when("I clear the footnote")
def when_I_clear_the_footnote(context: Context):
    context.footnote.clear()


@when("I delete the footnote")
def when_I_delete_the_footnote(context: Context):
    footnotes = list(context.document.footnotes)
    footnotes[0].delete()


# then =====================================================


@then("document.footnotes is a Footnotes object")
def then_document_footnotes_is_a_Footnotes_object(context: Context):
    assert type(context.document.footnotes) is Footnotes


@then("len(document.footnotes) == {count}")
def then_len_document_footnotes_eq(context: Context, count: str):
    actual = len(context.document.footnotes)
    expected = int(count)
    assert actual == expected, (
        f"expected len(document.footnotes) of {expected}, got {actual}"
    )


@then("the footnote has a single paragraph with FootnoteText style")
def then_footnote_has_single_paragraph_with_footnote_text_style(context: Context):
    footnote = context.footnote
    assert len(footnote.paragraphs) == 1
    assert footnote.paragraphs[0]._p.style == "FootnoteText"


@then('the footnote text is "{text}"')
def then_the_footnote_text_is(context: Context, text: str):
    actual = context.footnote.text
    assert actual == text, f"expected footnote text '{text}', got '{actual}'"


@then("the footnote text is empty")
def then_the_footnote_text_is_empty(context: Context):
    actual = context.footnote.text
    assert actual == "", f"expected empty footnote text, got '{actual}'"


@then("iterating document.footnotes yields {count} Footnote objects")
def then_iterating_footnotes_yields_count(context: Context, count: str):
    footnotes = list(context.document.footnotes)
    expected = int(count)
    assert len(footnotes) == expected, (
        f"expected {expected} footnotes, got {len(footnotes)}"
    )
    for fn in footnotes:
        assert isinstance(fn, Footnote)
