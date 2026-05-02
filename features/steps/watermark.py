"""Step implementations for watermark features."""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.watermark import Watermark

from helpers import test_docx


# given ===================================================


@given("a document having a text watermark")
def given_a_document_having_a_text_watermark(context: Context):
    context.document = Document(test_docx("wmk-text"))


@given("a document having an image watermark")
def given_a_document_having_an_image_watermark(context: Context):
    context.document = Document(test_docx("wmk-image"))


# when ====================================================


@when('I add a text watermark with text "{text}"')
def when_add_text_watermark(context: Context, text: str):
    section = context.document.sections[0]
    section.add_text_watermark(text=text)


@when("I call section.remove_watermark()")
def when_call_remove_watermark(context: Context):
    context.document.sections[0].remove_watermark()


# then ====================================================


@then("section.watermark is a Watermark object")
def then_section_watermark_is_watermark(context: Context):
    wm = context.document.sections[0].watermark
    assert isinstance(wm, Watermark), f"expected Watermark, got {wm!r}"


@then('section.watermark.type == "{value}"')
def then_section_watermark_type(context: Context, value: str):
    wm = context.document.sections[0].watermark
    assert wm is not None
    assert wm.type == value, f"expected {value!r}, got {wm.type!r}"


@then('section.watermark.text == "{value}"')
def then_section_watermark_text(context: Context, value: str):
    wm = context.document.sections[0].watermark
    assert wm is not None
    assert wm.text == value, f"expected {value!r}, got {wm.text!r}"


@then("section.watermark.text is None")
def then_section_watermark_text_is_none(context: Context):
    wm = context.document.sections[0].watermark
    assert wm is not None
    assert wm.text is None, f"expected None, got {wm.text!r}"


@then("section.watermark is None")
def then_section_watermark_is_none(context: Context):
    wm = context.document.sections[0].watermark
    assert wm is None, f"expected None, got {wm!r}"


@then('calling add_text_watermark with layout "{layout}" raises ValueError')
def then_add_text_watermark_raises(context: Context, layout: str):
    section = context.document.sections[0]
    try:
        section.add_text_watermark(text="X", layout=layout)
    except ValueError:
        return
    raise AssertionError(f"expected ValueError for layout={layout!r}")
