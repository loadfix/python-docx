"""Step implementations for font-table features."""

from __future__ import annotations

from behave import given, then
from behave.runner import Context

from docx import Document
from docx.font_table import FontMetadata

from helpers import test_docx


# given ===================================================


@given("a document having a font table")
def given_a_document_having_a_font_table(context: Context):
    context.document = Document(test_docx("fnt-table"))


# then ====================================================


@then("document.font_table is not None")
def then_document_font_table_not_none(context: Context):
    assert context.document.font_table is not None


@then("len(document.font_table) is at least {count:d}")
def then_font_table_len_at_least(context: Context, count: int):
    font_table = context.document.font_table
    assert font_table is not None
    actual = len(font_table)
    assert actual >= count, f"expected at least {count}, got {actual}"


@then('"{name}" is in document.font_table')
def then_name_in_font_table(context: Context, name: str):
    font_table = context.document.font_table
    assert font_table is not None
    assert name in font_table, f"{name!r} not in font_table"


@then('document.font_table["{name}"].name == "{expected}"')
def then_font_table_item_name(context: Context, name: str, expected: str):
    font_table = context.document.font_table
    assert font_table is not None
    actual = font_table[name].name
    assert actual == expected, f"expected {expected!r}, got {actual!r}"


@then('document.font_table["{name}"].panose has length {length:d}')
def then_font_table_item_panose_length(
    context: Context, name: str, length: int
):
    font_table = context.document.font_table
    assert font_table is not None
    panose = font_table[name].panose
    assert panose is not None and len(panose) == length, (
        f"panose was {panose!r}"
    )


@then('document.font_table.get("{name}") is None')
def then_font_table_get_none(context: Context, name: str):
    font_table = context.document.font_table
    assert font_table is not None
    assert font_table.get(name) is None


@then('document.font_table["{name}"] raises KeyError')
def then_font_table_lookup_raises_keyerror(context: Context, name: str):
    font_table = context.document.font_table
    assert font_table is not None
    try:
        font_table[name]
    except KeyError:
        return
    raise AssertionError(f"expected KeyError for {name!r}")


@then("iterating document.font_table yields only FontMetadata objects")
def then_iterating_font_table_yields_fontmetadata(context: Context):
    font_table = context.document.font_table
    assert font_table is not None
    for item in font_table:
        assert isinstance(item, FontMetadata), (
            f"expected FontMetadata, got {type(item).__name__}"
        )
