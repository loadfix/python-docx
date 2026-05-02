"""Step implementations for theme-part features."""

from __future__ import annotations

from behave import given, then
from behave.runner import Context

from docx import Document
from docx.shared import RGBColor

from helpers import test_docx


# given ===================================================


@given("a document having the default Office theme")
def given_a_document_having_the_default_office_theme(context: Context):
    context.document = Document(test_docx("thm-theme"))


# then ====================================================


@then('document.theme.name == "{name}"')
def then_document_theme_name(context: Context, name: str):
    theme = context.document.theme
    assert theme is not None
    assert theme.name == name, f"expected {name!r}, got {theme.name!r}"


@then("theme.colors.{slot} is a RGBColor")
def then_theme_colors_slot_is_rgbcolor(context: Context, slot: str):
    theme = context.document.theme
    assert theme is not None
    value = getattr(theme.colors, slot)
    assert isinstance(value, RGBColor), (
        f"expected RGBColor for colors.{slot}, got {value!r}"
    )


@then('theme.colors["{token}"] is a RGBColor')
def then_theme_colors_token_is_rgbcolor(context: Context, token: str):
    theme = context.document.theme
    assert theme is not None
    value = theme.colors[token]
    assert isinstance(value, RGBColor), (
        f"expected RGBColor for colors[{token!r}], got {value!r}"
    )


@then('theme.colors["{token}"] raises KeyError')
def then_theme_colors_token_raises_keyerror(context: Context, token: str):
    theme = context.document.theme
    assert theme is not None
    try:
        theme.colors[token]
    except KeyError:
        return
    raise AssertionError(f"expected KeyError for token {token!r}")


@then('theme.fonts.major_latin == "{typeface}"')
def then_theme_fonts_major_latin(context: Context, typeface: str):
    theme = context.document.theme
    assert theme is not None
    actual = theme.fonts.major_latin
    assert actual == typeface, f"expected {typeface!r}, got {actual!r}"


@then('theme.fonts.minor_latin == "{typeface}"')
def then_theme_fonts_minor_latin(context: Context, typeface: str):
    theme = context.document.theme
    assert theme is not None
    actual = theme.fonts.minor_latin
    assert actual == typeface, f"expected {typeface!r}, got {actual!r}"
