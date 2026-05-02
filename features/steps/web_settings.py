"""Step implementations for web-settings features."""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context

from docx import Document

from helpers import test_docx


# given ===================================================


@given("a document having non-default web settings")
def given_a_document_having_non_default_web_settings(context: Context):
    context.document = Document(test_docx("web-settings"))


# when ====================================================


@when("I assign web_settings.{flag} = True")
def when_assign_web_settings_true(context: Context, flag: str):
    web = context.document.web_settings
    assert web is not None
    setattr(web, flag, True)


@when("I assign web_settings.{flag} = None")
def when_assign_web_settings_none(context: Context, flag: str):
    web = context.document.web_settings
    assert web is not None
    setattr(web, flag, None)


# then ====================================================


@then("document.web_settings is not None")
def then_document_web_settings_not_none(context: Context):
    assert context.document.web_settings is not None


@then("web_settings.{flag} is True")
def then_web_settings_flag_true(context: Context, flag: str):
    web = context.document.web_settings
    assert web is not None
    actual = getattr(web, flag)
    assert actual is True, f"expected True, got {actual!r}"


@then("web_settings.{flag} is False")
def then_web_settings_flag_false(context: Context, flag: str):
    web = context.document.web_settings
    assert web is not None
    actual = getattr(web, flag)
    assert actual is False, f"expected False, got {actual!r}"
