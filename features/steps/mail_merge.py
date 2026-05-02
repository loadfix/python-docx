"""Step implementations for mail-merge feature coverage."""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.enum.text import (
    WD_MAIL_MERGE_DATA_TYPE,
    WD_MAIL_MERGE_DESTINATION,
    WD_MAIL_MERGE_TYPE,
)
from docx.settings import MailMerge

from helpers import test_docx

# -- values shared between the "enable with args" when-step and its thens ---
_CONNECT_STRING = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=contacts.xlsx"
_QUERY = "SELECT FirstName, Email FROM [Sheet1$]"

# -- map of textual tokens used in scenario outlines to their expected values --
_ENUM_LOOKUP: dict[str, object] = {
    "WD_MAIL_MERGE_TYPE.FORM_LETTERS": WD_MAIL_MERGE_TYPE.FORM_LETTERS,
    "WD_MAIL_MERGE_TYPE.EMAIL": WD_MAIL_MERGE_TYPE.EMAIL,
    "WD_MAIL_MERGE_TYPE.CATALOG": WD_MAIL_MERGE_TYPE.CATALOG,
    "WD_MAIL_MERGE_DESTINATION.EMAIL": WD_MAIL_MERGE_DESTINATION.EMAIL,
    "WD_MAIL_MERGE_DESTINATION.NEW_DOCUMENT": WD_MAIL_MERGE_DESTINATION.NEW_DOCUMENT,
    "WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET": WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET,
    "WD_MAIL_MERGE_DATA_TYPE.ODBC": WD_MAIL_MERGE_DATA_TYPE.ODBC,
    "True": True,
    "False": False,
    "None": None,
}


def _coerce(token: str) -> object:
    """Return the Python value represented by `token`.

    Tokens that match a known enum member or literal are returned from the
    lookup table; purely-integer tokens are coerced to |int|; everything else
    is returned as the raw string.
    """
    if token in _ENUM_LOOKUP:
        return _ENUM_LOOKUP[token]
    try:
        return int(token)
    except ValueError:
        return token


# given ====================================================


@given("a document having no mail-merge configuration")
def given_a_document_having_no_mail_merge_configuration(context: Context):
    context.document = Document(test_docx("doc-word-default-blank"))


@given("a document with mail-merge enabled")
def given_a_document_with_mail_merge_enabled(context: Context):
    context.document = Document(test_docx("mmg-enabled"))
    context.mail_merge = context.document.settings.mail_merge


# when =====================================================


@when("I call settings.enable_mail_merge()")
def when_I_call_settings_enable_mail_merge(context: Context):
    context.mail_merge = context.document.settings.enable_mail_merge()


@when("I call settings.enable_mail_merge() with realistic arguments")
def when_I_call_settings_enable_mail_merge_with_args(context: Context):
    context.mail_merge = context.document.settings.enable_mail_merge(
        main_document_type=WD_MAIL_MERGE_TYPE.EMAIL,
        destination=WD_MAIL_MERGE_DESTINATION.EMAIL,
        data_type=WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET,
        connect_string=_CONNECT_STRING,
        query=_QUERY,
        mail_subject="Quarterly update",
        address_field_name="Email",
    )


@when("I call settings.disable_mail_merge()")
def when_I_call_settings_disable_mail_merge(context: Context):
    context.document.settings.disable_mail_merge()


# then =====================================================


@then("settings.mail_merge is None")
def then_settings_mail_merge_is_None(context: Context):
    actual = context.document.settings.mail_merge
    assert actual is None, f"expected settings.mail_merge to be None, got {actual!r}"


@then("settings.mail_merge is a MailMerge object")
def then_settings_mail_merge_is_a_MailMerge_object(context: Context):
    mm = context.document.settings.mail_merge
    assert type(mm) is MailMerge, f"expected a MailMerge object, got {type(mm)!r}"


@then("mail_merge.main_document_type == WD_MAIL_MERGE_TYPE.FORM_LETTERS")
def then_mail_merge_main_document_type_is_form_letters(context: Context):
    actual = context.mail_merge.main_document_type
    assert actual == WD_MAIL_MERGE_TYPE.FORM_LETTERS, (
        f"expected main_document_type == FORM_LETTERS, got {actual!r}"
    )


@then("mail_merge.main_document_type == WD_MAIL_MERGE_TYPE.EMAIL")
def then_mail_merge_main_document_type_is_email(context: Context):
    actual = context.mail_merge.main_document_type
    assert actual == WD_MAIL_MERGE_TYPE.EMAIL, (
        f"expected main_document_type == EMAIL, got {actual!r}"
    )


@then("mail_merge.destination is None")
def then_mail_merge_destination_is_None(context: Context):
    actual = context.mail_merge.destination
    assert actual is None, f"expected destination is None, got {actual!r}"


@then("mail_merge.destination == WD_MAIL_MERGE_DESTINATION.EMAIL")
def then_mail_merge_destination_is_email(context: Context):
    actual = context.mail_merge.destination
    assert actual == WD_MAIL_MERGE_DESTINATION.EMAIL, (
        f"expected destination == EMAIL, got {actual!r}"
    )


@then("mail_merge.data_type is None")
def then_mail_merge_data_type_is_None(context: Context):
    actual = context.mail_merge.data_type
    assert actual is None, f"expected data_type is None, got {actual!r}"


@then("mail_merge.data_type == WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET")
def then_mail_merge_data_type_is_spreadsheet(context: Context):
    actual = context.mail_merge.data_type
    assert actual == WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET, (
        f"expected data_type == SPREADSHEET, got {actual!r}"
    )


@then("mail_merge.connect_string is None")
def then_mail_merge_connect_string_is_None(context: Context):
    actual = context.mail_merge.connect_string
    assert actual is None, f"expected connect_string is None, got {actual!r}"


@then("mail_merge.connect_string is the supplied connect_string")
def then_mail_merge_connect_string_matches(context: Context):
    actual = context.mail_merge.connect_string
    assert actual == _CONNECT_STRING, (
        f"expected connect_string {_CONNECT_STRING!r}, got {actual!r}"
    )


@then("mail_merge.query is None")
def then_mail_merge_query_is_None(context: Context):
    actual = context.mail_merge.query
    assert actual is None, f"expected query is None, got {actual!r}"


@then("mail_merge.query is the supplied query")
def then_mail_merge_query_matches(context: Context):
    actual = context.mail_merge.query
    assert actual == _QUERY, f"expected query {_QUERY!r}, got {actual!r}"


@then('mail_merge.mail_subject == "{subject}"')
def then_mail_merge_mail_subject_eq(context: Context, subject: str):
    actual = context.mail_merge.mail_subject
    assert actual == subject, f"expected mail_subject {subject!r}, got {actual!r}"


@then('mail_merge.address_field_name == "{field}"')
def then_mail_merge_address_field_name_eq(context: Context, field: str):
    actual = context.mail_merge.address_field_name
    assert actual == field, f"expected address_field_name {field!r}, got {actual!r}"


@then("mail_merge.{prop} == {expected}")
def then_mail_merge_property_eq(context: Context, prop: str, expected: str):
    mm = context.mail_merge
    assert mm is not None, "mail_merge proxy is None"
    actual = getattr(mm, prop)
    want = _coerce(expected)
    assert actual == want, (
        f"expected mail_merge.{prop} == {want!r} (from token {expected!r}), got {actual!r}"
    )
