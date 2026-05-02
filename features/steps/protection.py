"""Step implementations for document-protection features."""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.enum.text import WD_PROTECTION

from helpers import test_docx


_PROTECTION_BY_NAME: dict[str, WD_PROTECTION] = {
    member.name: member for member in WD_PROTECTION
}


# given ===================================================


@given("a document having comments-only protection")
def given_a_document_having_comments_only_protection(context: Context):
    context.document = Document(test_docx("doc-protection"))


# when ====================================================


@when("I call settings.disable_protection()")
def when_call_settings_disable_protection(context: Context):
    context.document.settings.disable_protection()


@when("I call settings.enable_protection(WD_PROTECTION.{mode:w}, enforce={enforce:w})")
def when_call_enable_protection_enforce(
    context: Context, mode: str, enforce: str
):
    context.document.settings.enable_protection(
        _PROTECTION_BY_NAME[mode], enforce=enforce == "True"
    )


@when(
    'I call settings.enable_protection(WD_PROTECTION.{mode:w}, password="{password}", enforce={enforce:w})'
)
def when_call_enable_protection_with_password(
    context: Context, mode: str, password: str, enforce: str
):
    context.document.settings.enable_protection(
        _PROTECTION_BY_NAME[mode],
        password=password,
        enforce=enforce == "True",
    )


# then ====================================================


@then("document_protection.mode is WD_PROTECTION.{mode}")
def then_document_protection_mode(context: Context, mode: str):
    dp = context.document.settings.document_protection
    expected = _PROTECTION_BY_NAME[mode]
    assert dp.mode == expected, f"expected {expected!r}, got {dp.mode!r}"


@then("document_protection.mode is None")
def then_document_protection_mode_none(context: Context):
    dp = context.document.settings.document_protection
    assert dp.mode is None, f"expected None, got {dp.mode!r}"


@then("document_protection.enforce is True")
def then_document_protection_enforce_true(context: Context):
    dp = context.document.settings.document_protection
    assert dp.enforce is True, f"expected True, got {dp.enforce!r}"


@then("document_protection.enforce is False")
def then_document_protection_enforce_false(context: Context):
    dp = context.document.settings.document_protection
    assert dp.enforce is False, f"expected False, got {dp.enforce!r}"


@then("document_protection.password_hash is not None")
def then_password_hash_not_none(context: Context):
    dp = context.document.settings.document_protection
    assert dp.password_hash is not None


@then("document_protection.password_hash is None")
def then_password_hash_none(context: Context):
    dp = context.document.settings.document_protection
    assert dp.password_hash is None, f"got {dp.password_hash!r}"


@then("document_protection.password_salt is not None")
def then_password_salt_not_none(context: Context):
    dp = context.document.settings.document_protection
    assert dp.password_salt is not None


@then("document_protection.spin_count == {count:d}")
def then_spin_count(context: Context, count: int):
    dp = context.document.settings.document_protection
    assert dp.spin_count == count, f"expected {count}, got {dp.spin_count!r}"


@then("document_protection.crypto_algorithm_sid == {sid:d}")
def then_crypto_algorithm_sid(context: Context, sid: int):
    dp = context.document.settings.document_protection
    assert dp.crypto_algorithm_sid == sid, (
        f"expected {sid}, got {dp.crypto_algorithm_sid!r}"
    )


@then('document_protection.crypto_provider_type == "{value}"')
def then_crypto_provider_type(context: Context, value: str):
    dp = context.document.settings.document_protection
    assert dp.crypto_provider_type == value, (
        f"expected {value!r}, got {dp.crypto_provider_type!r}"
    )
