"""Step implementations for permission-range features."""

from __future__ import annotations

import ast

from behave import given, then, when
from behave.runner import Context

from docx import Document

from helpers import test_docx


# given ===================================================


@given("a document having three permission ranges")
def given_a_document_having_three_permission_ranges(context: Context):
    context.document = Document(test_docx("prm-ranges"))


# when ====================================================


@when('I call paragraph.add_permission_range(edit_group="{edit_group}")')
def when_call_add_permission_range_edit_group(context: Context, edit_group: str):
    context.paragraph.add_permission_range(edit_group=edit_group)


@when('I call paragraph.add_permission_range(user="{user}")')
def when_call_add_permission_range_user(context: Context, user: str):
    context.paragraph.add_permission_range(user=user)


@when("I call permission_ranges[{idx:d}].delete()")
def when_call_permission_ranges_delete(context: Context, idx: int):
    context.document.permission_ranges[idx].delete()


# then ====================================================


@then("document.permission_ranges has length {count:d}")
def then_permission_ranges_length(context: Context, count: int):
    actual = len(context.document.permission_ranges)
    assert actual == count, f"expected {count}, got {actual}"


@then('permission_ranges[{idx:d}].edit_group == "{value}"')
def then_permission_range_edit_group(context: Context, idx: int, value: str):
    actual = context.document.permission_ranges[idx].edit_group
    assert actual == value, f"expected {value!r}, got {actual!r}"


@then('permission_ranges[{idx:d}].user == "{value}"')
def then_permission_range_user(context: Context, idx: int, value: str):
    actual = context.document.permission_ranges[idx].user
    assert actual == value, f"expected {value!r}, got {actual!r}"


@then("permission_ranges[{idx:d}].edit_group is None")
def then_permission_range_edit_group_none(context: Context, idx: int):
    actual = context.document.permission_ranges[idx].edit_group
    assert actual is None, f"expected None, got {actual!r}"


@then("permission_ranges[{idx:d}].user is None")
def then_permission_range_user_none(context: Context, idx: int):
    actual = context.document.permission_ranges[idx].user
    assert actual is None, f"expected None, got {actual!r}"


@then("permission_ranges have ids {ids}")
def then_permission_range_ids(context: Context, ids: str):
    expected = ast.literal_eval(ids)
    actual = [r.id for r in context.document.permission_ranges]
    assert actual == expected, f"expected {expected}, got {actual}"
