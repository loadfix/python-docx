"""Step implementations for custom-document-properties features."""

from __future__ import annotations

import ast

from behave import given, then, when
from behave.runner import Context

from docx import Document

from helpers import test_docx


# given ===================================================


@given("a document having known custom properties")
def given_a_document_having_known_custom_properties(context: Context):
    context.document = Document(test_docx("doc-customprops"))


@given("a fresh default document")
def given_a_fresh_default_document(context: Context):
    context.document = Document()


# when ====================================================


@when('I assign document.custom_properties["{name}"] = "{value}"')
def when_assign_custom_property_str(context: Context, name: str, value: str):
    context.document.custom_properties[name] = value


@when('I call document.custom_properties.add("{name}", "{value}")')
def when_call_custom_properties_add(context: Context, name: str, value: str):
    context.document.custom_properties.add(name, value)


@when('I delete document.custom_properties["{name}"]')
def when_delete_custom_property(context: Context, name: str):
    del context.document.custom_properties[name]


# then ====================================================


@then("document.custom_properties has length {count:d}")
def then_custom_properties_length(context: Context, count: int):
    actual = len(context.document.custom_properties)
    assert actual == count, f"expected {count}, got {actual}"


@then("document.custom_properties names are {names}")
def then_custom_properties_names(context: Context, names: str):
    expected = ast.literal_eval(names)
    actual = context.document.custom_properties.names()
    assert actual == expected, f"expected {expected}, got {actual}"


@then('document.custom_properties["{name}"] is {value}')
def then_custom_property_value(context: Context, name: str, value: str):
    actual = context.document.custom_properties[name]
    expected: object
    if value == "True":
        expected = True
    elif value == "False":
        expected = False
    elif value == "None":
        expected = None
    else:
        try:
            expected = int(value)
        except ValueError:
            try:
                expected = float(value)
            except ValueError:
                expected = value
    assert actual == expected, f"expected {expected!r}, got {actual!r}"


@then('"{name}" is in document.custom_properties')
def then_name_in_custom_properties(context: Context, name: str):
    assert name in context.document.custom_properties, (
        f"expected {name!r} to be in custom_properties"
    )


@then('"{name}" is not in document.custom_properties')
def then_name_not_in_custom_properties(context: Context, name: str):
    assert name not in context.document.custom_properties, (
        f"expected {name!r} not to be in custom_properties"
    )


@then('document.custom_properties.get("{name}") is None')
def then_custom_properties_get_none(context: Context, name: str):
    actual = context.document.custom_properties.get(name)
    assert actual is None, f"expected None, got {actual!r}"


@then("assigning a list to document.custom_properties raises TypeError")
def then_assigning_list_raises_typeerror(context: Context):
    try:
        context.document.custom_properties["Bad"] = [1, 2, 3]
    except TypeError:
        return
    raise AssertionError("expected TypeError")
