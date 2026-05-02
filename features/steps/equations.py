"""Step implementations for equation (OMML) features.

This module is shared by the equation-related behave feature files:
``equ-identifier.feature``, ``equ-fraction.feature``, and the scenarios that
other equation agents will add for superscript, subscript and radical
builders. Steps here MUST remain idempotent and free of feature-specific
coupling so the cherry-pick merge from separate agent branches composes
cleanly.
"""

from __future__ import annotations

from behave import given, then, when
from behave.runner import Context

from docx import Document
from docx.equations import Equation, build_fraction, build_identifier

from helpers import test_docx

# given ====================================================


@given("a document having identifier equations")
def given_a_document_having_identifier_equations(context: Context):
    context.document = Document(test_docx("equ-identifier"))


@given("a document having fraction equations")
def given_a_document_having_fraction_equations(context: Context):
    context.document = Document(test_docx("equ-fraction"))


@given("a fresh document")
def given_a_fresh_document(context: Context):
    context.document = Document()


# when =====================================================


@when("I assign equation = the first equation in the document")
def when_I_assign_equation_first(context: Context):
    context.equation = context.document.equations[0]


@when("I assign equation = the second equation in the document")
def when_I_assign_equation_second(context: Context):
    context.equation = context.document.equations[1]


@when('I assign xml = build_identifier("{text}")')
def when_I_assign_xml_eq_build_identifier(context: Context, text: str):
    context.xml = build_identifier(text)


@when('I assign xml = build_fraction("{num}", "{den}")')
def when_I_assign_xml_eq_build_fraction(context: Context, num: str, den: str):
    context.xml = build_fraction(num, den)


@when('I append an identifier equation for "{text}" to a new paragraph')
def when_I_append_identifier_equation(context: Context, text: str):
    paragraph = context.document.add_paragraph("prefix ")
    context.paragraph = paragraph
    context.equation = paragraph.add_equation(build_identifier(text))


@when(
    'I append a fraction equation with numerator "{num}" and denominator "{den}"'
    " to a new paragraph"
)
def when_I_append_fraction_equation(context: Context, num: str, den: str):
    paragraph = context.document.add_paragraph("prefix ")
    context.paragraph = paragraph
    context.equation = paragraph.add_equation(build_fraction(num, den))


# then =====================================================


@then("document.equations is a list of {count:d} Equation objects")
def then_document_equations_is_list_of_count(context: Context, count: int):
    equations = context.document.equations
    assert isinstance(equations, list), f"expected list, got {type(equations)}"
    assert len(equations) == count, (
        f"expected {count} equations, got {len(equations)}"
    )
    for eq in equations:
        assert isinstance(eq, Equation), (
            f"expected Equation, got {type(eq)}"
        )


@then('equation.text == "{text}"')
def then_equation_text_eq_text(context: Context, text: str):
    actual = context.equation.text
    assert actual == text, f"expected equation.text {text!r}, got {actual!r}"


@then("equation.is_display_mode is False")
def then_equation_is_display_mode_false(context: Context):
    assert context.equation.is_display_mode is False, (
        f"expected is_display_mode False, got {context.equation.is_display_mode}"
    )


@then("equation.is_display_mode is True")
def then_equation_is_display_mode_true(context: Context):
    assert context.equation.is_display_mode is True, (
        f"expected is_display_mode True, got {context.equation.is_display_mode}"
    )


@then('equation.raw_xml contains "{fragment}"')
def then_equation_raw_xml_contains(context: Context, fragment: str):
    raw = context.equation.raw_xml
    assert fragment.encode("utf-8") in raw, (
        f"expected {fragment!r} in raw_xml, got {raw!r}"
    )


@then('xml starts with "{prefix}"')
def then_xml_starts_with(context: Context, prefix: str):
    assert context.xml.startswith(prefix), (
        f"expected xml to start with {prefix!r}, got {context.xml[:60]!r}"
    )


@then('xml ends with "{suffix}"')
def then_xml_ends_with(context: Context, suffix: str):
    assert context.xml.endswith(suffix), (
        f"expected xml to end with {suffix!r}, got ...{context.xml[-60:]!r}"
    )


@then('xml contains "{fragment}"')
def then_xml_contains(context: Context, fragment: str):
    assert fragment in context.xml, (
        f"expected {fragment!r} in xml, got {context.xml!r}"
    )


@then("the paragraph has {count:d} equation")
@then("the paragraph has {count:d} equations")
def then_paragraph_has_count_equations(context: Context, count: int):
    equations = context.paragraph.equations
    assert len(equations) == count, (
        f"expected {count} equations on paragraph, got {len(equations)}"
    )


@then("xml contains a bar-type fraction property")
def then_xml_contains_bar_type(context: Context):
    fragment = '<m:type m:val="bar"/>'
    assert fragment in context.xml, (
        f"expected {fragment!r} in xml, got {context.xml!r}"
    )


@then('the appended equation.text == "{text}"')
def then_appended_equation_text(context: Context, text: str):
    actual = context.equation.text
    assert actual == text, (
        f"expected appended equation.text {text!r}, got {actual!r}"
    )
