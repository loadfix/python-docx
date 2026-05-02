"""Step implementations for chart-read behave features."""

from __future__ import annotations

import ast

from behave import given, then
from behave.runner import Context

from docx import Document
from docx.chart import Chart, WD_CHART_TYPE

from helpers import test_docx

# given ====================================================


@given("a document having three embedded charts")
def given_a_document_having_three_embedded_charts(context: Context):
    context.document = Document(test_docx("chart-has-charts"))


@given("a document having no charts")
def given_a_document_having_no_charts(context: Context):
    context.document = Document(test_docx("doc-default"))


@given("a document having a chart with no title")
def given_a_document_having_a_chart_with_no_title(context: Context):
    # -- build a document in-memory whose single chart has no c:title --
    document = Document()
    document.add_chart(
        WD_CHART_TYPE.PIE,
        ["A", "B", "C"],
        {"Slices": [1.0, 2.0, 3.0]},
    )
    context.document = document
    context.chart = document.charts[0]


# then =====================================================


@then("document.charts is a list of three Chart objects")
def then_document_charts_is_a_list_of_three_chart_objects(context: Context):
    charts = context.document.charts
    assert isinstance(charts, list), f"expected list, got {type(charts)}"
    assert len(charts) == 3, f"expected 3 charts, got {len(charts)}"
    for idx, chart in enumerate(charts):
        assert isinstance(chart, Chart), (
            f"expected Chart at index {idx}, got {type(chart)}"
        )


@then("iterating document.charts yields Chart objects in document order")
def then_iterating_document_charts_yields_charts_in_order(context: Context):
    charts_iter = iter(context.document.charts)
    expected_types = [
        WD_CHART_TYPE.COLUMN,
        WD_CHART_TYPE.BAR,
        WD_CHART_TYPE.LINE,
    ]
    for expected_type in expected_types:
        chart = next(charts_iter)
        assert isinstance(chart, Chart), f"expected Chart, got {type(chart)}"
        assert chart.chart_type == expected_type, (
            f"expected {expected_type}, got {chart.chart_type}"
        )


@then("charts[{idx}].chart_type == WD_CHART_TYPE.{member}")
def then_charts_idx_chart_type_eq_member(context: Context, idx: str, member: str):
    chart = context.document.charts[int(idx)]
    expected = WD_CHART_TYPE[member]
    assert chart.chart_type == expected, (
        f"expected {expected}, got {chart.chart_type}"
    )


@then('charts[{idx}].title == "{title}"')
def then_charts_idx_title_eq_title(context: Context, idx: str, title: str):
    chart = context.document.charts[int(idx)]
    assert chart.title == title, f"expected title {title!r}, got {chart.title!r}"


@then("chart.title is None")
def then_chart_title_is_None(context: Context):
    assert context.chart.title is None, (
        f"expected chart.title to be None, got {context.chart.title!r}"
    )


@then("[s.name for s in charts[{idx}].series] == {names_expr}")
def then_series_names_eq(context: Context, idx: str, names_expr: str):
    chart = context.document.charts[int(idx)]
    expected = ast.literal_eval(names_expr)
    actual = [s.name for s in chart.series]
    assert actual == expected, f"expected series names {expected}, got {actual}"


@then("charts[{idx}].series[{ser_idx}].values == {values_expr}")
def then_series_values_eq(
    context: Context, idx: str, ser_idx: str, values_expr: str
):
    chart = context.document.charts[int(idx)]
    series = chart.series[int(ser_idx)]
    expected = ast.literal_eval(values_expr)
    assert series.values == expected, (
        f"expected values {expected}, got {series.values}"
    )


@then("charts[{idx}].series[{ser_idx}].categories == {categories_expr}")
def then_series_categories_eq(
    context: Context, idx: str, ser_idx: str, categories_expr: str
):
    chart = context.document.charts[int(idx)]
    series = chart.series[int(ser_idx)]
    expected = ast.literal_eval(categories_expr)
    assert series.categories == expected, (
        f"expected categories {expected}, got {series.categories}"
    )


@then("document.charts is an empty list")
def then_document_charts_is_an_empty_list(context: Context):
    charts = context.document.charts
    assert charts == [], f"expected [], got {charts!r}"
