"""Step implementations for the smart-art-create behave feature."""

from __future__ import annotations

import ast
import io
import os
import zipfile

from behave import given, then, when
from behave.runner import Context

from docx import Document

from helpers import saved_docx_path


# given ====================================================


@given("a blank document for SmartArt authoring")
def given_a_blank_document_for_smartart_authoring(context: Context):
    context.document = Document()


# when =====================================================


@when('I add a SmartArt diagram of family "{family}"')
def when_i_add_a_smartart_of_family(context: Context, family: str):
    context.smart_art = context.document.add_smart_art(family)


@when('I add the SmartArt nodes {texts}')
def when_i_add_smartart_nodes(context: Context, texts: str):
    # -- parse "Alpha", "Beta", "Gamma" into a list --
    node_texts = [t.strip().strip('"') for t in texts.split(",")]
    for t in node_texts:
        context.smart_art.add_node(t)


@when("I save the SmartArt document to scratch")
def when_i_save_the_smartart_document_to_scratch(context: Context):
    os.makedirs(os.path.dirname(saved_docx_path), exist_ok=True)
    context.scratch_path = saved_docx_path
    context.document.save(context.scratch_path)


@when("I save and reopen the SmartArt document")
def when_i_save_and_reopen_the_smartart_document(context: Context):
    buf = io.BytesIO()
    context.document.save(buf)
    buf.seek(0)
    context.document = Document(buf)
    # -- locate the (first) SmartArt in the reopened document --
    diagrams = context.document.smart_art
    context.smart_art = diagrams[0] if diagrams else None


# then =====================================================


@then("document.smart_art has length {n:d}")
def then_document_smart_art_has_length(context: Context, n: int):
    actual = len(context.document.smart_art)
    assert actual == n, f"expected {n} SmartArt diagrams, got {actual}"


@then('the last SmartArt\'s data partname ends with "{suffix}"')
def then_last_smart_art_data_partname_ends_with(context: Context, suffix: str):
    partname = context.document.smart_art[-1].data_partname
    assert partname is not None and partname.endswith(suffix), (
        f"expected data_partname to end with {suffix!r}, got {partname!r}"
    )


@then("the SmartArt's node texts are {expr}")
def then_smart_art_node_texts_are(context: Context, expr: str):
    expected = ast.literal_eval(expr)
    actual = [n.text for n in context.smart_art.nodes]
    assert actual == expected, f"expected node texts {expected!r}, got {actual!r}"


@then("the package contains {partname}")
def then_package_contains_partname(context: Context, partname: str):
    with zipfile.ZipFile(context.scratch_path) as z:
        names = z.namelist()
    assert partname in names, (
        f"expected {partname!r} in package, got {names!r}"
    )


@then("word/_rels/document.xml.rels references a diagramData relationship")
def then_document_rels_references_diagram_data(context: Context):
    with zipfile.ZipFile(context.scratch_path) as z:
        rels = z.read("word/_rels/document.xml.rels").decode()
    assert "relationships/diagramData" in rels, (
        "expected a diagramData relationship in word/_rels/document.xml.rels"
    )


@then('add_smart_art("{family}") raises ValueError')
def then_add_smart_art_raises_valueerror(context: Context, family: str):
    try:
        context.document.add_smart_art(family)
    except ValueError:
        return
    raise AssertionError(f"expected ValueError for layout {family!r}")
