from __future__ import annotations

import tempfile
from pathlib import Path

from behave import given, then, when

from docx import Document


REAL_WORD_FIXTURE = Path("/mnt/data/Temp/microsoft/sample.docx")


@given("a Word-authored docx fixture")
def given_word_authored_fixture(context):
    if not REAL_WORD_FIXTURE.exists():
        context.scenario.skip(
            "Word-authored fixture not available at " + str(REAL_WORD_FIXTURE)
        )
        return
    context.original_path = REAL_WORD_FIXTURE
    context.tmpdir = Path(tempfile.mkdtemp())
    context.roundtrip_path = context.tmpdir / "rt.docx"


@when("I round-trip it through python-docx")
def when_roundtrip_through_python_docx(context):
    Document(str(context.original_path)).save(str(context.roundtrip_path))


@then("ooxml-validate reports zero new issues")
def then_zero_new_issues(context):
    try:
        from ooxml_validate import OoxmlValidateToolNotFound, validate_roundtrip
    except ImportError:
        context.scenario.skip("ooxml-validate not installed")
        return
    try:
        new_issues = validate_roundtrip(context.original_path, context.roundtrip_path)
    except OoxmlValidateToolNotFound as exc:
        context.scenario.skip(f"ooxml-validate dotnet runtime not available: {exc}")
        return
    assert len(new_issues) == 0, (
        f"round-trip introduced {len(new_issues)} new issues: "
        + ", ".join(f"{i.rule_id} in {i.part}" for i in new_issues[:5])
    )


@then("the LibreOffice PDF render is visually identical")
def then_libreoffice_visually_identical(context):
    try:
        from ooxml_validate import libreoffice_pdf_diff
    except ImportError:
        context.scenario.skip("ooxml-validate not installed")
        return
    result = libreoffice_pdf_diff(context.original_path, context.roundtrip_path)
    assert result.ok, (
        f"LibreOffice PDF diff failed: {result.report}, "
        f"pages_differing={result.pages_differing}, max_diff={result.max_difference}"
    )
