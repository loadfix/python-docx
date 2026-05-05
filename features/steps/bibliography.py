"""Step implementations for the bibliography/citation authoring feature.

Shared step definitions (``a fresh default document``, ``I save and reload
the document``) are declared in sibling step modules — see
``custom_properties.py`` and ``content_controls.py`` respectively.
"""

from __future__ import annotations

from behave import then, when
from behave.runner import Context

from docx.oxml.ns import qn


# when ====================================================


@when(
    'I call document.add_citation("{tag}", title="{title}", author="{author}", '
    'year={year:d})'
)
def _bib_when_add_citation(
    context: Context, tag: str, title: str, author: str, year: int
):
    context.document.add_citation(tag, title=title, author=author, year=year)


@when('I add a paragraph with a citation reference for "{tag}"')
def _bib_when_add_citation_reference(context: Context, tag: str):
    p = context.document.add_paragraph("See ")
    context.last_citation_sdt = p.add_citation_reference(tag)


# then ====================================================


@then("document.bibliography has length {count:d}")
def _bib_then_bibliography_length(context: Context, count: int):
    assert len(context.document.bibliography) == count, (
        f"expected {count} sources, got {len(context.document.bibliography)}"
    )


@then('document.bibliography.get_by_tag("{tag}").title is "{title}"')
def _bib_then_title(context: Context, tag: str, title: str):
    hit = context.document.bibliography.get_by_tag(tag)
    assert hit is not None, f"no source with tag {tag!r}"
    assert hit.title == title, (
        f"expected title {title!r}, got {hit.title!r}"
    )


@then('document.bibliography.get_by_tag("{tag}").year is "{year}"')
def _bib_then_year(context: Context, tag: str, year: str):
    hit = context.document.bibliography.get_by_tag(tag)
    assert hit is not None, f"no source with tag {tag!r}"
    assert hit.year == year, f"expected year {year!r}, got {hit.year!r}"


@then('the last paragraph contains a citation sdt referencing "{tag}"')
def _bib_then_citation_sdt_exists(context: Context, tag: str):
    last_p = context.document.paragraphs[-1]
    sdts = last_p._p.findall(qn("w:sdt"))
    for sdt in sdts:
        sdtPr = sdt.find(qn("w:sdtPr"))
        if sdtPr is None or sdtPr.find(qn("w:citation")) is None:
            continue
        sdtContent = sdt.find(qn("w:sdtContent"))
        if sdtContent is None:
            continue
        for instr in sdtContent.iter(qn("w:instrText")):
            if instr.text and tag in instr.text:
                return
    raise AssertionError(
        f"no <w:sdt><w:citation/> referencing tag {tag!r} in the last paragraph"
    )
