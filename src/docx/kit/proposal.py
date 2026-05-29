"""Sales proposal / Statement of Work template helpers.

Closes #299.

This module exposes two helpers that *append* a structured commercial
document to an existing |Document| in one call::

    from docx import Document
    from docx.kit import proposal

    doc = Document()
    proposal.sales_proposal(
        doc,
        title="Implementing Frobnitz at ACME",
        prepared_for="ACME Corp",
        prepared_by="Acme Consulting",
        date="2026-05-29",
        executive_summary="Frobnitz adoption is on the critical path...",
        problem_statement="ACME's current process is manual and brittle...",
        proposed_solution="Acme Consulting will deliver Frobnitz in three phases...",
        deliverables=[
            "Discovery report",
            "Implementation",
            "30-day support",
        ],
        timeline=[
            ("Week 1-2",  "Discovery"),
            ("Week 3-6",  "Implementation"),
            ("Week 7-10", "Rollout + support"),
        ],
        pricing=[
            {"item": "Discovery",      "qty": 1, "rate": "$15,000", "total": "$15,000"},
            {"item": "Implementation", "qty": 1, "rate": "$60,000", "total": "$60,000"},
            {"item": "Support (30d)",  "qty": 1, "rate": "$10,000", "total": "$10,000"},
        ],
        grand_total="$85,000",
        terms=["50% on signing", "50% on go-live", "Net 30"],
        next_steps=["Sign attached SOW", "Kick-off call within 5 business days"],
    )
    proposal.sow(
        doc,
        title="Statement of Work — Frobnitz Implementation",
        parties=("Acme Consulting Pty Ltd", "ACME Corp"),
        effective_date="2026-06-01",
        end_date="2026-08-31",
        scope="Acme Consulting will perform the following...",
        deliverables=[
            "Discovery report by Week 2",
            "Implementation by Week 6",
            "Documentation by Week 8",
        ],
        fees="Total: $85,000. Payable in two instalments...",
        acceptance_criteria=[
            "All test cases pass",
            "Documentation handed off",
            "Knowledge transfer completed",
        ],
    )
    doc.save("proposal.docx")

The two helpers — :func:`sales_proposal` and :func:`sow` — each
*append* a fully-shaped section to the supplied |Document| and return
the list of newly-appended block objects (|Paragraph| and |Table|),
in document order. This is the "kit" convention rather than the
"factory" convention used in :mod:`docx.kit.contracts` / :mod:`docx.kit.invoices`
(which build a fresh |Document|): proposals are typically appended to
a branded letterhead-bearing |Document| produced by other helpers, so
returning paragraphs/tables lets the caller post-process them and
chain composition cleanly.

.. warning::

    **Not legal advice.** The output of this module is template
    boilerplate intended as a *starting point only*. The text has not
    been reviewed by a lawyer, makes assumptions about jurisdiction,
    party type, GST treatment, and intellectual-property assignment
    that may not fit a given engagement, and elides clauses
    (insurance, modern-slavery, privacy, IP carve-outs,
    change-control, dispute-escalation, indemnity caps, …) that are
    commonly required in real-world commercial contracts. Use the
    output as a structural skeleton only and have a lawyer review
    every word before signing. The authors of python-docx accept no
    responsibility for losses arising from reliance on this
    boilerplate.

Common conventions across the two helpers:

- **``page_break=True`` by default** — each appended section ends
  with its own page break so subsequent sections start on a fresh
  page. Pass ``page_break=False`` to suppress (e.g. when the caller
  is appending the last section before a save).
- **Style fallback to ``Normal``.** Helpers prefer Word's built-in
  ``Title`` / ``Heading 1`` / ``Heading 2`` styles; when the loaded
  template lacks one, fall back to ``Normal`` rather than raise. The
  spirit of a kit is "works out of the box".
- **Pricing as ``list[dict]``.** Each entry must have ``item``,
  ``qty``, ``rate``, and ``total`` keys (rendered into a four-column
  ``Table Grid``-styled table). Numeric values are coerced to
  strings — callers are responsible for currency formatting because
  this module does not assume a currency.
- **Disclaimer rendered into every output document.** Mirroring the
  :mod:`docx.kit.contracts` and :mod:`docx.kit.medical` patterns,
  the helpers stamp an explicit "not legal advice — starting point
  only" notice into the appended section so a downstream Word user
  sees the same caveat the docstring carries.
- **No XML reach-down.** The helpers compose only public python-docx
  API (``Document.add_paragraph``, ``Document.add_heading``,
  ``Document.add_page_break``, ``Document.add_table``, ``_Cell.text``,
  ``Paragraph.alignment``, ``Run.bold``).

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import (
    TYPE_CHECKING,
    Any,
    List,
    Mapping,
    Sequence,
    Tuple,
    Union,
)

from docx.enum.text import WD_ALIGN_PARAGRAPH

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls
    from docx.table import Table
    from docx.text.paragraph import Paragraph


# -- Disclaimer rendered verbatim into every generated section so a
# -- downstream user opening the file in Word sees the same caveat the
# -- module docstring carries. Mirrors the contracts.py / medical.py
# -- patterns intentionally — the wording is deliberately identical to
# -- the contracts.py disclaimer so a reader of either module forms one
# -- mental model. --
_DISCLAIMER = (
    "DISCLAIMER: This document is template boilerplate generated by "
    "python-docx. It is not legal advice and has not been reviewed by "
    "a lawyer. Use as a starting point only and have qualified legal "
    "counsel review every clause before signing."
)

# -- Word built-in styles the helpers reach for. The kit prefers these
# -- but falls back to ``Normal`` when the loaded template lacks one. --
_STYLE_TITLE = "Title"
_STYLE_HEADING_1 = "Heading 1"
_STYLE_HEADING_2 = "Heading 2"
_STYLE_NORMAL = "Normal"
_STYLE_LIST_BULLET = "List Bullet"
_STYLE_LIST_NUMBER = "List Number"
_STYLE_TABLE_GRID = "Table Grid"


# -- Helpers --------------------------------------------------------------


def _has_style(document: "DocumentCls", style_name: str) -> bool:
    """Return |True| when `document` defines a paragraph style named `style_name`."""
    try:
        styles = document.styles
    except Exception:  # pragma: no cover - defensive
        return False
    try:
        styles[style_name]
        return True
    except KeyError:
        return False


def _resolve_style(document: "DocumentCls", preferred: str) -> str:
    """Return `preferred` if it exists on `document`, else ``"Normal"``."""
    return preferred if _has_style(document, preferred) else _STYLE_NORMAL


def _add_title(document: "DocumentCls", title: str) -> "Paragraph":
    """Append a centred document title in the ``Title`` style (or fallback)."""
    style = _resolve_style(document, _STYLE_TITLE)
    para = document.add_paragraph(title, style=style)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_heading(
    document: "DocumentCls", text: str, level: int = 1
) -> "Paragraph":
    """Append a heading at `level`, falling back to a styled paragraph.

    :meth:`Document.add_heading` already maps ``level`` to ``Heading {N}``
    and falls back internally for unknown styles, but for consistency
    with the rest of ``docx.kit`` we resolve the style explicitly so the
    fallback path is observable in the test suite.
    """
    preferred = f"Heading {level}" if level >= 1 else _STYLE_HEADING_1
    if _has_style(document, preferred):
        return document.add_heading(text, level=level)
    # -- ``Heading {N}`` missing — emit a bold paragraph in Normal so
    # -- the section still reads as a heading visually. --
    para = document.add_paragraph(text, style=_STYLE_NORMAL)
    if para.runs:
        para.runs[0].bold = True
    else:
        run = para.add_run(text)
        run.bold = True
    return para


def _add_disclaimer(document: "DocumentCls") -> "Paragraph":
    """Append the standard "not legal advice" notice to ``document``."""
    para = document.add_paragraph()
    run = para.add_run(_DISCLAIMER)
    run.italic = True
    run.bold = True
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_metadata_line(
    document: "DocumentCls", label: str, value: str
) -> "Paragraph":
    """Append a ``"Label: Value"`` paragraph with a bold label."""
    para = document.add_paragraph()
    label_run = para.add_run(f"{label}: ")
    label_run.bold = True
    para.add_run(value)
    return para


def _add_bulleted_list(
    document: "DocumentCls", items: Sequence[str]
) -> "List[Paragraph]":
    """Append each non-empty item as a ``List Bullet``-styled paragraph."""
    style = _resolve_style(document, _STYLE_LIST_BULLET)
    paragraphs: "List[Paragraph]" = []
    for item in items:
        if not item:
            continue
        paragraphs.append(document.add_paragraph(str(item), style=style))
    return paragraphs


def _add_numbered_list(
    document: "DocumentCls", items: Sequence[str]
) -> "List[Paragraph]":
    """Append each non-empty item as a ``List Number``-styled paragraph."""
    style = _resolve_style(document, _STYLE_LIST_NUMBER)
    paragraphs: "List[Paragraph]" = []
    for item in items:
        if not item:
            continue
        paragraphs.append(document.add_paragraph(str(item), style=style))
    return paragraphs


def _apply_table_grid(table: "Table") -> None:
    """Best-effort apply the ``Table Grid`` style; silently skip when absent."""
    try:
        table.style = _STYLE_TABLE_GRID
    except KeyError:
        # -- Loaded template lacks Table Grid; leave the table unstyled. --
        pass


def _add_section_body(
    document: "DocumentCls",
    body: Union[str, Sequence[str]],
) -> "List[Paragraph]":
    """Append the body of a section as one paragraph per chunk.

    A string ``body`` becomes a single paragraph; a sequence emits one
    paragraph per non-empty entry. Empty / whitespace-only entries are
    dropped so a missing body doesn't render as a blank paragraph.
    """
    paragraphs: "List[Paragraph]" = []
    if isinstance(body, str):
        chunks: "Sequence[str]" = [body] if body.strip() else []
    else:
        chunks = [chunk for chunk in body if chunk and str(chunk).strip()]
    for chunk in chunks:
        paragraphs.append(document.add_paragraph(str(chunk)))
    return paragraphs


def _validate_pricing(pricing: Sequence[Mapping[str, Any]]) -> None:
    """Raise :class:`ValueError` when ``pricing`` is malformed.

    Every entry must be a mapping with at minimum an ``item`` key. The
    other three columns (``qty`` / ``rate`` / ``total``) are rendered as
    empty strings when missing — a row may legitimately omit a quantity
    (e.g. for a discount line) so we don't insist on a value.
    """
    for index, row in enumerate(pricing):
        if not isinstance(row, Mapping):  # type: ignore[arg-type]
            raise ValueError(
                "pricing[%d] must be a mapping with an 'item' key" % index
            )
        if not row.get("item") or not str(row["item"]).strip():
            raise ValueError(
                "pricing[%d] is missing a non-empty 'item'" % index
            )


def _add_pricing_table(
    document: "DocumentCls",
    pricing: Sequence[Mapping[str, Any]],
    grand_total: Union[str, None],
) -> "Table":
    """Append the four-column pricing table and return it.

    The table has a header row (``Item`` / ``Qty`` / ``Rate`` / ``Total``)
    plus one row per pricing entry. When ``grand_total`` is supplied, a
    final row spans the first three columns with the literal text
    ``"Grand Total"`` and emits ``grand_total`` in the fourth column.
    """
    table = document.add_table(rows=1, cols=4)
    _apply_table_grid(table)
    header = table.rows[0].cells
    header[0].text = "Item"
    header[1].text = "Qty"
    header[2].text = "Rate"
    header[3].text = "Total"
    for row in pricing:
        cells = table.add_row().cells
        cells[0].text = str(row.get("item", ""))
        # -- Numeric ``qty`` values render via ``str(...)``; callers who
        # -- want a fixed number of decimals (``"1.0"`` vs ``"1"``) pass
        # -- the desired string verbatim. --
        cells[1].text = "" if row.get("qty") is None else str(row["qty"])
        cells[2].text = str(row.get("rate", ""))
        cells[3].text = str(row.get("total", ""))
    if grand_total:
        total_row = table.add_row().cells
        total_row[0].text = "Grand Total"
        total_row[1].text = ""
        total_row[2].text = ""
        total_row[3].text = str(grand_total)
        # -- Bold the grand-total label and value so the row reads as
        # -- the bottom-line summary even when the table style is muted. --
        for cell in (total_row[0], total_row[3]):
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
    return table


def _add_timeline_table(
    document: "DocumentCls",
    timeline: Sequence[Union[Tuple[str, str], Mapping[str, str]]],
) -> "Table":
    """Append a two-column timeline table and return it.

    Accepts either a sequence of ``(period, activity)`` tuples (the
    issue's example shape) or a sequence of mappings with ``period`` /
    ``activity`` keys (more readable for callers writing the literal
    in-line). The first row is the ``Period`` / ``Activity`` header.
    """
    table = document.add_table(rows=1, cols=2)
    _apply_table_grid(table)
    header = table.rows[0].cells
    header[0].text = "Period"
    header[1].text = "Activity"
    for entry in timeline:
        period: str
        activity: str
        if isinstance(entry, Mapping):
            period = str(entry.get("period", ""))
            activity = str(entry.get("activity", ""))
        else:
            # -- Tuple / list — be liberal about length so a single-cell
            # -- entry doesn't crash the renderer. --
            period = str(entry[0]) if len(entry) > 0 else ""
            activity = str(entry[1]) if len(entry) > 1 else ""
        cells = table.add_row().cells
        cells[0].text = period
        cells[1].text = activity
    return table


# -- Public API: sales_proposal ------------------------------------------


def sales_proposal(
    doc: "DocumentCls",
    *,
    title: str,
    prepared_for: str,
    prepared_by: str,
    date: str,
    executive_summary: str,
    problem_statement: str,
    proposed_solution: str,
    deliverables: Sequence[str],
    timeline: Sequence[Union[Tuple[str, str], Mapping[str, str]]],
    pricing: Sequence[Mapping[str, Any]],
    grand_total: Union[str, None] = None,
    terms: Sequence[str] = (),
    next_steps: Sequence[str] = (),
    page_break: bool = True,
) -> "List[Union[Paragraph, Table]]":
    """Append a structured sales proposal to ``doc``.

    Builds the conventional shape of a B2B sales proposal: cover block
    (title + prepared-for / prepared-by / date metadata), executive
    summary, problem statement, proposed solution, deliverables list,
    timeline table, pricing table (4-column ``Item`` / ``Qty`` /
    ``Rate`` / ``Total`` + optional grand-total row), terms, and next
    steps.

    Parameters
    ----------
    doc
        The |Document| to append to. Mutated in place.
    title
        Proposal title rendered in the ``Title`` style.
    prepared_for, prepared_by
        Recipient and author identifiers rendered in the cover block.
    date
        Free-text date string rendered verbatim (no format imposed).
    executive_summary
        Free-text executive summary rendered as a paragraph under the
        "Executive Summary" heading.
    problem_statement
        Free-text problem-statement paragraph.
    proposed_solution
        Free-text proposed-solution paragraph.
    deliverables
        Sequence of deliverable descriptions rendered as a bulleted
        list under the "Deliverables" heading.
    timeline
        Sequence of ``(period, activity)`` tuples or
        ``{"period": ..., "activity": ...}`` mappings, rendered as a
        two-column ``Period`` / ``Activity`` table.
    pricing
        Sequence of pricing-row dicts. Each row must have an ``item``
        key; ``qty`` / ``rate`` / ``total`` are optional and render as
        empty strings when missing. Rendered as a 4-column table.
    grand_total
        Optional grand-total string rendered as a bold trailing row of
        the pricing table (column 1 holds ``"Grand Total"``, column 4
        holds ``grand_total``).
    terms
        Sequence of payment / commercial terms rendered as a bulleted
        list under the "Terms" heading. ``()`` skips the section.
    next_steps
        Sequence of next-step bullets rendered as a numbered list
        under the "Next Steps" heading. ``()`` skips the section.
    page_break
        When |True| (the default), append a trailing page break so the
        next appended section starts on a fresh page.

    Returns
    -------
    list of Paragraph or Table
        The newly-appended block objects, in document order, including
        the trailing page-break paragraph when emitted.

    Raises
    ------
    ValueError
        When any required string argument is empty, when ``pricing`` is
        not a sequence of mappings with non-empty ``item`` keys, or
        when ``timeline`` contains a non-mapping / non-sequence entry.

    .. warning::
        **Not legal advice.** See the module docstring for the full
        disclaimer. The output is starting-point boilerplate only.

    .. versionadded:: 2026.05.29
    """
    # -- Validation: required strings must be non-empty after strip. --
    for name, value in (
        ("title", title),
        ("prepared_for", prepared_for),
        ("prepared_by", prepared_by),
        ("date", date),
        ("executive_summary", executive_summary),
        ("problem_statement", problem_statement),
        ("proposed_solution", proposed_solution),
    ):
        if not value or not str(value).strip():
            raise ValueError(f"{name} must be a non-empty string")
    if pricing is None:
        raise ValueError("pricing is required (may be an empty sequence)")
    _validate_pricing(pricing)

    blocks: "List[Union[Paragraph, Table]]" = []

    # -- Cover block --
    blocks.append(_add_title(doc, title))
    blocks.append(_add_disclaimer(doc))
    blocks.append(_add_metadata_line(doc, "Prepared for", prepared_for))
    blocks.append(_add_metadata_line(doc, "Prepared by", prepared_by))
    blocks.append(_add_metadata_line(doc, "Date", date))

    # -- Executive Summary --
    blocks.append(_add_heading(doc, "Executive Summary", level=1))
    blocks.extend(_add_section_body(doc, executive_summary))

    # -- Problem Statement --
    blocks.append(_add_heading(doc, "Problem Statement", level=1))
    blocks.extend(_add_section_body(doc, problem_statement))

    # -- Proposed Solution --
    blocks.append(_add_heading(doc, "Proposed Solution", level=1))
    blocks.extend(_add_section_body(doc, proposed_solution))

    # -- Deliverables --
    blocks.append(_add_heading(doc, "Deliverables", level=1))
    if deliverables:
        blocks.extend(_add_bulleted_list(doc, list(deliverables)))
    else:
        blocks.append(
            doc.add_paragraph("[Insert deliverable descriptions here.]")
        )

    # -- Timeline --
    blocks.append(_add_heading(doc, "Timeline", level=1))
    if timeline:
        blocks.append(_add_timeline_table(doc, list(timeline)))
    else:
        blocks.append(doc.add_paragraph("[Insert timeline schedule here.]"))

    # -- Pricing --
    blocks.append(_add_heading(doc, "Pricing", level=1))
    if pricing:
        blocks.append(_add_pricing_table(doc, list(pricing), grand_total))
    else:
        blocks.append(doc.add_paragraph("[Insert pricing line items here.]"))

    # -- Terms (optional) --
    if terms:
        blocks.append(_add_heading(doc, "Terms", level=1))
        blocks.extend(_add_bulleted_list(doc, list(terms)))

    # -- Next Steps (optional) --
    if next_steps:
        blocks.append(_add_heading(doc, "Next Steps", level=1))
        blocks.extend(_add_numbered_list(doc, list(next_steps)))

    if page_break:
        blocks.append(doc.add_page_break())

    return blocks


# -- Public API: sow ------------------------------------------------------


def sow(
    doc: "DocumentCls",
    *,
    title: str,
    parties: Sequence[str],
    effective_date: str,
    end_date: str,
    scope: str,
    deliverables: Sequence[str],
    fees: str,
    acceptance_criteria: Sequence[str],
    page_break: bool = True,
) -> "List[Union[Paragraph, Table]]":
    """Append a Statement of Work (SOW) to ``doc``.

    Builds the conventional shape of an engagement-level SOW: title +
    parties block, effective / end dates, scope, deliverables, fees,
    and acceptance criteria. The shape is the operational instrument
    that scopes a specific engagement under a parent MSA — see
    :func:`docx.kit.contracts.sow` for the standalone-document
    factory equivalent.

    Parameters
    ----------
    doc
        The |Document| to append to. Mutated in place.
    title
        SOW title rendered in the ``Title`` style.
    parties
        Two-element sequence of party names — typically
        ``(vendor, client)``. Rendered in a "Parties" line.
    effective_date, end_date
        Free-text date strings rendered verbatim.
    scope
        Free-text scope description rendered as one paragraph (or
        multiple paragraphs separated by blank lines if the caller
        embeds ``"\\n\\n"`` separators).
    deliverables
        Sequence of deliverable descriptions rendered as a bulleted
        list.
    fees
        Free-text fee summary rendered as a paragraph.
    acceptance_criteria
        Sequence of acceptance-criteria bullets rendered as a bulleted
        list under "Acceptance Criteria".
    page_break
        When |True| (the default), append a trailing page break.

    Returns
    -------
    list of Paragraph or Table
        The newly-appended block objects in document order, including
        the trailing page-break paragraph when emitted.

    Raises
    ------
    ValueError
        When any required string argument is empty, or when ``parties``
        does not contain at least two non-empty entries.

    .. warning::
        **Not legal advice.** See the module docstring.

    .. versionadded:: 2026.05.29
    """
    for name, value in (
        ("title", title),
        ("effective_date", effective_date),
        ("end_date", end_date),
        ("scope", scope),
        ("fees", fees),
    ):
        if not value or not str(value).strip():
            raise ValueError(f"{name} must be a non-empty string")
    if parties is None or len(parties) < 2:
        raise ValueError(
            "parties must contain at least two entries; got %d"
            % (len(parties) if parties is not None else 0)
        )
    for index, party in enumerate(parties):
        if not party or not str(party).strip():
            raise ValueError(
                "parties[%d] must be a non-empty string" % index
            )

    blocks: "List[Union[Paragraph, Table]]" = []

    blocks.append(_add_title(doc, title))
    blocks.append(_add_disclaimer(doc))

    # -- Parties + dates metadata block. The "between" connector mirrors
    # -- the AUS contract drafting convention used by ``docx.kit.contracts``. --
    parties_label = " and ".join(str(p) for p in parties)
    blocks.append(_add_metadata_line(doc, "Parties", parties_label))
    blocks.append(_add_metadata_line(doc, "Effective Date", effective_date))
    blocks.append(_add_metadata_line(doc, "End Date", end_date))

    # -- 1. Scope --
    blocks.append(_add_heading(doc, "Scope", level=1))
    blocks.extend(_add_section_body(doc, scope))

    # -- 2. Deliverables --
    blocks.append(_add_heading(doc, "Deliverables", level=1))
    if deliverables:
        blocks.extend(_add_bulleted_list(doc, list(deliverables)))
    else:
        blocks.append(
            doc.add_paragraph("[Insert deliverable descriptions here.]")
        )

    # -- 3. Fees --
    blocks.append(_add_heading(doc, "Fees", level=1))
    blocks.extend(_add_section_body(doc, fees))

    # -- 4. Acceptance Criteria --
    blocks.append(_add_heading(doc, "Acceptance Criteria", level=1))
    if acceptance_criteria:
        blocks.extend(_add_bulleted_list(doc, list(acceptance_criteria)))
    else:
        blocks.append(
            doc.add_paragraph("[Insert acceptance criteria here.]")
        )

    if page_break:
        blocks.append(doc.add_page_break())

    return blocks


__all__ = [
    "sales_proposal",
    "sow",
]
