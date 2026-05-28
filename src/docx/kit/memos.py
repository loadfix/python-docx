"""Investment memo / business case template family.

Closes #84.

This module exposes two template factories that build entire
strategy-document drafts in one call::

    from docx.kit.memos import investment_memo, business_case

    doc = investment_memo(
        company="Acme Corp",
        sector="SaaS",
        stage="Series B",
        ask="$25M for 18 months runway",
        situation="Acme is the leading mid-market revenue-ops platform...",
        complication="Despite product-market fit, Acme faces churn...",
        question="Should we lead the Series B?",
        answer="Yes, with terms of $25M at $200M post.",
        sections=[
            {"heading": "Market", "body": "..."},
            {"heading": "Team",   "body": "..."},
            {"heading": "Risks",  "body": "..."},
        ],
    )
    doc.save("acme-memo.docx")

    doc = business_case(
        project="Migration to AWS",
        sponsor="CTO",
        objectives=[
            "Reduce ops cost 40%",
            "Improve SLA to 99.95%",
        ],
        options=[
            {"name": "Status quo",         "cost": "$0", "pros": [...], "cons": [...]},
            {"name": "Rehost (lift+shift)", "cost": "$Ym", "pros": [...], "cons": [...]},
            {"name": "Replatform",         "cost": "$Zm", "pros": [...], "cons": [...]},
        ],
        recommendation="Replatform",
        risks=["Skill gap", "Vendor lock-in"],
        timeline="Q3 2026 - Q1 2027",
    )

The two factories — :func:`investment_memo` and :func:`business_case` —
each return a fresh |Document| pre-populated with the conventional
sections of the matching strategy document. Both lean on the
McKinsey-style **SCQA** (Situation / Complication / Question / Answer)
framework that anchors a memo's executive summary in the reader's
context before introducing recommendations.

Common conventions across the two factories:

- **Title page** — centred title plus a labelled metadata block (sector
  / stage / ask for memos; sponsor / timeline for business cases).
- **Executive summary** — for memos, the SCQA paragraphs (each headed
  by its label). For business cases, the recommendation rendered as a
  bold call-out at the top.
- **Body sections** — memos take an arbitrary ``sections`` list (each
  rendered as a ``Heading 1`` plus body paragraphs). Business cases
  emit a fixed structure: Objectives / Options / Recommendation /
  Risks / Timeline.
- **Options table** — business cases render the supplied ``options``
  list as a four-column table (Option / Cost / Pros / Cons) so a
  reader can compare at a glance.
- **Style fallback to ``Normal``.** Helpers prefer Word's built-in
  ``Title`` / ``Heading 1`` / ``Heading 2`` styles; when the loaded
  template lacks one, fall back to ``Normal`` rather than raise.
- **No XML reach-down** — the kit composes only public python-docx API
  (``Document.add_paragraph``, ``Document.add_heading``,
  ``Document.add_page_break``, ``Document.add_table``).

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Mapping, Optional, Sequence, Union

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls
    from docx.text.paragraph import Paragraph


# -- Helpers --------------------------------------------------------------


def _add_title(document: DocumentCls, title: str) -> Paragraph:
    """Append a centred document title in the ``Title`` style (or fallback)."""
    style = "Title"
    try:
        document.styles[style]
    except KeyError:
        style = "Normal"
    para = document.add_paragraph(title, style=style)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_subtitle(document: DocumentCls, text: str) -> Paragraph:
    """Append a centred subtitle in the ``Subtitle`` style (or fallback)."""
    style = "Subtitle"
    try:
        document.styles[style]
    except KeyError:
        style = "Normal"
    para = document.add_paragraph(text, style=style)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_metadata_line(
    document: DocumentCls, label: str, value: str
) -> Paragraph:
    """Append a centred ``"Label: Value"`` paragraph with a bold label."""
    para = document.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    label_run = para.add_run(f"{label}: ")
    label_run.bold = True
    para.add_run(value)
    return para


def _add_scqa_block(
    document: DocumentCls,
    situation: str,
    complication: str,
    question: str,
    answer: str,
) -> List[Paragraph]:
    """Append the four McKinsey-style SCQA paragraphs and return them.

    Each of the four pieces is rendered as its own paragraph led by a
    bold label run (``"Situation: "``, ``"Complication: "``, …) so the
    framework is legible at a glance. The block is preceded by a
    ``Heading 1`` "Executive Summary" so a reader scanning the doc tree
    can jump to it directly.
    """
    paragraphs: List[Paragraph] = []
    paragraphs.append(document.add_heading("Executive Summary", level=1))
    for label, body in (
        ("Situation", situation),
        ("Complication", complication),
        ("Question", question),
        ("Answer", answer),
    ):
        para = document.add_paragraph()
        label_run = para.add_run(f"{label}: ")
        label_run.bold = True
        para.add_run(body)
        paragraphs.append(para)
    return paragraphs


def _add_section(
    document: DocumentCls,
    heading: str,
    body: Union[str, Sequence[str]],
    level: int = 1,
) -> List[Paragraph]:
    """Append a heading + body paragraphs section and return its paragraphs."""
    paragraphs: List[Paragraph] = []
    paragraphs.append(document.add_heading(heading, level=level))
    if isinstance(body, str):
        chunks: Sequence[str] = [body]
    else:
        chunks = list(body)
    for chunk in chunks:
        if not chunk:
            continue
        paragraphs.append(document.add_paragraph(chunk))
    return paragraphs


def _add_bulleted_list(
    document: DocumentCls, items: Sequence[str]
) -> List[Paragraph]:
    """Append each item as a ``List Bullet``-styled paragraph (or fallback)."""
    style = "List Bullet"
    try:
        document.styles[style]
    except KeyError:
        style = "Normal"
    paragraphs: List[Paragraph] = []
    for item in items:
        if not item:
            continue
        paragraphs.append(document.add_paragraph(item, style=style))
    return paragraphs


def _format_pros_cons(values: Optional[Sequence[str]]) -> str:
    """Render a pros / cons list as a newline-joined cell value.

    Word table cells render ``\\n`` as a soft break inside the cell.
    Empty / missing input becomes an empty string so the table column
    still has a value.
    """
    if not values:
        return ""
    return "\n".join(str(v) for v in values if v)


# -- Investment memo ------------------------------------------------------


def investment_memo(
    company: str,
    sector: Optional[str] = None,
    stage: Optional[str] = None,
    ask: Optional[str] = None,
    situation: Optional[str] = None,
    complication: Optional[str] = None,
    question: Optional[str] = None,
    answer: Optional[str] = None,
    sections: Optional[Sequence[Mapping[str, Union[str, Sequence[str]]]]] = None,
    author: Optional[str] = None,
    date: Optional[str] = None,
) -> DocumentCls:
    """Build an investment memo and return the |Document|.

    Generates a complete investment memo with a McKinsey-style SCQA
    executive summary followed by free-form body sections supplied by
    the caller. The shape is the conventional venture / private-equity
    investment-committee memo: title page with company / sector /
    stage / ask metadata, executive summary (SCQA), then deep-dive
    sections (typical examples: Market, Team, Product, Traction,
    Financials, Risks, Recommendation).

    Parameters
    ----------
    company
        Subject company name. Required — rendered into the title.
    sector
        Industry sector (e.g. ``"SaaS"`` / ``"Fintech"``). Rendered in
        the metadata block when supplied.
    stage
        Funding stage (e.g. ``"Series B"`` / ``"Seed"``). Rendered in
        the metadata block when supplied.
    ask
        Free-text ask line (e.g. ``"$25M for 18 months runway"``).
        Rendered in the metadata block when supplied.
    situation, complication, question, answer
        The four pieces of the SCQA executive summary. Each is rendered
        as a paragraph led by a bold label run. ``None`` for any piece
        emits a placeholder so the writer can fill it in later.
    sections
        Sequence of section dicts. Each entry must have a ``heading``
        key; the optional ``body`` key can be a string (one paragraph)
        or a sequence of strings (one paragraph per item). ``None``
        leaves the body empty.
    author
        Optional memo author / sponsor name rendered under the title.
    date
        Optional ISO date rendered under the author line.

    Returns
    -------
    Document
        The freshly-built |Document|. Save with :meth:`Document.save`.

    Raises
    ------
    ValueError
        When ``company`` is empty, or when any entry in ``sections``
        lacks a ``heading``.

    .. versionadded:: 2026.05.29
    """
    if not company or not company.strip():
        raise ValueError("company is required")

    document = Document()

    # -- Title page --
    _add_title(document, f"Investment Memo: {company}")
    if author:
        _add_subtitle(document, author)
    if date:
        date_para = document.add_paragraph(date)
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # -- Metadata block (only labels with non-empty values) --
    if sector:
        _add_metadata_line(document, "Sector", sector)
    if stage:
        _add_metadata_line(document, "Stage", stage)
    if ask:
        _add_metadata_line(document, "Ask", ask)

    # -- Executive summary (SCQA) --
    _add_scqa_block(
        document,
        situation=situation
        or "[Situation: describe the relevant market context.]",
        complication=complication
        or "[Complication: describe the tension or challenge.]",
        question=question
        or "[Question: what decision is being asked of the reader?]",
        answer=answer
        or "[Answer: the recommendation, in one sentence.]",
    )

    # -- Body sections --
    if sections:
        for index, section in enumerate(sections):
            if not isinstance(section, Mapping):  # type: ignore[arg-type]
                raise ValueError(
                    "sections[%d] must be a mapping with a 'heading' key"
                    % index
                )
            heading = section.get("heading")
            if not heading or not str(heading).strip():
                raise ValueError(
                    "sections[%d] is missing a non-empty 'heading'" % index
                )
            body = section.get("body", "")
            if body is None:
                body = ""
            _add_section(document, str(heading), body)  # type: ignore[arg-type]

    return document


# -- Business case --------------------------------------------------------


def business_case(
    project: str,
    sponsor: Optional[str] = None,
    objectives: Optional[Sequence[str]] = None,
    options: Optional[Sequence[Mapping[str, Union[str, Sequence[str]]]]] = None,
    recommendation: Optional[str] = None,
    risks: Optional[Sequence[str]] = None,
    timeline: Optional[str] = None,
    date: Optional[str] = None,
) -> DocumentCls:
    """Build a business case document and return the |Document|.

    Generates a project / programme business case in the conventional
    options-analysis shape: title page with project / sponsor /
    timeline metadata, executive recommendation call-out, objectives,
    options table (one row per option with cost / pros / cons), risks,
    and timeline.

    The options table is the centre-piece — readers should be able to
    scan the four columns (Option / Cost / Pros / Cons) and see the
    trade-off space at a glance. The recommendation line at the top of
    the executive summary tells the reader which row of the table the
    author selected, before they read the analysis.

    Parameters
    ----------
    project
        Project / initiative name. Required — rendered into the title.
    sponsor
        Executive sponsor (e.g. ``"CTO"`` / ``"COO"``). Rendered in the
        metadata block when supplied.
    objectives
        Sequence of objective strings. Rendered as a bulleted list
        under the "Objectives" heading.
    options
        Sequence of option dicts. Each dict may have ``name`` (column
        1), ``cost`` (column 2), ``pros`` (sequence rendered into
        column 3), and ``cons`` (sequence rendered into column 4).
        Rendered as a four-column ``Table Grid``-styled table.
    recommendation
        Free-text recommendation line. Rendered both as a bold call-out
        in the executive summary and as the body of the
        "Recommendation" heading. Should match one of the option
        ``name`` values for the document to read coherently, but no
        validation enforces this — a recommendation that synthesises
        across options ("Hybrid: rehost + replatform") is also valid.
    risks
        Sequence of risk strings. Rendered as a bulleted list under
        the "Risks" heading.
    timeline
        Free-text timeline summary (e.g. ``"Q3 2026 - Q1 2027"``).
        Rendered in the metadata block and as the body of the
        "Timeline" heading.
    date
        Optional ISO date rendered under the title.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``project`` is empty, or when any entry in ``options``
        lacks a ``name``.

    .. versionadded:: 2026.05.29
    """
    if not project or not project.strip():
        raise ValueError("project is required")

    if options is not None:
        for index, option in enumerate(options):
            if not isinstance(option, Mapping):  # type: ignore[arg-type]
                raise ValueError(
                    "options[%d] must be a mapping with a 'name' key"
                    % index
                )
            if not option.get("name"):
                raise ValueError(
                    "options[%d] is missing a non-empty 'name'" % index
                )

    document = Document()

    # -- Title page --
    _add_title(document, f"Business Case: {project}")
    if date:
        date_para = document.add_paragraph(date)
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # -- Metadata block --
    if sponsor:
        _add_metadata_line(document, "Sponsor", sponsor)
    if timeline:
        _add_metadata_line(document, "Timeline", timeline)

    # -- Executive summary --
    document.add_heading("Executive Summary", level=1)
    summary_para = document.add_paragraph()
    summary_para.add_run(
        f"This business case evaluates {project} against the agreed "
        f"objectives and presents the recommended option."
    )
    if recommendation:
        rec_para = document.add_paragraph()
        label_run = rec_para.add_run("Recommendation: ")
        label_run.bold = True
        rec_run = rec_para.add_run(recommendation)
        rec_run.bold = True

    # -- Objectives --
    document.add_heading("Objectives", level=1)
    if objectives:
        _add_bulleted_list(document, list(objectives))
    else:
        document.add_paragraph(
            "[Insert measurable objectives here — one per line.]"
        )

    # -- Options analysis (table) --
    document.add_heading("Options", level=1)
    if options:
        table = document.add_table(rows=1, cols=4)
        try:
            table.style = "Table Grid"
        except KeyError:
            # -- Fall back silently when the template lacks Table Grid --
            pass
        header_cells = table.rows[0].cells
        header_cells[0].text = "Option"
        header_cells[1].text = "Cost"
        header_cells[2].text = "Pros"
        header_cells[3].text = "Cons"
        for option in options:
            row = table.add_row().cells
            row[0].text = str(option.get("name", ""))
            row[1].text = str(option.get("cost", ""))
            pros = option.get("pros")
            cons = option.get("cons")
            row[2].text = _format_pros_cons(
                pros if isinstance(pros, Sequence) and not isinstance(pros, str) else None
            )
            row[3].text = _format_pros_cons(
                cons if isinstance(cons, Sequence) and not isinstance(cons, str) else None
            )
    else:
        document.add_paragraph(
            "[Insert at least two options here — one per row, with "
            "cost, pros, and cons.]"
        )

    # -- Recommendation (full body) --
    document.add_heading("Recommendation", level=1)
    if recommendation:
        document.add_paragraph(recommendation)
    else:
        document.add_paragraph(
            "[State the recommended option and the rationale.]"
        )

    # -- Risks --
    document.add_heading("Risks", level=1)
    if risks:
        _add_bulleted_list(document, list(risks))
    else:
        document.add_paragraph(
            "[Identify the key risks and proposed mitigations.]"
        )

    # -- Timeline --
    document.add_heading("Timeline", level=1)
    if timeline:
        document.add_paragraph(timeline)
    else:
        document.add_paragraph(
            "[Insert milestones and target dates here.]"
        )

    return document


__all__ = [
    "investment_memo",
    "business_case",
]
