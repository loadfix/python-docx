"""Customer case-study / customer-story template helper.

Closes #300.

A *case study* (a.k.a. *customer story*) is the conventional B2B
marketing artefact: one customer's success story laid out so the
reader can quickly answer "who are they, what was the pain, what did
we ship, what was the impact?". Sales / marketing / customer-success
authors emit dozens of these per quarter — they're high-volume,
high-template, low-novelty content that benefits from a single-call
helper.

The :func:`case_study` helper appends a complete case-study section
to a |Document| in one call::

    from docx import Document
    from docx.kit import case_study

    doc = Document()
    case_study.case_study(
        doc,
        title="How ACME cut latency by 80% with FrobnitzPro",
        customer="ACME Corp",
        industry="Manufacturing",
        size="5,000 employees",
        location="Detroit, MI",
        summary="One-paragraph elevator pitch of the case study...",
        challenge="ACME's primary challenge was ...",
        solution="With FrobnitzPro, they were able to...",
        implementation="The rollout took 6 weeks across three phases...",
        results=[
            {"metric": "Latency",      "before": "500ms", "after": "100ms", "delta": "-80%"},
            {"metric": "Throughput",   "before": "1k/s",  "after": "5k/s",  "delta": "+400%"},
            {"metric": "Cost per req", "before": "$0.10", "after": "$0.04", "delta": "-60%"},
        ],
        customer_quote='"FrobnitzPro paid for itself in three months." -- Jane Doe, CTO',
        technologies=["FrobnitzPro 5", "Kubernetes", "PostgreSQL 17"],
        next_steps="ACME plans to expand to their EU region in Q3.",
    )
    doc.save("case_study.docx")

The rendered shape is::

    Title                              (Title style, centred)
    Customer name                      (Heading 2, centred)
    Customer profile (3-column strip)  (Industry / Size / Location)
    Summary                            (Heading 1 + paragraph)
    Challenge                          (Heading 1 + paragraph[s])
    Solution                           (Heading 1 + paragraph[s])
    Implementation                     (Heading 1 + paragraph[s])
    Results                            (Heading 1 + 4-column table)
    Customer Quote                     (Quote-styled block)
    Technologies                       (Heading 2 + bullet list)
    Next Steps                         (Heading 1 + paragraph[s])
    [page break]

Conventions:

- **Compose only python-docx public API** -- ``Document.add_paragraph``,
  ``Document.add_heading``, ``Document.add_table``, ``Document.add_page_break``,
  ``Paragraph.add_run``. No oxml / etree access.
- **Style fallback to ``Normal``** -- helpers prefer Word's built-in
  ``Title`` / ``Heading 1`` / ``Heading 2`` / ``Quote`` / ``List Bullet``
  styles; when the loaded template lacks one, fall back to ``Normal``
  rather than raise.
- **Returns the list of newly-appended objects** -- ``Paragraph`` and
  ``Table`` instances in document order (including the trailing
  page-break paragraph when ``page_break`` is true).

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Mapping, Optional, Sequence, Union

from docx.enum.text import WD_ALIGN_PARAGRAPH

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph


# -- Word built-in styles the helper reaches for. Each one falls back
# -- to "Normal" when the loaded template doesn't ship it. --
_STYLE_TITLE = "Title"
_STYLE_HEADING_1 = "Heading 1"
_STYLE_HEADING_2 = "Heading 2"
_STYLE_QUOTE = "Quote"
_STYLE_LIST_BULLET = "List Bullet"
_STYLE_NORMAL = "Normal"


def _has_style(document: "Document", style_name: str) -> bool:
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


def _resolve_style(document: "Document", preferred: str) -> str:
    """Return `preferred` if it exists on `document`, else ``"Normal"``."""
    return preferred if _has_style(document, preferred) else _STYLE_NORMAL


def _coerce_body(body: Union[str, Sequence[str]]) -> List[str]:
    """Return `body` as a list of paragraph strings.

    A string `body` is split on blank lines (``"\\n\\n"``) so callers
    can pass a single multi-paragraph string and get sensible
    paragraph breaks. A pre-split sequence is returned as a list
    unchanged. Empty strings are dropped to avoid producing empty
    paragraphs.
    """
    if isinstance(body, str):
        chunks = [chunk.strip() for chunk in body.split("\n\n")]
        return [chunk for chunk in chunks if chunk]
    return [chunk for chunk in body if chunk]


def _add_section(
    document: "Document",
    heading: str,
    body: Union[str, Sequence[str]],
    level: int = 1,
) -> List["Paragraph"]:
    """Append a heading + one paragraph per body chunk, return the new paragraphs."""
    paragraphs: List["Paragraph"] = [
        document.add_heading(heading, level=level)
    ]
    for chunk in _coerce_body(body):
        paragraphs.append(document.add_paragraph(chunk))
    return paragraphs


def _add_customer_profile(
    document: "Document",
    industry: Optional[str],
    size: Optional[str],
    location: Optional[str],
) -> List[Union["Paragraph", "Table"]]:
    """Render a 3-column customer-profile metadata strip.

    Renders a one-row, three-column table whose cells each carry a
    bold label + value pair (Industry / Size / Location). The strip
    is the conventional "at-a-glance customer profile" found at the
    top of vendor case studies, and the table layout keeps the three
    facts visually aligned regardless of label / value lengths.

    Cells whose corresponding value is ``None`` or empty render the
    label only (no value) so the column count stays stable across
    the rendered document. When all three values are missing the
    helper renders nothing and returns an empty list -- there's no
    information to display.
    """
    appended: List[Union["Paragraph", "Table"]] = []
    if not (industry or size or location):
        return appended

    table = document.add_table(rows=1, cols=3)
    try:
        table.style = "Table Grid"
    except KeyError:  # pragma: no cover - default template ships Table Grid
        pass

    cells = table.rows[0].cells
    for cell, label, value in (
        (cells[0], "Industry", industry),
        (cells[1], "Size", size),
        (cells[2], "Location", location),
    ):
        # -- Replace the empty default paragraph with a label / value
        # -- run pair. The label is bold, the value is plain. --
        para = cell.paragraphs[0]
        label_run = para.add_run(label)
        label_run.bold = True
        if value:
            para.add_run("\n")
            para.add_run(str(value))

    appended.append(table)
    return appended


def _add_results_table(
    document: "Document",
    results: Sequence[Mapping[str, str]],
) -> "Table":
    """Render the four-column metric / before / after / delta results table."""
    # -- One header row plus one row per result. --
    table = document.add_table(rows=1 + len(results), cols=4)
    try:
        table.style = "Table Grid"
    except KeyError:  # pragma: no cover - default template ships Table Grid
        pass

    header_cells = table.rows[0].cells
    for idx, label in enumerate(("Metric", "Before", "After", "Delta")):
        header_para = header_cells[idx].paragraphs[0]
        run = header_para.add_run(label)
        run.bold = True

    for row_idx, result in enumerate(results, start=1):
        row_cells = table.rows[row_idx].cells
        row_cells[0].text = str(result.get("metric", ""))
        row_cells[1].text = str(result.get("before", ""))
        row_cells[2].text = str(result.get("after", ""))
        row_cells[3].text = str(result.get("delta", ""))

    return table


def _add_customer_quote(
    document: "Document", quote: str
) -> "Paragraph":
    """Append the customer quote in the ``Quote`` style (or ``Normal`` fallback)."""
    style = _resolve_style(document, _STYLE_QUOTE)
    para = document.add_paragraph(quote, style=style)
    # -- belt-and-braces italic: Quote is italic in Word's default
    # -- style template, but custom templates may have stripped it. --
    for run in para.runs:
        run.italic = True
    return para


def _add_technologies(
    document: "Document",
    technologies: Sequence[str],
    bullet: bool,
) -> List["Paragraph"]:
    """Render the technologies list as bullets (or as one comma-joined paragraph)."""
    paragraphs: List["Paragraph"] = []
    items = [str(t) for t in technologies if t]
    if not items:
        return paragraphs

    if bullet:
        style = _resolve_style(document, _STYLE_LIST_BULLET)
        for item in items:
            paragraphs.append(document.add_paragraph(item, style=style))
    else:
        paragraphs.append(document.add_paragraph(", ".join(items)))
    return paragraphs


# -- Public API -----------------------------------------------------------


def case_study(
    document: "Document",
    *,
    title: str,
    customer: str,
    industry: Optional[str] = None,
    size: Optional[str] = None,
    location: Optional[str] = None,
    summary: Union[str, Sequence[str]] = "",
    challenge: Union[str, Sequence[str]] = "",
    solution: Union[str, Sequence[str]] = "",
    implementation: Union[str, Sequence[str]] = "",
    results: Optional[Sequence[Mapping[str, str]]] = None,
    customer_quote: Optional[str] = None,
    technologies: Optional[Sequence[str]] = None,
    next_steps: Union[str, Sequence[str]] = "",
    technologies_as_bullets: bool = True,
    page_break: bool = True,
) -> List[Union["Paragraph", "Table"]]:
    """Append a customer case-study section to `document`.

    The rendered shape is the conventional B2B case-study layout:
    a centred title, a customer-name subtitle, a 3-column
    customer-profile strip (industry / size / location), then the
    canonical narrative sections -- Summary / Challenge / Solution
    / Implementation / Results / Quote / Technologies / Next steps.
    The Results section is rendered as a four-column metric / before
    / after / delta table; the customer quote uses Word's ``Quote``
    style.

    Parameters
    ----------
    document
        The |Document| to append to. The helper is *additive* -- every
        existing block in the document is preserved.
    title
        Case-study headline. Required -- rendered into the centred
        ``Title`` paragraph at the top of the section.
    customer
        Customer / subject organisation name. Required -- rendered as
        a centred ``Heading 2`` paragraph immediately under the title.
    industry, size, location
        The three customer-profile facts rendered in the metadata
        strip. Each is optional; when all three are |None| / empty,
        the strip is skipped entirely.
    summary
        One-paragraph elevator pitch. Rendered under a "Summary"
        ``Heading 1``. Accepts either a single string (split on blank
        lines) or a sequence of strings (one paragraph per item).
        Empty / missing summary suppresses the entire section.
    challenge, solution, implementation, next_steps
        Free-text body sections. Each accepts a single string (split
        on blank lines) or a sequence of strings (one paragraph per
        item). Empty / missing input suppresses the section.
    results
        Sequence of result mappings. Each mapping should carry the
        keys ``metric`` / ``before`` / ``after`` / ``delta``; missing
        keys render as empty cells. Rendered as a four-column
        ``Table Grid`` table under a "Results" ``Heading 1``. Missing
        / empty input suppresses the section.
    customer_quote
        Free-text customer quote. Rendered in Word's built-in
        ``Quote`` style (italic) under a "Customer Quote"
        ``Heading 1``. Missing / empty input suppresses the section.
    technologies
        Sequence of technology / product names. Rendered as a
        bulleted list (when ``technologies_as_bullets`` is true, the
        default) or as a single comma-separated paragraph, under a
        "Technologies" ``Heading 2``. Missing / empty input
        suppresses the section.
    technologies_as_bullets
        When |True| (default), render `technologies` as a
        ``List Bullet`` list -- one paragraph per item. When |False|,
        render as a single comma-separated paragraph. Both forms are
        idiomatic; bullets read better for >5 items, comma-separated
        reads better for 2-3.
    page_break
        When |True| (default), append a trailing page break so the
        next content lands on a fresh page. Pass |False| to suppress
        when chaining multiple case studies on the same spread.

    Returns
    -------
    list
        The list of newly-appended ``Paragraph`` and ``Table``
        objects, in document order, including the trailing
        page-break paragraph when ``page_break`` is true. Callers
        post-process by iterating the returned list (attach
        bookmarks, tweak alignment, pin column widths) without
        having to rediscover the new objects via
        ``document.paragraphs[-N:]`` / ``document.tables[-1]``.

    Raises
    ------
    ValueError
        When `title` or `customer` is empty / whitespace-only, or
        when any entry in `results` is not a mapping.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title must be a non-empty string")
    if not customer or not customer.strip():
        raise ValueError("customer must be a non-empty string")
    if results is not None:
        for index, item in enumerate(results):
            if not isinstance(item, Mapping):  # type: ignore[arg-type]
                raise ValueError(
                    "results[%d] must be a mapping with metric/before/after/delta keys"
                    % index
                )

    appended: List[Union["Paragraph", "Table"]] = []

    # -- Title + customer name --
    title_style = _resolve_style(document, _STYLE_TITLE)
    title_para = document.add_paragraph(title, style=title_style)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    appended.append(title_para)

    customer_heading_style = _resolve_style(document, _STYLE_HEADING_2)
    customer_para = document.add_paragraph(customer, style=customer_heading_style)
    customer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    appended.append(customer_para)

    # -- Customer profile strip (3-column metadata) --
    appended.extend(_add_customer_profile(document, industry, size, location))

    # -- Narrative body sections --
    for heading, body in (
        ("Summary", summary),
        ("Challenge", challenge),
        ("Solution", solution),
        ("Implementation", implementation),
    ):
        if _coerce_body(body):
            appended.extend(_add_section(document, heading, body, level=1))

    # -- Results table --
    if results:
        appended.append(document.add_heading("Results", level=1))
        appended.append(_add_results_table(document, results))

    # -- Customer quote --
    if customer_quote and customer_quote.strip():
        appended.append(document.add_heading("Customer Quote", level=1))
        appended.append(_add_customer_quote(document, customer_quote))

    # -- Technologies --
    if technologies:
        items = [t for t in technologies if t]
        if items:
            appended.append(document.add_heading("Technologies", level=2))
            appended.extend(
                _add_technologies(document, items, technologies_as_bullets)
            )

    # -- Next steps --
    if _coerce_body(next_steps):
        appended.extend(
            _add_section(document, "Next Steps", next_steps, level=1)
        )

    if page_break:
        appended.append(document.add_page_break())

    return appended


__all__ = ["case_study"]
