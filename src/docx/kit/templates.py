"""Document templates registry — brief / CoE / RFP response / white paper.

Closes #39.

This module exposes four template factories that build entire
strategy-document drafts in one call::

    from docx.kit.templates import brief, coe, rfp_response, white_paper

    doc = brief(
        title="Q1 Strategy Brief",
        author="Strategy Team",
        sections=[
            {"heading": "Background",     "body": "..."},
            {"heading": "Recommendation", "body": "..."},
            {"heading": "Next Steps",     "body": "..."},
        ],
    )

    doc = coe(
        name="Cloud CoE",
        charter="...",
        governance=["Steering committee meets monthly.", "..."],
        services=["Platform engineering.", "Architecture review.", "..."],
    )

    doc = rfp_response(
        rfp_title="Cloud Migration Services RFP",
        company="Acme Corp",
        sections=[
            {"heading": "Executive Summary", "body": "..."},
            {"heading": "Approach",          "body": "..."},
        ],
        pricing_table=[
            {"item": "Discovery",  "quantity": 1, "unit_price": "$25k", "total": "$25k"},
            {"item": "Migration",  "quantity": 1, "unit_price": "$80k", "total": "$80k"},
        ],
    )

    doc = white_paper(
        title="The Future of OOXML",
        author="Ben Hooper",
        abstract="...",
        sections=[
            {"heading": "Introduction", "body": "..."},
            {"heading": "Background",   "body": "..."},
        ],
        references=["Hooper, B. (2026). ..."],
    )

The four factories — :func:`brief`, :func:`coe`, :func:`rfp_response`,
:func:`white_paper` — each return a fresh |Document| pre-populated with
the conventional sections of the matching document type. They are the
"template registry" sibling of :mod:`docx.kit.memos` (which covers
investment memos and business cases): same composition rules, same
fallback-to-``Normal``-style ethos, broader genre coverage.

Common conventions across the four factories:

- **Title page** — a centred title in the ``Title`` style (with
  fallback to ``Normal``), an optional centred ``Subtitle`` for the
  author / company, and an optional centred date paragraph.
- **Body sections** — each factory takes (or composes) a ``sections``
  list. Each entry is a mapping with a required ``heading`` key plus
  an optional ``body`` (string or sequence of strings — the latter
  rendered one paragraph per item).
- **Pricing table** — the RFP factory renders the supplied
  ``pricing_table`` as a four-column ``Table Grid``-styled table
  (Item / Quantity / Unit Price / Total). When the loaded template
  lacks the ``Table Grid`` style the factory falls back silently.
- **References** — the white-paper factory renders the supplied
  ``references`` as a numbered list under a "References" heading.
- **Style fallback to ``Normal``.** Helpers prefer Word's built-in
  ``Title`` / ``Subtitle`` / ``Heading 1`` / ``List Bullet`` /
  ``List Number`` styles; when the loaded template lacks one, fall
  back to ``Normal`` rather than raise.
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


def _add_centred_paragraph(document: DocumentCls, text: str) -> Paragraph:
    """Append a centred plain ``Normal``-styled paragraph."""
    para = document.add_paragraph(text)
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


def _add_section(
    document: DocumentCls,
    heading: str,
    body: Union[str, Sequence[str], None],
    level: int = 1,
) -> List[Paragraph]:
    """Append a heading + body paragraphs section and return its paragraphs."""
    paragraphs: List[Paragraph] = []
    paragraphs.append(document.add_heading(heading, level=level))
    if body is None:
        return paragraphs
    if isinstance(body, str):
        chunks: Sequence[str] = [body]
    else:
        chunks = list(body)
    for chunk in chunks:
        if not chunk:
            continue
        paragraphs.append(document.add_paragraph(chunk))
    return paragraphs


def _add_styled_list(
    document: DocumentCls, items: Sequence[str], style_name: str
) -> List[Paragraph]:
    """Append each item as a styled paragraph (fallback to ``Normal``)."""
    style = style_name
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


def _add_bulleted_list(
    document: DocumentCls, items: Sequence[str]
) -> List[Paragraph]:
    """Append each item as a ``List Bullet``-styled paragraph (or fallback)."""
    return _add_styled_list(document, items, "List Bullet")


def _add_numbered_list(
    document: DocumentCls, items: Sequence[str]
) -> List[Paragraph]:
    """Append each item as a ``List Number``-styled paragraph (or fallback)."""
    return _add_styled_list(document, items, "List Number")


def _validate_sections(
    sections: Optional[Sequence[Mapping[str, Union[str, Sequence[str]]]]],
    *,
    field_name: str = "sections",
) -> None:
    """Raise ``ValueError`` when any section entry is ill-shaped.

    A valid entry is a mapping with a non-empty ``heading``. ``body`` is
    optional. Anything else — a non-mapping element, or a mapping with
    an empty / missing heading — is reported with the offending index.
    """
    if sections is None:
        return
    for index, section in enumerate(sections):
        if not isinstance(section, Mapping):  # type: ignore[arg-type]
            raise ValueError(
                "%s[%d] must be a mapping with a 'heading' key"
                % (field_name, index)
            )
        heading = section.get("heading")
        if not heading or not str(heading).strip():
            raise ValueError(
                "%s[%d] is missing a non-empty 'heading'"
                % (field_name, index)
            )


def _emit_sections(
    document: DocumentCls,
    sections: Optional[Sequence[Mapping[str, Union[str, Sequence[str]]]]],
) -> None:
    """Emit each entry in ``sections`` as a heading + body paragraphs.

    Assumes :func:`_validate_sections` has already verified the shape.
    """
    if not sections:
        return
    for section in sections:
        heading = section.get("heading")
        body = section.get("body", "")
        _add_section(document, str(heading), body)  # type: ignore[arg-type]


# -- Brief -----------------------------------------------------------------


def brief(
    title: str,
    author: Optional[str] = None,
    sections: Optional[Sequence[Mapping[str, Union[str, Sequence[str]]]]] = None,
    date: Optional[str] = None,
) -> DocumentCls:
    """Build a short briefing document and return the |Document|.

    Generates a standard one-to-three-page strategy / decision brief:
    centred title, optional author / date subtitle, and a sequence of
    ``Heading 1`` body sections supplied by the caller. Typical
    examples of the body sections include ``Background``,
    ``Recommendation``, and ``Next Steps``.

    Parameters
    ----------
    title
        Brief title. Required — rendered as the centred document title.
    author
        Optional author / team name rendered as a subtitle under the
        title.
    sections
        Sequence of section mappings. Each entry must have a
        ``heading`` key; the optional ``body`` key can be a string
        (one paragraph) or a sequence of strings (one paragraph per
        item).
    date
        Optional ISO date rendered under the author line.

    Returns
    -------
    Document
        The freshly-built |Document|. Save with :meth:`Document.save`.

    Raises
    ------
    ValueError
        When ``title`` is empty, or when any entry in ``sections``
        lacks a non-empty ``heading``, or is not a mapping.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title is required")

    _validate_sections(sections)

    document = Document()

    # -- Title page --
    _add_title(document, title)
    if author:
        _add_subtitle(document, author)
    if date:
        _add_centred_paragraph(document, date)

    # -- Body sections --
    _emit_sections(document, sections)

    return document


# -- Centre of Excellence --------------------------------------------------


def coe(
    name: str,
    charter: Optional[str] = None,
    governance: Optional[Sequence[str]] = None,
    services: Optional[Sequence[str]] = None,
    sponsor: Optional[str] = None,
    date: Optional[str] = None,
) -> DocumentCls:
    """Build a Centre of Excellence brief and return the |Document|.

    Generates the conventional shape of a CoE charter document: title
    page with the CoE name and an optional sponsor, charter paragraph,
    governance bulleted list, and services bulleted list. The result
    is a starting point — most organisations layer their own
    operating-model details on top.

    Parameters
    ----------
    name
        Centre of Excellence name (e.g. ``"Cloud CoE"`` /
        ``"Data CoE"``). Required — rendered into the title.
    charter
        Free-text charter paragraph stating the CoE's mission and
        scope. Rendered under the "Charter" heading.
    governance
        Sequence of governance bullet items describing how the CoE is
        run (e.g. ``"Steering committee meets monthly."``). Rendered
        as a bulleted list under the "Governance" heading.
    services
        Sequence of service offerings (e.g. ``"Architecture review."``,
        ``"Platform engineering."``). Rendered as a bulleted list
        under the "Services" heading.
    sponsor
        Optional executive sponsor (e.g. ``"CTO"``). Rendered in the
        metadata block when supplied.
    date
        Optional ISO date rendered under the title.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``name`` is empty or whitespace-only.

    .. versionadded:: 2026.05.29
    """
    if not name or not name.strip():
        raise ValueError("name is required")

    document = Document()

    # -- Title page --
    _add_title(document, name)
    _add_subtitle(document, "Centre of Excellence Charter")
    if date:
        _add_centred_paragraph(document, date)

    # -- Metadata block --
    if sponsor:
        _add_metadata_line(document, "Sponsor", sponsor)

    # -- Charter --
    document.add_heading("Charter", level=1)
    if charter:
        document.add_paragraph(charter)
    else:
        document.add_paragraph(
            "[State the CoE's mission, scope, and success criteria.]"
        )

    # -- Governance --
    document.add_heading("Governance", level=1)
    if governance:
        _add_bulleted_list(document, list(governance))
    else:
        document.add_paragraph(
            "[Describe steering, cadence, and decision rights.]"
        )

    # -- Services --
    document.add_heading("Services", level=1)
    if services:
        _add_bulleted_list(document, list(services))
    else:
        document.add_paragraph(
            "[List the services this CoE offers to the wider business.]"
        )

    return document


# -- RFP response ----------------------------------------------------------


def _validate_pricing_rows(
    pricing_table: Optional[Sequence[Mapping[str, Union[str, int, float]]]],
) -> None:
    """Raise ``ValueError`` when any pricing row is ill-shaped.

    A valid entry is a mapping with a non-empty ``item``. Other keys
    (``quantity``, ``unit_price``, ``total``) are optional and default
    to an empty cell when absent.
    """
    if pricing_table is None:
        return
    for index, row in enumerate(pricing_table):
        if not isinstance(row, Mapping):  # type: ignore[arg-type]
            raise ValueError(
                "pricing_table[%d] must be a mapping with an 'item' key"
                % index
            )
        if not row.get("item") or not str(row.get("item")).strip():
            raise ValueError(
                "pricing_table[%d] is missing a non-empty 'item'" % index
            )


def _render_pricing_table(
    document: DocumentCls,
    pricing_table: Sequence[Mapping[str, Union[str, int, float]]],
) -> None:
    """Render ``pricing_table`` as a four-column ``Table Grid`` table.

    Header row: ``Item / Quantity / Unit Price / Total``. Each
    subsequent row is one entry. Missing keys render as empty cells.
    """
    table = document.add_table(rows=1, cols=4)
    try:
        table.style = "Table Grid"
    except KeyError:
        # -- Fall back silently when the template lacks Table Grid --
        pass
    header_cells = table.rows[0].cells
    header_cells[0].text = "Item"
    header_cells[1].text = "Quantity"
    header_cells[2].text = "Unit Price"
    header_cells[3].text = "Total"
    for row in pricing_table:
        cells = table.add_row().cells
        cells[0].text = str(row.get("item", ""))
        quantity = row.get("quantity", "")
        cells[1].text = "" if quantity == "" else str(quantity)
        cells[2].text = str(row.get("unit_price", ""))
        cells[3].text = str(row.get("total", ""))


def rfp_response(
    rfp_title: str,
    company: str,
    sections: Optional[Sequence[Mapping[str, Union[str, Sequence[str]]]]] = None,
    pricing_table: Optional[Sequence[Mapping[str, Union[str, int, float]]]] = None,
    contact: Optional[str] = None,
    date: Optional[str] = None,
) -> DocumentCls:
    """Build an RFP response document and return the |Document|.

    Generates the conventional shape of a vendor's response to an RFP:
    title page with the RFP title and the responding company,
    free-form body sections (typical examples: Executive Summary,
    Approach, Team, References), and a pricing table at the end.

    Parameters
    ----------
    rfp_title
        Title of the RFP being responded to (e.g.
        ``"Cloud Migration Services RFP"``). Required — rendered as
        the centred document title.
    company
        Name of the responding company. Required — rendered as a
        subtitle under the title.
    sections
        Sequence of section mappings. Each entry must have a
        ``heading`` key; the optional ``body`` key can be a string
        (one paragraph) or a sequence of strings (one paragraph per
        item). Typical headings include ``Executive Summary``,
        ``Approach``, ``Team``, ``Timeline``, ``References``.
    pricing_table
        Sequence of pricing-row mappings. Each row must have an
        ``item`` key; ``quantity``, ``unit_price``, and ``total`` are
        optional. Rendered as a four-column ``Table Grid`` table
        under a "Pricing" heading.
    contact
        Optional contact line (e.g. ``"Jane Doe, Account Director,
        jane@acme.com"``) rendered in the metadata block.
    date
        Optional ISO date rendered under the company subtitle.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``rfp_title`` or ``company`` is empty, when any entry in
        ``sections`` lacks a non-empty ``heading``, or when any entry
        in ``pricing_table`` lacks a non-empty ``item``.

    .. versionadded:: 2026.05.29
    """
    if not rfp_title or not rfp_title.strip():
        raise ValueError("rfp_title is required")
    if not company or not company.strip():
        raise ValueError("company is required")

    _validate_sections(sections)
    _validate_pricing_rows(pricing_table)

    document = Document()

    # -- Title page --
    _add_title(document, f"Response to: {rfp_title}")
    _add_subtitle(document, company)
    if date:
        _add_centred_paragraph(document, date)

    # -- Metadata block --
    if contact:
        _add_metadata_line(document, "Contact", contact)

    # -- Body sections --
    _emit_sections(document, sections)

    # -- Pricing --
    document.add_heading("Pricing", level=1)
    if pricing_table:
        _render_pricing_table(document, pricing_table)
    else:
        document.add_paragraph(
            "[Insert pricing line-items: item, quantity, unit price, total.]"
        )

    return document


# -- White paper -----------------------------------------------------------


def white_paper(
    title: str,
    author: Optional[str] = None,
    abstract: Optional[str] = None,
    sections: Optional[Sequence[Mapping[str, Union[str, Sequence[str]]]]] = None,
    references: Optional[Sequence[str]] = None,
    date: Optional[str] = None,
) -> DocumentCls:
    """Build a white-paper document and return the |Document|.

    Generates the conventional shape of a thought-leadership white
    paper: title page with author / date, an "Abstract" heading with
    the supplied abstract paragraph, free-form body sections (typical
    examples: Introduction, Background, Analysis, Conclusion), and a
    numbered "References" list at the end.

    Parameters
    ----------
    title
        White-paper title. Required — rendered as the centred document
        title.
    author
        Optional author name rendered as a subtitle under the title.
    abstract
        Free-text abstract paragraph rendered under an "Abstract"
        heading. Falls back to a placeholder when omitted.
    sections
        Sequence of section mappings. Each entry must have a
        ``heading`` key; the optional ``body`` key can be a string
        (one paragraph) or a sequence of strings (one paragraph per
        item).
    references
        Sequence of reference strings rendered as a numbered list
        under a "References" heading. ``None`` or an empty sequence
        suppresses the section entirely (white papers without
        citations are valid).
    date
        Optional ISO date rendered under the author line.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``title`` is empty, or when any entry in ``sections``
        lacks a non-empty ``heading``.

    .. versionadded:: 2026.05.29
    """
    if not title or not title.strip():
        raise ValueError("title is required")

    _validate_sections(sections)

    document = Document()

    # -- Title page --
    _add_title(document, title)
    if author:
        _add_subtitle(document, author)
    if date:
        _add_centred_paragraph(document, date)

    # -- Abstract --
    document.add_heading("Abstract", level=1)
    if abstract:
        document.add_paragraph(abstract)
    else:
        document.add_paragraph(
            "[Summarise the white paper's thesis, findings, and "
            "recommendations in a single paragraph.]"
        )

    # -- Body sections --
    _emit_sections(document, sections)

    # -- References (only when supplied — a paper without citations is valid) --
    if references:
        document.add_heading("References", level=1)
        _add_numbered_list(document, list(references))

    return document


__all__ = [
    "brief",
    "coe",
    "rfp_response",
    "white_paper",
]
