"""Invoice / quote / statement template family with AUS GST defaults.

Closes #64.

This module exposes three template factories that build complete
billing documents in one call::

    from docx.kit.invoices import invoice, quote, statement

    doc = invoice(
        invoice_number="INV-2026-0042",
        issue_date="2026-03-15",
        due_date="2026-04-14",
        seller={
            "name": "Acme Corp",
            "abn": "12 345 678 901",
            "address": "123 Pitt Street\\nSydney NSW 2000",
            "phone": "+61 2 1234 5678",
            "email": "billing@acme.com",
        },
        buyer={
            "name": "Beta Pty Ltd",
            "abn": "98 765 432 109",
            "address": "456 Collins Street\\nMelbourne VIC 3000",
        },
        items=[
            {"description": "Consulting (March)",
             "quantity": 40, "unit_price": 250, "gst_rate": 0.10},
            {"description": "Travel reimbursement",
             "quantity": 1, "unit_price": 580, "gst_rate": 0.10},
        ],
        payment_terms="Net 30",
        bank_details={
            "bsb": "062-001",
            "account": "1234 5678",
            "name": "Acme Corp Pty Ltd",
        },
    )
    doc.save("INV-2026-0042.docx")

The three factories — :func:`invoice`, :func:`quote`, :func:`statement`
— each return a fresh |Document| pre-populated with the conventional
shape of the matching billing artefact.

**Australian context.** Defaults follow ATO tax-invoice rules:

- A 10% GST rate is applied to any line item that omits ``gst_rate``.
- The header reads "Tax Invoice" when at least one line carries GST,
  or plain "Invoice" when every line is GST-free (the Australian
  GST-free / export case).
- The seller's ABN, when supplied, is rendered prominently in the
  header. The buyer's ABN is rendered when the ATO requires it (any
  invoice over A$1000 must show the buyer's ABN — but the helper
  emits it whenever supplied, regardless of total).
- Currency renders as ``"$"`` prefix without an ISO code, matching
  Australian convention.

**International callers** opt out of GST by setting ``gst_rate=0`` on
each line; the GST column then renders as ``"$0.00"`` and the totals
block omits the GST line. Override the default GST rate globally via
the per-call ``default_gst_rate`` keyword (e.g. ``default_gst_rate=0``
for a US sales-tax-free invoice, or ``default_gst_rate=0.15`` for a
NZ GST invoice).

**Totals.** Each factory auto-computes:

- ``subtotal`` — sum of (quantity * unit_price) across line items.
- ``gst_total`` — sum of (quantity * unit_price * gst_rate) across
  line items.
- ``grand_total`` — ``subtotal + gst_total``.

Numbers in the line-item table and totals block are right-aligned so
the currency column reads cleanly. All money values render with two
decimal places via the ``"%.2f"`` formatter.

**No XML reach-down** — every helper composes only public python-docx
API (``Document.add_paragraph``, ``Document.add_heading``,
``Document.add_table``, ``_Cell.text``, ``Paragraph.alignment``,
``Run.bold``).

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import (
    TYPE_CHECKING,
    Any,
    List,
    Mapping,
    Optional,
    Sequence,
    Tuple,
    Union,
)

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph


# -- Defaults -------------------------------------------------------------

#: Australian GST rate. Applied to any line item that omits ``gst_rate``.
DEFAULT_GST_RATE: float = 0.10

#: Currency prefix on rendered money values. Two decimals always.
_CURRENCY_PREFIX = "$"


# -- Style helpers --------------------------------------------------------


def _add_styled_paragraph(
    document: DocumentCls,
    text: str,
    style: str,
    fallback: str = "Normal",
) -> Paragraph:
    """Append a paragraph in ``style`` (or ``fallback`` when missing)."""
    try:
        document.styles[style]
    except KeyError:
        style = fallback
    return document.add_paragraph(text, style=style)


def _add_title(document: DocumentCls, title: str) -> Paragraph:
    """Append a centred document title in ``Title`` (or fallback)."""
    para = _add_styled_paragraph(document, title, "Title")
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_subtitle(document: DocumentCls, text: str) -> Paragraph:
    """Append a centred subtitle in ``Subtitle`` (or fallback)."""
    para = _add_styled_paragraph(document, text, "Subtitle")
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_heading(
    document: DocumentCls, text: str, level: int = 1
) -> Paragraph:
    """Append a heading; falls back to a bold paragraph when style absent."""
    try:
        return document.add_heading(text, level=level)
    except KeyError:
        para = document.add_paragraph()
        run = para.add_run(text)
        run.bold = True
        return para


def _add_label_value(
    document: DocumentCls,
    label: str,
    value: str,
    alignment: Optional[int] = None,
) -> Paragraph:
    """Append a ``"Label: value"`` paragraph with a bold label run."""
    para = document.add_paragraph()
    if alignment is not None:
        para.alignment = alignment
    label_run = para.add_run(f"{label}: ")
    label_run.bold = True
    para.add_run(value)
    return para


def _set_cell_text(
    cell: "_Cell",
    text: str,
    bold: bool = False,
    alignment: Optional[int] = None,
) -> None:
    """Write ``text`` into ``cell`` with optional bold + paragraph alignment.

    The first paragraph's existing runs are dropped (``cell.text =
    text`` replaces the cell's content with a single paragraph). The
    helper then post-processes that paragraph to set alignment and
    bolds the run when asked.
    """
    cell.text = text
    para = cell.paragraphs[0]
    if alignment is not None:
        para.alignment = alignment
    if bold:
        for run in para.runs:
            run.bold = True


def _apply_table_grid(table: "Table") -> None:
    """Apply ``Table Grid`` style; silently fall back when missing."""
    try:
        table.style = "Table Grid"
    except KeyError:
        pass


# -- Money formatting -----------------------------------------------------


def _format_money(value: float) -> str:
    """Render ``value`` as ``"$1,234.56"`` (two decimals, comma thousands)."""
    return f"{_CURRENCY_PREFIX}{value:,.2f}"


# -- Party rendering ------------------------------------------------------


def _render_party(
    document: DocumentCls, label: str, party: Optional[Mapping[str, Any]]
) -> List[Paragraph]:
    """Append a ``label`` heading followed by the party's contact lines.

    ``party`` is a free-shape mapping; recognised keys are rendered in
    a stable order with bold-labelled lines (``"Name"``, ``"ABN"``,
    ``"Address"``, ``"Phone"``, ``"Email"``). Unknown keys are
    appended verbatim after the recognised ones, in caller insertion
    order. Empty / |None| values are skipped.
    """
    paragraphs: List[Paragraph] = []
    paragraphs.append(_add_heading(document, label, level=2))
    if not party:
        paragraphs.append(
            document.add_paragraph(f"[{label} details]")
        )
        return paragraphs

    # -- Render the name first (no label) so it reads as a header line.
    name = party.get("name")
    if name:
        para = document.add_paragraph()
        run = para.add_run(str(name))
        run.bold = True
        paragraphs.append(para)

    # -- Recognised contact keys, in stable order. --
    for key, label_text in (
        ("abn", "ABN"),
        ("acn", "ACN"),
        ("address", "Address"),
        ("phone", "Phone"),
        ("email", "Email"),
        ("website", "Website"),
    ):
        value = party.get(key)
        if value is None or str(value).strip() == "":
            continue
        # -- Address may be multi-line; render each line under the bold label.
        if key == "address":
            lines = [
                ln for ln in str(value).splitlines() if ln.strip()
            ]
            if not lines:
                continue
            para = document.add_paragraph()
            label_run = para.add_run(f"{label_text}: ")
            label_run.bold = True
            para.add_run(lines[0])
            paragraphs.append(para)
            for extra in lines[1:]:
                paragraphs.append(document.add_paragraph(extra))
        else:
            paragraphs.append(
                _add_label_value(document, label_text, str(value))
            )

    # -- Trailing unknown keys, verbatim. --
    recognised = {"name", "abn", "acn", "address", "phone", "email", "website"}
    for key, value in party.items():
        if key in recognised:
            continue
        if value is None or str(value).strip() == "":
            continue
        paragraphs.append(
            _add_label_value(
                document, str(key).replace("_", " ").title(), str(value)
            )
        )
    return paragraphs


# -- Line-item normalisation + totals -------------------------------------


def _normalise_item(
    item: Mapping[str, Any],
    index: int,
    default_gst_rate: float,
) -> Tuple[str, float, float, float]:
    """Validate a single ``items`` entry and return ``(desc, qty, price, rate)``.

    Quantity / unit_price / gst_rate are all coerced to ``float`` so the
    arithmetic downstream is uniform. The helper raises
    :class:`ValueError` with the offending row index on missing
    ``description``, missing or non-numeric ``unit_price`` /
    ``quantity``, or out-of-range ``gst_rate``.
    """
    if not isinstance(item, Mapping):  # type: ignore[arg-type]
        raise ValueError(
            "items[%d] must be a mapping with at least 'description' "
            "and 'unit_price'" % index
        )

    description = item.get("description")
    if description is None or str(description).strip() == "":
        raise ValueError(
            "items[%d] is missing a non-empty 'description'" % index
        )

    if "unit_price" not in item or item.get("unit_price") is None:
        raise ValueError("items[%d] is missing 'unit_price'" % index)
    try:
        unit_price = float(item["unit_price"])
    except (TypeError, ValueError):
        raise ValueError(
            "items[%d]: 'unit_price' must be a number; got %r"
            % (index, item.get("unit_price"))
        ) from None

    quantity_raw = item.get("quantity", 1)
    try:
        quantity = float(quantity_raw)
    except (TypeError, ValueError):
        raise ValueError(
            "items[%d]: 'quantity' must be a number; got %r"
            % (index, quantity_raw)
        ) from None

    if "gst_rate" in item and item["gst_rate"] is not None:
        try:
            gst_rate = float(item["gst_rate"])
        except (TypeError, ValueError):
            raise ValueError(
                "items[%d]: 'gst_rate' must be a number; got %r"
                % (index, item.get("gst_rate"))
            ) from None
        if gst_rate < 0 or gst_rate > 1:
            raise ValueError(
                "items[%d]: 'gst_rate' must be between 0 and 1; got %r"
                % (index, gst_rate)
            )
    else:
        gst_rate = default_gst_rate

    return str(description), quantity, unit_price, gst_rate


def _compute_totals(
    items: Sequence[Tuple[str, float, float, float]],
) -> Tuple[float, float, float]:
    """Return ``(subtotal, gst_total, grand_total)`` for normalised items."""
    subtotal = 0.0
    gst_total = 0.0
    for _desc, qty, price, rate in items:
        line = qty * price
        subtotal += line
        gst_total += line * rate
    grand = subtotal + gst_total
    return subtotal, gst_total, grand


def _format_quantity(quantity: float) -> str:
    """Render quantity without trailing zeros (``40.0`` -> ``"40"``)."""
    if quantity == int(quantity):
        return str(int(quantity))
    return f"{quantity:g}"


# -- Line-item table ------------------------------------------------------


def _add_line_items_table(
    document: DocumentCls,
    items: Sequence[Tuple[str, float, float, float]],
    show_gst_column: bool,
) -> "Table":
    """Render the line-item table with right-aligned numeric cells.

    Columns: Description / Quantity / Unit Price / [GST] / Line Total.
    The GST column is only emitted when at least one item has a
    non-zero GST rate (``show_gst_column=True``); otherwise the table
    drops to four columns to keep the layout clean for GST-free
    invoices.
    """
    cols = 5 if show_gst_column else 4
    table = document.add_table(rows=1, cols=cols)
    _apply_table_grid(table)

    # -- Header row --
    header_cells = table.rows[0].cells
    headers = ["Description", "Quantity", "Unit Price"]
    if show_gst_column:
        headers.append("GST")
    headers.append("Line Total")
    for index, label in enumerate(headers):
        # -- Numbers right-align even in the header so the whole column
        # -- reads consistently. Description stays left-aligned.
        alignment = (
            WD_ALIGN_PARAGRAPH.LEFT
            if index == 0
            else WD_ALIGN_PARAGRAPH.RIGHT
        )
        _set_cell_text(
            header_cells[index], label, bold=True, alignment=alignment
        )

    # -- Data rows --
    for desc, qty, price, rate in items:
        row = table.add_row().cells
        line_total = qty * price
        gst_amount = line_total * rate
        _set_cell_text(row[0], desc, alignment=WD_ALIGN_PARAGRAPH.LEFT)
        _set_cell_text(
            row[1], _format_quantity(qty), alignment=WD_ALIGN_PARAGRAPH.RIGHT
        )
        _set_cell_text(
            row[2],
            _format_money(price),
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        )
        if show_gst_column:
            _set_cell_text(
                row[3],
                _format_money(gst_amount),
                alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            )
            total_idx = 4
        else:
            total_idx = 3
        _set_cell_text(
            row[total_idx],
            _format_money(line_total),
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        )

    return table


def _add_totals_block(
    document: DocumentCls,
    subtotal: float,
    gst_total: float,
    grand_total: float,
    show_gst: bool,
    grand_total_label: str = "Total",
) -> "Table":
    """Render the right-aligned totals box (subtotal / GST / total)."""
    rows: List[Tuple[str, str, bool]] = [
        ("Subtotal", _format_money(subtotal), False),
    ]
    if show_gst:
        rows.append(("GST", _format_money(gst_total), False))
    rows.append((grand_total_label, _format_money(grand_total), True))

    table = document.add_table(rows=len(rows), cols=2)
    _apply_table_grid(table)
    for row_index, (label, value, bold) in enumerate(rows):
        cells = table.rows[row_index].cells
        _set_cell_text(
            cells[0], label, bold=bold, alignment=WD_ALIGN_PARAGRAPH.RIGHT
        )
        _set_cell_text(
            cells[1], value, bold=bold, alignment=WD_ALIGN_PARAGRAPH.RIGHT
        )
    return table


# -- Bank details + payment terms -----------------------------------------


def _render_bank_details(
    document: DocumentCls, bank_details: Optional[Mapping[str, Any]]
) -> List[Paragraph]:
    """Render the bank details block (BSB / account / name)."""
    if not bank_details:
        return []
    paragraphs: List[Paragraph] = [_add_heading(document, "Payment Details", level=2)]
    # -- Stable AUS-friendly key order. Unknown keys appear verbatim. --
    for key, label in (
        ("name", "Account Name"),
        ("bsb", "BSB"),
        ("account", "Account Number"),
        ("bank", "Bank"),
        ("swift", "SWIFT/BIC"),
        ("iban", "IBAN"),
        ("reference", "Reference"),
    ):
        value = bank_details.get(key)
        if value is None or str(value).strip() == "":
            continue
        paragraphs.append(_add_label_value(document, label, str(value)))
    recognised = {"name", "bsb", "account", "bank", "swift", "iban", "reference"}
    for key, value in bank_details.items():
        if key in recognised:
            continue
        if value is None or str(value).strip() == "":
            continue
        paragraphs.append(
            _add_label_value(
                document, str(key).replace("_", " ").title(), str(value)
            )
        )
    return paragraphs


# -- Validation helpers ---------------------------------------------------


def _require_non_empty(value: Optional[str], name: str) -> str:
    """Return ``value.strip()``; raise ``ValueError`` when empty."""
    if value is None or not str(value).strip():
        raise ValueError(f"{name} is required")
    return str(value)


def _validate_default_gst_rate(rate: float) -> float:
    """Coerce + range-check ``rate`` (0..1)."""
    try:
        rate_f = float(rate)
    except (TypeError, ValueError):
        raise ValueError(
            "default_gst_rate must be a number between 0 and 1; got %r" % rate
        ) from None
    if rate_f < 0 or rate_f > 1:
        raise ValueError(
            "default_gst_rate must be between 0 and 1; got %r" % rate_f
        )
    return rate_f


# -- Public factory: invoice ---------------------------------------------


def invoice(
    invoice_number: str,
    issue_date: str,
    due_date: Optional[str] = None,
    seller: Optional[Mapping[str, Any]] = None,
    buyer: Optional[Mapping[str, Any]] = None,
    items: Optional[Sequence[Mapping[str, Any]]] = None,
    payment_terms: Optional[str] = None,
    bank_details: Optional[Mapping[str, Any]] = None,
    notes: Optional[str] = None,
    default_gst_rate: float = DEFAULT_GST_RATE,
) -> "DocumentCls":
    """Build a tax invoice and return the |Document|.

    The output complies with the ATO's tax-invoice rules when ``seller``
    carries an ``abn`` field and at least one line item carries GST.
    The header reads "Tax Invoice" in that case; when every line is
    GST-free it falls back to plain "Invoice".

    Parameters
    ----------
    invoice_number
        Invoice identifier (e.g. ``"INV-2026-0042"``). Required —
        rendered into the header.
    issue_date
        Issue date — free-text (typically ISO ``"YYYY-MM-DD"``).
        Required.
    due_date
        Optional due date. Rendered in the header metadata block when
        supplied.
    seller, buyer
        Free-shape party mappings. Recognised keys: ``name``, ``abn``,
        ``acn``, ``address`` (multi-line via ``"\\n"``), ``phone``,
        ``email``, ``website``. Unrecognised keys are appended
        verbatim.
    items
        Sequence of line-item mappings. Each entry needs at minimum a
        ``description`` and ``unit_price``. ``quantity`` defaults to
        ``1``; ``gst_rate`` defaults to ``default_gst_rate``.
    payment_terms
        Free-text payment terms (e.g. ``"Net 30"``). Rendered after
        the totals block when supplied.
    bank_details
        Bank-details mapping (``bsb``, ``account``, ``name``, ``bank``,
        ``swift``, ``iban``, ``reference``). Rendered as a labelled
        list under "Payment Details" when supplied.
    notes
        Optional trailing notes paragraph (rendered in italic-friendly
        ``Quote`` style when available, else ``Normal``).
    default_gst_rate
        GST rate applied to any line item that omits ``gst_rate``.
        Defaults to :data:`DEFAULT_GST_RATE` (10% AUS GST). Pass
        ``0`` for an international / GST-free invoice.

    Returns
    -------
    Document
        The freshly-built |Document|. Save with :meth:`Document.save`.

    Raises
    ------
    ValueError
        When ``invoice_number`` / ``issue_date`` are empty, when any
        line item is malformed, or when ``default_gst_rate`` is outside
        ``[0, 1]``.

    .. versionadded:: 2026.05.29
    """
    invoice_number = _require_non_empty(invoice_number, "invoice_number")
    issue_date = _require_non_empty(issue_date, "issue_date")
    default_rate = _validate_default_gst_rate(default_gst_rate)

    normalised: List[Tuple[str, float, float, float]] = []
    if items:
        for index, item in enumerate(items):
            normalised.append(_normalise_item(item, index, default_rate))

    has_gst = any(rate > 0 for _d, _q, _p, rate in normalised)
    show_gst_column = has_gst

    document = Document()

    # -- Header --
    title = "Tax Invoice" if has_gst else "Invoice"
    _add_title(document, title)
    _add_subtitle(document, invoice_number)

    # -- Header metadata (centred) --
    _add_label_value(
        document, "Issue Date", issue_date, alignment=WD_ALIGN_PARAGRAPH.CENTER
    )
    if due_date:
        _add_label_value(
            document, "Due Date", due_date, alignment=WD_ALIGN_PARAGRAPH.CENTER
        )

    # -- Parties --
    _render_party(document, "From", seller)
    _render_party(document, "Bill To", buyer)

    # -- Line items --
    _add_heading(document, "Items", level=2)
    if normalised:
        _add_line_items_table(document, normalised, show_gst_column)
    else:
        document.add_paragraph(
            "[No line items supplied — add rows describing the goods or "
            "services billed.]"
        )

    # -- Totals --
    subtotal, gst_total, grand_total = _compute_totals(normalised)
    _add_totals_block(
        document, subtotal, gst_total, grand_total, show_gst=has_gst
    )

    # -- Payment terms --
    if payment_terms:
        _add_heading(document, "Payment Terms", level=2)
        document.add_paragraph(payment_terms)

    # -- Bank details --
    _render_bank_details(document, bank_details)

    # -- Notes --
    if notes:
        _add_heading(document, "Notes", level=2)
        _add_styled_paragraph(document, notes, "Quote")

    return document


# -- Public factory: quote -----------------------------------------------


def quote(
    quote_number: str,
    issue_date: str,
    valid_until: Optional[str] = None,
    seller: Optional[Mapping[str, Any]] = None,
    buyer: Optional[Mapping[str, Any]] = None,
    items: Optional[Sequence[Mapping[str, Any]]] = None,
    notes: Optional[str] = None,
    default_gst_rate: float = DEFAULT_GST_RATE,
) -> "DocumentCls":
    """Build a quote / quotation and return the |Document|.

    A quote shares the structural skeleton of an invoice (parties,
    line items, totals) but advertises a *price offer* rather than an
    amount due. The header reads "Quote", and the totals block
    labels the grand total "Estimated Total" so the reader doesn't
    treat it as a binding bill. The ``valid_until`` date appears in
    the header metadata when supplied.

    Parameters
    ----------
    quote_number
        Quote identifier (e.g. ``"QU-2026-0099"``). Required.
    issue_date
        Issue date — free-text (typically ISO ``"YYYY-MM-DD"``).
        Required.
    valid_until
        Optional expiry date for the quote. Rendered in the header
        metadata block when supplied.
    seller, buyer
        See :func:`invoice` — same shape and recognised keys.
    items
        Sequence of line-item mappings. See :func:`invoice` for the
        recognised shape.
    notes
        Optional trailing notes paragraph.
    default_gst_rate
        GST rate applied to any line item that omits ``gst_rate``.
        Defaults to 10% AUS GST.

    Returns
    -------
    Document
        The freshly-built quote |Document|.

    Raises
    ------
    ValueError
        When ``quote_number`` / ``issue_date`` are empty, when any line
        item is malformed, or when ``default_gst_rate`` is outside
        ``[0, 1]``.

    .. versionadded:: 2026.05.29
    """
    quote_number = _require_non_empty(quote_number, "quote_number")
    issue_date = _require_non_empty(issue_date, "issue_date")
    default_rate = _validate_default_gst_rate(default_gst_rate)

    normalised: List[Tuple[str, float, float, float]] = []
    if items:
        for index, item in enumerate(items):
            normalised.append(_normalise_item(item, index, default_rate))

    has_gst = any(rate > 0 for _d, _q, _p, rate in normalised)

    document = Document()

    _add_title(document, "Quote")
    _add_subtitle(document, quote_number)
    _add_label_value(
        document, "Issue Date", issue_date, alignment=WD_ALIGN_PARAGRAPH.CENTER
    )
    if valid_until:
        _add_label_value(
            document,
            "Valid Until",
            valid_until,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )

    _render_party(document, "From", seller)
    _render_party(document, "Quote To", buyer)

    _add_heading(document, "Items", level=2)
    if normalised:
        _add_line_items_table(document, normalised, has_gst)
    else:
        document.add_paragraph(
            "[No line items supplied — add rows describing the goods or "
            "services quoted.]"
        )

    subtotal, gst_total, grand_total = _compute_totals(normalised)
    _add_totals_block(
        document,
        subtotal,
        gst_total,
        grand_total,
        show_gst=has_gst,
        grand_total_label="Estimated Total",
    )

    if notes:
        _add_heading(document, "Notes", level=2)
        _add_styled_paragraph(document, notes, "Quote")

    return document


# -- Public factory: statement -------------------------------------------


def _normalise_invoice_record(
    record: Mapping[str, Any], index: int
) -> Tuple[str, str, float, float, str]:
    """Validate a single ``statement.invoices[]`` record.

    Returns ``(invoice_number, date, amount, balance, status)``. The
    ``balance`` mirrors ``amount`` when the caller doesn't supply a
    separate balance; ``status`` defaults to an empty string.
    """
    if not isinstance(record, Mapping):  # type: ignore[arg-type]
        raise ValueError(
            "invoices[%d] must be a mapping with at least 'invoice_number' "
            "and 'amount'" % index
        )

    number = record.get("invoice_number") or record.get("number")
    if number is None or str(number).strip() == "":
        raise ValueError(
            "invoices[%d] is missing 'invoice_number'" % index
        )

    date = record.get("date") or record.get("issue_date") or ""
    if "amount" not in record or record.get("amount") is None:
        raise ValueError("invoices[%d] is missing 'amount'" % index)
    try:
        amount = float(record["amount"])
    except (TypeError, ValueError):
        raise ValueError(
            "invoices[%d]: 'amount' must be a number; got %r"
            % (index, record.get("amount"))
        ) from None

    balance_raw = record.get("balance", amount)
    try:
        balance = float(balance_raw)
    except (TypeError, ValueError):
        raise ValueError(
            "invoices[%d]: 'balance' must be a number; got %r"
            % (index, balance_raw)
        ) from None

    status = record.get("status") or ""
    return str(number), str(date), amount, balance, str(status)


def statement(
    period_start: str,
    period_end: str,
    seller: Optional[Mapping[str, Any]] = None,
    buyer: Optional[Mapping[str, Any]] = None,
    invoices: Optional[Sequence[Mapping[str, Any]]] = None,
    statement_number: Optional[str] = None,
    notes: Optional[str] = None,
) -> "DocumentCls":
    """Build a customer statement and return the |Document|.

    A statement summarises a buyer's outstanding invoices over a
    billing period. Each invoice row carries an invoice number, date,
    amount, balance owing, and optional status. The grand total at the
    bottom is the sum of every row's ``balance`` — what the buyer
    still owes overall.

    Parameters
    ----------
    period_start, period_end
        Statement period (free-text, typically ISO dates).
        Both required.
    seller, buyer
        Free-shape party mappings. See :func:`invoice` for the
        recognised keys.
    invoices
        Sequence of invoice records. Each must carry an
        ``invoice_number`` (or ``number``) and ``amount``. Optional
        keys: ``date`` / ``issue_date``, ``balance`` (defaults to
        ``amount``), ``status`` (free-text — e.g. ``"Paid"``,
        ``"Overdue"``).
    statement_number
        Optional statement identifier rendered as the subtitle.
    notes
        Optional trailing notes paragraph.

    Returns
    -------
    Document
        The freshly-built statement |Document|.

    Raises
    ------
    ValueError
        When ``period_start`` / ``period_end`` are empty, or when any
        invoice record is malformed.

    .. versionadded:: 2026.05.29
    """
    period_start = _require_non_empty(period_start, "period_start")
    period_end = _require_non_empty(period_end, "period_end")

    normalised: List[Tuple[str, str, float, float, str]] = []
    if invoices:
        for index, record in enumerate(invoices):
            normalised.append(_normalise_invoice_record(record, index))

    document = Document()

    _add_title(document, "Statement")
    if statement_number:
        _add_subtitle(document, statement_number)
    _add_label_value(
        document,
        "Period",
        f"{period_start} – {period_end}",
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    _render_party(document, "From", seller)
    _render_party(document, "Statement To", buyer)

    _add_heading(document, "Invoices", level=2)
    if not normalised:
        document.add_paragraph(
            "[No invoices recorded for this period.]"
        )
    else:
        # -- Columns: Invoice / Date / Amount / Balance / Status. The
        # -- status column is dropped when *no* row carries a status,
        # -- so a payments-pending statement reads cleanly.
        show_status = any(status for _n, _d, _a, _b, status in normalised)
        cols = 5 if show_status else 4
        table = document.add_table(rows=1, cols=cols)
        _apply_table_grid(table)
        header = ["Invoice", "Date", "Amount", "Balance"]
        if show_status:
            header.append("Status")
        for index, label in enumerate(header):
            alignment = (
                WD_ALIGN_PARAGRAPH.LEFT
                if index in (0, 1) or (show_status and index == 4)
                else WD_ALIGN_PARAGRAPH.RIGHT
            )
            _set_cell_text(
                table.rows[0].cells[index],
                label,
                bold=True,
                alignment=alignment,
            )
        for number, date, amount, balance, status in normalised:
            row = table.add_row().cells
            _set_cell_text(row[0], number, alignment=WD_ALIGN_PARAGRAPH.LEFT)
            _set_cell_text(row[1], date, alignment=WD_ALIGN_PARAGRAPH.LEFT)
            _set_cell_text(
                row[2],
                _format_money(amount),
                alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            )
            _set_cell_text(
                row[3],
                _format_money(balance),
                alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            )
            if show_status:
                _set_cell_text(
                    row[4], status, alignment=WD_ALIGN_PARAGRAPH.LEFT
                )

    # -- Total balance owing across every row in the period. --
    total_balance = sum(balance for _n, _d, _a, balance, _s in normalised)
    totals_table = document.add_table(rows=1, cols=2)
    _apply_table_grid(totals_table)
    cells = totals_table.rows[0].cells
    _set_cell_text(
        cells[0],
        "Total Balance Owing",
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.RIGHT,
    )
    _set_cell_text(
        cells[1],
        _format_money(total_balance),
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.RIGHT,
    )

    if notes:
        _add_heading(document, "Notes", level=2)
        _add_styled_paragraph(document, notes, "Quote")

    return document


__all__ = [
    "invoice",
    "quote",
    "statement",
    "DEFAULT_GST_RATE",
]
