"""Correction of Error / post-mortem template.

Closes #298.

This module exposes a single helper, :func:`coe`, that appends a
structured *Correction of Error* (also known as *post-mortem* or
*incident review*) document to an existing |Document|. The shape
follows the conventional Amazon / SRE-style COE: an incident
metadata block (title / date / severity / duration / customer
impact), a one-paragraph summary, a chronological timeline table,
the *Five Whys* cascade, a contributing-factors bullet list, an
action-items table (item / owner / due), and a lessons-learned
bullet list. Output is a *starting point* — the writer is expected
to fill in the detail; what the helper guarantees is that every
required section is present and ordered consistently across teams::

    from docx import Document
    from docx.kit import coe

    doc = Document()
    coe.coe(
        doc,
        title="DB-2026-05-29 — primary failover delay",
        incident_date="2026-05-29",
        severity="Sev2",
        duration="47 minutes",
        customer_impact="20% of users saw 5xx errors during the failover window.",
        summary="One-paragraph summary of what happened...",
        timeline=[
            ("14:32 UTC", "Heartbeat alert fires"),
            ("14:34 UTC", "On-call paged"),
            ("14:42 UTC", "Failover initiated"),
            ("14:55 UTC", "Failover failed; rollback initiated"),
            ("15:19 UTC", "Service restored"),
        ],
        five_whys=[
            ("Why did the service fail?", "The primary replica fell behind."),
            ("Why did the primary fail?", "Disk hit 100% util."),
            ("Why did the disk fill?", "A rogue analytics query backfilled to it."),
            ("Why did the query run on primary?", "Routing rule misconfigured 6 weeks ago."),
            ("Why was it not caught?", "We don't alert on routing rule changes."),
        ],
        contributing_factors=[
            "Routing rule misconfigured",
            "Lack of canary on rule changes",
        ],
        action_items=[
            {"item": "Add canary on routing rule changes", "owner": "SRE",
             "due": "2026-06-15"},
            {"item": "Backfill alert on primary disk util", "owner": "DBA",
             "due": "2026-06-08"},
        ],
        lessons_learned=[
            "Always canary routing changes",
            "Alert on every disk reaching > 80% utilisation",
        ],
    )
    doc.save("coe.docx")

The helper appends at the end of the body and returns the list of
newly-appended :class:`~docx.text.paragraph.Paragraph` and
:class:`~docx.table.Table` objects, in document order, so callers can
post-process them (attach bookmarks, tweak alignment, set run-level
formatting) without having to rediscover them by slicing the body.

Conventions:

- **No XML reach-down.** The kit composes only public python-docx API
  (``Document.add_paragraph`` / ``Document.add_heading`` /
  ``Document.add_page_break`` / ``Document.add_table``).
- **Style fallback to ``Normal``.** The helper prefers Word's built-in
  ``Title`` / ``Heading 1`` / ``List Bullet`` / ``Table Grid`` styles;
  when the loaded template lacks one, it falls back to ``Normal`` (or
  silently skips the table style) rather than raising.
- **Per-section page break.** ``page_break=True`` (the default) emits
  a trailing page break so the next content lands on a fresh page.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Mapping, Sequence, Tuple, Union

from docx.enum.text import WD_ALIGN_PARAGRAPH

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph


# -- Word built-in styles the kit reaches for -----------------------------
_STYLE_TITLE = "Title"
_STYLE_HEADING_1 = "Heading 1"
_STYLE_LIST_BULLET = "List Bullet"
_STYLE_TABLE_GRID = "Table Grid"
_STYLE_NORMAL = "Normal"


# -- Internal helpers -----------------------------------------------------


def _has_style(document: Document, style_name: str) -> bool:
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


def _resolve_style(document: Document, preferred: str) -> str:
    """Return `preferred` if it exists on `document`, else ``"Normal"``."""
    return preferred if _has_style(document, preferred) else _STYLE_NORMAL


def _add_centred_title(document: Document, title: str) -> Paragraph:
    """Append a centred document title in the ``Title`` style (or fallback)."""
    style = _resolve_style(document, _STYLE_TITLE)
    para = document.add_paragraph(title, style=style)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_metadata_line(
    document: Document, label: str, value: str
) -> Paragraph:
    """Append a ``"Label: Value"`` paragraph with a bold label run."""
    para = document.add_paragraph()
    label_run = para.add_run(f"{label}: ")
    label_run.bold = True
    para.add_run(value)
    return para


def _apply_table_grid(document: Document, table: Table) -> None:
    """Apply ``Table Grid`` to `table`, falling back silently if absent."""
    if not _has_style(document, _STYLE_TABLE_GRID):
        return
    try:
        table.style = _STYLE_TABLE_GRID
    except KeyError:  # pragma: no cover - belt-and-braces
        pass


def _add_bulleted_list(
    document: Document, items: Sequence[str]
) -> List[Paragraph]:
    """Append each item as a ``List Bullet``-styled paragraph (or fallback)."""
    style = _resolve_style(document, _STYLE_LIST_BULLET)
    paragraphs: List[Paragraph] = []
    for item in items:
        if not item:
            continue
        paragraphs.append(document.add_paragraph(item, style=style))
    return paragraphs


def _validate_pair_sequence(
    pairs: Sequence[Tuple[str, str]], *, field_name: str
) -> None:
    """Raise ``ValueError`` when `pairs` is not a sequence of 2-tuples."""
    for index, entry in enumerate(pairs):
        if isinstance(entry, str) or not hasattr(entry, "__len__"):
            raise ValueError(
                "%s[%d] must be a (str, str) pair, got %r"
                % (field_name, index, entry)
            )
        if len(entry) != 2:
            raise ValueError(
                "%s[%d] must be a (str, str) pair, got %d-tuple"
                % (field_name, index, len(entry))
            )


def _validate_action_items(
    action_items: Sequence[Mapping[str, str]],
) -> None:
    """Raise ``ValueError`` when any action-item is missing ``item``."""
    for index, row in enumerate(action_items):
        if not isinstance(row, Mapping):  # type: ignore[arg-type]
            raise ValueError(
                "action_items[%d] must be a mapping with an 'item' key"
                % index
            )
        if not row.get("item") or not str(row.get("item")).strip():
            raise ValueError(
                "action_items[%d] is missing a non-empty 'item'" % index
            )


def _render_two_column_table(
    document: Document,
    headers: Tuple[str, str],
    rows: Sequence[Tuple[str, str]],
) -> Table:
    """Render a two-column ``Table Grid`` table with `headers` and `rows`."""
    table = document.add_table(rows=1, cols=2)
    _apply_table_grid(document, table)
    header_cells = table.rows[0].cells
    header_cells[0].text = headers[0]
    header_cells[1].text = headers[1]
    for left, right in rows:
        cells = table.add_row().cells
        cells[0].text = str(left)
        cells[1].text = str(right)
    return table


def _render_action_items_table(
    document: Document,
    action_items: Sequence[Mapping[str, str]],
) -> Table:
    """Render a three-column ``Table Grid`` table for action items."""
    table = document.add_table(rows=1, cols=3)
    _apply_table_grid(document, table)
    header_cells = table.rows[0].cells
    header_cells[0].text = "Action Item"
    header_cells[1].text = "Owner"
    header_cells[2].text = "Due"
    for row in action_items:
        cells = table.add_row().cells
        cells[0].text = str(row.get("item", ""))
        cells[1].text = str(row.get("owner", ""))
        cells[2].text = str(row.get("due", ""))
    return table


# -- Public API -----------------------------------------------------------


def coe(
    document: Document,
    *,
    title: str,
    incident_date: str,
    severity: str,
    duration: str,
    customer_impact: str,
    summary: str,
    timeline: Sequence[Tuple[str, str]],
    five_whys: Sequence[Tuple[str, str]],
    contributing_factors: Sequence[str],
    action_items: Sequence[Mapping[str, str]],
    lessons_learned: Sequence[str],
    page_break: bool = True,
) -> List[Union[Paragraph, Table]]:
    """Append a Correction of Error / post-mortem to `document`.

    Renders, in order:

    1. A centred ``Title``-styled title paragraph.
    2. A metadata block (date / severity / duration / customer impact)
       with each row rendered as a bold-labelled paragraph.
    3. A ``Heading 1`` "Summary" with the supplied paragraph.
    4. A ``Heading 1`` "Timeline" with a two-column ``Table Grid``
       table (``Time`` / ``Event``).
    5. A ``Heading 1`` "Five Whys" with a two-column ``Table Grid``
       table (``Question`` / ``Answer``).
    6. A ``Heading 1`` "Contributing Factors" with a
       ``List Bullet``-styled list.
    7. A ``Heading 1`` "Action Items" with a three-column ``Table Grid``
       table (``Action Item`` / ``Owner`` / ``Due``).
    8. A ``Heading 1`` "Lessons Learned" with a
       ``List Bullet``-styled list.
    9. A trailing page break, when ``page_break`` is true (the default).

    Parameters
    ----------
    document
        The |Document| to append to. The helper *appends* — it never
        clears or reorders existing content.
    title
        Incident title (e.g. ``"DB-2026-05-29 — primary failover
        delay"``). Required.
    incident_date
        ISO-style incident date string. Rendered verbatim.
    severity
        Severity classification (e.g. ``"Sev1"`` / ``"Sev2"``).
    duration
        Free-text duration string (e.g. ``"47 minutes"``).
    customer_impact
        One-paragraph description of customer-visible impact.
    summary
        One-paragraph summary of what happened.
    timeline
        Sequence of ``(timestamp, event)`` 2-tuples rendered as a
        two-column table.
    five_whys
        Sequence of ``(question, answer)`` 2-tuples rendered as a
        two-column table.
    contributing_factors
        Sequence of contributing-factor strings rendered as a bulleted
        list.
    action_items
        Sequence of mappings with required ``item`` key plus optional
        ``owner`` / ``due`` keys. Rendered as a three-column table.
    lessons_learned
        Sequence of lesson strings rendered as a bulleted list.
    page_break
        When |True| (the default), append a trailing page break so the
        next content lands on a fresh page.

    Returns
    -------
    list[Paragraph | Table]
        The newly-appended paragraphs and tables in document order,
        including the trailing page-break paragraph when emitted.

    Raises
    ------
    ValueError
        When ``title`` is empty, when any timeline / five-whys entry is
        not a 2-tuple, or when any action-item lacks a non-empty
        ``item`` key.

    .. versionadded:: 2026.05.29
    """
    if not title or not str(title).strip():
        raise ValueError("title is required")

    _validate_pair_sequence(list(timeline), field_name="timeline")
    _validate_pair_sequence(list(five_whys), field_name="five_whys")
    _validate_action_items(list(action_items))

    appended: List[Union[Paragraph, Table]] = []

    # -- Title --
    appended.append(_add_centred_title(document, title))

    # -- Incident metadata block --
    appended.append(_add_metadata_line(document, "Date", incident_date))
    appended.append(_add_metadata_line(document, "Severity", severity))
    appended.append(_add_metadata_line(document, "Duration", duration))
    appended.append(
        _add_metadata_line(document, "Customer Impact", customer_impact)
    )

    # -- Summary --
    appended.append(document.add_heading("Summary", level=1))
    appended.append(document.add_paragraph(summary))

    # -- Timeline --
    appended.append(document.add_heading("Timeline", level=1))
    appended.append(
        _render_two_column_table(
            document, ("Time", "Event"), list(timeline)
        )
    )

    # -- Five Whys --
    appended.append(document.add_heading("Five Whys", level=1))
    appended.append(
        _render_two_column_table(
            document, ("Question", "Answer"), list(five_whys)
        )
    )

    # -- Contributing factors --
    appended.append(document.add_heading("Contributing Factors", level=1))
    appended.extend(_add_bulleted_list(document, list(contributing_factors)))

    # -- Action items --
    appended.append(document.add_heading("Action Items", level=1))
    appended.append(_render_action_items_table(document, list(action_items)))

    # -- Lessons learned --
    appended.append(document.add_heading("Lessons Learned", level=1))
    appended.extend(_add_bulleted_list(document, list(lessons_learned)))

    # -- Trailing page break --
    if page_break:
        appended.append(document.add_page_break())

    return appended


__all__ = [
    "coe",
]
