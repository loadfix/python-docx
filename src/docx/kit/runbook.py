"""Runbook / SOP / playbook helper.

Closes #297.

A *runbook* (also "playbook" / "standard operating procedure") is the
operational document an on-call engineer reaches for when a service
misbehaves at 3am. It captures what to do, who does it, and when to
escalate, in a shape designed for skim-reading under stress. This
module exposes a single :func:`runbook` helper that appends a
structured runbook section to an existing |Document|::

    from docx import Document
    from docx.kit import runbook

    doc = Document()
    runbook.runbook(
        doc,
        title="Database failover runbook",
        purpose="Recover the production primary after a node failure.",
        when_to_use="When the primary's heartbeat times out for >30s.",
        roles=["On-call SRE", "DBA", "Incident commander"],
        prerequisites=[
            "Pager access",
            "VPN connection",
            "Read access to dashboards",
        ],
        procedure=[
            {"step": "Confirm failure",
             "owner": "On-call SRE",
             "detail": "Check Grafana dashboard..."},
            {"step": "Trigger failover",
             "owner": "DBA",
             "detail": "Run `pg_ctl promote replica-1`..."},
            {"step": "Update DNS",
             "owner": "On-call SRE",
             "detail": "Promote replica-1 in Route 53..."},
            {"step": "Verify",
             "owner": "Incident commander",
             "detail": "Run smoke tests..."},
        ],
        escalation=[
            ("After 5min", "Page on-call DBA"),
            ("After 15min", "Page director of engineering"),
            ("After 30min", "Open incident with vendor"),
        ],
        rollback="If failover fails: revert DNS to old primary, page vendor.",
    )
    doc.save("runbook.docx")

The helper composes the conventional runbook skeleton:

- **Title** — ``Heading 1`` page-leader rendering ``title``.
- **Purpose** — short paragraph naming the document's reason for
  existing. Helps a reader decide in two seconds whether they're in
  the right document.
- **When to use** — one-line trigger / activation criterion. The
  matching predicate that says "yes, this runbook applies right now".
- **Roles** — bullet list naming the human roles that participate.
  Renders before the procedure so a reader sees who is involved
  before reading the steps.
- **Prerequisites** — checklist (``[ ]`` prefix) of access / tooling
  / state requirements. Rendered as separate paragraphs so a reader
  can tick them off mentally before kicking off the procedure.
- **Procedure** — numbered list (``List Number`` style with fallback)
  *plus* a three-column table (Step / Owner / Detail) covering the
  same content. The list lets a reader scan the procedure quickly;
  the table makes the role-column inspection fast. Both views are
  emitted because runbooks are read both linearly (bullets) and by
  role (table).
- **Escalation** — two-column table (Trigger / Action) listing the
  escalation ladder. Tuples preserve ordering — the first row is the
  first thing to do, the last row is the last resort.
- **Rollback** — final paragraph describing how to undo the
  procedure on partial failure. The single most-asked question after
  "what do I do?" is "what do I do when it doesn't work?".
- **Page break** — when ``page_break`` is true (the default), the
  helper appends a trailing page break so the next content lands on
  a fresh page.

The return value is a list of newly-appended objects in document
order — ``Paragraph`` objects for headings / bullets / body, ``Table``
objects for the procedure and escalation tables. Callers that want
to post-process specific items (attach bookmarks, tweak alignment)
can iterate the returned list rather than rediscovering the items
via ``document.paragraphs[-N:]`` / ``document.tables[-M:]``.

The helper composes only python-docx's *public* API
(``Document.add_heading``, ``Document.add_paragraph``,
``Document.add_table``, ``Document.add_page_break``,
``_Cell.text``, ``Run.bold``). Per kit conventions there is no XML
reach-down. When the loaded document lacks one of the conventional
list styles (``List Bullet`` / ``List Number``) the helper falls
back to ``Normal`` rather than raising.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Mapping, Optional, Sequence, Tuple, Union

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph


# -- Word built-in styles the helper reaches for. When the loaded
# -- template lacks one, the helper falls back to ``Normal`` rather
# -- than raising — kit helpers are best-effort cosmetic.
_STYLE_HEADING_1 = "Heading 1"
_STYLE_HEADING_2 = "Heading 2"
_STYLE_LIST_BULLET = "List Bullet"
_STYLE_LIST_NUMBER = "List Number"
_STYLE_TABLE_GRID = "Table Grid"
_STYLE_NORMAL = "Normal"


def _has_style(document: "Document", style_name: str) -> bool:
    """Return True when `document` defines a paragraph style named `style_name`."""
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


def _add_bulleted_list(
    document: "Document", items: Sequence[str]
) -> List["Paragraph"]:
    """Append each item as a ``List Bullet``-styled paragraph (or fallback)."""
    style = _resolve_style(document, _STYLE_LIST_BULLET)
    paragraphs: List["Paragraph"] = []
    for item in items:
        if not item:
            continue
        paragraphs.append(document.add_paragraph(str(item), style=style))
    return paragraphs


def _add_checklist(
    document: "Document", items: Sequence[str]
) -> List["Paragraph"]:
    """Append each item as a checkbox-prefixed paragraph (``[ ] item``).

    Word's built-in checkbox content control requires an SDT element
    that the kit deliberately does not reach for; instead the helper
    emits a literal ``[ ]`` prefix that prints and reads as a
    checkbox. The shape mirrors GitHub-flavoured Markdown task lists
    so a reader unfamiliar with the document recognises the pattern.
    """
    paragraphs: List["Paragraph"] = []
    for item in items:
        if not item:
            continue
        para = document.add_paragraph()
        prefix_run = para.add_run("[ ] ")
        prefix_run.font.name = "Courier New"
        para.add_run(str(item))
        paragraphs.append(para)
    return paragraphs


def _add_numbered_list(
    document: "Document", items: Sequence[str]
) -> List["Paragraph"]:
    """Append each item as a ``List Number``-styled paragraph (or fallback).

    When the template lacks ``List Number`` the helper falls back to a
    literal ``"N. item"`` prefix so the procedure remains scannable
    without the style. Returns the appended paragraphs in document
    order.
    """
    paragraphs: List["Paragraph"] = []
    if _has_style(document, _STYLE_LIST_NUMBER):
        style = _STYLE_LIST_NUMBER
        for item in items:
            if not item:
                continue
            paragraphs.append(
                document.add_paragraph(str(item), style=style)
            )
    else:
        for index, item in enumerate(items, start=1):
            if not item:
                continue
            paragraphs.append(
                document.add_paragraph(f"{index}. {item}")
            )
    return paragraphs


def _add_procedure_table(
    document: "Document",
    procedure: Sequence[Mapping[str, str]],
) -> "Table":
    """Append a ``Step / Owner / Detail`` three-column procedure table.

    One header row plus one row per procedure entry. Step ordering is
    implied by row order — the first procedure entry is row 1, etc.
    Owner column carries the human role responsible for the step;
    detail column carries the runbook prose for that step.
    """
    table = document.add_table(rows=1 + len(procedure), cols=3)
    try:
        table.style = _STYLE_TABLE_GRID
    except KeyError:
        # -- Fall back silently when the template lacks Table Grid --
        pass
    header_cells = table.rows[0].cells
    header_cells[0].text = "Step"
    header_cells[1].text = "Owner"
    header_cells[2].text = "Detail"
    # -- bold the header row so it reads as a header even when the
    # -- template lacks the Table Grid style --
    for cell in header_cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True

    for row_idx, entry in enumerate(procedure, start=1):
        row = table.rows[row_idx].cells
        row[0].text = str(entry.get("step", ""))
        row[1].text = str(entry.get("owner", ""))
        row[2].text = str(entry.get("detail", ""))
    return table


def _add_escalation_table(
    document: "Document",
    escalation: Sequence[Tuple[str, str]],
) -> "Table":
    """Append a two-column ``Trigger / Action`` escalation table.

    Tuple ordering is preserved: the first row is the first
    escalation step, the last row is the last resort.
    """
    table = document.add_table(rows=1 + len(escalation), cols=2)
    try:
        table.style = _STYLE_TABLE_GRID
    except KeyError:
        pass
    header_cells = table.rows[0].cells
    header_cells[0].text = "Trigger"
    header_cells[1].text = "Action"
    for cell in header_cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True

    for row_idx, entry in enumerate(escalation, start=1):
        trigger, action = entry
        row = table.rows[row_idx].cells
        row[0].text = str(trigger)
        row[1].text = str(action)
    return table


def _validate_procedure(
    procedure: Sequence[Mapping[str, str]],
) -> None:
    """Raise ``ValueError`` when any procedure entry is malformed.

    Each entry must be a mapping with a non-empty ``step`` key. The
    ``owner`` and ``detail`` keys are optional (rendered as empty
    cells in the table when omitted).
    """
    for index, entry in enumerate(procedure):
        if not isinstance(entry, Mapping):  # type: ignore[arg-type]
            raise ValueError(
                "procedure[%d] must be a mapping with a 'step' key" % index
            )
        step = entry.get("step")
        if step is None or not str(step).strip():
            raise ValueError(
                "procedure[%d] is missing a non-empty 'step'" % index
            )


def _validate_escalation(
    escalation: Sequence[Tuple[str, str]],
) -> None:
    """Raise ``ValueError`` when any escalation entry is malformed."""
    for index, entry in enumerate(escalation):
        if isinstance(entry, str) or isinstance(entry, Mapping):
            raise ValueError(
                "escalation[%d] must be a (trigger, action) tuple" % index
            )
        try:
            trigger, action = entry
        except (TypeError, ValueError) as exc:
            raise ValueError(
                "escalation[%d] must be a (trigger, action) tuple" % index
            ) from exc
        if not str(trigger).strip():
            raise ValueError(
                "escalation[%d] has an empty trigger" % index
            )
        if not str(action).strip():
            raise ValueError(
                "escalation[%d] has an empty action" % index
            )


def runbook(
    document: "Document",
    *,
    title: str,
    purpose: str,
    when_to_use: str,
    roles: Sequence[str],
    prerequisites: Sequence[str],
    procedure: Sequence[Mapping[str, str]],
    escalation: Optional[Sequence[Tuple[str, str]]] = None,
    rollback: Optional[str] = None,
    page_break: bool = True,
) -> List[Union["Paragraph", "Table"]]:
    """Append a structured runbook / SOP section to `document`.

    Renders, in order: a title heading, the purpose paragraph, the
    "when to use" trigger, a roles bullet list, a prerequisites
    checklist, a numbered procedure plus a Step / Owner / Detail
    table, an optional Trigger / Action escalation table, and an
    optional rollback paragraph. Returns the list of newly-appended
    |Paragraph| and |Table| objects in document order.

    Parameters
    ----------
    document
        The |Document| to append to. The helper appends at the end of
        the body — callers who need the runbook ahead of existing
        content should compose into a fresh ``Document()`` first.
    title
        Runbook title (e.g. ``"Database failover runbook"``). Required
        — rendered as a ``Heading 1`` paragraph at the top of the
        section.
    purpose
        One-paragraph statement of what the runbook achieves. Helps
        readers decide in two seconds whether they're in the right
        document.
    when_to_use
        Trigger / activation criterion in one sentence (e.g.
        ``"When the primary's heartbeat times out for >30s."``).
    roles
        Sequence of human-role names that participate in the
        procedure. Rendered as a bullet list under the "Roles"
        sub-heading.
    prerequisites
        Sequence of access / tooling / state requirements. Rendered
        as a ``[ ]`` checklist under the "Prerequisites" sub-heading.
    procedure
        Sequence of step mappings. Each mapping must have a
        non-empty ``step`` key; ``owner`` and ``detail`` keys are
        optional. The procedure is rendered both as a numbered list
        (one paragraph per step) and as a Step / Owner / Detail
        table.
    escalation
        Optional sequence of ``(trigger, action)`` tuples. Rendered
        as a two-column table when supplied; section is omitted when
        |None| or empty.
    rollback
        Optional rollback paragraph describing how to undo the
        procedure on partial failure. Rendered under a "Rollback"
        sub-heading when supplied; section is omitted when |None|.
    page_break
        Append a trailing page break so the next content lands on a
        fresh page. Default ``True``.

    Returns
    -------
    list
        Newly-appended |Paragraph| and |Table| objects, in document
        order, including the trailing page-break paragraph when
        emitted.

    Raises
    ------
    ValueError
        When ``title`` / ``purpose`` / ``when_to_use`` is empty, when
        any required collection is empty, when any ``procedure``
        entry lacks a non-empty ``step``, or when any ``escalation``
        entry is not a valid ``(trigger, action)`` tuple.

    .. versionadded:: 2026.05.29
    """
    if not title or not str(title).strip():
        raise ValueError("title is required")
    if not purpose or not str(purpose).strip():
        raise ValueError("purpose is required")
    if not when_to_use or not str(when_to_use).strip():
        raise ValueError("when_to_use is required")
    if not roles or len(list(roles)) == 0:
        raise ValueError("roles must contain at least one entry")
    if not prerequisites or len(list(prerequisites)) == 0:
        raise ValueError("prerequisites must contain at least one entry")
    if not procedure or len(list(procedure)) == 0:
        raise ValueError("procedure must contain at least one entry")
    _validate_procedure(procedure)
    if escalation:
        _validate_escalation(escalation)

    appended: List[Union["Paragraph", "Table"]] = []

    # -- Title --
    title_style = _resolve_style(document, _STYLE_HEADING_1)
    title_para = document.add_paragraph(title, style=title_style)
    appended.append(title_para)

    heading2 = _resolve_style(document, _STYLE_HEADING_2)

    # -- Purpose --
    appended.append(
        document.add_paragraph("Purpose", style=heading2)
    )
    appended.append(document.add_paragraph(purpose))

    # -- When to use --
    appended.append(
        document.add_paragraph("When to use", style=heading2)
    )
    appended.append(document.add_paragraph(when_to_use))

    # -- Roles --
    appended.append(
        document.add_paragraph("Roles", style=heading2)
    )
    appended.extend(_add_bulleted_list(document, list(roles)))

    # -- Prerequisites (checklist) --
    appended.append(
        document.add_paragraph("Prerequisites", style=heading2)
    )
    appended.extend(_add_checklist(document, list(prerequisites)))

    # -- Procedure (numbered list + table) --
    appended.append(
        document.add_paragraph("Procedure", style=heading2)
    )
    step_lines = [str(entry.get("step", "")) for entry in procedure]
    appended.extend(_add_numbered_list(document, step_lines))
    appended.append(_add_procedure_table(document, procedure))

    # -- Escalation (table) --
    if escalation:
        appended.append(
            document.add_paragraph("Escalation", style=heading2)
        )
        appended.append(_add_escalation_table(document, list(escalation)))

    # -- Rollback (paragraph) --
    if rollback and str(rollback).strip():
        appended.append(
            document.add_paragraph("Rollback", style=heading2)
        )
        appended.append(document.add_paragraph(rollback))

    if page_break:
        appended.append(document.add_page_break())

    return appended


__all__ = [
    "runbook",
]
