"""Medical clinical-note template family — opinionated authoring helpers.

Closes #83.

This module exposes three template factories that build entire
boilerplate-laden clinical-documentation drafts in one call::

    from docx.kit.medical import (
        soap_note,
        discharge_summary,
        referral_letter,
    )

    doc = soap_note(
        patient={"name": "John Doe", "dob": "1980-05-15", "mrn": "1234567"},
        encounter_date="2026-03-15",
        provider={
            "name": "Dr. Alice Smith",
            "role": "GP",
            "practice": "Sydney Medical Centre",
        },
        subjective="Patient reports persistent cough for 3 weeks...",
        objective={
            "vitals": {"bp": "120/80", "hr": 72, "temp": 36.8,
                       "rr": 18, "spo2": 98},
            "examination": "Chest clear, no rales...",
            "labs": ["FBC: WNL", "CRP: 12 mg/L"],
        },
        assessment=["Acute bronchitis (J20.9)", "Hypertension stable"],
        plan=[
            "Amoxicillin 500mg TDS for 7 days",
            "Review in 1 week",
            "Continue current antihypertensive regimen",
        ],
    )
    doc.save("doe-john-2026-03-15.docx")

The three factories — :func:`soap_note`, :func:`discharge_summary`,
:func:`referral_letter` — each return a fresh |Document| pre-populated
with the conventional sections of the matching clinical-document type.
The shapes follow common Australian primary-care drafting conventions
(MRN/DOB/Medicare-aware patient block, problem list with ICD-10-AM
codes, Australian medication shorthand such as ``TDS``/``BD``/``PRN``)
so a small AUS-based clinic can use the output as the *first draft* of
a real clinical note without re-typing boilerplate every time.

.. warning::

    **Template only — not a medical record and not legal advice.**
    Every document returned by this module is a structural skeleton
    intended to capture an actual clinical encounter. The text emitted
    by these factories is *not itself* a medical record: a real medical
    record requires an authentic clinician-patient encounter, accurate
    contemporaneous documentation, and storage in a system meeting the
    record-keeping obligations of the relevant jurisdiction (in
    Australia: the Medical Board's *Good medical practice: a code of
    conduct*; the *Privacy Act 1988* (Cth) and the Australian Privacy
    Principles; state-based health-records legislation such as
    Victoria's *Health Records Act 2001* and NSW's *Health Records and
    Information Privacy Act 2002*; My Health Record interoperability
    standards). The output of this module is also not legal advice —
    have a qualified practitioner author and sign every clinical note
    before relying on it. The authors of python-docx accept no
    responsibility for clinical, legal, or privacy-related losses
    arising from reliance on this boilerplate.

Common conventions across the three factories:

- **Patient identification block** — name, date of birth, medical
  record number (``MRN``), optional Medicare / address / phone, all on
  the first page so the document is self-identifying when printed.
- **Provider block** — clinician name, role (``GP`` / ``Specialist`` /
  ``Registrar`` / ``RN`` / …), practice name, and optional provider
  number / contact.
- **Section structure** — each factory uses the conventional headings
  for its document type. SOAP notes use Subjective / Objective /
  Assessment / Plan; discharge summaries use Admission /
  Investigations / Procedures / Discharge medications / Follow-up;
  referral letters use a salutation / clinical question / history /
  examination / requested action / closing structure.
- **Vitals as a structured table** — when ``vitals`` is supplied to
  :func:`soap_note` or :func:`discharge_summary`, the helper renders a
  two-column "Parameter / Value" table with one row per recognised
  vital (``bp``, ``hr``, ``temp``, ``rr``, ``spo2``, ``weight``,
  ``height``, ``bmi``, ``pain``) — the structured form makes the
  numbers scannable at a glance and trivially extractable for an EHR
  import. Unrecognised keys are appended verbatim after the recognised
  ones so unusual measurements (e.g. ``cvp``, ``fio2``) round-trip.
- **AUS defaults** — date format is ISO ``YYYY-MM-DD``; medication
  shorthand follows AUS prescribing convention (``TDS`` / ``QID`` /
  ``BD`` / ``PRN`` / ``mane`` / ``nocte``); where a code is supplied
  alongside a diagnosis the helper renders ``"Diagnosis text (CODE)"``
  matching the ICD-10-AM presentation.
- **No XML reach-down** — the kit composes only public python-docx API
  (``Document.add_paragraph``, ``Document.add_heading``,
  ``Document.add_table``, ``Document.add_page_break``).

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
    Union,
)

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls
    from docx.table import Table
    from docx.text.paragraph import Paragraph


# -- Constants used across the three factories. AUS-biased. --
_DEFAULT_DATE_LABEL = "Encounter Date"

# -- Disclaimer rendered verbatim into every generated document so a
# -- downstream user opening the file in Word sees the same caveat the
# -- module docstring carries. --
_DISCLAIMER = (
    "TEMPLATE ONLY: This document is a structural template generated by "
    "python-docx. It is not itself a medical record, not medical advice, "
    "and not legal advice. Use only as a starting point — a qualified "
    "clinician must author and sign the actual clinical record."
)

# -- The recognised vitals keys, in the order Word readers conventionally
# -- expect them in a vitals row of an inpatient observation chart. Keys
# -- not on this list are still rendered, in caller-insertion order, after
# -- the recognised ones — see ``_render_vitals_table`` for the algorithm. --
_VITALS_LABELS: "List[tuple[str, str, str]]" = [
    # (key, label, unit)
    ("bp", "Blood pressure", "mmHg"),
    ("hr", "Heart rate", "bpm"),
    ("temp", "Temperature", "°C"),
    ("rr", "Respiratory rate", "breaths/min"),
    ("spo2", "SpO2", "%"),
    ("weight", "Weight", "kg"),
    ("height", "Height", "cm"),
    ("bmi", "BMI", "kg/m²"),
    ("pain", "Pain score", "/10"),
]


# -- Helpers --------------------------------------------------------------


def _add_title(document: "DocumentCls", title: str) -> "Paragraph":
    """Append a centred document title in ``Title`` style (or fallback)."""
    style = "Title"
    try:
        document.styles[style]
    except KeyError:
        style = "Normal"
    para = document.add_paragraph(title, style=style)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _add_disclaimer(document: "DocumentCls") -> "Paragraph":
    """Append the standard "template only" notice to ``document``."""
    para = document.add_paragraph()
    run = para.add_run(_DISCLAIMER)
    run.italic = True
    run.bold = True
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return para


def _validate_patient(patient: Optional[Mapping[str, Any]]) -> None:
    """Raise :class:`ValueError` when ``patient`` is missing or malformed.

    Every clinical document this module emits requires at minimum a
    patient ``name``. ``dob`` and ``mrn`` are conventional but not
    enforced — a real EHR would enforce MRN, but the kit is permissive
    so callers can use it for templating before an MRN is assigned.
    """
    if patient is None:
        raise ValueError("patient is required")
    if not isinstance(patient, Mapping):  # type: ignore[arg-type]
        raise ValueError("patient must be a mapping with a 'name' key")
    if not patient.get("name"):
        raise ValueError("patient is missing a non-empty 'name'")


def _validate_provider(provider: Optional[Mapping[str, Any]]) -> None:
    """Raise :class:`ValueError` when ``provider`` is malformed.

    Provider is required (a clinical document with no signing clinician
    is not a clinical document) and must have a non-empty ``name``.
    ``role`` and ``practice`` are conventional but optional.
    """
    if provider is None:
        raise ValueError("provider is required")
    if not isinstance(provider, Mapping):  # type: ignore[arg-type]
        raise ValueError("provider must be a mapping with a 'name' key")
    if not provider.get("name"):
        raise ValueError("provider is missing a non-empty 'name'")


def _format_diagnosis(item: Any) -> str:
    """Format a diagnosis entry — string or ``{text, code}`` mapping.

    Strings are rendered verbatim. Mappings get ``"text (CODE)"``
    formatting where ``code`` is the ICD-10-AM (or any other coding
    system) identifier. An entry like ``"Acute bronchitis (J20.9)"``
    that already has the parenthesised code is returned as-is.
    """
    if isinstance(item, str):
        return item
    if isinstance(item, Mapping):
        text = str(item.get("text", "")).strip()
        code = str(item.get("code", "")).strip()
        if text and code:
            return f"{text} ({code})"
        if text:
            return text
        if code:
            return code
        raise ValueError(
            "diagnosis mapping must include 'text' or 'code'"
        )
    return str(item)


def _add_patient_block(
    document: "DocumentCls",
    patient: Mapping[str, Any],
) -> "Table":
    """Append the patient identification table.

    Renders a two-column "Field / Value" table with one row per
    recognised key (``name``, ``dob``, ``mrn``, ``medicare``, ``sex``,
    ``address``, ``phone``). Unrecognised keys are appended verbatim.
    The structured form is more scannable than free-text and trivially
    extractable for EHR import.
    """
    document.add_heading("Patient", level=1)
    recognised = [
        ("name", "Name"),
        ("dob", "Date of birth"),
        ("mrn", "MRN"),
        ("medicare", "Medicare"),
        ("sex", "Sex"),
        ("address", "Address"),
        ("phone", "Phone"),
    ]
    rows: "List[tuple[str, str]]" = []
    for key, label in recognised:
        if key in patient and patient[key] not in (None, ""):
            rows.append((label, str(patient[key])))
    extras = [
        k for k in patient.keys() if k not in {r[0] for r in recognised}
    ]
    for key in extras:
        if patient[key] in (None, ""):
            continue
        rows.append((str(key).replace("_", " ").title(), str(patient[key])))

    table = document.add_table(rows=len(rows), cols=2)
    try:
        table.style = "Table Grid"
    except KeyError:  # pragma: no cover -- default template ships Table Grid
        pass
    for index, (label, value) in enumerate(rows):
        cells = table.rows[index].cells
        cells[0].text = label
        cells[1].text = value
    return table


def _add_provider_block(
    document: "DocumentCls", provider: Mapping[str, Any]
) -> "List[Paragraph]":
    """Append a "Provider" heading + provider-detail paragraphs."""
    paragraphs: "List[Paragraph]" = []
    paragraphs.append(document.add_heading("Provider", level=1))
    name = str(provider.get("name", "")).strip()
    role = str(provider.get("role", "")).strip()
    practice = str(provider.get("practice", "")).strip()
    provider_no = str(
        provider.get("provider_number", provider.get("provider_no", ""))
    ).strip()
    contact = str(provider.get("contact", provider.get("phone", ""))).strip()

    line = name
    if role:
        line = f"{line} ({role})"
    paragraphs.append(document.add_paragraph(line))
    if practice:
        paragraphs.append(document.add_paragraph(practice))
    if provider_no:
        paragraphs.append(
            document.add_paragraph(f"Provider No: {provider_no}")
        )
    if contact:
        paragraphs.append(document.add_paragraph(f"Contact: {contact}"))
    return paragraphs


def _render_vitals_table(
    document: "DocumentCls", vitals: Mapping[str, Any]
) -> "Table":
    """Render a "Vitals" parameter/value table.

    Recognised keys (``bp``, ``hr``, ``temp``, ``rr``, ``spo2``,
    ``weight``, ``height``, ``bmi``, ``pain``) are emitted in the order
    declared in :data:`_VITALS_LABELS`, with the unit appended in the
    "Value" column. Unrecognised keys are emitted verbatim after the
    recognised ones, in caller-insertion order, so unusual measurements
    (e.g. ``cvp``, ``fio2``) round-trip.
    """
    rows: "List[tuple[str, str]]" = []
    recognised_keys = {entry[0] for entry in _VITALS_LABELS}
    for key, label, unit in _VITALS_LABELS:
        if key in vitals and vitals[key] not in (None, ""):
            value = vitals[key]
            value_text = (
                f"{value} {unit}".strip()
                if unit and not str(value).endswith(unit)
                else str(value)
            )
            rows.append((label, value_text))
    extras = [k for k in vitals.keys() if k not in recognised_keys]
    for key in extras:
        if vitals[key] in (None, ""):
            continue
        rows.append(
            (
                str(key).replace("_", " ").title(),
                str(vitals[key]),
            )
        )

    table = document.add_table(rows=1 + len(rows), cols=2)
    try:
        table.style = "Table Grid"
    except KeyError:  # pragma: no cover
        pass
    header = table.rows[0].cells
    header[0].text = "Parameter"
    header[1].text = "Value"
    for index, (label, value) in enumerate(rows, start=1):
        cells = table.rows[index].cells
        cells[0].text = label
        cells[1].text = value
    return table


def _add_bullet_list(
    document: "DocumentCls",
    items: Sequence[Any],
    formatter: Optional[Any] = None,
) -> "List[Paragraph]":
    """Append one bulleted paragraph per item.

    ``formatter`` is an optional callable that turns each item into the
    paragraph's text — defaults to :func:`str`. Uses Word's built-in
    ``List Bullet`` style with a ``Normal`` fallback (a custom template
    may not ship the bullet style).
    """
    style = "List Bullet"
    try:
        document.styles[style]
    except KeyError:
        style = "Normal"
    fmt = formatter or str
    paragraphs: "List[Paragraph]" = []
    for item in items:
        paragraphs.append(
            document.add_paragraph(fmt(item), style=style)
        )
    return paragraphs


def _add_numbered_list(
    document: "DocumentCls",
    items: Sequence[Any],
    formatter: Optional[Any] = None,
) -> "List[Paragraph]":
    """Append one numbered paragraph per item.

    Falls back to ``"N. <text>"`` flat numbering inside the ``Normal``
    style when ``List Number`` is unavailable, so the helper degrades
    cleanly on minimal templates.
    """
    style = "List Number"
    fmt = formatter or str
    paragraphs: "List[Paragraph]" = []
    try:
        document.styles[style]
        for item in items:
            paragraphs.append(
                document.add_paragraph(fmt(item), style=style)
            )
    except KeyError:
        for index, item in enumerate(items, start=1):
            paragraphs.append(
                document.add_paragraph(f"{index}. {fmt(item)}")
            )
    return paragraphs


def _add_section_text(
    document: "DocumentCls",
    heading: str,
    body: Union[str, Sequence[str]],
    level: int = 1,
) -> "List[Paragraph]":
    """Append a heading + one paragraph per item in ``body``."""
    paragraphs: "List[Paragraph]" = [
        document.add_heading(heading, level=level)
    ]
    if isinstance(body, str):
        chunks: Sequence[str] = [body]
    else:
        chunks = list(body)
    for chunk in chunks:
        if chunk:
            paragraphs.append(document.add_paragraph(chunk))
    return paragraphs


def _add_signature_line(
    document: "DocumentCls", provider: Mapping[str, Any]
) -> "List[Paragraph]":
    """Append the clinician signature stub."""
    paragraphs: "List[Paragraph]" = []
    paragraphs.append(document.add_heading("Signed", level=1))
    name = str(provider.get("name", "Clinician"))
    paragraphs.append(
        document.add_paragraph(
            "Signature: ______________________________"
        )
    )
    paragraphs.append(document.add_paragraph(f"Clinician: {name}"))
    paragraphs.append(
        document.add_paragraph(
            "Date: ______________________________"
        )
    )
    return paragraphs


# -- SOAP note ------------------------------------------------------------


def soap_note(
    patient: Optional[Mapping[str, Any]] = None,
    encounter_date: Optional[str] = None,
    provider: Optional[Mapping[str, Any]] = None,
    subjective: Optional[Union[str, Sequence[str]]] = None,
    objective: Optional[Mapping[str, Any]] = None,
    assessment: Optional[Sequence[Any]] = None,
    plan: Optional[Sequence[Any]] = None,
) -> "DocumentCls":
    """Build a SOAP-format clinical note and return the |Document|.

    SOAP — Subjective / Objective / Assessment / Plan — is the standard
    progress-note structure used in primary care worldwide and the
    default record format for the Australian Medicare Benefits Schedule
    item descriptors that require "appropriate clinical record
    documentation".

    Parameters
    ----------
    patient
        Mapping with at minimum a ``name``; commonly ``dob``, ``mrn``,
        ``medicare``, ``sex``, ``address``, ``phone``. Unrecognised keys
        are rendered verbatim after the recognised ones.
    encounter_date
        ISO date string (``"2026-03-15"``) rendered verbatim under the
        title.
    provider
        Mapping with at minimum a ``name``; commonly ``role``,
        ``practice``, ``provider_number``, ``contact``.
    subjective
        Free-text patient history / chief complaint. Pass a string for a
        single paragraph, or a sequence of strings for one paragraph per
        item.
    objective
        Mapping that may include ``vitals`` (a dict — see the module
        docstring), ``examination`` (free text), ``labs`` (a sequence of
        strings rendered as a bulleted list).
    assessment
        Sequence of diagnoses. Each item is either a string (rendered
        verbatim) or a mapping ``{"text": ..., "code": ...}`` (rendered
        as ``"text (CODE)"``).
    plan
        Sequence of plan items (medications, follow-up, education).
        Rendered as a numbered list.

    Returns
    -------
    Document
        The freshly-built |Document|. Save with :meth:`Document.save`.

    Raises
    ------
    ValueError
        When ``patient`` is missing / malformed, or when ``provider`` is
        missing / malformed.

    .. warning::
        **Template only — not a medical record.** See the module
        docstring for the full disclaimer.

    .. versionadded:: 2026.05.29
    """
    _validate_patient(patient)
    _validate_provider(provider)
    assert patient is not None and provider is not None  # for type-checker

    document = Document()
    _add_title(document, "Clinical Note (SOAP)")
    _add_disclaimer(document)

    if encounter_date:
        date_para = document.add_paragraph(
            f"{_DEFAULT_DATE_LABEL}: {encounter_date}"
        )
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    _add_patient_block(document, patient)
    _add_provider_block(document, provider)

    # -- Subjective --
    document.add_heading("Subjective", level=1)
    if subjective:
        chunks: Sequence[str] = (
            [subjective] if isinstance(subjective, str) else list(subjective)
        )
        for chunk in chunks:
            if chunk:
                document.add_paragraph(chunk)
    else:
        document.add_paragraph(
            "[Insert patient-reported history and chief complaint here.]"
        )

    # -- Objective --
    document.add_heading("Objective", level=1)
    obj = objective or {}
    vitals = obj.get("vitals") if isinstance(obj, Mapping) else None
    if vitals:
        document.add_heading("Vitals", level=2)
        _render_vitals_table(document, vitals)
    examination = obj.get("examination") if isinstance(obj, Mapping) else None
    if examination:
        document.add_heading("Examination", level=2)
        if isinstance(examination, str):
            document.add_paragraph(examination)
        else:
            for chunk in examination:
                if chunk:
                    document.add_paragraph(str(chunk))
    labs = obj.get("labs") if isinstance(obj, Mapping) else None
    if labs:
        document.add_heading("Investigations", level=2)
        _add_bullet_list(document, list(labs))
    if not vitals and not examination and not labs:
        document.add_paragraph(
            "[Insert vital signs, examination findings, and "
            "investigation results here.]"
        )

    # -- Assessment --
    document.add_heading("Assessment", level=1)
    if assessment:
        _add_bullet_list(
            document, list(assessment), formatter=_format_diagnosis
        )
    else:
        document.add_paragraph(
            "[Insert clinical impression / problem list here.]"
        )

    # -- Plan --
    document.add_heading("Plan", level=1)
    if plan:
        _add_numbered_list(document, list(plan))
    else:
        document.add_paragraph(
            "[Insert management plan — medications, follow-up, "
            "patient education — here.]"
        )

    document.add_page_break()
    _add_signature_line(document, provider)

    return document


# -- Discharge summary ----------------------------------------------------


def discharge_summary(
    patient: Optional[Mapping[str, Any]] = None,
    admission_date: Optional[str] = None,
    discharge_date: Optional[str] = None,
    provider: Optional[Mapping[str, Any]] = None,
    presenting_complaint: Optional[str] = None,
    history: Optional[Union[str, Sequence[str]]] = None,
    investigations: Optional[Sequence[Any]] = None,
    procedures: Optional[Sequence[Any]] = None,
    diagnoses: Optional[Sequence[Any]] = None,
    discharge_medications: Optional[Sequence[Any]] = None,
    follow_up: Optional[Union[str, Sequence[str]]] = None,
    discharge_vitals: Optional[Mapping[str, Any]] = None,
) -> "DocumentCls":
    """Build a hospital discharge summary and return the |Document|.

    A discharge summary is the document the inpatient team sends to the
    patient's GP at discharge. It records the reason for admission,
    investigations, procedures, final diagnoses, the medications the
    patient is going home on, and the follow-up plan. The shape here
    follows the Australian Commission on Safety and Quality in Health
    Care's "National guidelines for on-screen presentation of discharge
    summaries" (2017) section ordering.

    Parameters
    ----------
    patient
        Patient mapping (see :func:`soap_note`).
    admission_date, discharge_date
        ISO date strings rendered verbatim.
    provider
        Mapping for the discharging clinician (or team contact).
    presenting_complaint
        Free-text reason for admission.
    history
        History of presenting complaint plus relevant background. Pass a
        string for a single paragraph, or a sequence for multiple.
    investigations
        Sequence of investigation entries. Each item is either a string
        or a mapping ``{"name": ..., "result": ...}``.
    procedures
        Sequence of procedure entries (string or
        ``{"name": ..., "date": ...}`` mapping).
    diagnoses
        Sequence of final diagnoses (string or
        ``{"text": ..., "code": ...}`` mapping).
    discharge_medications
        Sequence of discharge-medication entries (string or
        ``{"name": ..., "dose": ..., "frequency": ..., "duration": ...}``
        mapping; missing fields are elided).
    follow_up
        Follow-up plan — string or sequence of strings.
    discharge_vitals
        Optional mapping of vital signs at discharge — same shape as
        :func:`soap_note`'s ``objective.vitals``.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``patient`` or ``provider`` is missing / malformed.

    .. warning::
        **Template only — not a medical record.** See module docstring.

    .. versionadded:: 2026.05.29
    """
    _validate_patient(patient)
    _validate_provider(provider)
    assert patient is not None and provider is not None

    document = Document()
    _add_title(document, "Hospital Discharge Summary")
    _add_disclaimer(document)

    if admission_date or discharge_date:
        if admission_date and discharge_date:
            line = (
                f"Admission: {admission_date}  ·  "
                f"Discharge: {discharge_date}"
            )
        elif admission_date:
            line = f"Admission: {admission_date}"
        else:
            line = f"Discharge: {discharge_date}"
        date_para = document.add_paragraph(line)
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    _add_patient_block(document, patient)
    _add_provider_block(document, provider)

    # -- Presenting complaint --
    document.add_heading("Presenting Complaint", level=1)
    if presenting_complaint:
        document.add_paragraph(presenting_complaint)
    else:
        document.add_paragraph("[Insert reason for admission here.]")

    # -- History --
    if history:
        chunks: Sequence[str] = (
            [history] if isinstance(history, str) else list(history)
        )
        document.add_heading("History", level=1)
        for chunk in chunks:
            if chunk:
                document.add_paragraph(chunk)

    # -- Investigations --
    document.add_heading("Investigations", level=1)
    if investigations:
        _add_bullet_list(
            document,
            list(investigations),
            formatter=_format_investigation,
        )
    else:
        document.add_paragraph(
            "[Insert pathology, imaging, and other investigation "
            "results here.]"
        )

    # -- Procedures --
    if procedures:
        document.add_heading("Procedures", level=1)
        _add_bullet_list(
            document,
            list(procedures),
            formatter=_format_procedure,
        )

    # -- Diagnoses --
    document.add_heading("Final Diagnoses", level=1)
    if diagnoses:
        _add_numbered_list(
            document, list(diagnoses), formatter=_format_diagnosis
        )
    else:
        document.add_paragraph(
            "[Insert principal and additional diagnoses here.]"
        )

    # -- Discharge medications --
    document.add_heading("Discharge Medications", level=1)
    if discharge_medications:
        _add_bullet_list(
            document,
            list(discharge_medications),
            formatter=_format_medication,
        )
    else:
        document.add_paragraph("[Insert discharge medication list here.]")

    # -- Discharge vitals (optional) --
    if discharge_vitals:
        document.add_heading("Vitals at Discharge", level=1)
        _render_vitals_table(document, discharge_vitals)

    # -- Follow-up --
    document.add_heading("Follow-up Plan", level=1)
    if follow_up:
        chunks = (
            [follow_up] if isinstance(follow_up, str) else list(follow_up)
        )
        for chunk in chunks:
            if chunk:
                document.add_paragraph(chunk)
    else:
        document.add_paragraph(
            "[Insert GP review, specialist appointments, and "
            "patient education items here.]"
        )

    document.add_page_break()
    _add_signature_line(document, provider)

    return document


# -- Referral letter ------------------------------------------------------


def referral_letter(
    patient: Optional[Mapping[str, Any]] = None,
    referrer: Optional[Mapping[str, Any]] = None,
    recipient: Optional[Mapping[str, Any]] = None,
    encounter_date: Optional[str] = None,
    reason: Optional[str] = None,
    history: Optional[Union[str, Sequence[str]]] = None,
    examination: Optional[Union[str, Sequence[str]]] = None,
    investigations: Optional[Sequence[Any]] = None,
    medications: Optional[Sequence[Any]] = None,
    allergies: Optional[Sequence[str]] = None,
    requested_action: Optional[Union[str, Sequence[str]]] = None,
    closing: Optional[str] = None,
) -> "DocumentCls":
    """Build a clinician-to-clinician referral letter and return the |Document|.

    A referral letter is the standard handover document a GP sends to a
    specialist (or a junior clinician sends to a senior). The shape
    follows the Royal Australian College of General Practitioners
    (RACGP) "Standards for general practices" guidance on referral
    content: identification, clinical question, relevant history,
    examination, investigations to date, current medications, known
    allergies, and the specific action requested.

    Parameters
    ----------
    patient
        Patient mapping (see :func:`soap_note`).
    referrer
        Mapping for the referring clinician — same shape as
        :func:`soap_note`'s ``provider``.
    recipient
        Mapping for the addressed clinician. Optional; when missing the
        salutation falls back to ``"Dear Colleague"``. Recognised keys
        ``name``, ``role``, ``practice``, ``address``.
    encounter_date
        ISO date string for the consultation that prompted the referral.
    reason
        One- or two-sentence statement of the clinical question. Use
        plain language — "Please assess for X and advise on Y".
    history
        Relevant history of presenting complaint plus background.
    examination
        Examination findings.
    investigations
        Sequence of investigations already performed (string or
        ``{"name": ..., "result": ...}`` mapping).
    medications
        Sequence of current medications (string or
        ``{"name": ..., "dose": ..., "frequency": ...}`` mapping).
    allergies
        Sequence of known allergies (rendered as a single bulleted
        paragraph; ``["Nil known"]`` is the conventional null entry).
    requested_action
        Specific action requested of the recipient (string or sequence
        of strings — one per request).
    closing
        Free-text closing line. Defaults to a polite RACGP-style sign-off.

    Returns
    -------
    Document
        The freshly-built |Document|.

    Raises
    ------
    ValueError
        When ``patient`` or ``referrer`` is missing / malformed.

    .. warning::
        **Template only — not a medical record.** See module docstring.

    .. versionadded:: 2026.05.29
    """
    _validate_patient(patient)
    _validate_provider(referrer)
    assert patient is not None and referrer is not None

    document = Document()
    _add_title(document, "Referral Letter")
    _add_disclaimer(document)

    if encounter_date:
        date_para = document.add_paragraph(
            f"{_DEFAULT_DATE_LABEL}: {encounter_date}"
        )
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # -- Recipient block --
    document.add_heading("To", level=1)
    if recipient and recipient.get("name"):
        recipient_name = str(recipient.get("name", "")).strip()
        recipient_role = str(recipient.get("role", "")).strip()
        recipient_practice = str(recipient.get("practice", "")).strip()
        recipient_address = str(recipient.get("address", "")).strip()
        line = recipient_name
        if recipient_role:
            line = f"{line} ({recipient_role})"
        document.add_paragraph(line)
        if recipient_practice:
            document.add_paragraph(recipient_practice)
        if recipient_address:
            document.add_paragraph(recipient_address)
        salutation = f"Dear {recipient_name},"
    else:
        document.add_paragraph("[Insert recipient clinician details here.]")
        salutation = "Dear Colleague,"

    document.add_paragraph(salutation)

    # -- Patient identification --
    _add_patient_block(document, patient)

    # -- Reason for referral --
    document.add_heading("Reason for Referral", level=1)
    if reason:
        document.add_paragraph(reason)
    else:
        document.add_paragraph(
            "[Insert the specific clinical question or reason for "
            "referral here.]"
        )

    # -- History --
    if history:
        _add_section_text(document, "History", history)

    # -- Examination --
    if examination:
        _add_section_text(document, "Examination", examination)

    # -- Investigations to date --
    if investigations:
        document.add_heading("Investigations to Date", level=1)
        _add_bullet_list(
            document,
            list(investigations),
            formatter=_format_investigation,
        )

    # -- Current medications --
    if medications:
        document.add_heading("Current Medications", level=1)
        _add_bullet_list(
            document,
            list(medications),
            formatter=_format_medication,
        )

    # -- Allergies --
    if allergies:
        document.add_heading("Allergies", level=1)
        _add_bullet_list(document, list(allergies))

    # -- Requested action --
    document.add_heading("Requested Action", level=1)
    if requested_action:
        if isinstance(requested_action, str):
            document.add_paragraph(requested_action)
        else:
            _add_numbered_list(document, list(requested_action))
    else:
        document.add_paragraph(
            "[Insert the specific action requested of the recipient — "
            "e.g. assessment, investigation, treatment, ongoing care.]"
        )

    # -- Closing --
    closing_text = closing or (
        "Thank you for seeing this patient. Please feel free to contact "
        "me if you require any further information."
    )
    document.add_paragraph(closing_text)
    document.add_paragraph("Kind regards,")

    # -- Referrer block --
    _add_provider_block(document, referrer)

    document.add_page_break()
    _add_signature_line(document, referrer)

    return document


# -- Per-domain item formatters -------------------------------------------


def _format_investigation(item: Any) -> str:
    """Format an investigation entry — string or ``{name, result}`` mapping."""
    if isinstance(item, str):
        return item
    if isinstance(item, Mapping):
        name = str(item.get("name", "")).strip()
        result = str(item.get("result", "")).strip()
        if name and result:
            return f"{name}: {result}"
        return name or result or ""
    return str(item)


def _format_procedure(item: Any) -> str:
    """Format a procedure entry — string or ``{name, date}`` mapping."""
    if isinstance(item, str):
        return item
    if isinstance(item, Mapping):
        name = str(item.get("name", "")).strip()
        date = str(item.get("date", "")).strip()
        if name and date:
            return f"{name} ({date})"
        return name or date or ""
    return str(item)


def _format_medication(item: Any) -> str:
    """Format a medication entry — string or
    ``{name, dose, frequency, duration}`` mapping.

    Renders ``"Amoxicillin 500mg TDS for 7 days"`` from the structured
    form. Missing fields are elided so partial input still produces
    legible output.
    """
    if isinstance(item, str):
        return item
    if isinstance(item, Mapping):
        name = str(item.get("name", "")).strip()
        dose = str(item.get("dose", "")).strip()
        frequency = str(item.get("frequency", "")).strip()
        duration = str(item.get("duration", "")).strip()
        parts: "List[str]" = []
        if name:
            parts.append(name)
        if dose:
            parts.append(dose)
        if frequency:
            parts.append(frequency)
        line = " ".join(parts)
        if duration:
            line = (
                f"{line} for {duration}".strip()
                if line
                else f"for {duration}"
            )
        return line
    return str(item)


__all__ = [
    "soap_note",
    "discharge_summary",
    "referral_letter",
]
