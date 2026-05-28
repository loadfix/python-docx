"""``docx.kit`` — opinionated authoring helpers built on top of python-docx.

The :mod:`docx.kit` submodule is a curated collection of high-level
"recipe" helpers that compose the existing python-docx public API into
common document-authoring patterns. Each helper is a thin wrapper that
adds a styled, shaped chunk of content to a |Document| in one call —
the kind of boilerplate every report / book / proposal author writes
by hand the first time and copy-pastes thereafter. The kit is **opt-in**
via the ``[kit]`` extras flag (``pip install python-docx[kit]``) so
callers who use only the core authoring API pay nothing for it.

The kit lives **inside** :mod:`docx` (not as a sibling package) per the
project's Wave-4 scoping memo: kit content is Python helpers tightly
coupled to a single parent's public API with no cross-parent reuse
story, so an in-parent submodule under a ``[kit]`` extras flag is the
right home (the ``[kit]`` extras list is currently empty — kit modules
are pure-Python compositions of existing python-docx surface and add no
new runtime dependencies; the flag exists as a versioning hook so
future kit issues can declare deps without touching the core dependency
list).

Available kit submodules:

* :mod:`docx.kit.front_matter` — title / copyright / dedication /
  preface / TOC / list-of-figures / list-of-tables helpers.
* :mod:`docx.kit.chapter` — chapter opener pages (large title +
  decorative image + drop cap).
* :mod:`docx.kit.dividers` — section-divider / chapter-ornament
  helpers (``add_divider`` / ``add_fleuron`` / ``add_three_stars`` /
  ``add_chapter_break``) for inserting fleurons and decorative
  breaks between long-form-document sections.
* :mod:`docx.kit.back_matter` — appendix / glossary / index /
  bibliography helpers.
* :mod:`docx.kit.brand` — :func:`~docx.kit.brand.validate_brand` brand-
  guideline linter. Walks a document and surfaces a list of
  :class:`~docx.kit.brand.BrandFinding` records covering five rules
  (``font-not-on-brand`` / ``color-not-on-brand`` / ``wrong-logo`` /
  ``heading-style-mismatch`` / ``inconsistent-spacing``) against a
  YAML / dict / ``BrandAssets``-shaped palette.
* :mod:`docx.kit.letterhead` — branded header (logo + return address)
  and footer (phone / email / website) with three built-in styles
  (``modern`` / ``classic`` / ``minimal``).
* :mod:`docx.kit.resume` — resume / CV template family
  (``resume_chronological`` / ``resume_functional`` / ``resume_technical``)
  with three built-in styles (``modern`` / ``classic`` / ``minimal``).
* :mod:`docx.kit.contracts` — contract / NDA template family
  (``nda`` / ``msa`` / ``sow`` / ``contractor_agreement``) with
  AUS-default boilerplate. Output is a *starting point only* — the
  module docstring carries an explicit "not legal advice" disclaimer.
* :mod:`docx.kit.invoices` — invoice / quote / statement template
  family (``invoice`` / ``quote`` / ``statement``) with AUS GST
  defaults (10%, override per-line via ``gst_rate=0`` for
  international callers), auto-computed subtotal / GST / grand total,
  and a right-aligned line-item table. Output complies with ATO
  tax-invoice rules when the seller carries an ABN.
* :mod:`docx.kit.mail_merge` — bulk render N personalised documents
  from a single template + iterable of records, composing the
  smart-placeholder machinery from #68 with an ergonomic
  one-line API.
* :mod:`docx.kit.memos` — investment memo / business case template
  family (``investment_memo`` / ``business_case``) with McKinsey-style
  SCQA (Situation / Complication / Question / Answer) structure for
  memos and an options-analysis table for business cases.
* :mod:`docx.kit.templates` — generic document template registry
  (``brief`` / ``coe`` / ``rfp_response`` / ``white_paper``) covering
  short briefs, Centre of Excellence charters, RFP responses with a
  pricing table, and white papers with abstract and references.
* :mod:`docx.kit.scientific` — scientific paper template family
  (``ieee_paper`` / ``acm_paper`` / ``apa_paper`` / ``nature_paper``)
  applying each venue's structural skeleton (IEEE two-column compact,
  ACM ``sigconf``, APA double-spaced single column, Nature compact
  display style).
* :mod:`docx.kit.legal` — legal industry template family
  (``court_paper`` / ``brief`` / ``declaration`` / ``table_of_authorities``)
  with Federal Court of Australia / NSW Supreme Court front-sheet
  layout, Word built-in line numbering (``w:sectPr/w:lnNumType``),
  and a live ``TOA`` complex field. Output is a *starting point only*
  — the module docstring carries an explicit "not legal advice"
  disclaimer.
* :mod:`docx.kit.medical` — medical clinical-note template family
  (``soap_note`` / ``discharge_summary`` / ``referral_letter``) with a
  Subjective / Objective / Assessment / Plan structure, structured
  vitals table, and an explicit "template only — not a medical record"
  disclaimer rendered into every output document.
* :mod:`docx.kit.brand` — :class:`~docx.kit.brand.BrandAssets`,
  a YAML-driven manifest loader for corporate brand colours
  (RGB triples), font pairs (heading / body), logo path variants
  (full-colour / monochrome / reverse), and conventional spacing
  values. Composes with the rest of the kit (``set_letterhead`` /
  ``add_chapter_opener`` / ``invoice``) so an organisation declares
  its brand once and reuses it across every authored document.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from docx.kit import (
    back_matter,
    brand,
    chapter,
    contracts,
    dividers,
    front_matter,
    invoices,
    legal,
    letterhead,
    mail_merge,
    medical,
    memos,
    resume,
    scientific,
    templates,
)

__all__ = [
    "back_matter",
    "brand",
    "chapter",
    "contracts",
    "dividers",
    "front_matter",
    "invoices",
    "legal",
    "letterhead",
    "mail_merge",
    "medical",
    "memos",
    "resume",
    "scientific",
    "templates",
]
