# TODO

Fork-specific feature backlog across the loadfix OOXML trio. Each item is
a candidate for a future implementation wave. Grouped by repo.

## docx

- **Real XLookup / complex-field evaluation.** Current fork can read
  complex field codes and resolve REF/PAGEREF/DOCPROPERTY, but most
  other field types (IF, HYPERLINK, MERGEFIELD with conditions, formula
  fields) are returned as raw field-code + cached result. Implement a
  proper evaluator.

## pptx

- **Transitions authoring.** Read-side exists; no API to set a
  transition type on a slide (fade/push/wipe/etc.) programmatically.
- **Animation timelines.** `timing.xml` is the hardest OOXML format
  to author. Read support partial; no write support.
- **Slide master cascade edit.** Edit layout + master theme inheritance
  and propagate to slides.

## xlsx

- **Formula evaluation.** Workbook reads the string `=SUM(A1:A10)` but
  does not evaluate it. Need a minimal calc engine covering the common
  ~50 functions.
- **Pivot table create.** Read-side exists via `pivot/builder.py`;
  writer is scaffold-only. Complete the builder so new pivots can be
  authored.
- **Conditional formatting writer.** Read exists; create API partial
  (rule.py added accessors but limited coverage).

## Cross-series

- **MS Word / PowerPoint / Excel interop testing.** Test each library
  against a corpus of real-world files authored by Office. Save-reload-
  save cycles to detect fidelity loss. Particular attention to charts,
  equations, images, and content that uses `mc:AlternateContent`.
- **Periodic upstream sync.** Both `scanny/*` and `openpyxl` keep
  moving. Decide a cadence for pulling in upstream bug fixes without
  reverting fork additions.

---

## Completed items (for reference)

All 2026.05.0 feature work. See `HISTORY.rst` and `FEATURES.md` for the
shipped surface.

- **Bibliography / citation support.** Shipped in 2026.05.7: `Document.bibliography`
  (read + write), `Document.add_citation(tag, ...)`, `Paragraph.add_citation_reference(tag)`,
  and the backing `/customXml/item{N}.xml` part with a `<b:Sources>` root plus a
  sibling `itemProps{N}.xml`. See `FEATURES.md` § "Bibliography and citations".

---

## Conformance gaps (auto-filed from corpus 2026-05-04 overnight run)

The 950-case OOXML reference corpus
(`loadfix/ooxml-reference-corpus` built against python-docx at
`d75cfc7`) surfaced one authoring-side API gap. Linked to the driving
corpus manifest on GitHub with an actionable fix hypothesis.

- **`Section.vertical_alignment` property is missing.** Driving
  manifest:
  [features/docx/vertical-alignment.json](https://github.com/loadfix/ooxml-validate/blob/master/features/docx/vertical-alignment.json)
  (P25 finding). Authoring a section with vertical alignment other
  than the `top` default currently requires falling back to raw
  `OxmlElement("w:vAlign")` access on `Section._sectPr` because the
  `Section` proxy (in `src/docx/section.py`) exposes no accessor for
  `w:vAlign`. Fix: add a `Section.vertical_alignment` property with
  getter+setter backed by a new `WD_SECTION_VERTICAL_ALIGNMENT` enum
  (values `TOP=0`, `CENTER=1`, `BOTH=2`, `BOTTOM=3`, mapping to XML
  `top` / `center` / `both` / `bottom` respectively; see
  `spec/xsd/wml.xsd` `ST_VerticalJc`). Plumb it through `CT_SectPr`
  in `src/docx/oxml/section.py` as a `ZeroOrOne("w:vAlign",
  successors=(...))` following the existing successor-ordering
  pattern used by siblings like `titlePg` / `docGrid`. Add unit tests
  under `tests/unit/test_section.py` following the existing
  `Describe*` / `it_*` BDD convention, a behave scenario under
  `features/sct-*.feature`, and the FEATURES.md entry.
