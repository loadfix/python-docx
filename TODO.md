# TODO

Fork-specific feature backlog across the loadfix OOXML trio. Each item is
a candidate for a future implementation wave. Grouped by repo.

## docx

- **Real XLookup / complex-field evaluation.** Current fork can read
  complex field codes and resolve REF/PAGEREF/DOCPROPERTY, but most
  other field types (IF, HYPERLINK, MERGEFIELD with conditions, formula
  fields) are returned as raw field-code + cached result. Implement a
  proper evaluator.
- **Bibliography / citation support.** `Document.bibliography` read
  exists but not create. Need `Document.add_citation()` wiring the
  `customXml/bibliography.xml` part.
- **Emit `w:bCs` / `w:iCs` alongside `w:b` / `w:i` for CS-affecting
  runs.** Three-way diff of the corpus bold-text manifest shows Word
  writes both `<w:b/>` and `<w:bCs/>` (complex-script bold). python-docx
  emits only `<w:b/>`, so bold on Arabic/Hebrew/Thai text is dropped by
  Word on load. Same issue for italic (`<w:i/>` vs `<w:iCs/>`).
  Surfaced by `ooxml-reference-corpus/features/docx/bold-text.json`
  three-way comparison on 2026-05-04.

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
