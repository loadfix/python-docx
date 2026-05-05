# TODO

Fork-specific feature backlog across the loadfix OOXML trio. Each item is
a candidate for a future implementation wave. Grouped by repo.

## Audit findings 2026-05-05

- [ ] **Remove shipped `Section.vertical_alignment` entry from "Conformance gaps".** Already resolved by commit `1657c0ef`; move to a "Resolved" block or delete the stale open entry below.
- [ ] **Close GitHub issue #171 (vt:date custom-properties).** Partially addressed by `c3edf01b` in 2026.05.8 (accept `datetime.date`, serialise as `vt:date`); verify round-trip fully works and close the issue.
- [ ] **Close GitHub issue #172 (bare `KeyError('[Content_Types].xml')` on missing part).** Wrap in a typed `PackageNotFoundError` (or similar) at the OPC load boundary.
- [ ] **Bump README "Current version" string.** Currently reads `2026.05.0`; should be `2026.05.8`.
- [x] **Git-tag all untagged releases.** Only `v2026.05.0` has a git tag; `v2026.05.1` through `v2026.05.8` have HISTORY.rst entries but no git tags. Add all 8 missing tags.
- [x] **Land W11-D `UPSTREAM_SYNC.md` onto master.** Commit `d2d5cdcf` lives only on branch `feat/w11-d-upstream-sync`; merged to master via merge commit `721b7753` so upstream divergence is documented on the canonical tree.
- [ ] **Seal submodule oxml leakage.** 590 `CT_*`/`ST_*` names reachable via `docx.oxml.*` / `docx.dml.color.*` / `docx.drawing.*` / etc. Add explicit `__all__` to public submodules (without breaking internal re-imports).
- [x] **Fix dev-extras portability.** `pyproject.toml` declared `ooxml-validate @ file:///home/ben/code/ooxml-validate`, a host-specific absolute path. Moved out of `[dev]` into a new opt-in `[conformance]` extra pointing at the GitHub VCS URL so `pip install -e '.[dev]'` works for external contributors.
- [x] **Delete obsolete `.travis.yml`.** 243 bytes of dead config; `.github/` was deliberately removed.
- [x] **Move scratch audit artefacts to `audits/`.** DOCS_AUDIT.md (49KB), DOCS_SIBLING_AUDIT.md (47KB), FEATURES_AUDIT.md (42KB), TEST_AUDIT.md (30KB), INTEROP_REPORT.md (17KB), SCALE_NOTES.md (5KB), real-world-audit-findings.md (11KB) accumulating at repo root — move to an `audits/` subdirectory or delete after resolution.
- [x] **Prune merged remote branches.** ~14 feature branches on origin (`feat/w10-*`, `fix/w8-*`, `chore/overnight-*`, `worktree-agent-*`) merged but not deleted.
- [ ] **Fix API-addition version markers in FEATURES.md.** Many entries uniformly say `[Added in 2026.05.0]` regardless of the release that actually introduced them; audit and correct to reflect .1–.8 where appropriate.

## docx

_No open docx items — bibliography and complex-field evaluation both
shipped in 2026.05.8 (see "Completed items" below)._

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

### Performance fixes

- **W11-A: O(N^2) indexing on `_Rows[i]` and `Document.paragraphs[i]`**
  (closed 2026-05-05 on `fix/w11-a-indexing-perf`). `_Rows.__getitem__`
  and `BlockItemContainer.paragraphs` both materialised the entire
  child list on every call; a naive indexed loop was O(N^2). Replaced
  `_Rows.__getitem__` with direct `tr_lst[idx]` access and replaced
  `BlockItemContainer.paragraphs` with a lazy
  `_ParagraphsView(Sequence[Paragraph])` that memoises `p_lst` on
  first access. Cached-idiom access dropped from ~1.53 ms/access to
  ~0.0007 ms/access at N=5 000 paragraphs. See `SCALE_NOTES.md` for
  methodology and post-fix numbers.

### Authoring features (2026.05.8)

- **Bibliography / citation support.** `Document.bibliography`
  (read + write), `Document.add_citation(tag, ...)`, `Paragraph.add_citation_reference(tag)`,
  and the backing `/customXml/item{N}.xml` part with a `<b:Sources>` root plus a
  sibling `itemProps{N}.xml`. See `FEATURES.md` § "Bibliography and citations".
- **SmartArt authoring.** `Document.add_smart_art(layout_name)` and
  `SmartArt.add_node(text)` with three built-in layouts (list, cycle,
  process). See `FEATURES.md` § "SmartArt".
- **Complex-field evaluation.** `Field.evaluate(context)` and
  `Document.evaluate_fields(context)` now evaluate `IF` (with nested
  `{MERGEFIELD}`), `MERGEFIELD`, `HYPERLINK`, `= <expr>` arithmetic
  formulas, and the runtime-dynamic `PAGE` / `NUMPAGES` / `DATE` /
  `TIME` placeholders. Deferred: string-function formulas (`=SUM()`,
  `=AVERAGE()`, etc. beyond arithmetic), nested `IF`, `QUOTE`, `FILLIN`,
  and the full date-picture/numeric-format switch grammar.

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
