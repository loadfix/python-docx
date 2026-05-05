# Behave Acceptance-Tests Audit (`features/`)

This report surveys the state of the `features/` behave acceptance-tests suite
on the `loadfix/python-docx` fork with three aims:

1. Document what the suite covers today.
2. Map every shipped fork-era feature to "has behave coverage?" / "no behave
   coverage".
3. Propose prioritised follow-ups.

All measurements were taken at commit `50c2078` (`master`, 2026-05-01).

---

## 1. Summary

The project uses **behave** (Gherkin BDD) as its acceptance-test framework,
living entirely under `features/`. The configuration is minimal: there is no
`behave.ini`, no `.behaverc`, no tags, and no wiring into CI; contributors run
it locally with `uv run behave features/`.

- **67 `.feature` files** (2570 lines total)
- **239 `Scenario`/`Scenario Outline` blocks** in the source files, expanding to
  **650 scenarios** at run-time (the outlines produce 411 additional rows from
  `Examples` tables)
- **22 step-definition modules** under `features/steps/` (4103 lines)
- **53 fixture files** under `features/steps/test_files/` (`.docx`, `.png`,
  `.jpg`, `.jpeg`, `.tif`, `.bmp`, `.gif`)
- **1856 steps** executed end-to-end in ~2.0 s

**The suite has not been meaningfully extended in this fork.** The five most
recent commits touching `features/` are:

| SHA | Subject | Date |
|---|---|---|
| `874c1d5` | fix: ensure run.add_picture() produces Word-compatible inline images (#31) (#78) | 2026-04-05 |
| `a809d6c` | comments: add Comment.text | 2025-06-09 |
| `66da522` | xfail: acceptance test for Document.add_comment() | 2025-06-09 |
| `761f4cc` | comments: add Comment.author, .initials setters | 2025-06-09 |
| `8ac9fc4` | comments: add Comments.add_comment() | 2025-06-09 |

The only topic area given fresh behave coverage in this fork is **comments**
(`features/cmt-mutations.feature`, `features/cmt-props.feature`,
`features/doc-add-comment.feature`, `features/doc-comments.feature` — 18
scenarios). Everything else in `features/` pre-dates the fork.

Approximately **55 Microsoft-Word features** have shipped in this fork since
the June 2025 comments work landed. **None** of them have acceptance coverage.
The behave suite as it stands describes only the upstream API surface; the
loadfix extensions (footnotes, endnotes, bookmarks, fields, tracked changes,
content controls, charts, etc.) are exclusively covered by pytest units in
`tests/`.

---

## 2. Layout

```
features/
├── *.feature            # 67 Gherkin spec files
├── environment.py       # behave hooks (before_all only)
├── _scratch/            # run output (gitignored)
└── steps/
    ├── *.py             # 22 step-definition modules
    └── test_files/      # 53 fixture files (.docx + image files)
```

### `features/environment.py` (10 lines)

Only `before_all(context)` is defined: creates `features/_scratch/` if it does
not exist. There is no `after_scenario` cleanup, no tags wiring, no shared
setup. Adding new features that need per-scenario state teardown will require
enlarging this file.

### `features/_scratch/`

Not tracked. `.gitignore` has a line `_scratch/` that correctly covers both
`features/_scratch/` and any other `_scratch/` directory. At the time of
writing `features/_scratch/test_out.docx` exists locally (leftover from a run)
but is correctly ignored. **No hygiene action needed.**

### Step modules (`features/steps/*.py`)

22 modules, 4103 lines total. Listed in descending size with a one-line scope
summary:

| Module | Lines | Feature areas |
|---|---:|---|
| `table.py` | 558 | Tables, rows, columns, cells — spans, props, add/access |
| `styles.py` | 548 | Style access, add/delete, latent styles, style props |
| `text.py` | 322 | Run properties (breaks, char style, inner content) and add-picture |
| `comments.py` | 284 | Comments API (`cmt-*`, `doc-comments`, `doc-add-comment`) |
| `section.py` | 265 | Section iteration, `sct-*.feature`, odd/first-page header/footer |
| `paragraph.py` | 256 | Paragraph access, inner content, set-text, insert, style |
| `document.py` | 259 | `Document.add_*`, `Document.sections/styles/inline_shapes/tables` |
| `font.py` | 227 | Font property matrix (colour, highlight, bold, italic, etc.) |
| `parfmt.py` | 210 | `ParagraphFormat` on/off props, line spacing, alignment |
| `shape.py` | 151 | InlineShape access and size |
| `tabstops.py` | 141 | Tab-stop collection and props |
| `pagebreak.py` | 135 | Rendered page-break splitting (`pbk-split-para`) |
| `hdrftr.py` | 134 | Header/footer iteration, linked-to-previous |
| `coreprops.py` | 117 | `CoreProperties` read/write (title, author, created, etc.) |
| `hyperlink.py` | 116 | Hyperlink properties and fragments |
| `block.py` | 100 | `BlockItemContainer` iteration (`blk-*`) |
| `image.py` | 74 | Pure image-file characterisation (dimensions / DPI / MIME) |
| `api.py` | 59 | `docx.Document` open/save API |
| `settings.py` | 55 | `Settings` object read access |
| `helpers.py` | 34 | `test_docx()` / `test_file()` path helpers + `bool_vals` maps |
| `numbering.py` | 32 | Only one Given for accessing `document.part.numbering_part` |
| `shared.py` | 26 | Trivial `Given a blank document` / `Given a document` impl |

Two modules are effectively shim-only: `numbering.py` (32 lines, one step) and
`shared.py` (26 lines).

### `features/steps/test_files/`

53 fixtures:
- **40 `.docx` files** used by `test_docx(name)` in step modules — one file per
  "preset scenario state" (e.g. `tbl-2x2-table.docx`, `par-known-styles.docx`)
- **13 image files** (`.jpg`, `.jpeg`, `.png`, `.tif`, `.bmp`, `.gif`) used by
  `img-characterize-image.feature` and a few `run.add_picture` steps

One fixture (`doc-odd-even-hdrs.docx`) is tracked but not referenced from any
step module or feature — see §6.

---

## 3. Current coverage

### 3.1 Scenarios per step-module domain

Counted at the source-file level (before `Scenario Outline` expansion). Feature
files are mapped to step modules by filename prefix / content.

| Step module | Feature files covered | Scenarios |
|---|---|---:|
| `table.py` | tbl-*.feature (11 files) | 36 |
| `styles.py` | sty-*.feature (8 files) | 33 |
| `document.py` | doc-access-*, doc-add-*, doc-styles (9 files) | 24 |
| `font.py` | txt-font-*, run-access-font, sty-access-font (4 files) | 21 |
| `comments.py` | cmt-*, doc-add-comment, doc-comments (4 files) | 18 |
| `section.py` | sct-section (1 file) | 17 |
| `text.py` | run-access-inner-content, run-add-*, run-char-style, run-clear-run, run-enum-props, txt-add-break (7 files) | 15 |
| `parfmt.py` | par-access-parfmt, txt-parfmt-props (2 files) | 13 |
| `paragraph.py` | par-access-*, par-add-run, par-*-prop, par-clear, par-insert, par-set-text (7 files) | 13 |
| `tabstops.py` | tab-access-tabs, tab-tabstop-props | 11 |
| `hdrftr.py` | hdr-header-footer | 10 |
| `block.py` | blk-* (3 files) | 7 |
| `hyperlink.py` | hlk-props | 6 |
| `shape.py` | shp-inline-shape-access, shp-inline-shape-size | 4 |
| `pagebreak.py` | pbk-split-para | 4 |
| `coreprops.py` | doc-coreprops | 3 |
| `settings.py` | doc-settings | 3 |
| `api.py` | api-open-document | 2 |
| `numbering.py` | num-access-numbering-part | 1 |
| `image.py` | img-characterize-image (1 scenario outline × 11 rows) | 1 |

### 3.2 Scenario-type split

- `Scenario:` blocks: **115**
- `Scenario Outline:` blocks: **124**
- Total source blocks: **239**
- Runtime expansion: **650** scenarios (outlines contribute 535; each outline
  has 1..N Examples rows, averaging ~4.3)

Heavy use of `Scenario Outline` + `Examples` keeps the feature files compact
and makes them look smaller on paper than they are at runtime.

---

## 4. Build health

```
$ uv run behave features/ 2>&1 | tail -5
67 features passed, 0 failed, 0 skipped
650 scenarios passed, 0 failed, 0 skipped
1856 steps passed, 0 failed, 0 skipped, 0 undefined
Took 0m2.010s
```

- **0 failed / 0 skipped / 0 undefined.** Clean green.
- **Run time**: 2.0 s. Fast enough that adding dozens of new feature files
  would not materially hurt the developer loop.
- **No tags in use.** `grep -rE "^\s*@" features/*.feature` produces no
  output — there are no `@xfail`, `@wip`, `@slow`, or topical tags anywhere in
  the suite.
- **No undefined / pending steps.** Every Gherkin phrase in the suite resolves
  to exactly one step implementation.
- **No xfail / skip markers.** All `xfail:` commits from June 2025 were
  resolved by the follow-up implementation commits the same day.

The suite has zero observable warts — but its size has not kept pace with the
production code, which is the problem this audit exists to quantify.

---

## 5. Coverage gaps — feature-by-feature matrix

This section walks every shipped fork-era issue (as labelled `word-feature-gap`
or `phase-*` on GitHub) and reports whether behave coverage exists. **Most
entries are "No".** Issue numbers, commit SHAs, and suggested `.feature`
filenames follow.

Column key:
- **Existing .feature?** — answer from keyword-grep across `features/*.feature`
- **Suggested filename** — 3-letter prefix + kebab-case, matching the
  repository's established naming convention (`cmt-`, `hlk-`, `par-`, etc.)
- **Effort**: **S** = reuse existing fixtures/steps, **M** = one new `.docx`
  fixture + one new step module or extension, **L** = multiple fixtures or
  complex setup (e.g. numeric data for charts, mail-merge data source)

### 5.1 Phase A — Footnotes / endnotes

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 1 | Phase A.1: Footnotes Part class and relationship management | `5d0178e` | No | (covered by A.2/A.3/A.4) | — |
| 2 | Phase A.2: High-level footnotes API — document.footnotes.add() | `d589c71` | No | `fnt-add-footnote.feature` | M |
| 3 | Phase A.3: Read and iterate existing footnotes | `904ccf4` | No | `fnt-read-footnotes.feature` | M |
| 4 | Phase A.4: Delete and modify footnotes | `86dcafa` | No | `fnt-mutate-footnotes.feature` | M |
| 5 | Phase A.5: Endnotes support (mirror footnotes API) | `9390293` | No | `end-*.feature` (mirror `fnt-*`) | M |
| 17 | Phase A.6: Footnote and endnote properties (numbering, restart, position) | `8bf6011` | No | `fnt-numbering-props.feature` | M |

Sources: `src/docx/footnotes.py`, `src/docx/oxml/footnotes.py`,
`src/docx/parts/footnotes.py`, `src/docx/endnotes.py`,
`src/docx/parts/endnotes.py`. Grep of `features/` for `footnote`, `endnote`
returns **zero hits**.

Scenarios would cover: `Document.footnotes.add()`, `Footnote.text` read,
`Footnote.delete()`, `Footnotes.__iter__`, restart-numbering + format
attributes on the footnote properties element. One shared `fnt-has-footnotes.docx`
fixture could serve all of A.2/A.3/A.4/A.6.

### 5.2 Phase B — Track changes

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 6 | Phase B.1: Read tracked insertions and deletions | `caff0e6` | No | `trk-read-ins-del.feature` | M |
| 7 | Phase B.2: Accept and reject tracked changes | `25c5951` | No | `trk-accept-reject.feature` | L |
| 8 | Phase B.3: Read formatting track changes (rPrChange, pPrChange, sectPrChange) | `1e7d64a` | No | `trk-format-changes.feature` | M |
| 134 | Move revisions (`w:moveFrom`, `w:moveTo`) | `aef523c` | No | `trk-move-revisions.feature` | M |
| 135 | Cell and row-level tracked changes (`w:cellIns`, `w:cellDel`, `w:trPrChange`, `w:tcPrChange`) | `dfa7daf` | No | `trk-table-changes.feature` | M |
| 136 | Revision IDs (`w:rsid`, `w:rsidRoot`) | `28c05dc` | No | `trk-rsid.feature` | S |
| 163 | Revision marks viewer mode (`revision_marks_text()`) | `18ca8af` | No | `trk-marks-text.feature` | S |

Sources: `src/docx/tracked_changes.py`, `src/docx/oxml/tracked_changes.py`.
Grep of `features/` for `track`, `tracked`, `revision`, `w:ins`, `w:del`
returns **zero hits**.

Accept/reject (#7) is an **L** because it deserves a full matrix of
scenarios — insertion vs deletion, inside-run vs whole-run vs whole-paragraph,
accept-all vs reject-all vs specific-revision — and needs distinct "before"
and "after" fixture pairs to assert round-tripping.

### 5.3 Phase C — Fields / bookmarks / cross-references

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 9 | Phase C.1: Bookmarks — create, read, delete | `19db82e` | No (only "linkedBookmark" in hyperlink fixture) | `bmk-create-read.feature` | M |
| 10 | Phase C.2: Simple and complex field codes | `c708ad5` | No | `fld-simple.feature`, `fld-complex.feature` | M |
| 115 | Cross-references (`REF`/`PAGEREF` resolution) | `42e76f5` | No | `fld-cross-ref.feature` | M |
| 116 | Table of Contents generation | `cdf178c` | No | `toc-generate.feature` | M |

Sources: `src/docx/bookmarks.py`, `src/docx/fields.py`,
`src/docx/oxml/fields.py`, `src/docx/toc.py`. Grep of `features/` for
`bookmark` finds only an example-table value (`linkedBookmark`) inside
`hlk-props.feature`, not bookmark coverage.

### 5.4 Content controls + custom XML

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 27 | Phase D.14: Content controls (structured document tags) | `fda39ef` | No | `sdt-content-controls.feature` | M |
| 131 | Custom XML data binding (`w:dataBinding`) | `c079126` | No | `sdt-data-binding.feature` | M |

Sources: `src/docx/content_controls.py`, `src/docx/oxml/content_controls.py`,
`src/docx/parts/custom_xml.py`. Grep of `features/` for `content.control`,
`sdt`, `custom_xml` returns **zero hits**.

### 5.5 Custom properties

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 14 | Phase D.4: Custom document properties | `adc0485` | No | `doc-customprops.feature` | S |

Sources: `src/docx/custom_properties.py`,
`src/docx/parts/custom_properties.py`. Existing `doc-coreprops.feature`
(3 scenarios) is the closest analogue; a sibling `doc-customprops.feature` with
typed-value round-trip scenarios would fit alongside it. The existing
`coreprops.py` step module is a natural home for the new steps — hence **S**.

### 5.6 Numbering

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 22 | Phase D.9: Numbering style control (restart, custom lists, nested lists) | `738258a` | No (only `num-access-numbering-part.feature`, 1 scenario, read-only) | `num-define-and-apply.feature` | M |

Sources: `src/docx/numbering.py`, `src/docx/oxml/numbering.py`. The existing
`num-access-numbering-part.feature` only asserts that
`document.part.numbering_part` is accessible — it does not exercise
`Numbering.add_numbering_definition` or `NumberingDefinition.apply_to`
(flagged as uncovered in `TEST_AUDIT.md §2`).

### 5.7 Page / section layout

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 19 | Phase D.8: Character spacing and kerning (pre-fork already) | `0bc5c30` | No (kerning has zero hits) | extend `txt-font-props.feature` | S |
| 32 | Phase D.19: Multi-column section layout | `3db6754` | No | extend `sct-section.feature` | S |
| 121 | Page borders (`w:pgBorders`) | `e1e9b69` | No | `sct-page-borders.feature` | S |
| 122 | Line numbering (`w:lnNumType`) | `621eddc` | No | `sct-line-numbering.feature` | S |
| 146 | Paper source (`w:paperSrc`) | `90c5d00` | No | `sct-paper-source.feature` | S |
| 147 | Document grid (`w:docGrid`) | `b190220` | No | `sct-document-grid.feature` | S |
| 148 | Asian typography on section (`w:textDirection`, `w:bidi`) | `429c93a` | No | `sct-text-direction.feature` | S |
| 149 | Section odd-page vs even-page header/footer | `608695b` | No | extend `hdr-header-footer.feature` | S |

Sources: `src/docx/section.py`, `src/docx/oxml/section.py`. All section-level
additions are simple attribute reads/writes that fit the existing
`sct-section.feature`/`hdr-header-footer.feature` step module — **S** across
the board.

### 5.8 Tables — extensions

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 28 | Phase D.15: Row.height — set and get table row height | `819ed67` | No (only height-rule / min-height scenarios in `tbl-row-props`) | extend `tbl-row-props.feature` | S |
| 39 | Phase D.26: Table autofit and column width control | `e5d88e3` | Partial — legacy `Table.autofit` boolean covered in `tbl-props.feature`; new `autofit_behavior`, `preferred_width`, `allow_autofit` not covered | extend `tbl-props.feature` | S |
| 142 | Vertical text direction in cells (`w:textDirection`) | `e0c79cb` | No | `tbl-cell-text-direction.feature` | S |
| 143 | Cell margins per-cell (`w:tcMar`) | `08dcbeb` | No | `tbl-cell-margins.feature` | S |
| 144 | Banded rows / columns (`w:tblLook`) | `2bfe7c4` | No | `tbl-style-flags.feature` | S |
| 145 | Merged cell read robustness (Cell.is_merge_origin / .merge_origin) | `bf44d78` | Partial — `tbl-merge-cells.feature` covers `merge()`; `is_merge_origin`/`merge_origin` getters are not asserted | extend `tbl-merge-cells.feature` | S |

All reuse the existing `table.py` step module and a handful of existing
`tbl-*.docx` fixtures. Most are **S**; #144 may need a new "banded" fixture.

### 5.9 Text / font — extensions

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 33 | Phase D.20: Font.shading — run-level background color | `64ce4aa` | No | extend `txt-font-color.feature` | S |
| 120 | Run border (`w:bdr`) | `bf99398` | No | `txt-run-border.feature` | S |
| 127 | Right-to-left / bidirectional text | `c754772` | No (`bidi` appears only in a table-direction context) | extend `txt-font-props.feature` | S |
| 128 | East Asian typography features | `c754772` / `c0b9f32` | No | `txt-east-asian.feature` | S |
| 129 | Ruby (`w:ruby`) | `c0b9f32` | No | `txt-ruby.feature` | S |
| 160 | Language tags on runs / paragraphs | `74671ce` | No | extend `txt-font-props.feature` | S |

Sources: `src/docx/text/font.py`, `src/docx/oxml/text/font.py`. All are
single-attribute property-matrix tests; all **S**.

### 5.10 Paragraph helpers

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 26 | Phase D.13: Insert paragraph/table at arbitrary position | `50c2078` | No (only `par-insert-paragraph.feature`, which covers *before* an existing paragraph only) | extend `par-insert-paragraph.feature` + `blk-insert.feature` | S |
| 126 | Frames (`w:framePr`) — `paragraph.paragraph_format.frame` | `9924bc2` | No | extend `txt-parfmt-props.feature` | S |

### 5.11 Drawing / shapes

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 30 | Phase D.17: Floating images (non-inline positioning) | `f51e7a9` | No | `shp-floating-images.feature` | M |
| 36 | Phase D.23: Watermark support (text and image) | `0036485` | No | `wmk-watermark.feature` | M |
| 111 | Charts (read + create) | `8f426b4` | No (only the fixture-table string "a chart" in `shp-inline-shape-access.feature` identifying a chart shape) | `chart-read.feature` + `chart-create.feature` | L |
| 112 | SmartArt | `3c04c90` | No | `sart-read.feature` | M |
| 137 | Full shape creation (DrawingML `wps:wsp`) — `Paragraph.add_shape()` | `e0e7b52` | No | `shp-add-preset-shape.feature` | M |
| 138 | Group shapes (`wpg:grpSp`) | `da3e3f1` | No | `shp-group-shapes.feature` | M |
| 139 | Ink annotations (`w:ink`) | `4791c1c` | No | `shp-ink.feature` | S |
| 140 | Embedded object insertion (`w:object`) | `f3b0937` | No | `shp-ole-embed.feature` | M |
| 141 | Chart / picture captions | `e9420f1` | No | `cap-caption.feature` | S |
| 158 | Alt text for images / shapes | `0359f68` | No | `shp-alt-text.feature` | S |

Sources: `src/docx/shape.py`, `src/docx/oxml/shape.py`, `src/docx/charts.py`,
`src/docx/smart_art.py`, `src/docx/captions.py`. Chart create (#111) is the
only **L** because it needs numeric `.docx` fixtures for every category (bar,
line, pie) and the scenarios have to assert both the drawing inline and the
embedded `chart*.xml`.

### 5.12 Headers / footers / settings

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 13 | Phase D.3: Extended document settings | `175d13d` | Partial — `doc-settings.feature` has 3 scenarios covering basic read access; new settings (compat, etc.) not covered | extend `doc-settings.feature` | S |
| 118 | Document background color / watermark color | `71023bd` | No | `doc-background-color.feature` | S |
| 125 | Protection modes beyond read-only | `4e64e3c` | No | `doc-protection.feature` | S |
| 130 | Mail merge directives (`w:mailMerge`) | `6c73ea5` | No | `mmg-mail-merge.feature` | M |
| 133 | Building block gallery categories | `f375d0c` | No | `glo-building-blocks.feature` | M |
| 136 | Revision IDs (`w:rsid`, `w:rsidRoot`) | `28c05dc` | No | (see §5.2 — `trk-rsid.feature`) | S |
| 156 | Compatibility mode flags (`w:compat`) | `4bf4fc4` | No | extend `doc-settings.feature` | S |
| 157 | Web settings (`webSettings.xml`) | `2f99194` | No | `web-web-settings.feature` | S |
| 162 | Style mapping / "keep a style the same as X" — `Style.link_style`, `next_style`, `is_redefined` | `008dcd1` | No | extend `sty-style-props.feature` | S |
| 164 | Draft / normal / outline / print layout hints | `40a8679` | No | extend `doc-settings.feature` | S |

### 5.13 Accessibility

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 158 | Alt text for images / shapes (duplicate with §5.11) | `0359f68` | No | `shp-alt-text.feature` | S |
| 159 | Heading structure validation | `218f756` | No | `acc-heading-structure.feature` | S |
| 161 | Word count / statistics | `5c5cb4b` | No | `doc-statistics.feature` | S |

### 5.14 Search / navigation

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 23 | Phase D.10: Search and replace with formatting preservation | `2bbedd2` | No | `srh-search-replace.feature` | M |
| 153 | Regex search/replace | `eb3e9b7` | No | extend `srh-search-replace.feature` | S |
| 154 | Search across tables / headers / footers / footnotes | `03728a8` | No | extend `srh-search-replace.feature` | M |
| 155 | Stable element IDs (`stable_id` on Paragraph, Run, Table, Cell) | `70fad92` | No | `doc-stable-ids.feature` | S |

### 5.15 Packaging

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 150 | Digital signatures (`_xmlsignatures/`) | `7c67167` | No | `pkg-signatures.feature` | S |
| 151 | Document.xml content recovery (recover=True mode) | `17fa36f` | No | `pkg-recover-mode.feature` | M |
| 152 | Password-encrypted .docx (detection + EncryptedDocumentError) | `68ea68b` | No | `pkg-encrypted-docx.feature` | S |

### 5.16 Other

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 113 | Math / equation (OMML) | `9140728` | No | `equ-equation.feature` | L |
| 114 | Symbols (`w:sym`) | `b4c4f92` | No | extend `run-add-content.feature` | S |
| 117 | Themes (`theme1.xml`) | `342d2d3` | No (only appears as a fixture-table string `'Inspiration Theme Colour 2'` in `txt-font-color.feature`) | `thm-theme.feature` | S |
| 119 | Font table reference (`fontTable.xml`) | `3ee4969` | No | `fnt-font-table.feature` | S |
| 123 | Legacy form fields (`w:ffData`) | no PR number found | No | `frm-form-fields.feature` | M |
| 124 | Rich-text ranges (`w:permStart` / `w:permEnd`) | `0965a08` | No | `prm-perm-ranges.feature` | S |
| 132 | Glossary document (`glossaryDocument.xml`) | `8dd08c8` | No | `glo-glossary.feature` | M |

Equations (#113) is **L** because the OMML create path touches a different
builder (the math-XML namespace) than any existing feature exercises, and
scenarios would want to round-trip several operator types.

### 5.17 Pre-fork phase-D items

Some phase-D issues were resolved by commits that pre-date fork extensions but
remain un-behave-covered in the same way:

| # | Title | Commit | Existing .feature? | Suggested | Effort |
|---:|---|---|:---:|---|:---:|
| 11 | Phase D.1: Hyperlink creation API | `58f27da` | Partial — `hlk-props.feature` covers read; create API not covered | extend `hlk-props.feature` | S |
| 12 | Phase D.2: Comment replies (threaded comments) | `90ef316` | Partial — `cmt-*.feature` covers base comments only, not replies | extend `cmt-mutations.feature` | M |
| 15 | Phase D.5: Table and cell border control | `1fd205c` | No (known to have a production bug — see `TEST_AUDIT.md §3`) | `tbl-borders.feature` | M (blocked on bug) |
| 16 | Phase D.6: Cell shading and background color | `64ce4aa` | Partial — `tbl-style.feature` has 1 shading scenario, coarse | extend `tbl-cell-props.feature` | S |
| 18 | Phase D.7: Paragraph borders | `26c91d9` | No | `par-borders.feature` | M |
| 20 | Page break insert and delete API | `50e2dc2` | Partial — `doc-add-page-break.feature` covers Document.add_page_break; new `Paragraph.add_page_break_before` / `Run.add_break(WD_BREAK.PAGE)` etc. vary | extend `doc-add-page-break.feature` | S |
| 21 | Section break insert and delete API | `527aade` | Partial — `doc-add-section.feature` covers add; no delete | extend `doc-add-section.feature` | S |
| 24 | Phase D.11: Paragraph.delete() and Run.delete() (and Table.delete()) | `90c7c3d` | No | `par-delete.feature`, `run-delete.feature`, `tbl-delete.feature` | S |
| 25 | Phase D.12: Table header row repeat on page break (`is_header`) | `248a932` | No | extend `tbl-row-props.feature` | S |
| 29 | Phase D.16: Row.allow_break_across_pages | `9572f10` | No | extend `tbl-row-props.feature` | S |
| 31 | Phase D.18: Fix run.add_picture() not inserting image | `874c1d5` | Yes — `run-add-picture.feature` updated, is the one fork commit that touched `features/` for a non-comments feature | — | — |
| 34 | Phase D.21: Run splitting at character position | `e432519` | No | `run-split.feature` | S |
| 35 | Phase D.22: SVG image support | `cc8b202` | No | extend `img-characterize-image.feature` + new `run-add-picture` SVG row | S |
| 37 | Phase D.24: .docm macro-enabled file support | `7ca9d2c` | No | `api-docm.feature` | S |
| 38 | Phase D.25: Font.name_far_east — East Asian font support | `3116676` | No | extend `txt-font-props.feature` | S |
| 40 | Phase D.27: DrawingML shapes and text box content access | `362554a` | No | `shp-text-box.feature` | M |
| 41 | Phase D.28: Fix core_properties.last_modified_by making document invalid | `df8f833` | Partial — `doc-coreprops.feature` does not assert round-trip validity | extend `doc-coreprops.feature` | S |

### 5.18 Audit / infrastructure issues (skipped per spec)

Issues #82, #83, and #165 are not feature work and are intentionally excluded
from the audit.

### 5.19 Tally

- **66 feature-delivery issues** surveyed (excluding infra)
- **1 has post-fork behave coverage** (#31 — a one-file tweak to
  `run-add-picture.feature` as part of a bug fix)
- **~11 have partial coverage** because a pre-fork feature file exists in the
  same area but does not exercise the new API (#11, #12, #13, #15, #16, #20,
  #21, #25, #38, #39, #41, #145)
- **~54 have no behave coverage at all**

The 18 comments scenarios (`cmt-props.feature`, `cmt-mutations.feature`,
`doc-add-comment.feature`, `doc-comments.feature`) pre-date the
word-feature-gap issue tracking and therefore do not appear in the matrix.

---

## 6. Fixture-file catalog (`features/steps/test_files/`)

53 files. Column key:
- **References**: step modules under `features/steps/` that name the fixture
  (via `test_docx("name")` or `test_file("name.ext")`). For image fixtures,
  `img-characterize-image.feature` references them via its Examples table, not
  via a step module.

### 6.1 `.docx` fixtures (40 files)

| Filename | Size | Referenced by |
|---|---:|---|
| blk-containing-table.docx | 25 142 | `block.py`, `table.py` |
| blk-paras-and-tables.docx | 15 649 | `block.py` |
| comments-rich-para.docx | 20 023 | `comments.py` |
| doc-access-sections.docx | 25 591 | `document.py`, `section.py` |
| doc-add-section.docx | 17 956 | `document.py` |
| doc-coreprops.docx | 11 992 | `coreprops.py` |
| doc-default.docx | 21 366 | `api.py`, `comments.py` |
| doc-no-coreprops.docx | 11 394 | `coreprops.py` |
| doc-odd-even-hdrs.docx | 17 711 | **ORPHAN** (referenced by neither step nor feature) |
| doc-word-default-blank.docx | 21 309 | `document.py`, `settings.py` |
| fnt-color.docx | 15 846 | `font.py` |
| hdr-header-footer.docx | 18 079 | `hdrftr.py` |
| num-having-numbering-part.docx | 24 334 | `numbering.py` |
| par-alignment.docx | 15 126 | `paragraph.py` |
| par-hlink-frags.docx | 12 071 | `hyperlink.py` |
| par-hyperlinks.docx | 12 385 | `hyperlink.py`, `paragraph.py` |
| par-known-paragraphs.docx | 27 969 | `paragraph.py` |
| par-known-styles.docx | 20 901 | `paragraph.py` |
| par-rendered-page-breaks.docx | 12 244 | `pagebreak.py`, `paragraph.py`, `text.py` |
| run-char-style.docx | 26 645 | `text.py` |
| run-enumerated-props.docx | 14 645 | `text.py` |
| sct-first-page-hdrftr.docx | 14 849 | `section.py` |
| sct-inner-content.docx | 12 051 | `section.py` |
| sct-section-props.docx | 28 168 | `section.py` |
| set-no-settings-part.docx | 10 760 | `settings.py` |
| shp-inline-shape-access.docx | 122 610 | `document.py`, `shape.py` |
| sty-behav-props.docx | 12 195 | `styles.py` |
| sty-having-no-styles-part.docx | 8 358 | `styles.py` |
| sty-having-styles-part.docx | 21 573 | `document.py`, `styles.py` |
| sty-known-styles.docx | 13 560 | `parfmt.py`, `styles.py` |
| tab-stops.docx | 13 170 | `parfmt.py`, `tabstops.py` |
| tbl-2x2-table.docx | 25 129 | `table.py` |
| tbl-cell-access.docx | 36 051 | `table.py` |
| tbl-cell-props.docx | 13 773 | `table.py` |
| tbl-col-props.docx | 13 654 | `table.py` |
| tbl-having-applied-style.docx | 50 294 | `table.py` |
| tbl-having-tables.docx | 32 010 | `document.py` |
| tbl-on-off-props.docx | 16 017 | `table.py` |
| tbl-props.docx | 20 419 | `table.py` |
| txt-font-highlight-color.docx | 12 859 | `font.py` |
| txt-font-props.docx | 37 924 | `font.py` |

### 6.2 Image fixtures (13 files)

| Filename | Size | Referenced by |
|---|---:|---|
| court-exif.jpg | 80 603 | `hyperlink.py` (hlk-props fixture URL), `img-characterize-image.feature` row |
| jfif-300-dpi.jpg | 355 196 | `img-characterize-image.feature` row |
| jpeg420exif.jpg | 768 608 | `img-characterize-image.feature` row |
| lena_std.jpg | 104 428 | `img-characterize-image.feature` row |
| python-icon.jpeg | 3 277 | **ORPHAN** (not cited by any feature or step module — available via `test_file()` only if a scenario passes the literal name) |
| monty-truth.png | 64 276 | `document.py`, `text.py` (run.add_picture), `img-characterize-image.feature` |
| test.png | 146 892 | `hdrftr.py`, `img-characterize-image.feature` |
| lena.tif | 786 572 | `img-characterize-image.feature` row |
| sample.tif | 10 409 | `img-characterize-image.feature` row |
| lena.bmp | 263 222 | `img-characterize-image.feature` row |
| mountain.bmp | 308 280 | `img-characterize-image.feature` row |
| lena.gif | 72 985 | `img-characterize-image.feature` row |

### 6.3 Orphan summary

- **`doc-odd-even-hdrs.docx`** (17 711 bytes) — neither step modules nor
  feature files reference it. It was likely staged for a future odd-even
  headers scenario (see issue #149 — still no behave coverage). Keep or
  delete? — deleting needs only git and costs nothing; keeping it signals
  intent for #149.
- **`python-icon.jpeg`** (3 277 bytes) — orphan; likely a spare for
  `run.add_picture()` exercises that did not land.

Total fixture footprint: ~3.4 MB, dominated by the image files (the TIFF
`lena.tif` alone is 786 KB). None are egregiously large; no LFS needed.

---

## 7. Infrastructure notes

### `features/environment.py`

Minimal. Only `before_all` is defined, and it merely ensures that
`features/_scratch/` exists. There is no `after_scenario`, no `before_tag`, no
`after_all`. If future work introduces tags (e.g. `@slow`, `@fixture-heavy`),
the wiring to gate them goes here.

### Tags

Grep finds no tags in any `.feature` file:

```
$ grep -rE "^\s*@[a-z]" features/*.feature | wc -l
0
```

The behave `-t`/`--tags` facility is therefore unused. There are no `@wip`
scaffolds waiting to be filled in.

### Scratch files

- `.gitignore` includes `_scratch/` — correctly covers both root and
  `features/_scratch/`.
- `git ls-files features/_scratch/` returns nothing. **No scratch files are
  tracked.** The local run-output file `features/_scratch/test_out.docx`
  exists only in the working tree.

### behave configuration

No `behave.ini`, `.behaverc`, `setup.cfg [behave]`, or `pyproject.toml`
`[tool.behave]` table. Everything runs with defaults.

### CI wiring

`uv run behave features/` is documented in `CLAUDE.md` and referenced in
`TEST_AUDIT.md`, but behave is **not run by any GitHub Actions workflow**
under `.github/workflows/` (pytest is). Adding `uv run behave features/` to
the test workflow would cost 2 s and lock the current green baseline in
place — a purely-upside change that is out of scope for this audit but worth
mentioning.

### `.rgignore` / `.fdignore`

Nothing in these files pertains to `features/`.

---

## 8. Step-definition reuse

If a contributor writes a new `.feature`, what's the rough amortised cost of
writing the scenarios versus the steps? The top 10 recurring Given/When/Then
phrases (exact-string match, counts of literal occurrences in `.feature`
files):

| Count | Phrase |
|---:|---|
| 9 | `Given a run` (`text.py:given_a_run`) |
| 8 | `Given a blank document` (`shared.py:step_given_blank_document`) |
| 7 | `Given a Comment object` (`comments.py`) |
| 6 | `Given a Section object as section` (`section.py`) |
| 5 | `Given a font having <type> color` (`font.py`) |
| 4 | `When I merge from cell <origin> to cell <other>` (`table.py`) |
| 4 | `Then the row cells text is <expected-text>` (`table.py`) |
| 4 | `Then the picture appears at the end of the run` (`text.py`) |
| 4 | `Given a paragraph` (`paragraph.py`) |
| 4 | `Given a document having known styles` (`styles.py`) |

The suite makes **heavy use** of `Scenario Outline` + `Examples` plus a small
set of `Given a <thing>` priming steps — 124 outlines among 239 source blocks.
That pattern translates directly to the new-feature opportunities in §5: the
typical cost of writing a new attribute-matrix scenario (e.g. "set
`font.border_color` across eight values") is:

- **Gherkin lines**: ~15 (feature header + one outline + 8-row examples)
- **New steps**: 0-3 (`Given`/`When`/`Then` specific to the property)

because `Given a run` / `Given a paragraph` / `Given a document having known
styles` are already available.

Based on these numbers, most of the **S**-labelled recommendations in §5 will
come in at **3-5 new step definitions and a single fixture extension**, and
land in the existing step module for that domain (e.g. `font.py` for font
extensions, `table.py` for table extensions).

---

## 9. Recommendations (follow-up issue backlog)

Effort labels: **S** ≤ 1 day, **M** 1-3 days, **L** > 3 days.

### 9.1 Quick wins (S — each reuses existing fixtures / steps)

1. **[S] `fnt-font-table.feature`** (#119) — 1 scenario: "Document.font_table
   reports the fonts referenced in fontTable.xml".
2. **[S] `doc-statistics.feature`** (#161) — 1 outline × 3 rows: word count /
   character count / paragraph count on `par-known-paragraphs.docx`.
3. **[S] `acc-heading-structure.feature`** (#159) — 1 outline × 3 rows: valid
   doc, missing-H2, skipped-level.
4. **[S] `shp-alt-text.feature`** (#158) — 1 outline × 4 rows: get, set, clear,
   empty on `shp-inline-shape-access.docx`.
5. **[S] `trk-rsid.feature`** (#136) — 1 outline: rsidRoot, rsid_lst,
   per-paragraph rsid on an rsid-tagged fixture.
6. **[S] Extend `txt-font-props.feature`** for kerning (#19), bidi (#127),
   language (#160), `name_far_east` (#38), and Font.shading (#33) — a single
   commit adding ~25 examples rows to the existing outlines.
7. **[S] Extend `tbl-row-props.feature`** for `Row.height` (#28),
   `allow_break_across_pages` (#29), and `is_header` (#25).
8. **[S] Extend `tbl-props.feature`** for `autofit_behavior`,
   `preferred_width`, `allow_autofit` (#39).
9. **[S] `srh-search-replace.feature`** (#23, #153) — 2 scenarios, regex and
   non-regex, using an existing fixture containing the target text.
10. **[S] `doc-stable-ids.feature`** (#155) — 1 scenario: every paragraph /
    run / table has a stable id across save/load.

Combined effort for these ten items: ~3 dev-days. Would move ~10 of the
coverage gaps from "no behave" to "at least one scenario".

### 9.2 Core missing coverage (M — each justifies its own fixture + feature file)

11. **[M] Footnotes + endnotes** — `fnt-*.feature` (#2, #3, #4, #17) and
    `end-*.feature` (#5). Four fixtures (`fnt-empty.docx`,
    `fnt-has-footnotes.docx`, `end-empty.docx`, `end-has-endnotes.docx`); a
    new `footnotes.py` step module; ~25 scenarios total.
12. **[M] Tracked changes** — `trk-read-*.feature` (#6, #8) and
    `trk-accept-reject.feature` (#7). At least three fixture pairs
    (`trk-simple.docx` + expected-accepted + expected-rejected). New
    `tracked_changes.py` step module; ~30 scenarios.
13. **[M] Bookmarks and fields** — `bmk-create-read.feature` (#9),
    `fld-simple.feature` / `fld-complex.feature` (#10), `fld-cross-ref.feature`
    (#115), `toc-generate.feature` (#116). Share fixtures across the group.
14. **[M] Content controls + data binding** — `sdt-content-controls.feature`
    (#27), `sdt-data-binding.feature` (#131). Leverage existing
    `tests/test_content_controls.py` fixtures as a starting point.
15. **[M] Charts read** — `chart-read.feature` (#111). Start with read-only;
    create goes into §9.3.
16. **[M] Paragraph borders** (#18) — `par-borders.feature` with one
    `par-borders.docx` fixture.
17. **[M] Floating images** (#30) — `shp-floating-images.feature`. Cover the
    anchor-position matrix exercised in `tests/oxml/test_shape.py`.
18. **[M] Watermark** (#36) — `wmk-watermark.feature`. Two scenarios: text
    watermark, image watermark.

### 9.3 Larger efforts (L — multi-day)

19. **[L] Tracked-changes accept/reject matrix (#7)** — see item 12 above;
    the `trk-accept-reject.feature` file alone justifies an L budget.
20. **[L] Charts create (#111)** — needs numeric fixtures for bar/line/pie,
    output-validation steps that compare the generated `chart1.xml`.
21. **[L] Math / equation (#113)** — OMML builder is a separate namespace
    tree; may warrant a new `equations.py` step module.
22. **[L] Mail merge (#130) + glossary (#132)** — both require bespoke
    fixtures and likely a new step module each. They sit at the rim of the
    feature set and can wait until the core gaps are addressed.

### 9.4 Infrastructure hygiene (S)

23. **[S] Wire `uv run behave features/` into `.github/workflows/` alongside
    pytest.** The suite takes 2 s and is stable; adding it locks in the
    current green state at negligible cost. (Out of scope for this
    report-only audit but trivially cheap.)
24. **[S] Delete orphan fixture `doc-odd-even-hdrs.docx`** (or use it to
    implement #149 and remove the orphan flag). Same for `python-icon.jpeg`.
25. **[S] Introduce at least one tag (`@slow`, `@fixture-heavy`)** and extend
    `environment.py` with an `after_scenario` hook that cleans up scratch
    files. The scratch directory is already gitignored and in practice
    contains only one file, so this is primarily forward-looking.
26. **[S] Add a module-level note in `features/environment.py`** pointing
    readers to the conventions in `features/steps/helpers.py`
    (`test_docx()`, `test_file()`, `saved_docx_path`). New contributors
    currently have to read the helpers module to discover them.

---

## 10. Relationship to `TEST_AUDIT.md` and `DOCS_AUDIT.md`

Three audits share overlapping territory. `TEST_AUDIT.md` covers the pytest
unit suite in depth and mentions behave only briefly (two lines in §1, and
items 12-14 in §10's recommendations proposing `.feature` files for
footnotes/endnotes/tracked-changes/numbering/TOC/watermark/fields — which this
audit expands into §5.1–§5.16). `DOCS_AUDIT.md` (being written in parallel)
covers the Sphinx reference under `docs/`. **`FEATURES_AUDIT.md` (this
document)** covers behave in depth.

Cross-references:

- Closing the **~54 missing `.feature`** entries catalogued in §5 would
  retire recommendations **12-14 of `TEST_AUDIT.md`** (the only behave items
  on that list) and remove the "behave coverage has not kept pace" caveat
  from its §1 framing.
- Issue **#165** (`WD_BORDER_STYLE`/`CT_Border` duplicate — a production
  bug) is called out as the only *correctness* blocker in
  `TEST_AUDIT.md §3` and surfaces in §5.17 here as a block on the
  `tbl-borders.feature` recommendation (#15). Both audits agree: fix the
  bug first.
- `DOCS_AUDIT.md` will likely identify the same ~55 undocumented features,
  since a feature shipping without docs tends to ship without behave. The
  two audits' Quick-wins lists should be treated as complementary: each
  behave scenario written for §9.1 doubles as a small worked example the
  docs audit can reference, and each `.rst` page added by the docs work
  makes behave scenarios easier to write.

A combined, single follow-up issue per feature (docstring + `.feature` file +
API doc page) is probably the most efficient unit of remediation work.
