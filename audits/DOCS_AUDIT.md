# Sphinx Documentation Audit

This report surveys the state of the Sphinx-based reference docs under `docs/`
against the current `loadfix/python-docx` source tree at commit `50c2078`
(`master`). It is a companion to `TEST_AUDIT.md`. The goal is to document what
is currently shipped, what has drifted since upstream, and to give a
prioritised punch-list of concrete follow-ups.

This report is **advisory only** — no `.rst`, `.py`, or workflow files were
modified while producing it.

---

## 1. Summary

- `docs/` contains **75 reStructuredText files, 12 005 lines total**
  (`find docs -name '*.rst' | xargs wc -l`). Of those:
  - 11 API-reference pages under `docs/api/` (538 lines, `document.rst` the
    biggest at 117)
  - 16 enum pages + 1 index under `docs/api/enum/` (982 lines)
  - 12 user-guide pages under `docs/user/` (2337 lines)
  - 36 developer analysis pages under `docs/dev/analysis/` (8148 lines) —
    these are pre-existing XSD / feature analyses, largely untouched since
    upstream.
- Build system: **Sphinx with the `armstrong` HTML theme** vendored under
  `docs/_themes/armstrong` (`docs/conf.py:236`). The CLAUDE.md note saying
  the theme is `alabaster` is **incorrect** — `alabaster` appears only in
  `requirements-docs.txt:3` as a pin that is never read because the active
  theme is `armstrong`.
- Configuration: `docs/conf.py` — Sphinx 1.0-style, pinned to
  `Sphinx==1.8.6` / `Jinja2==2.11.3` / `MarkupSafe==0.23` in
  `requirements-docs.txt`. None of those install on Python ≥3.10 without
  patches.
- Last meaningful doc-tree commit: **`4fbe1f6` "docs: add Comments docs"
  (2025-06-11)**, five days before upstream's `1.2.0` release. There have
  been **zero docs commits since the fork began shipping the Phase A / B / C
  / D feature additions** in this fork (`git log --since=2025-06-12 --
  docs/` is empty). Every one of the ~55 new features (fork-specific
  commits `#1`..`#165`) relies exclusively on docstrings for discovery.
- HISTORY.rst (repo root) stops at `1.2.0 (2025-06-16)`. It contains no
  entries for any fork phase — Phase A, B, C, or D.

### Top-3 findings

1. **24 new proxy modules ship with zero `docs/api/*.rst` coverage.** Every
   module listed in section 4.1 below (`accessibility`, `captions`, `chart`,
   `content_controls`, `custom_properties`, `custom_xml`, `embedded_objects`,
   `equations`, `fields`, `font_table`, `form_fields`, `glossary`, `ids`,
   `ink`, `numbering`, `permissions`, `ruby`, `signatures`, `smart_art`,
   `statistics`, `theme`, `toc`, `tracked_changes`, `watermark`,
   `web_settings`, plus `bookmarks`, `footnotes`, `endnotes`, `search`) is
   invisible to an autodoc build — they are neither imported by
   `docs/api/*.rst` nor wired into `docs/index.rst`'s toctree.
2. **`docs/conf.py:69-203` `rst_epilog` is the single-largest source of
   build errors.** A non-strict Sphinx 6 build emits **101 warnings**, **96 of
   which are `Undefined substitution referenced:`** for 53 distinct
   `|ClassName|` substitutions that were never added to `rst_epilog` when
   the new proxy modules landed. These show up in almost every rendered
   docstring — for example, 27 `Document.*` properties render "Undefined
   substitution referenced" boxes instead of class cross-references.
3. **Zero user-guide narrative for any fork feature.** Of the thirty-plus
   feature areas shipped in Phases A/B/C/D (tracked changes, footnotes,
   endnotes, fields, bookmarks, content controls, form fields, numbering,
   captions, TOC, watermarks, …), only one — comments — got a
   `docs/user/*.rst` page. That page was delivered by upstream.

### Sphinx build result

A build with **Sphinx 6.2.1** (the oldest version that still installs on
Python 3.11) completes with exit-status 0 but emits 101 warnings when
`conf.py` is left untouched. The strict `-W` variant fails at the
configuration step because `intersphinx_mapping`'s value is in the
pre-Sphinx-1.0 format. Concrete numbers are in section 3.

---

## 2. Docs layout

| Path | Purpose |
|---|---|
| `docs/index.rst` (114 lines) | Landing page: the "What it can do" code sample and the three top-level toctrees (User Guide, API Documentation, Contributor Guide). |
| `docs/conf.py` (394 lines) | Sphinx config. Theme, extensions (`autodoc`, `intersphinx`, `todo`, `coverage`, `viewcode`), and the big `rst_epilog` substitutions block (lines 69-203). |
| `docs/user/*.rst` (12 files, 2337 lines) | Narrative user guide — quickstart, install, documents, tables, text, sections, headers/footers, api-concepts, styles-understanding, styles-using, comments, shapes. |
| `docs/api/*.rst` (10 top-level pages, 538 lines) | API reference. One page per major module group: `document`, `settings`, `style`, `text`, `table`, `section`, `comments`, `shape`, `dml`, `shared`. Typically `.. autoclass:: Foo :members:`. |
| `docs/api/enum/*.rst` (16 enum pages + `index.rst`, 982 lines) | Hand-written enum reference pages (title, alias, intro, flat list of members). Not autogenerated — `:ref:` targets in docstrings point at these. |
| `docs/dev/analysis/*.rst` (36 files, ~8100 lines) | XSD / feature analyses written during upstream development. Not user-facing; linked from the "Contributor Guide" toctree. |
| `docs/_themes/armstrong/` | Vendored HTML theme. A fork of the old "armstrong" sidebar theme. |
| `docs/_static/img/` | Four PNGs: `comment-parts.png`, `example-docx-01.png`, `hdrftr-01.png`, `hdrftr-02.png`. All four are referenced and present on disk (no broken `.. image::` links). |
| `docs/_templates/` | Empty (the directory is registered in `conf.py:42` but nothing lives there). |

---

## 3. Build health

### 3.1 Prerequisites

`requirements-docs.txt` pins:

```
Sphinx==1.8.6
Jinja2==2.11.3
MarkupSafe==0.23
alabaster<0.7.14
-e .
```

None of these install on modern Python (`MarkupSafe==0.23` fails with
`ImportError: cannot import name 'Mapping' from 'collections'` on 3.10+).
On Read-the-Docs it builds because `.readthedocs.yaml` still targets
`python: "3.9"` — but a local `make html` on any modern checkout fails
before the first file is read.

### 3.2 Build command and result

Running with a modern Sphinx 6:

```
python -m sphinx -b html docs docs/_build/html
```

- **Exit status: 0** (build succeeded).
- **Warnings emitted: 101.**

Strict mode (`-W`) fails immediately:

```
Failed to read intersphinx_mapping[http://docs.python.org/3/], ignored:
SphinxWarning('The pre-Sphinx 1.0 intersphinx_mapping format is deprecated
and will be removed in Sphinx 8. [...]')
```

Source of the problem: `docs/conf.py:393`

```python
intersphinx_mapping = {"http://docs.python.org/3/": None}
```

The modern format is `intersphinx_mapping = {"python": ("http://docs.python.org/3/", None)}`.

### 3.3 Warning breakdown

| Count | Category |
|---:|---|
| 96 | `ERROR: Undefined substitution referenced:` — a `|SomeClass|` substitution used in a docstring that has no entry in `conf.py`'s `rst_epilog` (see 3.4 below) |
| 3 | `WARNING: undefined label:` — `wdtextdirection` (×2, `docx.section.Section.text_direction`, `docx.table._Cell.text_direction`), `wdborderstyle` (×1, `docx.text.run.Font.border_style`). No `docs/api/enum/WdTextDirection.rst` / `WdBorderStyle.rst` exist. |
| 1 | `ERROR: Unknown target name: "container"` — `docs/user/comments.rst:134` uses Markdown-style `_block-item container_` emphasis; Sphinx parses the second `_container_` as a reference target. |
| 1 | `WARNING: The pre-Sphinx 1.0 'intersphinx_mapping' format is deprecated` — `docs/conf.py:393` |

### 3.4 First ten warnings (representative)

```
WARNING: The pre-Sphinx 1.0 'intersphinx_mapping' format is deprecated [...]
ERROR: Undefined substitution referenced: "EndnoteProperties"
    (src/docx/document.py::Document.add_endnote_properties)
ERROR: Undefined substitution referenced: "FootnoteProperties"
    (src/docx/document.py::Document.add_footnote_properties)
ERROR: Undefined substitution referenced: "Bookmarks"
    (src/docx/document.py::Document.bookmarks)
ERROR: Undefined substitution referenced: "Chart"
    (src/docx/document.py::Document.charts)
ERROR: Undefined substitution referenced: "ContentControl"
    (src/docx/document.py::Document.content_controls)
ERROR: Undefined substitution referenced: "CustomProperties"
    (src/docx/document.py::Document.custom_properties)
ERROR: Undefined substitution referenced: "CustomXmlPart"
    (src/docx/document.py::Document.custom_xml_parts)
ERROR: Undefined substitution referenced: "EmbeddedObject"
    (src/docx/document.py::Document.embedded_objects)
ERROR: Undefined substitution referenced: "Endnotes"
    (src/docx/document.py::Document.endnotes)
```

### 3.5 Missing `|Name|` substitutions

The full set of substitutions referenced in docstrings but not declared in
`conf.py`'s `rst_epilog` (alphabetical):

```
Bookmarks, CellBorders, CellMargins, CellShading, Chart, ContentControl,
CustomProperties, CustomXmlPart, DocumentGrid, DocumentStatistics, Drawing,
EastAsianLayout, EmbeddedObject, EndnoteProperties, Endnotes, Equation,
Field, FloatingImage, FontTable, FootnoteProperties, Footnotes,
FormattingChange, FormField, Glossary, HeadingIssue, InkAnnotation, Level,
LineNumbering, MailMerge, MoveRevision, Numbering, PageBorder, PageBorders,
ParagraphBorders, PermissionRange, RubyAnnotation, SearchMatch,
SectionColumns, SignatureInfo, SmartArt, Symbol, TableBorders,
TableStyleFlags, TextFrame, Theme, TrackedChange, Watermark, WD_ANCHOR_H,
WD_ANCHOR_V, WD_BORDER_STYLE, WD_TABLE_AUTOFIT, WD_VIEW, WD_WRAP_TYPE,
WebSettings
```

53 symbols total. Adding them to `rst_epilog` is a mechanical change —
one `.. |X| replace:: :class:`.X`` line each — and will immediately
retire 96 of the 101 build warnings.

### 3.6 Broken / stale refs

- `:ref:`wdtextdirection`` referenced from `Section.text_direction` and
  `_Cell.text_direction` — no `docs/api/enum/WdTextDirection.rst` page.
- `:ref:`wdborderstyle`` referenced from `Font.border_style` — no
  `docs/api/enum/WdBorderStyle.rst` page.
- `docs/user/comments.rst:134` uses Markdown-style italic
  `_block-item container_`, which Sphinx interprets as a malformed
  reference. (One-character RST fix: replace with `*block-item container*`
  or quote it.)

### 3.7 Duplicated labels

None detected by this build.

### 3.8 Cleanup

`docs/_build/` was removed after the build completed.

---

## 4. API reference gaps

### 4.1 New proxy modules with **no** `docs/api/*.rst` page

Every module below exists in `src/docx/` as of `50c2078`, but none is
referenced from any page under `docs/api/`. Verified with
`grep -rn 'docx\.<module>' docs/api/`, which returns zero matches for each.

| Module | Issue(s) | Public surface | Suggested rst | Summary |
|---|---|---|---|---|
| `src/docx/accessibility.py` | #159 | `HeadingIssue` (class, ln 32); `validate_heading_structure` (fn, ln 64) | `docs/api/accessibility.rst` | Heading-structure validator: flags skipped levels, multiple `Heading 1`s, empty headings. |
| `src/docx/bookmarks.py` | #52, #82 | `Bookmarks` (ln 14), `Bookmark` (ln 42) | `docs/api/bookmarks.rst` | `w:bookmarkStart` / `w:bookmarkEnd` create / read / delete. |
| `src/docx/captions.py` | #141 | `new_caption_paragraph` (fn, ln 39) | `docs/api/captions.rst` | Builds a `SEQ`-field caption paragraph styled `Caption`. Module is function-only. |
| `src/docx/chart.py` | #111 | `WD_CHART_TYPE` (enum, ln 20), `ChartSeries` (ln 73), `Chart` (ln 96), `_chart_type_for` (private) | `docs/api/chart.rst` | Read + minimal create for embedded charts. |
| `src/docx/content_controls.py` | #27, #131 | `ContentControlType` (enum, ln 16), `ContentControl` (ln 68), `DataBinding` (ln 263), `new_sdt` (fn, ln 330) | `docs/api/content-controls.rst` | Structured document tags (rich-text, plain-text, date, checkbox, combo, dropdown, picture). |
| `src/docx/custom_properties.py` | #14, #82 | `CustomProperties` (ln 29) | `docs/api/custom-properties.rst` | `docProps/custom.xml` typed name/value pairs. |
| `src/docx/custom_xml.py` | #131 | `CustomXmlPart` (ln 33), `iter_custom_xml_parts` (fn, ln 129) | `docs/api/custom-xml.rst` | Custom XML data parts backing data-bound SDTs. |
| `src/docx/embedded_objects.py` | #140 | `EmbeddedObject` (ln 27) | `docs/api/embedded-objects.rst` | Read-only OLE object info (Excel, equations, …). |
| `src/docx/endnotes.py` | #17, #82, #96 | `Endnotes` (ln 19), `Endnote` (ln 58), `EndnoteProperties` (ln 139) | `docs/api/endnotes.rst` | Mirror of the footnotes API. |
| `src/docx/equations.py` | #113 | `Equation` (ln 36), `build_identifier`, `build_fraction`, `build_superscript`, `build_subscript`, `build_radical` (builder fns, ln 123-197) | `docs/api/equations.rst` | Read OMML; minimal build helpers. |
| `src/docx/fields.py` | #10, #115 | `WD_FIELD_TYPE` (ln 33), `Field` (ln 58) | `docs/api/fields.rst` | Simple + complex field codes; REF / PAGEREF resolution. |
| `src/docx/font_table.py` | #119 | `FontTable` (ln 22), `FontMetadata` (ln 68) | `docs/api/font-table.rst` | Read-only `word/fontTable.xml`. |
| `src/docx/footnotes.py` | #3, #17, #46, #48, #56, #82 | `Footnotes` (ln 19), `Footnote` (ln 58), `FootnoteProperties` (ln 139) | `docs/api/footnotes.rst` | High-level footnotes API. |
| `src/docx/form_fields.py` | #123 | `WD_FORM_FIELD_TYPE` (ln 38), `TextInputFormField` (ln 96), `CheckboxFormField` (ln 127), `DropdownFormField` (ln 151), `FormField` (ln 183) | `docs/api/form-fields.rst` | Legacy `w:ffData` form fields. |
| `src/docx/glossary.py` | #132, #133 | `Glossary` (ln 36), `BuildingBlock` (ln 154), `BuildingBlockCategory` (ln 243) | `docs/api/glossary.rst` | Read-only glossary document (AutoText / Quick Parts / cover pages). |
| `src/docx/ids.py` | #155 | `compute_stable_id` (fn, ln 56) — the API is `stable_id` on Paragraph, Run, Table, Cell | `docs/api/stable-ids.rst` | Pragmatic mostly-stable identifiers via `w:rsidR` + element-hash. |
| `src/docx/ink.py` | #139 | `InkAnnotation` (ln 25) | `docs/api/ink.rst` | Read-only ink / stylus annotations (InkML). |
| `src/docx/numbering.py` | #22, #82, #108 | `Numbering` (ln 119), `NumberingDefinition` (ln 221), `Level` (ln 269) | `docs/api/numbering.rst` | List-numbering control: restart, custom definitions, nested levels, `apply_to`. |
| `src/docx/permissions.py` | #124 | `PermissionRange` (ln 13) | `docs/api/permissions.rst` | `w:permStart` / `w:permEnd` ranges. |
| `src/docx/ruby.py` | #129 | `RubyAnnotation` (ln 13) | `docs/api/ruby.rst` | Read-only ruby (phonetic) annotations. |
| `src/docx/search.py` | #82, #91, #153, #154 | `SearchMatch` (ln 17), `search_paragraphs`, `replace_in_paragraphs`, `search_paragraphs_regex`, `replace_in_paragraphs_regex` (fns at ln 113, 147, 246, 281) | `docs/api/search.rst` | Text search / replace with regex + all-stories variants. |
| `src/docx/signatures.py` | #150 | `SignatureInfo` (ln 32) | `docs/api/signatures.rst` | Detection + minimal metadata for digital signatures (no verify). |
| `src/docx/smart_art.py` | #112 | `SmartArtNode` (ln 41), `SmartArt` (ln 85), `smart_art_for_drawing` (fn, ln 223) | `docs/api/smart-art.rst` | Read-only SmartArt diagram detection + node text. |
| `src/docx/statistics.py` | #161 | `DocumentStatistics` (NamedTuple, ln 28), `compute_statistics` (fn, ln 47) | `docs/api/statistics.rst` | Word / character / paragraph counts matching Word's dialog. |
| `src/docx/theme.py` | #117 | `Theme` (ln 48), `ThemeColors` (ln 89), `ThemeFonts` (ln 199) | `docs/api/theme.rst` | Read-only `word/theme/theme1.xml` access. |
| `src/docx/toc.py` | #116 | `build_toc_instruction` (fn, ln 103), `populate_toc_paragraph` (fn, ln 151) — public surface is `Document.add_table_of_contents()` | `docs/api/toc.rst` | TOC field builder. |
| `src/docx/tracked_changes.py` | #7, #8, #53, #134, #135, #163 | `TrackedChange` (ln 30), `MoveRevision` (ln 97), `FormattingChange` (ln 160) | `docs/api/tracked-changes.rst` | Insertions, deletions, moves, formatting changes; accept / reject; cell / row changes. |
| `src/docx/watermark.py` | #36 | `Watermark` (ln 17) | `docs/api/watermark.rst` | Read side of text / image watermarks. Section.add_text_watermark / add_image_watermark live on Section. |
| `src/docx/web_settings.py` | #157 | `WebSettings` (ln 27) | `docs/api/web-settings.rst` | Read-only `word/webSettings.xml`. |

29 modules, no page each. All would need to be added to the
`docs/index.rst` "API Documentation" toctree once the `.rst` stubs exist.

### 4.2 Existing API pages that are stale

Most existing API pages use `.. autoclass:: X :members:`, which means new
methods on already-documented classes appear automatically. The gaps
below are new **top-level** proxies that share a module with an already-
documented class but have no directive pointing at them, plus a handful
of new narrative sections that an autogenerated page will not cover.

#### `docs/api/document.rst` (117 lines)

Uses `:members:` on `docx.document.Document` so new methods auto-render.
What is missing is the **return-type substitutions** (see 3.5) and the
human-written narrative on the new return objects. New `Document`
properties / methods added in this fork that now render as raw
`|Name|` placeholders:

- `Document.bookmarks` (`src/docx/document.py:326`)
- `Document.charts` (ln 333), `Document.add_chart(...)` (Phase D charts)
- `Document.content_controls` (ln 385)
- `Document.endnotes` (ln 394), `Document.add_endnote_properties` (ln 529),
  `Document.endnote_properties`
- `Document.equations` (ln 399)
- `Document.form_fields` (ln 420)
- `Document.signatures` (ln 455)
- `Document.font_table` (ln 466)
- `Document.footnotes` (ln 477), `Document.add_footnote_properties` (ln 516),
  `Document.footnote_properties`
- `Document.custom_properties` (ln 539)
- `Document.custom_xml_parts` (ln 549)
- `Document.numbering` (ln 561)
- `Document.ink_annotations` (ln 571)
- `Document.embedded_objects` (ln 588)
- `Document.smart_art` (ln 608)
- `Document.replace` / `replace_all` / `replace_regex` / `replace_regex_all`
  (ln 667-761)
- `Document.search` / `search_all` / `search_regex` / `search_regex_all`
  (ln 829-895)
- `Document.validate_heading_structure` (ln 916)
- `Document.statistics` (ln 934)
- `Document.glossary` (ln 957)
- `Document.theme` (ln 970)
- `Document.web_settings` (ln 981)
- `Document.add_table_of_contents` (ln 261)

Additionally, these exist but do not appear in the API pages because no
narrative points at them and their return types have no `.. autoclass::`
elsewhere. Even when autodoc picks them up, a reader cannot click through.

#### `docs/api/settings.rst` (13 lines)

Page is **one autoclass line** (`docs/api/settings.rst:9`) with
`:inherited-members:`. New proxy types exposed from `Settings` but with no
dedicated section:

- `DocumentProtection` (`src/docx/settings.py:446`) — Phase D.3, #125
- `CompatSettings` (ln 674), `CompatFlags` (ln 754) — #156
- `MailMerge` (ln 847) — #130
- `WD_VIEW` (enum) — #164 — `Settings.view` references this, but
  `docs/api/enum/WdView.rst` does not exist.
- `WD_PROTECTION` (enum) — referenced by `DocumentProtection.protection_type`;
  no enum page.

These need their own `.. autoclass::` blocks so the "Document Protection"
and "Mail Merge" object surfaces are reachable.

#### `docs/api/section.rst` (41 lines)

Uses `:members:` on `Section` so instance methods auto-render, but the
following **new proxy classes** referenced from `Section` have no
`.. autoclass::` block anywhere:

- `Column` (`src/docx/section.py:824`) — #60
- `SectionColumns` (ln 855) — `Section.columns` returns this
- `PageBorder` (ln 940), `PageBorders` (ln 1046) — #121
- `LineNumbering` (ln 1125) — #122
- `DocumentGrid` (ln 1187) — #147

New `Section` methods that do auto-render, but with broken substitution
/ enum refs:

- `Section.columns`, `Section.set_columns(...)` (Phase D.19)
- `Section.page_borders`, `Section.set_page_border(...)`,
  `Section.remove_page_borders(...)` (ln 294-342)
- `Section.line_numbering`, `Section.set_line_numbering`,
  `Section.remove_line_numbering` (ln 343-385)
- `Section.first_page_paper_source`, `Section.other_pages_paper_source`
  (ln 387-441) — #146
- `Section.document_grid`, `Section.set_document_grid`,
  `Section.remove_document_grid` (ln 443-482) — #147
- `Section.right_to_left`, `Section.text_direction` (ln 494-521)
- `Section.footnote_properties` / `.add_footnote_properties` /
  `.remove_footnote_properties` (ln 548-575)
- `Section.endnote_properties` / `.add_endnote_properties` /
  `.remove_endnote_properties` (ln 578-605)
- `Section.add_text_watermark(...)`, `Section.add_image_watermark(...)`,
  `Section.remove_watermark`, `Section.watermark` (ln 609-763)
- `Section.different_odd_and_even_pages_header_footer`,
  `Section.different_first_page_header_footer` (ln 78-127)
- `Section.first_page_header` / `.first_page_footer` /
  `.even_page_header` / `.even_page_footer` (ln 129-168)
- `Section.formatting_change` (ln 91)

#### `docs/api/table.rst` (55 lines)

Uses `:members:` + `:inherited-members:`, so `Table` / `_Cell` / `_Row`
method additions auto-render. Missing `.. autoclass::` blocks for new
proxy types in the same module:

- `CellShading` (`src/docx/table.py:832`) — #63
- `BorderElement` (ln 898), `TableBorders` (ln 997) — #102
- `TableStyleFlags` (ln 1068) — #144
- `CellBorders` (ln 1148), `CellMargins` (ln 1204) — #102, #143

New methods that render with undefined substitutions:

- `Table.borders`, `_Cell.borders`, `_Cell.margins`, `_Cell.shading`,
  `Table.style_flags`, `Table.autofit` setter (#39), `Table.column_width`
  helpers, `_Cell.is_merge_origin` / `.merge_origin` (#145),
  `_Row.height` / `.height_rule` (#28), `_Row.allow_break_across_pages`
  (#51), `_Row.is_header` (#93), `_Row.grid_cols_before` /
  `grid_cols_after`, `_Cell.grid_span`, `_Cell.text_direction` (#142),
  `CT_Tc.grid_offset` (low-level).

#### `docs/api/text.rst` (63 lines)

Uses `:members:` on all proxy classes. Missing `.. autoclass::` for new
types in the same module:

- `docx.text.run._Text` (`src/docx/text/run.py:377`) — internal but
  referenced from the `|_Text|` substitution in `conf.py:198`
- `docx.text.symbol.Symbol` (`src/docx/text/symbol.py:11`) — #114
- `docx.text.font.EastAsianLayout` (`src/docx/text/font.py:889`) — #128
- `docx.text.parfmt.ParagraphBorders` (`src/docx/text/parfmt.py:452`),
  `docx.text.parfmt.Border` (ln 492), `docx.text.parfmt.TextFrame`
  (ln 602) — #126, Phase D.7

Auto-rendering but with broken substitutions / labels:

- `Font.shading_color` (#20 / #33), `Font.border_*` (#120),
  `Font.language` / `east_asian_language` / `bidi_language` (#160),
  `Font.character_spacing`, `Font.kerning` (#19 / #95),
  `Font.highlight_color`
- `Run.add_symbol`, `Run.symbols` (#114)
- `Run.split(offset)` (#34 / #94)
- `Run.bidi`, `Paragraph.bidi` (#127)
- `ParagraphFormat.frame` (#126)
- `Paragraph.insert_paragraph_before`, `Paragraph.delete`, `Run.delete`,
  `Table.delete` (#50)
- `Paragraph.stable_id`, `Run.stable_id` (#155)
- `Paragraph.rsid`, `Run.rsid` (#136)

#### `docs/api/shape.rst` (31 lines)

Missing `.. autoclass::` for:

- `docx.shape.FloatingImage` (`src/docx/shape.py:132`) — #30
- Any of the WD_ANCHOR / WD_WRAP enum pages (all missing —
  see section 5).

`InlineShape` declares `:members: height, type, width` — an explicit
allowlist — so `InlineShape.alt_text`, `.title` (#158), and any future
additions will *not* render.

#### `docs/api/comments.rst` (27 lines)

Up to date for 1.2.0 — the only fork-scope question is whether
`CommentReplies` (Phase D.2, #67) now merits its own subsection. `Comments`
and `Comment` both use `:members:` so new methods auto-render, but
thread-reply narrative is missing.

### 4.3 `docs/api/style.rst` minor gaps

`docs/api/style.rst:21` uses `:members:` on `BaseStyle` / subclasses. New
style attrs from #162 (`Style.link_style`, `Style.next_style`,
`Style.is_redefined`) auto-render but have no narrative section. No missing
classes here.

---

## 5. Enum coverage gaps

`docs/api/enum/` contains 16 hand-written pages (see section 2). Each page
is short (20-70 lines). `src/docx/enum/*.py` defines **37 enum classes**
(listed below). 16 have a page, **21 do not**.

### 5.1 Enums present in `src/docx/enum/` (37 total)

From `grep -E '^class (WD|MSO)_' src/docx/enum/*.py`:

| Enum | File:line | Doc page | Status |
|---|---|---|---|
| `MSO_COLOR_TYPE` | `dml.py:6` | `MsoColorType.rst` | covered |
| `MSO_THEME_COLOR_INDEX` | `dml.py:30` | `MsoThemeColorIndex.rst` | covered |
| `WD_ALIGN_PARAGRAPH` (alias of `WD_PARAGRAPH_ALIGNMENT`) | `text.py:10` | `WdAlignParagraph.rst` | covered |
| `WD_ANCHOR_H` | `shape.py:33` | — | **missing** |
| `WD_ANCHOR_V` | `shape.py:45` | — | **missing** |
| `WD_BORDER_DISPLAY` | `section.py:6` | — | **missing** |
| `WD_BORDER_OFFSET_FROM` | `section.py:29` | — | **missing** |
| `WD_BORDER_STYLE` | `text.py:274` | — | **missing** (referenced by `Font.border_style`) |
| `WD_BREAK_TYPE` | `text.py:70` | — | **missing** |
| `WD_BUILDING_BLOCK_GALLERY` | `text.py:752` | — | **missing** |
| `WD_BUILTIN_STYLE` | `style.py:6` | `WdBuiltinStyle.rst` | covered |
| `WD_CELL_VERTICAL_ALIGNMENT` | `table.py:6` | `WdCellVerticalAlignment.rst` | covered |
| `WD_COLOR_INDEX` | `text.py:92` | `WdColorIndex.rst` | covered |
| `WD_DOC_GRID_TYPE` | `section.py:74` | — | **missing** |
| `WD_DRAWING_TYPE` | `shape.py:22` | — | **missing** |
| `WD_ENDNOTE_POSITION` | `text.py:564` | — | **missing** |
| `WD_FOOTNOTE_POSITION` | `text.py:547` | — | **missing** |
| `WD_FOOTNOTE_RESTART` | `text.py:531` | — | **missing** |
| `WD_FRAME_DROP_CAP` | `text.py:689` | — | **missing** |
| `WD_FRAME_H_ALIGN` | `text.py:705` | — | **missing** |
| `WD_FRAME_H_ANCHOR` | `text.py:632` | — | **missing** |
| `WD_FRAME_V_ALIGN` | `text.py:727` | — | **missing** |
| `WD_FRAME_V_ANCHOR` | `text.py:648` | — | **missing** |
| `WD_FRAME_WRAP` | `text.py:664` | — | **missing** |
| `WD_HEADER_FOOTER_INDEX` | `section.py:108` | — | **missing** |
| `WD_INLINE_SHAPE_TYPE` | `shape.py:6` | — | **missing** |
| `WD_LINE_NUMBERING_RESTART` | `section.py:49` | — | **missing** |
| `WD_LINE_SPACING` | `text.py:159` | `WdLineSpacing.rst` | covered |
| `WD_MAIL_MERGE_DATA_TYPE` | `text.py:944` | — | **missing** |
| `WD_MAIL_MERGE_DESTINATION` | `text.py:925` | — | **missing** |
| `WD_MAIL_MERGE_TYPE` | `text.py:900` | — | **missing** |
| `WD_NUMBER_FORMAT` | `text.py:476` | — | **missing** |
| `WD_ORIENTATION` | `section.py:132` | `WdOrientation.rst` | covered |
| `WD_PARAGRAPH_ALIGNMENT` | `text.py:10` | `WdAlignParagraph.rst` | covered (as alias) |
| `WD_PROTECTION` | `text.py:605` | — | **missing** |
| `WD_ROW_HEIGHT_RULE` | `table.py:51` | `WdRowHeightRule.rst` | covered |
| `WD_SECTION_START` | `section.py:157` | `WdSectionStart.rst` | covered |
| `WD_SHADING_PATTERN` | `table.py:110` | — | **missing** |
| `WD_SHAPE` | `shape.py:74` | — | **missing** |
| `WD_STYLE_TYPE` | `style.py:426` | `WdStyleType.rst` | covered |
| `WD_TAB_ALIGNMENT` | `text.py:208` | `WdTabAlignment.rst` | covered |
| `WD_TAB_LEADER` | `text.py:247` | `WdTabLeader.rst` | covered |
| `WD_TABLE_ALIGNMENT` / `WD_ROW_ALIGNMENT` alias | `table.py:85` | `WdRowAlignment.rst` | covered (shared) |
| `WD_TABLE_AUTOFIT` | `table.py:249` | — | **missing** |
| `WD_TABLE_DIRECTION` | `table.py:348` | `WdTableDirection.rst` | covered |
| `WD_TEXT_DIRECTION` | `table.py:287` | — | **missing** (`wdtextdirection` label referenced by `Section.text_direction` and `_Cell.text_direction`) |
| `WD_UNDERLINE` | `text.py:380` | `WdUnderline.rst` | covered |
| `WD_VIEW` | `text.py:577` | — | **missing** |
| `WD_WRAP_TYPE` | `shape.py:57` | — | **missing** |

### 5.2 21 enum pages to add

Stub pages needed (pattern from `docs/api/enum/WdAlignParagraph.rst` —
title, alias, one-line intro, bulleted member list). Each is 20-40 lines.

```
WdAnchorH.rst                (WD_ANCHOR_H)
WdAnchorV.rst                (WD_ANCHOR_V)
WdBorderDisplay.rst          (WD_BORDER_DISPLAY)
WdBorderOffsetFrom.rst       (WD_BORDER_OFFSET_FROM)
WdBorderStyle.rst            (WD_BORDER_STYLE)          ← fixes :ref:`wdborderstyle`
WdBreakType.rst              (WD_BREAK_TYPE)
WdBuildingBlockGallery.rst   (WD_BUILDING_BLOCK_GALLERY)
WdDocGridType.rst            (WD_DOC_GRID_TYPE)
WdDrawingType.rst            (WD_DRAWING_TYPE)
WdEndnotePosition.rst        (WD_ENDNOTE_POSITION)
WdFootnotePosition.rst       (WD_FOOTNOTE_POSITION)
WdFootnoteRestart.rst        (WD_FOOTNOTE_RESTART)
WdFrameDropCap.rst           (WD_FRAME_DROP_CAP)
WdFrameHAlign.rst            (WD_FRAME_H_ALIGN)
WdFrameHAnchor.rst           (WD_FRAME_H_ANCHOR)
WdFrameVAlign.rst            (WD_FRAME_V_ALIGN)
WdFrameVAnchor.rst           (WD_FRAME_V_ANCHOR)
WdFrameWrap.rst              (WD_FRAME_WRAP)
WdHeaderFooterIndex.rst      (WD_HEADER_FOOTER_INDEX)
WdInlineShapeType.rst        (WD_INLINE_SHAPE_TYPE)
WdLineNumberingRestart.rst   (WD_LINE_NUMBERING_RESTART)
WdMailMergeDataType.rst      (WD_MAIL_MERGE_DATA_TYPE)
WdMailMergeDestination.rst   (WD_MAIL_MERGE_DESTINATION)
WdMailMergeType.rst          (WD_MAIL_MERGE_TYPE)
WdNumberFormat.rst           (WD_NUMBER_FORMAT)
WdProtection.rst             (WD_PROTECTION)
WdShadingPattern.rst         (WD_SHADING_PATTERN)
WdShape.rst                  (WD_SHAPE)
WdTableAutofit.rst           (WD_TABLE_AUTOFIT)
WdTextDirection.rst          (WD_TEXT_DIRECTION)        ← fixes :ref:`wdtextdirection`
WdView.rst                   (WD_VIEW)
WdWrapType.rst               (WD_WRAP_TYPE)
```

32 entries (more than 21 because several enums that aren't yet warned
about — `WD_FRAME_DROP_CAP`, mail-merge family, etc. — are used in
public signatures that currently document their parameter types as raw
enum names).

`docs/api/enum/index.rst` has a 16-entry toctree that would need to grow
accordingly.

### 5.3 `WD_TABLE_ALIGNMENT` note

`WdRowAlignment.rst` exists; `WdTableAlignment.rst` does not. In
`src/docx/enum/table.py:85` the class is `WD_TABLE_ALIGNMENT`, aliased to
`WD_ROW_ALIGNMENT`. The page title and alias name should be flipped (the
canonical name is `WD_TABLE_ALIGNMENT`), or duplicated. Minor.

---

## 6. User-guide gaps

`docs/user/` has 12 pages totalling 2337 lines:

```
api-concepts.rst          31 lines
comments.rst             168      ← only fork-era narrative
documents.rst             94
hdrftr.rst               166
install.rst               38      ← out of date (section 7.3)
quickstart.rst           328      ← pre-fork (section 7.2)
sections.rst             121
shapes.rst                27
styles-understanding.rst 382
styles-using.rst         391
tables.rst               202
text.rst                 389
```

Per the feature list (121 distinct fork-scope `feat:` commits, grouped by
`^Phase` subject or `#NNN`), the following narrative pages should exist
and do not:

### 6.1 Missing user-guide topics

| Feature area | Issues | Suggested page |
|---|---|---|
| Tracked changes — accept/reject, inspect, move revisions, cell/row changes, formatting changes | #7, #8, #53, #134, #135, #163 | `docs/user/track-changes.rst` |
| Fields — simple, complex, REF, PAGEREF, add_field | #10, #115 | `docs/user/fields.rst` |
| Content controls (SDTs) + data binding | #27, #131 | `docs/user/content-controls.rst` |
| Footnotes + endnotes — add, delete, modify, content, properties | #3, #17, #46, #48, #56, #96 | `docs/user/footnotes.rst` |
| Bookmarks — create, read, delete | #52 | `docs/user/bookmarks.rst` |
| Numbering / lists — apply_to, custom definitions, restart, nested | #22, #108 | `docs/user/numbering.rst` |
| Tables — autofit, borders, cell margins, banded rows, text direction, merged-cell helpers, style flags | #15/#102, #39, #63, #143, #144, #145, #142 | `docs/user/tables-advanced.rst` (or extend `tables.rst`) |
| Section — page borders, line numbering, document grid, paper source, columns, formatting changes, odd/even headers, RTL | #19, #60, #121, #122, #146, #147, #148, #149 | `docs/user/sections-advanced.rst` (or extend `sections.rst`) |
| Charts — read + minimal create | #111 | `docs/user/charts.rst` |
| Watermarks (text + image) | #36 | `docs/user/watermarks.rst` |
| Captions | #141 | `docs/user/captions.rst` |
| Table of contents | #116 | `docs/user/toc.rst` |
| Form fields (legacy) | #123 | `docs/user/form-fields.rst` |
| Permission ranges + document protection | #124, #125 | `docs/user/permissions.rst` |
| Glossary / building blocks | #132, #133 | `docs/user/glossary.rst` |
| Themes | #117 | `docs/user/themes.rst` |
| Mail merge | #130 | `docs/user/mail-merge.rst` |
| Custom document properties + custom XML | #14, #82, #131 | `docs/user/custom-properties.rst` |
| Font — shading, borders, language, East Asian layout, symbols, ruby | #19, #20/#33, #114, #120, #128, #129, #160 | `docs/user/text-advanced.rst` |
| Paragraph — frames, RTL, insert-at-position | #26, #126, #127 | (extend `text.rst`) |
| Accessibility — alt text, heading validation, language tags | #158, #159, #160 | `docs/user/accessibility.rst` |
| Statistics (word count) | #161 | `docs/user/statistics.rst` |
| Search — regex, all-stories | #91, #153, #154 | `docs/user/search.rst` |
| Drawing — floating images, shape creation, group shapes, SVG, alt text | #30, #75, #76, #137, #138, #158 | `docs/user/drawing.rst` (or extend `shapes.rst`, currently 27 lines) |
| Equations (OMML) | #113 | `docs/user/equations.rst` |
| Digital signatures, encrypted/recoverable docs, macro-enabled (.docm) | #150, #151, #152, #65 | `docs/user/document-safety.rst` |

### 6.2 Feature areas already covered (partial check)

- Comments — `docs/user/comments.rst` (168 lines, covers 1.2.0 scope;
  missing threaded-reply narrative for #67).
- Shapes — `docs/user/shapes.rst` exists but is **27 lines** and predates
  all of Phase D shape work.

---

## 7. Front-page + quickstart

### 7.1 `docs/index.rst`

- The "What it can do" code sample (lines 18-65) hasn't been touched since
  upstream — it still demonstrates only `add_heading`, `add_paragraph`,
  `add_picture`, `add_table`, `add_page_break`, and `Inches`. None of the
  fork-era features appear.
- The "API Documentation" toctree (lines 91-104) lists 11 pages. **There
  is no entry for any of the 29 new modules catalogued in section 4.1.**
  Even if the `.rst` stubs existed, a reader would not find them.
- Toctree does not have a `:maxdepth:` that would show second-level
  anchors — new features stay hidden from the sidebar navigator.

### 7.2 `docs/user/quickstart.rst`

- 328 lines, last touched in upstream's 1.2.0 cycle. None of the fork
  features appear — no example of a footnote, tracked change, bookmark,
  content control, field, watermark, TOC, or form field.
- Sections that are still accurate for 1.2.0 but lag the fork: "Adding a
  paragraph", "Adding a heading", "Adding a page break", "Adding a
  picture", "Adding a table", "Applying a character style".
- Feature additions that a modern quickstart should mention (one paragraph
  each, not exhaustive): footnotes, comments, tracked changes, search /
  replace, stable_id, statistics.

### 7.3 `docs/user/install.rst`

- `docs/user/install.rst:37` claims:

  ```
  * Python 2.6, 2.7, 3.3, or 3.4
  * lxml >= 2.3.2
  ```

  Both lines are **wrong**. The current `pyproject.toml` has
  `requires-python = ">=3.9"` and a `lxml>=3.1.0` runtime dep. The
  `easy_install` paragraph (lines 16-19) is obsolete.
- The file still recommends `python setup.py install` (lines 21-27) — this
  has been ineffective since the project switched to a `pyproject.toml`
  build.

---

## 8. HISTORY.rst

`HISTORY.rst` (repo root) stops at `1.2.0 (2025-06-16)`. Every
fork-specific commit — ~121 distinct `feat:` entries — is absent. Before
any release cut, the following release-note skeleton would be needed,
grouped by rough phase:

```
1.3.0.dev0 (unreleased)
+++++++++++++++++++++++

Phase A — Footnotes and endnotes
  - Add Document.footnotes and Footnotes / Footnote / FootnoteProperties (#1, #3, #17, #46, #48, #56, #82)
  - Add Document.endnotes mirror API (#17, #96)
  - Add Section.footnote_properties / endnote_properties (#17)

Phase B — Tracked changes
  - Add read of tracked insertions and deletions (#53)
  - Add accept / reject tracked changes (#7)
  - Add read of formatting changes (#8)
  - Add move revisions (w:moveFrom / w:moveTo) (#134)
  - Add cell and row-level tracked changes (#135)
  - Add revision_marks_text() for CLI previews (#163)

Phase C — Bookmarks and fields
  - Add bookmarks create / read / delete (#52)
  - Add simple and complex field codes (#10)
  - Add REF / PAGEREF cross-reference resolution (#115)

Phase D — Miscellaneous OOXML feature coverage
  - D.1  Hyperlink creation API (#97)
  - D.2  Comment replies (threaded) (#67)
  - D.3  Extended document settings + DocumentProtection (#66, #125)
  - D.4  Custom document properties (#14)
  - D.6  Cell shading and background color (#63)
  - D.7  Paragraph borders (#109)
  - D.9  Numbering style control (#22)
  - D.10 Search and replace with formatting preservation (#91)
  - D.13 Insert paragraph / table at arbitrary position (#26)
  - D.14 Content controls (SDTs) (#27)
  - D.15 Row.height setter (#28)
  - D.16 Row.allow_break_across_pages (#51)
  - D.17 Floating images with wp:anchor positioning (#30)
  - D.19 Multi-column section layout (#60)
  - D.20 Font.shading — run-level background color (#33)
  - D.22 SVG image support (#76)
  - D.23 Watermark support (text and image) (#36)
  - D.24 .docm macro-enabled file support (#65)
  - D.26 Table autofit and column-width control (#39)
  - D.27 DrawingML shapes and text-box content access (#75)

Other feature additions
  - Charts read + add_chart() (#111)
  - SmartArt detection and node text (#112)
  - Equation read + minimal create API (#113)
  - Add Run.add_symbol and Run.symbols (#114)
  - Add Section.page_borders (#121)
  - Add Section.line_numbering (#122)
  - Add Section.document_grid (#147)
  - Add Section.first_page / other_pages_paper_source (#146)
  - Add Section.text_direction / right_to_left (#148)
  - Add Section odd/even page header-footer (#149)
  - Add Font.border_* properties (#120)
  - Add Font.language / east_asian_language / bidi_language (#160)
  - Add East Asian typography (kinsoku, word_wrap, east_asian_layout) (#128)
  - Add RTL / bidi on Paragraph and Run (#127)
  - Add paragraph_format.frame for text frames (#126)
  - Add ParagraphBorders / Border (#109)
  - Add read-only ruby (#129)
  - Add read-only ink (#139)
  - Add read-only embedded OLE objects (#140)
  - Add read-only grouped shapes (#138)
  - Add read-only SmartArt (#112)
  - Add read-only Document.glossary (#132, #133)
  - Add read-only Document.theme (#117)
  - Add read-only Document.web_settings (#157)
  - Add Document.font_table (#119)
  - Add Document.background_color (#118)
  - Add Document.statistics (#161)
  - Add Document.search_regex / replace_regex / search_all / replace_all (#153, #154)
  - Add Document.add_table_of_contents (#116)
  - Add caption helpers (#141)
  - Add permission ranges (#124)
  - Add Settings.mail_merge (#130)
  - Add Settings.compat_flags / compat_settings (#156)
  - Add Settings.view (#164)
  - Add Style.link_style / next_style / is_redefined (#162)
  - Add Table.borders / _Cell.borders (#102)
  - Add Cell.margins (#143)
  - Add Table.style_flags (#144)
  - Add Cell.text_direction (#142)
  - Add Cell.is_merge_origin / merge_origin (#145)
  - Add _Row.is_header (#93)
  - Add Run.split (#94)
  - Add Paragraph.delete / Run.delete / Table.delete (#50)
  - Add alt_text / title on InlineShape and FloatingImage (#158)
  - Add stable_id on Paragraph / Run / Table / Cell (#155)
  - Add Paragraph.insert_paragraph_before arbitrary positioning (#26)
  - Add legacy form fields (#123)
  - Add heading-structure accessibility validator (#159)

Reliability / safety
  - Add recover=True mode for malformed .docx (#151)
  - Add EncryptedDocumentError for password-protected .docx (#152)
  - Add digital signature detection (#150)

Dev / tooling
  - Add py.typed, improve public types
  - Add AI-agent CI pipeline (Product / Develop / Review / Security / Revise
    / Merge / Debug / Watchdog)
```

Numbers above are taken directly from `git log --oneline --all --grep='^feat'`.
Writing the real HISTORY entries should take ~2 hours if done as one pass.

---

## 9. Other issues

### 9.1 Broken image links

**None.** All four `.. image::` references point at files that exist in
`docs/_static/img/`:

```
_static/img/comment-parts.png    used in docs/user/comments.rst:16 + dev/analysis/features/comments.rst:15
_static/img/example-docx-01.png  used in docs/index.rst:14
_static/img/hdrftr-01.png        used in docs/user/hdrftr.rst:66
_static/img/hdrftr-02.png        used in docs/user/hdrftr.rst:92
```

### 9.2 References to removed modules / classes

Grepping the `.rst` tree for stale `.. currentmodule::` or dotted paths
against the current source tree:

```
grep -rn 'docx\.' docs/api/*.rst | cut -d: -f3 | grep -oE 'docx\.[a-z_.]+' | sort -u
```

returns 18 distinct dotted paths, every one of which resolves under
`src/docx/`. No dead references.

### 9.3 `.. todo::` / `:deprecated:` markers

`grep -rn '.. todo::\|.. deprecated::' docs/ src/docx/` returns zero hits.
There are no dangling to-dos in docstrings or RST files.

### 9.4 `docs/conf.py`

Observations:

- No `autodoc_default_options` is set. Every `.. autoclass::` directive
  must spell out `:members:` / `:inherited-members:` individually. This is
  why pages like `docs/api/text.rst` need 10 `:members:` tokens — a
  project-wide `autodoc_default_options = {"members": True, "undoc-members": False, "show-inheritance": True}` would halve the size of each page.
- `intersphinx_mapping = {"http://docs.python.org/3/": None}` (line 393)
  is in Sphinx 1.0 format — deprecated in Sphinx 4, breaks in Sphinx 8.
- `html_theme = "armstrong"` (line 236) points at a vendored fork of a
  long-abandoned theme (`docs/_themes/armstrong/`). `alabaster` (the note
  in CLAUDE.md) is pinned in `requirements-docs.txt` but never consumed.
  The project would benefit from migrating to a maintained theme
  (`furo`, `pydata-sphinx-theme`, or `sphinx_rtd_theme`).
- `exclude_patterns = [".build"]` (line 208) — the actual build dir is
  `_build`, not `.build`, so this exclusion does nothing.
- `copyright = "2013, Steve Canny"` (line 55) has not been updated in
  12 years.
- Typo in docstring of `add_endnote_properties` etc. would not be fixed
  by any conf change — the `|EndnoteProperties|` etc. substitutions are
  simply absent from `rst_epilog` (section 3.5).
- `sphinx.ext.todo` is enabled but never used (`grep -rn '.. todo::' docs/`
  returns 0).
- `sphinx.ext.coverage` is enabled but has no `coverage_*` options set.

---

## 10. Recommendations

Prioritised punch list. Effort labels: **S** = < 2 hr, **M** = 1/2 day,
**L** = 1-2 days.

### Sphinx build hygiene

1. **S — Add the 53 missing `|Name|` substitutions to `docs/conf.py:rst_epilog`.**
   One line per symbol (`.. |Foo| replace:: :class:`.Foo``). Retires
   96 of the 101 build warnings. See section 3.5. ~30 min.
2. **S — Fix `intersphinx_mapping` format** (`docs/conf.py:393`). Change
   `{"http://docs.python.org/3/": None}` to
   `{"python": ("https://docs.python.org/3/", None)}`. Unblocks `-W`. ~5 min.
3. **S — Fix `docs/user/comments.rst:134`** Markdown underscore to RST
   emphasis. ~1 min.
4. **S — Update `requirements-docs.txt`** to modern Sphinx (≥5, <8),
   drop `MarkupSafe==0.23` / `Jinja2==2.11.3` pins, drop `alabaster`
   pin (unused). ~10 min.
5. **S — Fix `exclude_patterns`** in `conf.py:208` (`.build` → `_build`).

### New API stubs

6. **S — Add missing enum reference pages (21-32 files).** Follow
   `docs/api/enum/WdAlignParagraph.rst` pattern. Each page is 20-40
   lines of title / alias / one-line intro / member list. Update
   `docs/api/enum/index.rst` toctree. See section 5.2. ~2 hr total
   (scriptable from `enum/base.py`'s `DocsPageFormatter`, which already
   emits exactly this format).
7. **S — Add autoclass stubs for the 29 new proxy modules** (section 4.1).
   Each is a ~10-line file (`.. _X_api:`, title, `.. currentmodule::`,
   `.. autoclass:: Foo :members:`). Add to `docs/index.rst` toctree. ~2 hr.

### Existing API pages

8. **M — Update `document.rst`, `settings.rst`, `section.rst`, `table.rst`,
   `text.rst`, `shape.rst`** to surface the new return-type classes and
   narrative sections per section 4.2. Most `:members:` directives already
   cover method additions; the missing pieces are new top-level classes
   in the same module (`CellShading`, `TableBorders`, `LineNumbering`,
   `DocumentGrid`, `PageBorder`, `PageBorders`, `FloatingImage`,
   `EastAsianLayout`, `TextFrame`, `ParagraphBorders`, `Symbol`, etc.) plus
   re-ordered subsection titles. ~1 day of careful writing to match the
   existing narrative tone.
9. **S — Drop the explicit `:members: height, type, width` allowlist on
   `docs/api/shape.rst:31`.** New `InlineShape.alt_text` and `.title`
   (#158) will start rendering automatically. ~2 min.
10. **S — Rename `docs/api/enum/WdRowAlignment.rst`** (or duplicate it) to
    expose `WD_TABLE_ALIGNMENT` as the canonical name. See 5.3. ~10 min.

### Conf-file improvements

11. **S — Add `autodoc_default_options`** in `conf.py` so new autoclass
    directives inherit `members: True`, `show-inheritance: True`. Shrinks
    each API page, makes new features visible by default. ~15 min.
12. **M — Migrate theme** from vendored `armstrong` to a maintained theme
    (`furo` or `pydata-sphinx-theme`). Delete `docs/_themes/armstrong/`.
    ~1 hr + QA pass.

### User guide

13. **L — Write user-guide narratives for fork feature areas.** Section 6.1
    lists 26 missing topics. Pair each with a 100-200 line page in the
    existing narrative style (compare `docs/user/comments.rst`, 168
    lines). Prioritise in this order based on commit frequency: tracked
    changes → fields → content controls → footnotes → bookmarks →
    numbering → watermarks → search → captions → TOC → form fields →
    charts → equations → rest. ~1-2 weeks at normal pace.
14. **M — Rewrite `docs/user/install.rst`.** Drop Python 2.x and easy_install,
    update `requires-python = ">=3.9"`, update `lxml` to current floor,
    mention `pip install python-docx`. ~20 min.
15. **M — Modernise `docs/user/quickstart.rst`.** Add short-paragraph
    sections for the five or six most common fork features (footnote,
    bookmark, find/replace, comment, tracked change, stable_id). ~2 hr.
16. **S — Update the `docs/index.rst` "What it can do" code sample** to
    include at least one fork-era call (e.g. a `document.footnotes.add()`
    line). ~10 min.

### History / release prep

17. **M — Write `HISTORY.rst` entries** for the unreleased version. The
    skeleton in section 8 can be pasted in and polished. ~2 hr.

### Longer-term

18. **L — Audit every public docstring** for OOXML-term consistency
    (`w:rPr`, `w:tc`, ...) and add "Added in version x.y.z" directives
    so the release-notes skeleton can be auto-generated from code.
    ~1 week, on and off.
19. **M — Add a `docs/user/api-concepts.rst` extension** covering the
    three-layer proxy / part / oxml architecture (already described in
    `CLAUDE.md` but not visible to end-users). ~3 hr.
20. **S — Replace `sphinx.ext.todo`** (unused) with `sphinx.ext.napoleon`
    (Google / NumPy docstring support). The codebase mostly uses plain
    reStructuredText docstrings but a few newer modules use NumPy-style
    Returns sections that currently render as plain paragraphs.

---

## Appendix A — counts at a glance

| Metric | Count |
|---:|---|
| `.rst` files under `docs/` | 75 |
| Lines across all `.rst` files | 12 005 |
| `.rst` files under `docs/api/` (top level) | 10 |
| `.rst` files under `docs/api/enum/` (incl. index) | 17 |
| `.rst` files under `docs/user/` | 12 |
| `.rst` files under `docs/dev/analysis/` | 36 |
| Python modules in `src/docx/` (top level) | 44 |
| Python submodule files under `src/docx/{text,styles,enum,drawing,...}` | additional 29 |
| Enum classes defined in `src/docx/enum/` | 37 |
| Enum classes with a `docs/api/enum/*.rst` page | 16 |
| `|Substitution|` tokens defined in `conf.py:rst_epilog` | 74 |
| `|Substitution|` tokens referenced but undefined | 53 |
| Distinct `feat:` commits in `git log --all` (fork scope) | 121 |
| Doc commits since 2025-06-11 | 0 |
| Sphinx build warnings (non-strict) | 101 |

## Appendix B — Sphinx build command used

```
python -m sphinx -b html docs docs/_build/html
```

Ran with Sphinx 6.2.1 inside a throwaway Python 3.11 virtualenv.
Build result: `build succeeded, 101 warnings.` Exit status 0.
`docs/_build` and the virtualenv were removed after the run.
