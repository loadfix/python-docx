# Sibling Documentation Audit — loadfix OOXML Series

Observational audit of three sibling projects under the `loadfix` org that
together aim at a uniform OOXML-in-Python experience:

- `loadfix/python-docx` — Word `.docx` (this repo)
- `loadfix/python-pptx` — PowerPoint `.pptx`
- `loadfix/python-xlsx` — Excel `.xlsx` (fork of `openpyexcel`, which in turn
  is a snapshot of `openpyxl` ~2.5.14)

The audit looks only at *documentation surfaces* — READMEs, Sphinx builds,
`docs/` trees, history/changelog files, contributor-facing files, docstring
quality and build health. No runtime or test-coverage claims are made here.
The audit is read-only against the two sibling repos; only this file is
written.

---

## 1. Summary

The three projects present three very different levels of documentation
maturity and consistency:

- **python-docx (this repo)** is the most mature documentation surface of
  the three. Sphinx builds with `furo` (modern theme), exactly one Sphinx
  warning (`undefined label: 'wdoutlinelvl'`), a 37-page user guide
  mirroring a 38-file API reference, a 49-file enum reference, a freshly
  written fork-centric `README.md` using Markdown, an 81 KB `FEATURES.md`
  acting as the authoritative fork-feature inventory, a 12 KB `HISTORY.rst`
  ordered by phase (A/B/C/D), and a pre-napoleon hook in `conf.py` that
  warrants its own ~80-line comment block explaining why it exists. A
  previous run noted **17 build warnings** for this repo; current grep of
  `/tmp/sphinx-docx.log` shows 1 WARNING and 16 ERRORs all of the
  `Undefined substitution referenced` class, total 17 — the previous
  agent counted these together.

- **python-pptx** is in a middle state: an inherited Sphinx 1.8.6 /
  Jinja2 2.11.3 / `alabaster<0.7.14` / `armstrong` theme stack, an
  `rst_epilog`-based substitution catalogue that is showing its age
  (many class references are no longer resolved, hence 140 warnings — of
  which **127 are `ERROR: Undefined substitution referenced`** of the
  form `|ErrorBars|`, `|AnimationEffect|`, `|PathGeometry|`, `|Path|`,
  `|Sound|`, `|Section|`, and so on). Docstring coverage in `src/pptx` is
  larger than docx's but the reference tree is thin (14 API pages, 36
  enum pages, 19 user-guide pages). `README.rst` is 25 lines and still
  matches the upstream form — no fork narrative. `HISTORY.rst` is 1768
  lines and still active (an unreleased block at the top describes recent
  refactors). There is no `FEATURES.md`, no `CONTRIBUTING.md` (only
  `CLAUDE.md`), and the user guide has no fork-feature pages.

- **python-xlsx** is the outlier. Its Sphinx build **fails outright**:
  `doc/conf.py` line 25 does `import openpyxl`, but the actual top-level
  package in this fork is `xlsx` (under `src/xlsx/`) with a legacy
  `openpyexcel/` tree alongside — nothing named `openpyxl` exists to
  import, so Sphinx dies before the first source file is read
  (`ModuleNotFoundError: No module named 'openpyxl'`). There is no
  `requirements-docs.txt`, no `furo`/`alabaster` pin, no `docs/`
  directory (the docs live in `doc/` — a different convention — with 39
  `.rst` files as a flat list, no `api/` and no `user/` split). There is
  a `TODO.md` (582 lines), `CONTRIBUTING.md` (17 lines, stub), `CLAUDE.md`
  (334 lines, much more prose than python-docx's 166 lines), a 23 KB
  `README.md`, and a `doc/changes.rst` (1376 lines — still the openpyxl
  2.5.14 changelog, not a fork-era history). No `FEATURES.md`.

Net: docx has by far the most *current* documentation, xlsx the most
*broken*, pptx the most *inherited*. The three do not yet look like a
series.



## 2. Layout matrix

Raw numbers, gathered from filesystem state as of 2026-05-02.

| Surface                              | python-docx         | python-pptx         | python-xlsx                              |
|--------------------------------------|---------------------|---------------------|------------------------------------------|
| Top-level README                     | `README.md` (131 L) | `README.rst` (25 L) | `README.md` (190 L)                      |
| Top-level HISTORY                    | `HISTORY.rst` (438 L) | `HISTORY.rst` (1768 L) | `doc/changes.rst` (1376 L) — openpyxl-era |
| Top-level CHANGELOG                  | (none, in HISTORY)  | (none, in HISTORY)  | (none, openpyxl-era `changes.rst`)       |
| Top-level FEATURES                   | `FEATURES.md` (1791 L / 81 KB) | (absent) | (absent; `TODO.md` 582 L) |
| Top-level CLAUDE.md                  | 166 L (7.3 KB)      | 256 L (15 KB)       | 334 L (23 KB)                            |
| Top-level CONTRIBUTING               | (absent)            | (absent)            | `CONTRIBUTING.md` (17 L — stub)          |
| Top-level LICENSE                    | `LICENSE` (MIT)     | `LICENSE` (MIT)     | `LICENCE.md` (British spelling)          |
| Top-level AUTHORS                    | (absent)            | (absent)            | `AUTHORS.md`                             |
| Sphinx source dir                    | `docs/`             | `docs/`             | `doc/` (singular)                        |
| Sphinx config                        | `docs/conf.py` (627 L) | `docs/conf.py` (587 L) | `doc/conf.py` (314 L)                |
| Sphinx theme                         | `furo`              | `armstrong`         | `nature` (or `default`)                  |
| Total `.rst` files under docs/       | 162                 | 180                 | 39 (flat)                                |
| `docs/api/*.rst` (non-enum)          | 38                  | 14                  | 0 (no api/ dir)                          |
| `docs/api/enum/*.rst`                | 49                  | 36                  | 0 (no enum/ dir)                         |
| `docs/user/*.rst`                    | 38                  | 19                  | 0 (no user/ dir; flat `.rst` instead)    |
| `docs/dev/` directory                | `docs/dev/analysis/` | `docs/dev/` (6 files) | (absent)                              |
| `docs/community/`                    | (absent)            | `docs/community/` (3 files) | (absent)                         |
| `requirements-docs.txt`              | yes (3 lines, Sphinx>=6,<8 / furo / -e .) | yes (5 lines, Sphinx==1.8.6 pinned) | **absent**              |
| Sphinx build result                  | success, 17 warnings | success, 140 warnings | **fails** (ModuleNotFoundError)       |
| `versionadded::` in source           | 885 occurrences     | 0                   | 0                                        |
| Public `.py` modules in src top-lvl  | 46 (`src/docx/*.py`) | 15 (`src/pptx/*.py`) | ~5 (`src/xlsx/*.py`) plus `openpyexcel/` |
| `spec/` directory                    | yes (ISO-IEC-29500 PDFs + xsd + rnc + styles.xml) | yes (ISO-IEC-29500 tree + gen_spec) | (absent) |

Line-count and file-count figures are raw `wc -l` / `ls | wc -l` — not
weighted by content quality.



## 3. Surface-by-surface comparison

### 3.1 README / top-of-project landing

**python-docx — `README.md` (131 lines, Markdown).**
Opens with a two-sentence description, then a "Based on..." paragraph that
is explicit about the fork provenance ("Based on python-openxml/python-docx
by Steve Canny and contributors. Forked at upstream `1.2.0` (2025-06-16)
and extended with 100+ additional OOXML features — footnotes and
endnotes, tracked changes, bookmarks, fields, content controls, charts,
equations, SmartArt, watermarks, digital signatures, accessibility
tooling, cross-document operations, and more."). Subsequent sections
appear to cover "Status" (unstable, not yet on PyPI, CalVer
`2026.05.0`), "Installation" from source, and a short narrative pitch.
The tone is fork-aware and version-dated.

**python-pptx — `README.rst` (25 lines, reStructuredText).**
Opens with the traditional upstream blurb: *python-pptx is a Python
library for creating, reading, and updating PowerPoint (.pptx) files.*
No mention of the fork, no mention of `loadfix`, no mention of CalVer,
no mention of `2026.05.0`. Essentially the upstream README preserved
verbatim. The README is also `include`d into `docs/index.rst`, so the
Sphinx landing page inherits the same upstream-flavoured intro.

**python-xlsx — `README.md` (190 lines, Markdown).**
Opens with the same shape as `python-docx`'s README — "A Python library
for reading, creating, and updating Microsoft Excel 2007+ (.xlsx /
.xlsm) files" — followed by a multi-step "Based on openpyxl via
sciris/openpyexcel. Forked from openpyexcel (which tracks openpyxl
~2.5.14, circa 2019) and extended with modern Excel 365 capabilities..."
paragraph naming dynamic arrays, threaded comments, rich text on
cells, sparklines, modern chart types (treemap, sunburst, funnel,
waterfall, box-whisker, histogram, map), SHA-512 protection, encrypted
file I/O, first-class shape models. Status section says "Unstable. Not
yet published to PyPI. Install from source only. Current version:
`2026.05.0` (first release as an independent fork)." — identical
wording to docx's.

Observation: the docx and xlsx READMEs are visibly the work of the same
hand, using the same CalVer pitch, same "Unstable" status boilerplate,
and same "Based on ... Forked at / from ..." pattern. pptx's README has
not been updated for the fork at all.



### 3.2 Sphinx build setup

**python-docx — Sphinx >=6,<8, `furo` theme.**
- `requirements-docs.txt` is three lines: `Sphinx>=6,<8`, `furo`, `-e .`.
- `docs/conf.py` is 627 lines. The top of the file loads `docx.__version__`.
  Extensions enabled: `sphinx.ext.autodoc`, `sphinx.ext.intersphinx`,
  `sphinx.ext.napoleon`, `sphinx.ext.viewcode`. Napoleon is configured
  with Google- and NumPy-style docstring support plus `include_special_with_doc`.
- The largest novelty is a ~80-line `setup(app)` function that installs a
  *pre-napoleon snapshot / post-napoleon restore* hook at priorities 100
  and 900. The comment block explains it is there to prevent napoleon
  from mis-parsing attribute docstrings that reference OOXML element
  names like `w:moveFrom` — napoleon would otherwise split on the colon
  and emit ~42 bogus "Unknown target name: 'w:moveFrom'" docutils
  warnings across 11 files.
- `intersphinx_mapping` is the modern `{'python': ('https://docs.python.org/3/', None)}` form.

**python-pptx — Sphinx 1.8.6, `armstrong` theme.**
- `requirements-docs.txt` pins `Sphinx==1.8.6`, `Jinja2==2.11.3`,
  `MarkupSafe==0.23`, `alabaster<0.7.14`, `-e .` — a 2019-era Sphinx
  stack preserved verbatim. `armstrong` is a custom theme (no evidence
  of a `_themes/` directory vendoring it, so it is expected to be pulled
  from PyPI — the build log reports the theme is resolved).
- `docs/conf.py` is 587 lines. Extensions enabled: autodoc, doctest,
  inheritance_diagram, intersphinx, todo, coverage, ifconfig, viewcode —
  a broader set than docx but no napoleon. Also monkey-patches
  `sphinx.environment.BuildEnvironment.warn_node` to suppress "nonlocal
  image URI found:" warnings for an old travis-ci status badge.
- The bulk of `conf.py` is a massive `rst_epilog` declaring dozens of
  `|ClassName|` substitutions — the source of the 127 "Undefined
  substitution referenced" errors in the current build, since the
  substitution table has not kept pace with renamed/new classes
  (`ErrorBars`, `AnimationEffect`, `PathGeometry`, `Path`, `Sound`,
  `Section`, `_HeaderFooter`, `Transition`, `SmartArt`, …).

**python-xlsx — openpyxl-era Sphinx config, build fails.**
- No `requirements-docs.txt`. Docs dependencies are not separately
  declared anywhere the audit located.
- `doc/conf.py` (singular `doc/`, not `docs/`) is 314 lines and still
  attributed to "openpyxl documentation build configuration". The module
  import at line 25 reads `import openpyxl`, but the top-level package
  name in this fork is `xlsx` (under `src/xlsx/`) with an `openpyexcel/`
  sibling tree. Nothing ever renames `openpyxl` to match, so Sphinx
  aborts: `ModuleNotFoundError: No module named 'openpyxl'`.
- Theme selection is `'default'` in the `'on_rtd'` branch, else
  `'nature'` — pre-ReadTheDocs-era.
- Extensions: `sphinx.ext.autodoc`, `ifconfig`, `viewcode`, `doctest`,
  `coverage`. No napoleon, no intersphinx.
- The module-patch dance at top of `conf.py` (AliasProxyGet /
  NumberFormatGet / StyleDescriptorGet monkey-patches behind
  `APIDOC=True`) is a flag that `openpyxl` had its own autodoc
  workarounds that were never ported either.

Observation: the three configs span three Sphinx generations (1.8.6,
openpyxl-era pre-RTD, and 7.x + furo). Only docx is on the current
stack.



### 3.3 docs/ layout

**python-docx — `docs/` (plural).**
- Top-level files: `index.rst`, `conf.py`, `Makefile`, `_static/`,
  `api/`, `dev/`, `user/`.
- `docs/api/` has 38 non-enum `.rst` files plus an `enum/` subtree of
  49 files — so 87 reference pages in total.
- `docs/user/` has 38 user-guide pages (see §3.5 for the list).
- `docs/dev/` contains an `analysis/` subtree — analysis notes for
  reverse-engineered features, inherited from upstream and extended.
- No `docs/community/`, no FAQ, no support page.
- `docs/_static/` holds images referenced by `index.rst`.

**python-pptx — `docs/` (plural).**
- Top-level files: `index.rst`, `conf.py`, `Makefile`, `_static/`,
  `_templates/`, `api/`, `community/`, `dev/`.
- `docs/api/` has 14 non-enum `.rst` files plus 36 enum pages — 50
  reference pages total. About 57% the size of docx's reference tree.
- `docs/user/` has 19 user-guide pages. About half the size of docx's
  user guide.
- `docs/dev/` has 6 files: `analysis/` (subtree), `development_practices.rst`,
  `philosophy.rst`, `resources/`, `runtests.rst`, `security.rst`,
  `xmlchemy.rst`. Richer than docx's `dev/`.
- `docs/community/` has `faq.rst`, `support.rst`, `updates.rst` —
  community-facing pages absent from docx and xlsx.
- `docs/_templates/` is present (docx does not have one).

**python-xlsx — `doc/` (singular).**
- Flat structure: no `api/` or `user/` split, no `enum/` subtree.
- 39 `.rst` files at the top of `doc/`, including `index.rst`,
  `tutorial.rst`, `usage.rst`, `changes.rst`, `development.rst`,
  `worksheet_tables.rst`, `filters.rst`, `formula.rst`, `formatting.rst`,
  `charts/` (subdirectory — one of only two subdirs), plus
  `comments.rst`, `defined_names.rst`, `editing_worksheets.rst`,
  `optimized.rst`, `pandas.rst`, `performance.rst`, `pivot.rst`,
  `print_settings.rst`, `protection.rst`, `styles.rst`, `validation.rst`,
  `windows-development.rst`, `worksheet_properties.rst`.
- Also contains Python example source alongside the RSTs (`example.py`,
  `filters.py`, `format_merged_cells.py`, `table.py`) and PNG image
  assets (`filters.png`, `logo.png`, `table.png`) mixed into the same
  flat directory rather than isolated under `_static/` — the `_static/`
  dir exists too, so the layout is half-converted.
- `changes.rst`, `read_performance.txt`, `write_performance.txt` are
  all in `doc/` — a `txt` with perf numbers committed alongside docs.

Observation: docx and pptx share the `docs/` name, an `api/` /
`api/enum/` split, a `user/` dir, and a Makefile. xlsx does none of
that.



### 3.4 API reference coverage

**python-docx — 38 non-enum API pages + 49 enum pages = 87 files.**

Non-enum API pages (`docs/api/*.rst`):
`accessibility`, `bookmarks`, `captions`, `chart`, `comments`,
`content-controls`, `custom-properties`, `custom-xml`, `dml`, `document`,
`embedded-objects`, `endnotes`, `equations`, `fields`, `font-table`,
`footnotes`, `form-fields`, `glossary`, `ink`, `numbering`, `permissions`,
`ruby`, `search`, `section`, `settings`, `shape`, `shared`, `signatures`,
`smart-art`, `stable-ids`, `statistics`, `style`, `table`, `text`,
`theme`, `toc`, `tracked-changes`, `watermark`, `web-settings`.

This mirrors almost 1:1 the fork's phase-A/B/C/D feature list. Every
fork-added subsystem has its own reference page.

Enum pages (49) include every upstream `WD_*` enum plus the fork's
new additions — e.g. `WdAnchorH`, `WdAnchorV`, `WdBorderDisplay`,
`WdBorderOffsetFrom`, `WdBorderStyle`, `WdBreakType`,
`WdBuildingBlockGallery`, `WdBuiltinStyle`, `WdCellVerticalAlignment`,
`WdColorIndex`, `WdDocGridType`, `WdDrawingType`, `WdEndnotePosition`,
`WdFootnotePosition`, `WdFootnoteRestart`, `WdFrameDropCap`,
`WdFrameHAlign`, plus two MSO enums (`MsoColorType`,
`MsoThemeColorIndex`). Both fork-specific (e.g. `WdAnchorH`,
`WdFootnoteRestart`) and inherited enums are documented alongside.

**python-pptx — 14 non-enum API pages + 36 enum pages = 50 files.**

Non-enum API pages (`docs/api/*.rst`):
`action`, `chart-data`, `chart`, `comments`, `dml`, `exc`, `image`,
`placeholders`, `presentation`, `shapes`, `slides`, `table`, `text`,
`util`.

This is substantially smaller than docx's 38 pages and looks closer to
the upstream python-pptx 1.0.x reference tree. Recent Wave-5 fork
additions — animation, sections, transitions, smart-art, tags — do not
yet have dedicated API pages despite having source modules, which
matches the pattern of "Undefined substitution referenced" errors
piling up in the Sphinx build (§3.11).

Enum pages (36) include the MSO/PP/XL trinity expected by the upstream
python-pptx design — `MsoAutoShapeType`, `MsoAutoSize`, `MsoColorType`,
`MsoConnectorType`, `MsoFillType`, `MsoLanguageId`, `MsoLineDashStyle`,
`MsoLineEndLength`, `MsoLineEndType`, `MsoLineEndWidth`,
`MsoPatternType`, `MsoShapeType`, `MsoTextStrikeType`,
`MsoTextUnderlineType`, `MsoThemeColorIndex`, `MsoVerticalAnchor`,
`PpActionType`, `PpAutoNumberScheme`, `PpMediaType`, plus
`ExcelNumFormat` (unusual — a cross-sibling name appearing in the pptx
enum tree) and others.

**python-xlsx — no API reference tree at all.**

No `api/` directory, no `enum/` subtree, no per-class reference pages.
The top-level `.rst` files are all task/topic-oriented (tutorial,
usage, filters, formula, pivot, …) — autodoc usage across them has not
been spot-checked in this audit, but the flat layout means there is no
place dedicated to class-by-class reference.



### 3.5 User-guide coverage

**python-docx — 38 user-guide pages in `docs/user/`.**

Full list: `accessibility`, `api-concepts`, `bookmarks`, `captions`,
`charts`, `comments`, `content-controls`, `custom-properties`,
`document-safety`, `documents`, `drawing`, `endnotes`, `equations`,
`fields`, `footnotes`, `form-fields`, `glossary`, `hdrftr`, `install`,
`mail-merge`, `numbering`, `permissions`, `quickstart`, `search`,
`sections`, `sections-advanced`, `shapes`, `statistics`,
`styles-understanding`, `styles-using`, `tables`, `tables-advanced`,
`text`, `text-advanced`, `themes`, `toc`, `track-changes`, `watermarks`.

Pattern: for the big subsystems (`text`, `tables`, `sections`) there
are both a core page and an `-advanced` page. Near-1:1 parity between
user-guide pages and API reference pages — almost every API area has a
prose companion.

**python-pptx — 19 user-guide pages in `docs/user/`.**

Full list: `autoshapes`, `charts`, `comments`, `concepts`, `install`,
`intro`, `math-equations`, `media`, `notes`, `ole-objects`,
`placeholders-understanding`, `placeholders-using`, `presentations`,
`quickstart`, `slides`, `table`, `text`, `understanding-shapes`,
`use-cases`.

No dedicated page for: animation, sections, transitions, smart-art,
tags, custom properties, extended properties, field manipulation,
accessibility. For a fork that has added substantial new capability
(`pptx.animation`, `pptx.slide.AnimationEffectView`, Wave-5 animation
API, section API #256), the user-guide has not grown correspondingly.

**python-xlsx — no `user/` dir; 39 flat `.rst` files.**

The flat layout mixes reference-ish (`comments.rst`, `styles.rst`,
`protection.rst`) with task-oriented (`editing_worksheets.rst`,
`optimized.rst`, `performance.rst`, `read_performance.txt`,
`write_performance.txt`, `windows-development.rst`) and a `charts/`
subdir plus `tutorial.rst` and `usage.rst`. The content is mostly
openpyxl heritage; there are no fork-era pages documenting dynamic
arrays, threaded comments, rich text on cells, sparklines, or modern
chart types (which the README pitches as the fork's main value
proposition).



### 3.6 HISTORY / CHANGELOG

**python-docx — `HISTORY.rst`, 438 lines.**

Opens with a bold fork-transition header:

> `2026.05.0 (unreleased) — first release as independent fork`
>
> This release marks the project's split from upstream
> `python-openxml/python-docx`. Versioning switches to CalVer
> (YYYY.MM.patch) from this point forward. The previous upstream line
> stops at `1.2.0` (2025-06-16); everything below is new to this fork.

The changelog is organised by **development phase** (`Phase A —
Footnotes and endnotes`, `Phase B — Tracked changes`, `Phase C —
Bookmarks and fields`, `Phase D — Numbering / lists / misc`) with
issue numbers next to each bullet (e.g. `(#1, #3, #17, #46, #48, #56,
#82)`). This is a format unique to docx in the series — it reads like
an engineering phase plan collapsed into a release-note shape.

**python-pptx — `HISTORY.rst`, 1768 lines.**

Opens with an `Unreleased` section that reads like an active working
log: refactor notes, issue verifications, deprecation warnings for
`pptx.slide.AnimationEffect → AnimationEffectView`, docs closes, API
resolution notes ("`#357 (customize pie-chart slice colors) resolved
by …`"). Very different tone from docx — granular, incremental,
issue-tracked, written for contributors as much as users. No
"first release as independent fork" banner — the fork transition is
not called out anywhere visible at the top.

**python-xlsx — `doc/changes.rst`, 1376 lines.**

Starts with `2.5.14 (2019-01-23)` — *this is still the openpyxl 2.5.14
changelog*. Issue links point at `bitbucket.org/openpyxl/openpyxl` —
Bitbucket URLs for an upstream that moved off Bitbucket. There is no
fork-era release banner, no `2026.05.0`, no CalVer. The fork's own
history appears to live only in git log, not in a changelog file.
There is also no `HISTORY.rst` at project root — `changes.rst` inside
`doc/` is the only changelog surface.

Observation: docx is "phase-banded CalVer", pptx is "engineering
working-log", xlsx is "upstream snapshot not updated". Three formats,
three audiences.



### 3.7 Contributor docs (CLAUDE.md, CONTRIBUTING)

**python-docx — `CLAUDE.md`, 166 lines.**

A compact, code-heavy reference. Opens one-line: "python-docx fork
(loadfix/python-docx) — extending python-docx with footnotes,
endnotes, track changes, fields, bookmarks, and other missing OOXML
capabilities." Immediately acknowledges the sibling series and tells
the reader that when implementing a cross-sibling feature they should
"consult the sibling repos for naming and API-shape precedent". Then
it launches into the three-layer architecture diagram and worked
code snippets — `CT_Footnote`, `FootnotesPart`, `ZeroOrOne` /
`ZeroOrMore` with `successors=(...)`. Testing conventions (Describe,
`it_*`/`its_*`/`they_*`, `cxml.element(...)` snippets). No
`CONTRIBUTING.md`, no `AUTHORS`.

**python-pptx — `CLAUDE.md`, 256 lines.**

More prose-oriented. Same "Guidance for Claude Code (and other AI
assistants)" framing as xlsx. Numbered sections (`## 1. Project
summary`). Mentions the sibling series explicitly: "python-pptx is
part of a family of Python libraries for reading/writing Office Open
XML formats. Each targets a different Office application but shares
the same design philosophy (lxml-backed, no Office install required,
round-trip fidelity, src-layout + strict tooling)". Lists the two
siblings under "Sibling projects" but names them `scanny/python-docx`
and (presumably) `sciris/...` — **upstream names, not `loadfix/` fork
names**. No `CONTRIBUTING.md`.

**python-xlsx — `CLAUDE.md`, 334 lines. `CONTRIBUTING.md`, 17 lines.**

The most extensive CLAUDE.md of the three. Same numbered-sections
format and opening line as pptx. Section 1 is richer than the other
two — describes the double-fork history (`python-xlsx → openpyexcel
→ openpyxl (~2.5.14, circa 2019)`), names remotes (`origin` is
`loadfix/python-xlsx`), identifies descriptor-layer inheritance,
notes the post-fork focus areas, confirms MIT/Expat, `2026.05.0`
CalVer, and runtime deps `jdcal` / `et_xmlfile`. Most detailed of
the three.

`CONTRIBUTING.md` exists but is a 17-line stub — no counterpart in
docx or pptx. `AUTHORS.md` exists alongside — a format absent from
the others.

Observation: the three CLAUDE.md files share a family resemblance
(sibling-series acknowledgement, three-layer architecture diagram,
lxml note, Python version note) but disagree on whether the sibling
org is `loadfix/` (docx) or upstream names (pptx). xlsx's is the
longest and most prose-heavy.



### 3.8 Inline / docstring quality

Measured by two proxies — counts of `versionadded::` directives and
Sphinx's own complaints about docstring content.

**`versionadded::` directives in `src/`**
- python-docx: **885 occurrences**
- python-pptx: **0 occurrences**
- python-xlsx: **0 occurrences**

docx is the only project that tracks per-feature addition version in
the docstring. Given the "first release as independent fork" nature of
the 2026.05.0 cut, almost all 885 entries are plausibly tagged
`.. versionadded:: 2026.05.0` and will become historically meaningful
after the second release — but the habit is already instrumented.

**Sphinx's complaints**
- docx: 17 warnings (16 are `Undefined substitution referenced` for
  class names, one is an `undefined label: 'wdoutlinelvl'`). The class
  substitutions resolve elsewhere because docx does not use an
  `rst_epilog` substitution table at all — the few `|AltChunk|`,
  `|Attachment|`, `|ExtendedProperties|`, `|DocVars|`, `|ImagePart|`,
  `|Image|`, `|StoryPart|`, `|TableCellMargins|`, `|IndexError|`,
  `|FloatingImage|` references are ad-hoc and don't have definitions.
- pptx: 140 warnings, of which 127 are `Undefined substitution
  referenced` caused by the `rst_epilog` substitution table in
  `conf.py` going stale relative to class renames and new additions
  (big offenders: `|ErrorBars|` ×26, `|AnimationEffect|` ×11,
  `|ErrorBarType|` ×9, `|ErrorBarInclude|` ×9, `|ErrorBarDirection|`
  ×9, `|PathGeometry|` ×8, `|Path|` ×8). Non-substitution warnings: 4
  `Title underline too short`, 2 `Explicit markup ends without a blank
  line`, 1 `document isn't included in any toctree`, 1 `Literal block
  ends without a blank line`, 1 `Malformed table`.
- xlsx: build fails before docstrings are ever read, so docstring
  quality cannot be sampled this way.

**napoleon-compatibility evidence (docx only)**

docx's `conf.py` explicitly annotates — in prose — that ~11 files
reference OOXML element names using single-backtick inline code
(e.g. `\`w:moveFrom\``) and that this would produce ~42 bogus
docutils warnings without the pre-napoleon/post-napoleon hook. This
is a strong indicator that docx has been through at least one
deliberate docstring-cleanup pass; pptx has no equivalent hook and
no equivalent cleanup.



### 3.9 Spec / reference material

**python-docx — `spec/` at project root.**
Contains the ISO/IEC 29500 parts 1–4 as PDFs (`ISO-IEC-29500-1.pdf`
through `ISO-IEC-29500-4.pdf`), an `xsd/` tree (presumably the XML
schema definitions referenced during element implementation), an
`rnc/` tree (Relax-NG-compact equivalents), and a `styles.xml`
reference file. The spec tree is not exposed through Sphinx — it's a
developer-side resource checked in alongside source.

**python-pptx — `spec/` at project root.**
Organised one level deeper: `ISO-IEC-29500-1/`, `ISO-IEC-29500-2/`,
`ISO-IEC-29500-3/`, `ISO-IEC-29500-4/`, plus `gen_spec/`. The per-part
directories likely contain the PDFs plus extracted/searchable
artifacts — richer than docx's flat PDFs. `gen_spec/` suggests tooling
for regenerating spec-derived test fixtures or docs, which docx does
not appear to have.

**python-xlsx — no `spec/` at project root.**
No ISO-IEC-29500 PDFs, no schema tree, no RNC tree checked in. The
project carries its OOXML-aware descriptor layer (`src/xlsx/descriptors/`)
but not the normative spec alongside it. Given the fork targets Excel
365 features that are ECMA-376-second-edition-plus, the absence of the
spec tree is conspicuous.



### 3.10 FEATURES.md-equivalent

**python-docx — `FEATURES.md`, 1791 lines (81 KB).**
Plus a companion `FEATURES_AUDIT.md`, 41 KB. These two files between
them appear to be the authoritative inventory of what the fork has
delivered (`FEATURES.md`) and what's been validated against upstream
intent (`FEATURES_AUDIT.md`). This is a documentation surface docx has
uniquely invested in.

**python-pptx — no FEATURES.md.**
The closest equivalent is the `Feature Support` list inside
`docs/index.rst`, which is inherited from upstream: round-trip PPTX,
add slides, populate text placeholders, add image, add textbox, add
table, add auto shapes, toggle bullet formatting, add/manipulate
column/bar/line/pie charts, discover 2016+ extended charts and
preserve on round-trip, core document properties, header/footer/slide
number/date placeholder toggles, etc. The list does not advertise
fork-specific work (animation API, section API, transitions, smart-art,
tags, threaded comments).

**python-xlsx — no FEATURES.md, but `TODO.md` (582 lines).**
Role is different from docx's `FEATURES.md`: it is forward-looking
(what remains to do), not retrospective (what has been delivered). As
such it plays the part of a project board rather than a feature
inventory. The README's "extended with ... dynamic arrays and spill
semantics, threaded comments, rich text on cells, sparklines, modern
chart types..." pitch is the de-facto feature inventory for the fork.

Observation: only docx has a retrospective feature inventory, and it
has two (the main file + an audit). This is one of the biggest
documentation-surface divergences in the series.



### 3.11 Build health

**python-docx — build succeeds, 17 warnings.**

From `/tmp/sphinx-docx.log`: `build succeeded, 17 warnings.`

The warning shape (per the partial log sampled and the full sphinx log):
- 16 `ERROR: Undefined substitution referenced` for class names used
  with `|Name|` syntax that has no substitution definition — the
  offenders are `|AltChunk|` (×2), `|Attachment|`, `|ExtendedProperties|`,
  `|DocVars|`, `|ImagePart|` (multiple), `|Image|`, `|StoryPart|`
  (×2), `|TableCellMargins|` (×2), `|IndexError|`, and a handful
  more. All in docstrings in `src/docx/document.py`, `settings.py`,
  `shape.py`, `table.py`.
- 1 `WARNING: undefined label: 'wdoutlinelvl'` in
  `src/docx/text/parfmt.py` — a dangling `:ref:` target.
- 1 deprecation hint about `intersphinx_mapping` format (though
  `conf.py` already uses the new form — this appears to be a false
  positive from Sphinx 7 detection).

All 17 are cosmetic / dangling — no broken autodoc, no missing
modules, no malformed tables.

**python-pptx — build succeeds, 140 warnings.**

From `/tmp/sphinx-pptx.log`: `build succeeded, 140 warnings.`

The warning shape is dominated by the `rst_epilog` drift described in
§3.2:
- 127 `ERROR: Undefined substitution referenced` — top offenders are
  `|ErrorBars|` (26), `|AnimationEffect|` (11),
  `|XL_ERROR_BAR_TYPE|` (9), `|XL_ERROR_BAR_INCLUDE|` (9),
  `|XL_ERROR_BAR_DIRECTION|` (9), `|PathGeometry|` (8), `|Path|` (8),
  `|Sound|` (6), `|Section|` (5), `|Comments|` (4), `|Audio|` (4),
  `|_HeaderFooter|` (3), `|Transition|` (3),
  `|ConnectorAdjustmentCollection|` (3), `|SmartArt|` (2),
  `|PROG_ID|` (2), `|CommentAuthors|` (2), `|CommentAuthor|` (2),
  `|Comment|` (2), and long tail of singletons
  (`|_Field|`, `|XL_CROSS_BETWEEN|`, `|XL_AXIS_POSITION|`,
  `|TagsPart|`, `|SlideTags|`, `|Sections|`, `|Movie|`, `|MediaPart|`,
  `|LinePlot|`, `|ExtendedPropertiesPart|`, `|CustomProperties|`,
  `|AnimationEffectView|`, `|bool|`).
- 4 `Title underline too short` warnings.
- 2 `Explicit markup ends without a blank line; unexpected unindent`.
- 1 `document isn't included in any toctree`.
- 1 `Literal block ends without a blank line; unexpected unindent`.
- 1 `Malformed table`.

Every one of those substitution errors is a visible render defect in
the HTML output (literal `|ClassName|` will appear in prose where a
cross-reference should be).

**python-xlsx — build fails.**

Error text (from `/tmp/sphinx-xlsx.log`):

```
Running Sphinx v7.4.7

Configuration error:
There is a programmable error in your configuration file:

Traceback (most recent call last):
  File "/tmp/sphinx-venv-xlsx/lib/python3.14/site-packages/sphinx/config.py",
    line 529, in eval_config_file
    exec(code, namespace)
  File "/home/ben/code/python-xlsx/doc/conf.py", line 25, in <module>
    import openpyxl
ModuleNotFoundError: No module named 'openpyxl'
```

Root cause: the Sphinx config still expects an `openpyxl` top-level
package, but the fork renamed the package. `src/xlsx/` and
`openpyexcel/` coexist at project root; neither is named `openpyxl`.
The `conf.py` reads `release = openpyxl.__version__` and uses
`openpyxl.__author__` in its `copyright`, so even if the import were
shimmed, multiple downstream lines would need to follow.

No partial build output is produced — the tool exits at configuration
stage, before source files are read. That means no HTML landing page,
no rendered tutorials, no published reference. The README's claim
that the project is installable and usable is plausibly independent
of this — the Sphinx docs surface is simply broken.

Secondary risk: there is no `requirements-docs.txt`, so there is no
pinned way to reproduce this build reliably even once the import is
fixed.



### 3.12 Code samples

**python-docx.** The README has no inline code block surfaced in the
first 20 lines sampled. The Sphinx `docs/index.rst` landing page,
however, carries an embedded quickstart snippet inside a two-column
layout — an example image on the left, Python code on the right. The
code goes beyond upstream's quickstart by adding fork-only calls:

```python
# -- fork feature: attach a footnote to a run --
document.footnotes.add(p.runs[0], 'Footnote body text.')

# -- fork feature: attach a comment to a range of runs --
document.add_comment(
    runs=p.runs,
    text='A reviewer comment.',
    author='Editor',
    ...
)
```

The code-sample strategy here is *lead with fork-value immediately on
the landing page*.

**python-pptx.** `docs/index.rst` relies on `include:: ../README.rst`
for its intro — so the Sphinx landing page is essentially the 25-line
upstream README. A `Feature Support` bullet list (inherited from
upstream) gives functional scope but no `.py` sample. There is a
`lab/` directory at project root, suggesting experimentation fixtures;
not examined in this audit.

**python-xlsx.** `doc/` contains working `.py` files alongside the
`.rst` documents — `example.py`, `filters.py`, `format_merged_cells.py`,
`table.py`. These are presumably referenced from the corresponding
`.rst` pages via `.. literalinclude::` or similar. This is a different
and arguably more honest approach (the example source is executable
and linted), but it puts `.py` files in a `doc/` tree — architecturally
unusual.



## 4. Divergences worth aligning

The following are observed differences across the three projects that
stand out as *avoidable* — places where the siblings disagree on
format or presence of a surface without any obvious per-project reason.

1. **`docs/` vs `doc/`.** docx and pptx use `docs/` (plural). xlsx
   uses `doc/` (singular). One of the two conventions is inherited
   from openpyxl's 2019-era layout; the other is from the
   docx/pptx upstream lineage.

2. **README format.** docx and xlsx use `README.md`. pptx uses
   `README.rst`. Consequence: pptx's `docs/index.rst` can and does
   `include:: ../README.rst`; docx's `index.rst` cannot do the same
   trivially.

3. **LICENSE filename.** docx and pptx use `LICENSE`. xlsx uses
   `LICENCE.md` (British spelling + `.md` extension). Packaging
   tools that look for `LICENSE` or `LICENSE.*` may miss it.

4. **Sphinx stack generation.** docx is on `Sphinx>=6,<8` + `furo`.
   pptx is pinned at `Sphinx==1.8.6` + `alabaster<0.7.14` + `armstrong`.
   xlsx's `conf.py` was written for an openpyxl-era Sphinx (pre-RTD).
   Three generations of Sphinx in one "series".

5. **`requirements-docs.txt` presence.** docx and pptx have one. xlsx
   does not. This is what prevented reproducing the xlsx build
   reliably.

6. **`versionadded::` usage.** docx has 885 occurrences in `src/`.
   pptx and xlsx have zero. Only docx will be able to auto-render a
   "new in 2026.05" badge in its docs.

7. **Fork banner in README.** docx and xlsx both lead with "A Python
   library for reading, creating, and updating Microsoft <app> 2007+
   files." followed by a "Based on ... Forked at / from ...
   Unstable. Not yet published to PyPI. Current version: 2026.05.0"
   block. pptx's `README.rst` retains the upstream opening
   ("*python-pptx* is a Python library for creating, reading, and
   updating PowerPoint (.pptx) files.") and never mentions the fork.

8. **Fork-era changelog.** docx `HISTORY.rst` starts with a "first
   release as independent fork" banner. pptx `HISTORY.rst` starts
   with an `Unreleased` section that reads as a contributor log and
   has no fork-transition marker. xlsx's `doc/changes.rst` starts
   at `2.5.14 (2019-01-23)` — it's the unmodified openpyxl log.

9. **Sibling-org names.** docx's `CLAUDE.md` names the sibling org
   as `loadfix/python-docx`, `loadfix/python-pptx`,
   `loadfix/python-xlsx`. pptx's `CLAUDE.md` names them with
   upstream maintainers (`scanny/python-docx`, etc.). The three
   CLAUDE.md files do not agree on what this series is called.

10. **`api/` and `api/enum/` subtrees.** docx and pptx both split
    reference docs into `api/` + `api/enum/`. xlsx has no such
    split — all 39 `.rst` files are flat. There is no class-by-class
    reference surface for xlsx.

11. **`FEATURES.md`.** Only docx has one. For a series whose pitch is
    "these forks add N new features", the absence of retrospective
    feature inventories on two of the three projects matters.

12. **`spec/` presence.** docx and pptx both check in ISO-IEC-29500
    spec PDFs alongside source. xlsx does not. For a project whose
    value proposition is fidelity to Excel-365 OOXML parts, this is
    surprising.

13. **Napoleon docstring handling.** Only docx has configured
    Napoleon (and even installed a hook to tame it against OOXML
    colon-in-backtick attribute docstrings). pptx uses autodoc
    without Napoleon. xlsx uses autodoc without Napoleon. Docstring
    conventions in the codebase are therefore different.

14. **Build health.** docx: 17 cosmetic warnings. pptx: 140 mostly-
    substitution errors (128 of them visible as literal `|Name|` in
    rendered HTML). xlsx: does not build at all.

15. **CONTRIBUTING.md / AUTHORS.md.** xlsx has both. docx and pptx
    have neither. (xlsx's `CONTRIBUTING.md` is a 17-line stub; the
    presence is the point, not the content depth.)



## 5. Conventions that should be identical across the series

These are the things that the audit found are *already* identical in
at least two of the three projects, and which therefore implicitly
define the intended series-wide convention.

- **Top-level layout skeleton:** `src/<pkg>/`, `tests/`, `features/`
  (behave acceptance), `spec/`, `docs/`, `CLAUDE.md`, `HISTORY.rst`,
  `LICENSE`, `MANIFEST.in`, `Makefile`, `pyproject.toml`, `tox.ini`,
  `requirements-*.txt`. docx and pptx match this in full. xlsx
  diverges on `doc/` (not `docs/`), `LICENCE.md` (not `LICENSE`),
  absence of `spec/`, absence of `features/`, presence of legacy
  `openpyexcel/` directory alongside `src/xlsx/`, and `pytest.ini` +
  `setup.cfg` + `setup.py` living alongside a nominally modern layout.

- **`docs/` structure:** `docs/index.rst`, `docs/conf.py`,
  `docs/Makefile`, `docs/_static/`, `docs/api/`, `docs/api/enum/`,
  `docs/user/`, `docs/dev/`. docx and pptx match. xlsx does not have
  any of these subdirs.

- **API-reference organisation:** one `.rst` per proxy class/topic
  under `docs/api/`, one `.rst` per enum under `docs/api/enum/`,
  filename matches class name (docx: `WdAnchorH.rst`; pptx:
  `MsoAutoShapeType.rst`). xlsx absent.

- **User-guide page topic naming:** both docx and pptx have
  `quickstart`, `install`, `charts`, `comments`, `text`, `table`
  (sometimes `tables`). docx additionally splits `-advanced` for the
  larger sections. xlsx has `tutorial.rst` / `usage.rst` instead —
  different convention, no "quickstart" page.

- **Sphinx autodoc + intersphinx:** all three enable `autodoc` and
  `viewcode`. docx adds `intersphinx` and `napoleon`; pptx adds
  `intersphinx`, `doctest`, `coverage`, `inheritance_diagram`,
  `todo`, `ifconfig`; xlsx adds `doctest`, `coverage`, `ifconfig`.
  The minimum common set is `autodoc` + `viewcode`.

- **CalVer version `2026.05.0`:** docx and xlsx. pptx: not confirmed
  by the audit — pptx `HISTORY.rst` opens with `Unreleased` without a
  CalVer label, so either pptx is still on semver or has not yet cut
  the first CalVer tag.

- **Language framing of the CLAUDE.md intro:** "Guidance for Claude
  Code (and other AI assistants) working in this repository" —
  pptx and xlsx open identically. docx opens differently
  ("python-docx fork (loadfix/python-docx) — extending python-docx
  with ..."). If docx is the "canonical" one in terms of fork
  maturity, its CLAUDE.md intro is the odd one out stylistically.

- **Three-layer architecture diagram in CLAUDE.md:** docx, pptx, and
  xlsx all describe their architecture as a three-layer stack
  (Document API / Parts Layer / oxml Layer over lxml).



## 6. What this repo (python-docx) should consider

Observations specific to this repo (no prescription of action — just
what the audit noticed).

- The 17 Sphinx warnings are all cosmetic (dangling `|Name|`
  substitutions + one stale `:ref:` target, `wdoutlinelvl`). They are
  the smallest warning count in the series and individually small,
  but they are visible in rendered HTML as literal `|AltChunk|`,
  `|Attachment|`, `|ImagePart|` text. Fixing them would bring the
  build to zero-warning, which pptx would then be the lone warning
  carrier.

- `docs/conf.py` includes an 80-line prose explanation of the
  pre-/post-napoleon hook. The explanation is thorough and worth
  preserving; it also makes this `conf.py` the obvious template for
  porting to pptx if/when pptx moves off Sphinx 1.8.6.

- `CLAUDE.md` is 166 lines — the shortest of the three, by a large
  margin. It is also the only one that opens with fork narrative
  rather than the "Guidance for Claude Code" boilerplate. If the
  series aims for consistent AI-assistant entry points, this repo's
  opening section is stylistically distinct.

- `FEATURES.md` (1791 lines) and `FEATURES_AUDIT.md` (41 KB) are
  unique to this repo. They represent substantial curation effort;
  the audit file suggests a second round of self-verification has
  already been done.

- No `CONTRIBUTING.md` and no `AUTHORS` file at project root. xlsx
  has both (even if stub-sized). For an "independent fork" status,
  the lack of a contribute-here doorway is a notable surface gap.

- `docs/user/` has 38 pages and covers almost every fork-era
  subsystem in prose — the most complete user guide in the series.

- `docs/dev/` has only `analysis/` underneath it. pptx's `docs/dev/`
  has `development_practices.rst`, `philosophy.rst`, `security.rst`,
  `xmlchemy.rst` in addition to `analysis/`. These documents describe
  the *how and why* of contributing — surface absent here.

- The README is 131 lines of Markdown, fork-focused. It does not
  appear to point readers at the Sphinx docs (no "Full docs at
  readthedocs/..." line visible in the first 20 lines). The Sphinx
  build is unpublished as far as the audit could tell.



## 7. What the series as a whole should consider

Observations about the three-project ensemble, not any one member.

- **There is not yet a single series-level landing page** that names
  all three projects and routes users to the right one. docx's
  `CLAUDE.md` names the series `loadfix/python-{docx,pptx,xlsx}` but
  pptx's `CLAUDE.md` names the siblings with upstream-maintainer
  prefixes (`scanny/...`, `sciris/...`). A reader arriving from
  search has no way to discover the other two projects from the
  README of any one of them.

- **Three different Sphinx theme choices** (`furo`, `armstrong`,
  `nature`) across three projects with identical intended audience.
  The rendered HTML will look unrelated.

- **Three different changelog formats** — phase-banded CalVer (docx),
  engineering-working-log with `Unreleased` section (pptx),
  openpyxl-2019-snapshot (xlsx). A user tracking across the three
  sees three entirely different release narratives.

- **Three different fork-transition postures** in the READMEs —
  explicit "first release as independent fork" banner (docx, xlsx
  with nearly identical wording) vs. no fork mention whatsoever
  (pptx).

- **Spec material included in docx and pptx, absent in xlsx.** The
  ISO-IEC-29500 PDFs plus `xsd/`, `rnc/` trees anchor the OOXML
  implementation in both existing `spec/` dirs; xlsx's work on Excel
  365 features is done without this anchor.

- **`versionadded::` annotations are docx-only (885 of them).** The
  two siblings cannot render "new in 2026.05.0" badges because no
  directives exist in their source. Cross-sibling consistency on
  fork-era version tagging would require pptx and xlsx to backfill.

- **Build health is uneven.** docx 17 warnings (mostly cosmetic),
  pptx 140 warnings (mostly substitution drift, visible in HTML),
  xlsx build fails at `conf.py`. For a reader trying to tell whether
  a given library is production-grade, the rendered-docs quality is
  the first contact point, and the three contact points are in three
  different states.

- **Enum documentation convention is present in docx and pptx
  (one-file-per-enum under `api/enum/`) but absent in xlsx.** The
  convention is consistent between the two that use it.

- **Behave acceptance tests** (`features/` dir) are in docx and pptx.
  xlsx has no `features/` dir — it uses pytest-only, though `tests/`
  does exist. Testing philosophy is not uniform, and that shows up
  indirectly in documentation (what kind of examples the project
  treats as normative).

- **`CLAUDE.md` line counts** — 166 (docx), 256 (pptx), 334 (xlsx).
  xlsx's is the richest, docx's is the terse "shape-only" version.
  For AI agents bouncing between the three repos, the guidance-depth
  is not symmetric.

- **Documentation-as-source (executable `.py` examples in `doc/`)**
  is an xlsx-only convention; the other two put examples as
  RST-inline snippets. This is a cross-sibling style divergence at
  the smallest scale.

- **Each project's docstring substitution strategy differs.** docx
  writes inline class references directly (`:class:`.Name``) and
  gets substitution drift only in a handful of places (~17). pptx
  relies on `rst_epilog` in `conf.py` declaring `|ClassName|` style
  shortcuts, and that list has gone substantially stale relative to
  the code (~127 undefined). xlsx does neither because it does not
  build. The three projects have three different docstring styles as
  a side effect.

