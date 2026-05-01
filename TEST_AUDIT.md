# Test Suite Audit (Issue #83)

This report surveys the state of the `loadfix/python-docx` test suite with three aims:

1. Document what is well covered and what is not.
2. Identify pre-existing latent defects and anti-patterns.
3. Propose concrete, prioritised follow-ups.

All tests were run with `uv run pytest` against commit `50c2078` (`master`). The
baseline is **4058 passed, 1 skipped, 28 deselected** (the 28 deselected
failures are the pre-existing `CT_Border` / `BorderElement` tests investigated
in section 3 below).

Behave acceptance tests pass cleanly: **67 features / 650 scenarios / 1856
steps, 0 failures** (`uv run behave features/` in ~2s).

---

## 1. Coverage summary

Run:

```
uv run pytest --cov=docx --cov-report=term-missing tests/ \
  --deselect tests/test_table.py::DescribeBorderElement \
  --deselect tests/oxml/test_table.py::DescribeCT_Border \
  --deselect tests/oxml/test_table.py::DescribeCT_TblBorders \
  --deselect tests/oxml/test_table.py::DescribeCT_TcBorders
```

**Overall coverage: 97 %** (15 901 statements, 489 missed).

Counting tests, the suite comprises 404 `Describe*` classes across 104 test
modules and ~37 241 lines of test code (`tests/` tree, excluding fixtures
directories).

### 1.1 Lowest-coverage production modules

The modules with the lowest coverage percentages and a terse note about the
missing lines.

| % | module | stmt / miss | uncovered (representative) |
|---:|---|---:|---|
| 68 | `src/docx/enum/base.py` | 71/23 | `DocsPageFormatter` (lines 88-150) — RST doc-generation tool, exercised only by the `Makefile` docs target |
| 83 | `src/docx/numbering.py` | 154/26 | error branches in `_normalize_format` (69-71) and `_normalize_level_spec` (93-95, 104, 114); `Numbering.element`/`part` props (145, 149); all of `_num_id_for` reuse loop (211-218); `NumberingDefinition.element` (239); all of `apply_to` including bounds guard (259-266); `Level.indent` None branches (308, 311); `Level.element` (316) |
| 85 | `src/docx/oxml/shared.py` | 20/3 | `CT_String.new` classmethod (50-52) |
| 86 | `src/docx/image/svg.py` | 85/12 | UTF-8 decode failure (31-32), XML parse failure (43-44), viewBox value errors (66-67), unit match miss (78), `pt`/`cm`/`mm` unit branches (86, 89-92) |
| 88 | `src/docx/ids.py` | 17/2 | paragraph-without-`w:id` fallback (51-52) |
| 88 | `src/docx/oxml/content_controls.py` | 181/22 | `tag_val`/`alias_val` None-return branches (62, 75, 88, 91, 94, 97-98, 113, 124, 130, 134, 180); `CT_SdtPr.tag_val` / `alias_val` None-remove branches (231-233, 243-245); `CT_SdtContent.text` nested-SDT branch (319, 324, 341-343) |
| 88 | `src/docx/oxml/simpletypes.py` | 282/35 | many `@classmethod validate` error branches (e.g. `ST_EighthPointMeasure.validate`, `ST_HexColor` with missing hex forms, `ST_DecimalNumber` negative paths); lines 282-289 cover `ST_Merge.validate` error paths |
| 88 | `src/docx/parts/story.py` | 52/6 | SVG floating-image path (99-102) and `_new_svg_pic_inline` (126-128) |
| 90 | `src/docx/oxml/endnotes.py` | 59/6 | `next_available_id`: unsigned wrap-around (72), and full enumerate-for-hole fallback (78-83) — unreachable in practice |
| 90 | `src/docx/shape.py` | 155/15 | `FloatingShape.horizontal_offset`/`vertical_offset` `ValueError` branches (178-179, 192-193); `alt_text` / `title` `docPr is None` branches (229, 247); type-dispatch fallbacks for `CHART`, `SMART_ART`, `NOT_IMPLEMENTED` (268, 270-274) |
| 90 | `src/docx/signatures.py` | 70/7 | `_extract_signer` / `_extract_signed_at` exception branches (107-108, 119, 123-124); ISO-8601 `Z`-fallback ValueError (142-143) |
| 91 | `src/docx/form_fields.py` | 255/22 | `_val_attr` absent-attr branch (60); `_bool_val` true/false-no-val branch (73); `_int_val` default branches (83, 86, 89-90); many type-specific property None branches in `TextInputFormField`/`CheckboxFormField`/`DropdownFormField` (117, 124, 148, 177, 202, 208, 223, 230, 239, 247, 255, 267, 275); `result` dropdown out-of-range (323, 331); trailing run-result text branch (404) |
| 91 | `src/docx/oxml/footnotes.py` | 67/6 | symmetric with `endnotes.py` (72, 78-83) — unreachable fallback |
| 92 | `src/docx/oxml/styles.py` | 227/18 | `next_available_numId` / `next_available_num_* ` helpers (145-148, 167-170, 192-195); `_update_num_val` None branches (216, 258, 264-267) |
| 92 | `src/docx/oxml/text/pagebreak.py` | 90/7 | `preceding_paragraph_fragment` / `following_paragraph_fragment` edge-cases when there is no sibling content (140, 151, 164, 179, 192, 215, 244) |
| 92 | `src/docx/package.py` | 98/8 | VBA-project / macro wiring branches (50, 76, 81-82, 85, 92); `_next_partname` numeric reuse (156-157) |

### 1.2 Notable modules at 93 %-96 %

A handful of proxy/oxml modules sit in the 93-96 % band. The missing lines
are typically defensive None-returning branches and a few edge cases:

- `src/docx/oxml/shape.py` (93 %, 17/241 missed) — lines 626-678 are the
  anchor-position reset paths when positional attributes are absent (Phase
  D.17 floating-image code). Worth a couple of additional parametrized tests.
- `src/docx/tracked_changes.py` (93 %, 11/157) — lines 152, 154, 313, 326,
  335, 364-376: uncovered setter branches when an attribute is being removed.
- `src/docx/fields.py` (94 %, 11/186) — `Field.result_text` setter with an
  empty field (line 148, 169), `add_field` whitespace branches (240-347).

### 1.3 Highest-risk "dense green" modules

Several large modules sit above 95 % but have large uncovered *ranges* worth
double-checking:

- `src/docx/section.py` (96 %, 24/653 missed): lines 995-998, 1019-1022,
  1039-1042, 1118-1121 are contiguous blocks — usually a signal that one
  branch of an entire `if` was never hit. Worth a quick read.
- `src/docx/table.py` (97 %, 27/781 missed): line range 977-991 (14 lines) is
  one contiguous dead region in the border-style write path — almost
  certainly related to the `BorderElement`/`CT_Border` bug in section 3.
- `src/docx/text/paragraph.py` (96 %, 20/448 missed): scattered but includes
  831-842 and 886-895 which look like two whole branches.

---

## 2. Coverage gaps — module-by-module notes

Ranked by severity (production impact × uncovered lines):

**`docx/numbering.py` (83 %)** — tested happy-path construction of numbering
definitions only (`tests/test_numbering.py`, 199 lines). `apply_to()` — the
*only* API exposed for applying a numbering definition to a paragraph — is
completely uncovered (src lines 253-266), as is the num-id reuse loop in
`_num_id_for` (211-218), the level-out-of-range error guard, and all of the
level-spec validation error paths. Highest-value coverage gap in the repo.

**`docx/form_fields.py` (91 %)** — all four form-field proxy classes have
"absent element returns X" branches that are untested. Given form-fields are
read-mostly from third-party documents, these fallbacks are load-bearing.
22 missed lines spread across 10+ small accessors.

**`docx/shape.py` (90 %)** — `FloatingShape` type-dispatch for
CHART/SMART_ART/NOT_IMPLEMENTED (268-274) is uncovered, and both
`horizontal_offset`/`vertical_offset` `ValueError` paths (178-179, 192-193)
are uncovered — these are the fallbacks when a document contains a malformed
`wp:posOffset` text value. Worth an adversarial-input test.

**`docx/signatures.py` (90 %)** — every `except Exception` clause in
`_extract_signer` / `_extract_signed_at` is untested. The code explicitly
swallows exceptions for robustness; parametrised tests with broken XML would
pin those down.

**`docx/oxml/content_controls.py` (88 %)** — repeated absent-child
`return None` branches on every getter. Covering them requires only a
single bare-`w:sdt` element fixture.

**`docx/image/svg.py` (86 %)** — non-UTF-8 stream, malformed XML,
non-numeric viewBox values, and `pt`/`cm`/`mm` length units are all
uncovered. These matter because `Document.add_picture` delegates to this
branchy parser.

**`docx/enum/base.py` (68 %)** — the `DocsPageFormatter` class is used only
by `Makefile` docs targets; it's legitimately internal tooling. Either move
it into `docs/` (where it won't pollute coverage stats) or add a smoke test
calling `DocsPageFormatter("WD_FOO", WD_FOO.__dict__).page_str`.

**`docx/parts/story.py` (88 %)** — the `_new_svg_pic_inline` / floating-
image SVG-fallback paths (99-102, 126-128) are only reachable from
`Document.add_picture` with an SVG file. A small integration test covering
this matters because the SVG-fallback pipeline is fragile.

**`docx/oxml/footnotes.py` / `docx/oxml/endnotes.py`** — in both modules,
lines 72, 78-83 are the "all 2**31 ids used, enumerate to find the hole"
fallback. Effectively unreachable at real-world scale; not worth covering.

**`docx/oxml/simpletypes.py` (88 %)** — `validate` error branches
(`ST_EighthPointMeasure`, `ST_HexColor`, `ST_Merge`). These are reached only
from malformed XML input; parametrised `raises` tests would plug most gaps
trivially.

Well-covered areas worth calling out as healthy: `docx.text.parfmt` (100 %),
`docx.text.font` (99 %), `docx.opc.*` (99 %+), `docx.image.jpeg`, `.png`,
`.tiff`, `.bmp`, `.gif`, `.helpers` all 100 %, `docx.oxml.settings` (100 %),
`docx.search` (100 %), `docx.oxml.theme` (100 %), and most of the recently-
added Phase-D modules (watermark, web_settings, ruby, ink, embedded_objects,
captions, content_controls at 95 %).

---

## 3. Pre-existing deselected failures

**28 tests fail** in the border-element tests:

- `tests/oxml/test_table.py::DescribeCT_Border` (parametrised; all 4 param
  groups fail — `val`, `sz`, `color`, `space`)
- `tests/oxml/test_table.py::DescribeCT_TblBorders`
- `tests/oxml/test_table.py::DescribeCT_TcBorders`
- `tests/test_table.py::DescribeBorderElement`

### 3.1 Root cause

There are **two distinct `WD_BORDER_STYLE` enums** in the codebase:

- `src/docx/enum/table.py:243` — `SINGLE=1, DOUBLE=2, DOTTED=3, ...`
- `src/docx/enum/text.py:274` — `NIL=0, NONE=1, SINGLE=2, THICK=3, DOUBLE=4, ...`

Two distinct `CT_Border` classes also exist, each binding one of the enums via
`OptionalAttribute("w:val", WD_BORDER_STYLE)`:

- `src/docx/oxml/table.py:72` uses the **table** enum.
- `src/docx/oxml/text/parfmt.py:45` uses the **text** enum.

In `src/docx/oxml/__init__.py`, the `w:top`, `w:left`, `w:bottom`, `w:right`
element-class registrations are made **twice**:

- lines 377-382 bind them to `oxml.table.CT_Border`.
- lines 544-564 (later in the same file) **overwrite** the same tag names
  with `oxml.text.parfmt.CT_Border`, because the `w:pBdr` block also uses
  those tags.

Because `register_element_cls` last-write-wins, every `w:top` (etc.) element
parsed from XML becomes an instance of `parfmt.CT_Border`, which uses the
`enum.text.WD_BORDER_STYLE`. The tests in `tests/oxml/test_table.py` import
`WD_BORDER_STYLE` from `docx.enum.table` and compare against objects whose
`val` comes back as `enum.text.WD_BORDER_STYLE.SINGLE` — same name, different
integer value, different class.

Verified at runtime:

```
>>> from docx.oxml.parser import parse_xml
>>> from docx.oxml.ns import nsdecls
>>> el = parse_xml(f'<w:top {nsdecls("w")} w:val="single"/>')
>>> type(el.val).__module__
'docx.enum.text'
>>> el.val
<WD_BORDER_STYLE.SINGLE: 2>
```

### 3.2 Proposed fix (not implemented)

Two viable approaches:

1. **Consolidate** the two enums into one module (either `enum/table.py` or a
   new `enum/border.py`), re-export from both modules for back-compat, and
   unify the two `CT_Border` classes into one element class. This is the
   cleanest but requires either renumbering (a semver breaking change) or
   picking one numbering and deprecating the other.

2. **Namespace-separate** the element registrations: since `w:top` has the
   *same tag name* but a *different parent* in the two cases (`w:tblBorders`
   / `w:tcBorders` vs. `w:pBdr`), introduce two `lxml.CustomElementClass`
   lookups discriminated by parent — e.g. register them as distinct
   `class_lookup` entries keyed on `parent.tag`. `lxml` supports this via
   `ElementNamespaceClassLookup` fallbacks, but the current codebase uses a
   flat `element_class_lookup` (see `src/docx/oxml/parser.py`). This would
   require a parser refactor.

Short-term pragmatic fix: repoint `tests/oxml/test_table.py` and
`tests/test_table.py::DescribeBorderElement` to import `WD_BORDER_STYLE` from
`docx.enum.text`, and adjust the expected XML values accordingly. This would
un-deselect the tests but permanently enshrines the `enum.text` numbering as
the canonical one — which the `enum/table.py:243` author presumably did not
intend. Not recommended without maintainer sign-off.

**Recommended follow-up:** open a dedicated GitHub issue for the
CT_Border/WD_BORDER_STYLE conflict; the proper resolution is a design
decision, not a test fix.

---

## 4. Test quality findings

### 4.1 Tautological / low-value tests

The code review spot-checked ~20 test modules for tautologies (tests that
only assert what a mock was configured to return). The suite is, on the
whole, **clean** — the cxml helper plus real-element fixtures keep most
tests grounded in behaviour. A few borderline cases:

- **`tests/test_document.py:231-235`** — `it_provides_access_to_the_comments`
  sets `document_part_.comments = comments_` and then asserts
  `document.comments is comments_`. This is a pure "the property forwards"
  test; value is low but not zero (confirms the getter isn't hardcoded).
  Pattern repeats ~5 times in that file (e.g. `its_core_properties`,
  `its_settings`).

- **`tests/test_section.py:1011-1026`** — the `_get_or_add_definition` test
  cluster mocks four of the method's collaborators (`_has_definition_prop_`,
  `_prior_headerfooter_prop_`, `_add_definition_`, and returns a mocked
  `header_part_`) and then asserts the correct collaborator was called.
  This is an "interaction test" — brittle to internal refactoring. Worth
  leaving as-is for legacy code but not a pattern to propagate.

- **`tests/test_accessibility.py:217-249`** —
  `DescribeDocument_validate_heading_structure::it_calls_the_module_function_with_document_paragraphs`
  patches `docx.accessibility.validate_heading_structure` and asserts the
  mock was called with the right arguments *and* `result is mock.return_value`.
  This is a pure-delegation test; the `return_value` check is a tautology.
  Dropping the `result == mock.return_value` assertion would tighten it.

### 4.2 Over-mocked tests

- **`tests/test_section.py:980-1026`** — `DescribeBaseHeaderFooter::_get_or_add_definition`
  replaces four collaborators with property/method mocks. Refactoring the
  class will break these tests even if behaviour is preserved.

- **`tests/test_custom_xml.py:270-279`** — a class-level `import pytest`
  inside the test class followed by a `document_part_` fixture nested within
  the class. Awkward placement; should be a top-level fixture or (better) a
  conftest fixture (see section 7).

### 4.3 Order-dependent tests

Spot-check found **none**. The `fake_parent` and `blank_document` fixtures in
`tests/conftest.py` are function-scoped, and the cxml helpers produce fresh
elements per test. The `tmp_docx_path` fixture uses `tempfile.mkstemp` with
per-test cleanup.

### 4.4 Misc test-hygiene notes

- **`tests/oxml/test__init__.py:147-148`** — contains a stray `class
  CustElmCls(BaseOxmlElement): pass` inside a test module. It's a fixture
  for the tests above but is placed in an out-of-the-way "static fixture"
  comment block at the bottom; looks accidental.

- **`tests/opc/test_pkgreader.py:479-481`** — `try/except: pass` block. On
  inspection it is *intentionally* swallowing an expected `TypeError` — but
  `pytest.raises(TypeError)` would be clearer.

- **`tests/helpers/libreoffice.py:57`, `helpers/validate.py:158`,
  `helpers/schema.py:63,108`, `helpers/roundtrip.py:37,59`** — all use
  `try/except` but these are *helpers*, not test assertions; appropriate.

- **`tests/test_strategy.py:305,351,369`** — `try/.../finally: os.unlink`
  patterns where a context manager or `tmp_docx_path` fixture would be
  cleaner. These could be collapsed to the existing fixture.

---

## 5. Outdated patterns

**`unittest.TestCase` usage:** **none**. Only `unittest.mock` is imported,
which is the pytest-idiomatic use. The codebase is cleanly pytest-native.

- `tests/test_custom_xml.py:8` — `from unittest.mock import MagicMock` —
  fine; could use the repo's own `unitutil.mock` helpers instead (only
  place `MagicMock` is imported directly).
- `tests/test_accessibility.py:222` — `from unittest.mock import patch`
  inside a method. Should be at module top.
- `tests/unitutil/mock.py:6` — appropriate (utility wrapper).

**`setUp`/`tearDown`:** **none**.

**Try/except where `pytest.raises` would be cleaner:**

- `tests/test_docm.py:32-48`, `50-64` — both wrap `tempfile.NamedTemporaryFile`
  output in `try:/finally: os.unlink(tmp_path)` to clean up. Trivial: use the
  existing `tmp_docx_path` fixture (rename for `.docm`) or
  `tmp_path_factory`.

- `tests/test_strategy.py:302-319` — two `tempfile.mkstemp` calls with a
  `try:/finally: os.unlink`. Could use `tmp_path_factory` fixture.

No true "`try/except: fail('expected exception')`" anti-patterns were found
in the test suite.

**Deprecated pytest idioms:** none found. `pytest.mark.parametrize` is used
consistently; `pytest.raises` is the norm for exception assertions.

---

## 6. Behave acceptance tests

Location: `features/` (70 feature files plus `steps/` with 22 step files,
~4103 lines). All 67 features / 650 scenarios / 1856 steps pass in ~2s.

### 6.1 Coverage map

Features are grouped by subject prefix:

| prefix | topic | files | notes |
|---|---|---:|---|
| `api-` | top-level `docx.Document` API | 1 | smoke |
| `blk-` | `BlockItemContainer` | 3 | core; fine |
| `cmt-` | Comments | 2 | Phase D |
| `doc-` | Document API (sections, add-X, collections, settings, comments) | 12 | strong |
| `hdr-` | Header/Footer | 1 | ok |
| `hlk-` | Hyperlink | 1 | ok |
| `img-` | Image characterisation | 1 | ok |
| `num-` | Numbering | 1 | only `num-access-numbering-part.feature` — **no coverage** of Phase-D.9 `apply_to` / `add_numbering_definition` |
| `par-` | Paragraph | 9 | strong |
| `pbk-` | Page break | 1 | ok |
| `run-` | Run | 6 | strong |
| `sct-` | Section | 1 | ok |
| `shp-` | Inline shape | 2 | ok |
| `sty-` | Styles | 7 | strong |
| `tab-` | Tabs / tab-stops | 2 | ok |
| `tbl-` | Table | 9 | strong |
| `txt-` | Text/font | 4 | strong |

### 6.2 Gaps — acceptance tests that do **not** exist

No behave coverage for the following newer features, despite each being a
headline Phase-D/Phase-B addition:

- **Footnotes** (`docx.footnotes`) — no `fnt-*.feature`.
- **Endnotes** (`docx.endnotes`) — no `ent-*.feature`.
- **Bookmarks** (`docx.bookmarks`) — no `bkm-*.feature`.
- **Tracked changes** (`docx.tracked_changes`) — no `trk-*.feature`.
- **Fields** (`docx.fields`, legacy form fields) — no `fld-*.feature`.
- **Table of contents** (`docx.toc`) — no `toc-*.feature`.
- **Watermarks** (`docx.watermark`) — no `wmk-*.feature`.
- **Content controls / SDTs** (`docx.content_controls`) — no `sdt-*.feature`.
- **Custom XML parts** (`docx.custom_xml`) — no behave coverage.
- **Custom properties** (`docx.custom_properties`) — no behave coverage.
- **Ruby** (`docx.ruby`) — no behave coverage.
- **Ink annotations** (`docx.ink`) — no behave coverage.
- **Digital signatures** (`docx.signatures`) — no behave coverage.
- **Numbering.add_numbering_definition** / `apply_to` — only the legacy
  "access numbering part" feature exists.
- **Floating images** (Phase D.17) — covered only by `shp-inline-shape-*.feature`.
- **Watermark** (Phase D.23), **Table autofit / column widths** (Phase D.26),
  **Insert paragraph/table at position** (Phase D.13) — all untested by
  behave.

### 6.3 Behave fixture state

`features/steps/test_files/` contains the test-fixture `.docx` files. Size
(number of files) is reasonable and the content is tracked in git.
`features/environment.py` is a minimal `before_feature`/`after_feature`
boilerplate — **not** a hook for setup that requires external tools, so
behave runs in-process and quickly (~2 s total).

No obvious stale scenarios were detected. The 650 scenarios all exercise the
pre-fork legacy API surface; nothing newer than 2019 pytest-port era.

### 6.4 Recommendation

Behave is **healthy but stale**. Rather than retrofit scenarios for every
Phase-D feature (expensive and low-marginal-value given the strong unit+XML
tests), pick 2-3 flagship features (footnotes, tracked changes, TOC) and add
a small `*-props.feature` + `*-mutations.feature` pair for each, mirroring
the comments pattern (`cmt-props.feature`, `cmt-mutations.feature`).

---

## 7. Flaky-test risk analysis

Grep for common flake sources:

### 7.1 Wall-clock dependency

- **`tests/test_comments.py:117,122`** — `datetime.now(...)` bracketing the
  under-test call. Safe: the test asserts the timestamp falls in the
  `[before, after]` interval, which is robust to clock jitter.
- **`tests/opc/parts/test_coreprops.py:41`** — `dt.datetime.now(...) - core_properties.modified`
  used to assert recency. Similar shape; safe.

No `time.sleep` calls in `tests/`. No `monotonic()`, `perf_counter`, or
explicit deadline checks.

### 7.2 Filesystem state

`tempfile.mkstemp`/`NamedTemporaryFile` usage in:

- `tests/conftest.py:29-36` — the `tmp_docx_path` fixture. Correct: closes
  the fd, yields, unlinks after.
- `tests/test_docm.py:32-48,53-64` — two inlined variants; should use the
  shared fixture.
- `tests/test_strategy.py:302-319` — inlined `mkstemp` with
  `try/finally: os.unlink`; should use the shared fixture.
- `tests/helpers/roundtrip.py`, `helpers/libreoffice.py` — helpers, not
  tests.

No tests depend on the CWD. No tests depend on environment variables.

### 7.3 External resources

- **LibreOffice-backed tests** (`test_strategy.py::DescribeLayer5_LibreOfficeValidation`)
  — gated behind `pytest.mark.libreoffice` and `is_libreoffice_available()`
  checks; correctly `pytest.skip`s on machines without LibreOffice. 1 skip
  in the baseline run (likely this fixture under CI).
- **Reference-doc-dependent tests** (`test_strategy.py::DescribeLayer4_ReferenceComparison`)
  — gated behind `ref_docx_exists()` and skip cleanly when the reference
  `.docx` file is absent. The `tests/ref-docs/` directory currently
  contains only a README listing "planned" reference files.

### 7.4 Network

Zero tests touch the network. `grep -rn "urlopen\|requests\|http" tests/`
returns nothing test-relevant.

### 7.5 Summary

Flake risk: **low**. The two real concerns are (1) the redundant temp-file
boilerplate in `test_docm.py` and `test_strategy.py` (cleanliness, not
flakiness), and (2) the LibreOffice-gated tests, which are already
pragmatically skipped.

---

## 8. Missing conftest fixtures — duplication hotspots

The most-duplicated per-class fixture patterns across `tests/`:

| fixture shape | approx. duplicate count | suggestion |
|---|---:|---|
| `def document_part_(self, request): return instance_mock(request, DocumentPart)` | 25+ (e.g. `test_document.py`, `test_custom_xml.py:276`, `test_section.py:1039`, `parts/test_story.py:137`) | promote to `tests/conftest.py` as `document_part_`|
| `def parent_(self, request): return instance_mock(request, Table)` / `...Paragraph)` / `...BlockItemContainer)` | 12+ (e.g. `test_table.py:1120`, `test_blkcntnr.py:150`) | possibly per-subpackage conftest (table/block) |
| `def paragraph_(self, request): return instance_mock(request, Paragraph)` | 10+ (e.g. `test_blkcntnr.py:150`, `text/test_paragraph.py:*`) | conftest in `tests/text/` |
| `def part_(self, request): return instance_mock(request, XmlPart)` | 6+ (e.g. `opc/test_package.py:233`, `test_custom_properties.py:174`) | package-local conftest in `tests/opc/` |
| `def paragraph_format_(self, request): ...` | 3+ (`styles/test_style.py:742`) | local conftest |

Additionally:

- `tests/conftest.py` already exposes `fake_parent`, `tmp_docx_path`,
  `blank_document`. A natural extension is:
  - `document_part_` (mock of `DocumentPart`)
  - `paragraph_` (mock of `Paragraph`)
  - `run_` (mock of `Run`)
  - `part_` (mock of generic `XmlPart`)

These would not break any tests (the local overrides would still win); they
would remove ~200 lines of boilerplate from the test suite.

**Builder helpers** are already consolidated in
`tests/unitutil/cxml.py` (cxml element / xml expressions) and
`tests/unitutil/mock.py` (class_mock / instance_mock / method_mock /
property_mock). Those are in good shape.

---

## 9. Dead code / skipped tests

- **`tests/ref-docs/`** — documented but the directory contains no
  reference `.docx` files. All Layer-4 tests in `test_strategy.py` skip.
  Either commit the reference files (see `tests/ref-docs/README.md`), or
  remove the Layer-4 tests.

- **`tests/test_strategy.py::DescribeLayer5_LibreOfficeValidation`** —
  only runs when LibreOffice is available. CI either needs to install
  libreoffice-headless or mark this class as "local dev only".

- **`tests/oxml/test__init__.py:147-148`** — the `class CustElmCls` stub is
  still in the module; harmless but worth a comment explaining why.

No commented-out tests or orphaned `def test_*` lines were detected.

---

## 10. Recommendations (follow-up issue backlog)

Effort labels: **S** ≤ 1 day, **M** 1-3 days, **L** > 3 days.

### Correctness / blocker

1. **[L] Resolve the duplicate `WD_BORDER_STYLE` / `CT_Border` conflict.**
   Re-read section 3 — either consolidate the enums or parent-discriminate
   the element-class registration. Re-enable the 28 deselected tests once
   the underlying bug is fixed. This is the only item that represents a
   *production bug*, not just a test gap.

2. **[S] Move `BorderElement`, `BordersCollection`, and any other
   table-border writer code into a single module importing the canonical
   `WD_BORDER_STYLE`.** Part of #1, small when #1 is accepted.

### Coverage fills (mostly small, high ROI)

3. **[M] `docx.numbering.apply_to` has zero tests.** Add unit tests
   covering: paragraph-to-definition attach, level range validation,
   matching-num-id reuse vs. new-num-id creation. Also cover the
   positional/mapping `LevelSpec` error paths (`_normalize_format`
   TypeError; short positional tuple ValueError).

4. **[S] `docx.form_fields.*FormField` read-path None branches.** Add a
   single "bare `w:ffData`" XML fixture and parametrise across
   TextInput/Checkbox/Dropdown properties. Covers ~20 of the missed
   lines.

5. **[S] `docx.image.svg` parametric tests for units (`pt`, `in`, `cm`,
   `mm`), non-UTF-8 streams, and malformed XML.** Each case is a one-line
   fixture feeding `Svg.from_stream`.

6. **[S] `docx.shape.FloatingShape` `alt_text`/`title`/type-dispatch
   branches** (`tests/test_shape.py`) — add parametrised tests for CHART,
   SMART_ART, NOT_IMPLEMENTED; assert alt_text returns None when `docPr`
   absent.

7. **[S] `docx.signatures` malformed-XML paths** — feed a non-XML string
   and broken `<X509SubjectName>`/`<SigningTime>` fragments and assert the
   `except Exception` clauses return None rather than propagating.

8. **[M] `docx.oxml.content_controls` None-branch parametric coverage**
   (`tests/test_content_controls.py`). Ten or so one-liner parametrised
   cases would lift the module to ~99 %.

9. **[S] `docx.parts.story._new_svg_pic_inline` floating-image SVG test.**
   One integration test calling `Document.add_picture("sample.svg")` with
   a floating anchor.

10. **[S] `docx.oxml.simpletypes.validate` error branches.** Drive through
    `pytest.raises` for each simpletype's `validate` method. 30+ tests,
    each one-liner.

11. **[S] `docx.enum.base.DocsPageFormatter` smoke test** — a single test
    instantiating it against `WD_PARAGRAPH_ALIGNMENT.__dict__` and
    asserting the returned string starts with `.. _`.

### Behave fills

12. **[M] Add behave coverage for footnotes, endnotes, and tracked
    changes.** Mirror `cmt-props.feature` + `cmt-mutations.feature`.
    ~150 lines of `.feature` files + ~200 lines of steps per topic.

13. **[S] Add a `num-define-and-apply.feature` for
    `Numbering.add_numbering_definition` and `NumberingDefinition.apply_to`.**
    Hooks straight into the Phase-D.9 gap flagged in section 2.

14. **[S] Add `toc-*.feature`, `wmk-*.feature`, `fld-*.feature`.**
    Smoke-coverage-only; each 1-2 scenarios.

### Test infrastructure

15. **[S] Promote common mock fixtures to `tests/conftest.py`.** Add
    `document_part_`, `paragraph_`, `run_`, `part_` fixtures. Delete the
    local duplicates in `test_document.py`, `test_section.py`,
    `test_custom_xml.py`, `test_blkcntnr.py`, etc. Expected diff: -200
    lines of boilerplate.

16. **[S] Replace ad-hoc `tempfile.mkstemp` in `test_docm.py` and
    `test_strategy.py` with the existing `tmp_docx_path` fixture (or add
    `tmp_docm_path`).**

17. **[S] Populate `tests/ref-docs/` with the planned reference files**
    (comments-simple, comments-threaded, comments-multi-author,
    comments-formatted). Either commit them or remove the Layer-4 scaffolding.

18. **[S] Add `libreoffice-headless` to the CI runner**, or remove the
    Layer-5 tests. Today, the 2 Layer-5 tests always skip in CI.

### Hygiene

19. **[S] Fix the 28 deselected tests after #1 lands** — they will start
    passing; remove the `--deselect` lines from CI configs.

20. **[S] Move `from unittest.mock import patch` from
    `tests/test_accessibility.py:222` to the module top.** Trivial.

21. **[S] Convert `try/except: pass` in `tests/opc/test_pkgreader.py:479-481`
    to `pytest.raises(TypeError)`.** Trivial; improves readability.

22. **[S] Audit `tests/test_document.py`'s "forward-and-assert" tests**
    (the `comments`, `core_properties`, `settings` group). Consider
    collapsing to a single parametrised "proxies expose the expected
    `part` attributes" test.

---

## Appendix A — full coverage output

See `pyproject.toml` for test configuration. To reproduce:

```
uv pip install pytest-cov
uv run pytest --cov=docx --cov-report=term-missing tests/ \
  --deselect tests/test_table.py::DescribeBorderElement \
  --deselect tests/oxml/test_table.py::DescribeCT_Border \
  --deselect tests/oxml/test_table.py::DescribeCT_TblBorders \
  --deselect tests/oxml/test_table.py::DescribeCT_TcBorders
```

Expected outcome: `4058 passed, 1 skipped, 28 deselected in ~33s`, overall
**97 %** line coverage.

To reproduce behave:

```
uv run behave features/
```

Expected: `67 features passed, 0 failed, 0 skipped / 650 scenarios passed /
1856 steps passed` in ~2s.
