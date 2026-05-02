# Upstream Issues & PRs Audit — python-openxml/python-docx vs loadfix/python-docx

This report classifies every open issue and pull request from the upstream
repository `python-openxml/python-docx` against the current state of the
loadfix fork at commit 37442d6.

Snapshot taken: 2026-05-02T05:14Z.

---

## Summary

| Verdict | Issues | PRs | Total |
|---|---:|---:|---:|
| resolved-in-loadfix | 217 | 47 | 264 |
| new-feature-needed | 118 | 22 | 140 |
| needs-investigation | 110 | 21 | 131 |
| out-of-scope | 113 | 20 | 133 |
| new-bug-needed | 26 | 15 | 41 |
| **Total** | 584 | 125 | **709** |

---

## Part 1 — Open Issues (584)

### upstream#1549 — is there an approach to reading semi-structured documents?

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** User asks how to walk paragraphs and tables in document order to extract semi-structured content from legislative bill analyses (upstream maintainer pointed to `Document.iter_inner_content()`).

**Evidence:**
- `src/docx/document.py:701` `Document.iter_inner_content()` delegates to `src/docx/blkcntnr.py:77`; `_Cell.iter_inner_content()` also exists.
- Upstream maintainer's answer already resolves it; no fork change required.

**TODO (if applicable):** none.

---

### upstream#1547 — Insert Picture Properties and change it to Top and Bottom

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** User wants to control picture wrap mode (topAndBottom) instead of inline when inserting images.

**Evidence:**
- `src/docx/enum/shape.py:66` defines `WD_WRAP_TYPE` with `TOP_AND_BOTTOM`.
- `src/docx/text/paragraph.py:442` `Paragraph.add_floating_image(..., position={"wrap": WD_WRAP_TYPE.TOP_AND_BOTTOM, ...})` — shipped in D.17 (#30).
- `HISTORY.rst` 1.3.0.dev0 lists "Floating images with wp:anchor positioning (#30)".

**TODO (if applicable):** none — may want a small docs example showing wrap=TOP_AND_BOTTOM.

---

### upstream#1546 — AI provenance metadata in DOCX custom XML parts

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** Asks whether customXml parts can be read/written, and whether a dedicated AI-provenance helper is desired.

**Evidence:**
- `src/docx/custom_xml.py` + `src/docx/parts/custom_xml.py` — `Document.custom_xml_parts` list with `item_id`, `schema_refs`, XML read access (`D.14` ecosystem).
- Dedicated AI-provenance helper is an application-level concern; maintainer's answer pattern is "use custom XML parts" — already provided.

**TODO (if applicable):** none (out-of-scope for a generic OOXML library to ship AI-specific helpers).

---

### upstream#1544 — Support for pathlib.Path

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** `Document.add_picture(path)` and related APIs should accept `pathlib.Path`, not just `str`.

**Evidence:**
- `src/docx/text/run.py:62` signature is `image_path_or_stream: str | IO[bytes]` — no `os.PathLike`.
- `src/docx/image/image.py:39` only checks `isinstance(image_descriptor, str)` to open from disk; a Path would be treated as a stream and fail.
- Fork hasn't widened any path-accepting APIs to `PathLike`.

**TODO (if applicable):** Widen `add_picture` / `Image.from_file` / `Document()` constructor to accept `os.PathLike`; apply `os.fspath()` at entry points. Effort: S.

---

### upstream#1542 — python-docx Document Timezone Issue

**Verdict:** `needs-investigation`

**Ask summary (1-2 sentences):** Report that core_properties.created/modified stored as UTC without tz conversion; Windows Explorer shows +8h. Essentially user-error (pass tz-aware datetime), but fork could normalize naive datetimes.

**Evidence:**
- `src/docx/oxml/coreprops.py` handles dcterms datetimes. No explicit fork fix for naive-datetime UTC tagging seen.
- Upstream has not fixed; debate about whether this is a bug or user error.

**TODO (if applicable):** Consider documenting that naive datetimes are treated as UTC; optionally emit DeprecationWarning for naive datetimes passed to `core_properties.created/modified`. Effort: S.

---

### upstream#1541 — NotImplementedError on numbering_part when file has no bullets

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** Opening a doc with no numbering part and accessing `doc.part.numbering_part` raises because upstream `NumberingPart.new()` is a stub.

**Evidence:**
- `src/docx/parts/numbering.py:24-34` — fork implements `NumberingPart.new()` returning a real part with default empty `w:numbering` root. Also `NumberingPart.default(package)` at line 37.
- Lazy creation in `src/docx/parts/document.py:240-246` works.

**TODO (if applicable):** none.

---

### upstream#1540 — Unable to detect bullet point style in a docx file

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** Paragraphs with bullets show style "Body Text" — user wants to detect list/bullet paragraphs reliably.

**Evidence:**
- `src/docx/text/paragraph.py:941` `Paragraph.list_level` returns `w:numPr/w:ilvl` int or None — presence indicates a numbered/bulleted paragraph regardless of style name.
- `src/docx/numbering.py` + `Paragraph.num_format` / `apply_to` give full numbering introspection (D.9, #22).

**TODO (if applicable):** Possibly add a convenience `Paragraph.is_list_item` boolean. Effort: S.

---

### upstream#1539 — ValueError: invalid literal for int() accessing right_margin

**Verdict:** `new-bug-needed`

**Ask summary (1-2 sentences):** Section margin attrs fail with `invalid literal for int()` when the document stores fractional twip values (e.g. `0.218505859375`). Similar to closed #1335/#1475.

**Evidence:**
- `src/docx/oxml/simpletypes.py:443-447` `ST_TwipsMeasure.convert_from_xml` does `int(str_value)` — no float tolerance.
- `ST_SignedTwipsMeasure` at line 407 does tolerate floats via `int(round(float(str_value)))` — the unsigned version should mirror that.
- `pgMar/@w:right` is typed as `ST_TwipsMeasure` (unsigned) → crash.

**TODO (if applicable):** Tolerate float in `ST_TwipsMeasure.convert_from_xml` (round to int), mirroring the signed variant. Effort: S.

---

### upstream#1535 — Word Header Style / heading numbering format

**Verdict:** `needs-investigation`

**Ask summary (1-2 sentences):** User wants to transform heading numbering format from CJK style ("一、") to Arabic-decimal ("1 Title" / "1.1 Title") — essentially editing the heading numbering definition.

**Evidence:**
- `src/docx/numbering.py` + `WD_NUMBER_FORMAT` already allow changing a numbering level's format via `NumberingDefinition.apply_to` / level creation.
- Not a one-liner; the user could already achieve this with existing Numbering API, but no higher-level helper "apply a heading number scheme to Heading 1..N".

**TODO (if applicable):** Consider a `Numbering.set_heading_scheme(...)` helper that applies a multi-level scheme to Heading1..Heading9 styles. Effort: M.

---

### upstream#1533 — Support Timezone Specification in Comments Feature

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** `Comments.add_comment` and replies always write `datetime.now(UTC)` with no way to pass a different tz-aware datetime.

**Evidence:**
- `src/docx/comments.py:61` and `:167` hard-code `dt.datetime.now(dt.timezone.utc)` for comment and reply date.
- No `date=` parameter on `add_comment()` / `reply()`.

**TODO (if applicable):** Add optional `date: datetime | None = None` parameter to `Comments.add_comment()` and `Comment.add_reply()`; default to `datetime.now(tz=timezone.utc)` but accept any tz-aware datetime. Effort: S.

---

### upstream#1532 — ValueError when loading .dotx Word template

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** Opening a `.dotx` template fails with "file is not a Word file, content type is template.main+xml". Request: native .dotx support (like upstream PR #1537) or a `Document.from_template(path)` helper.

**Evidence:**
- `src/docx/api.py:39` whitelists only `CT.WML_DOCUMENT_MAIN` and `CT.WML_DOCUMENT_MACRO` — no template type.
- `src/docx/opc/constants.py` has no `WML_TEMPLATE` / `WML_TEMPLATE_MACRO` constant.
- No .dotx handling anywhere in the fork.

**TODO (if applicable):** Add `CT.WML_TEMPLATE` (+ macro variant) constants, register `DocumentPart` for those, accept in `api.Document`, and optionally rewrite content type to DOCUMENT_MAIN on save. Effort: M.

---

### upstream#1528 — A way to retrieve non-standard characters (w:sym)

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** When a run contains `<w:sym w:font="STEDT" w:char="F0DC"/>`, the glyph is lost during text extraction because Run.text skips it.

**Evidence:**
- `src/docx/oxml/text/run.py:357` `CT_Sym` has `font`/`char` attrs but no `__str__`, unlike `CT_Text`/`CT_Br`, so it contributes nothing to `Run.text`.
- Fork has `Run.symbols` (getter) and `Run.add_symbol` (writer) but no integration with `Run.text`.

**TODO (if applicable):** Make `CT_Sym.__str__` return the character derived from its `char` hex code (user can then post-process with the font), or add `Run.text_with_symbols` variant. Effort: S.

---

### upstream#1520 — Does the library support Strict Open XML docx?

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** Strict Open XML .docx files raise "no relationship of type officeDocument" because namespaces differ from Transitional OOXML.

**Evidence:**
- `grep` for "strict"/"transitional"/"ISO/IEC 29500" in src/docx returns only unrelated docstrings — no namespace remapping.
- `src/docx/opc/constants.py` and `src/docx/oxml/ns.py` declare only Transitional namespaces.

**TODO (if applicable):** On package open, detect Strict namespaces and rewrite to Transitional (a content-type + namespace remap at zip ingest), so the rest of the code doesn't have to care. Effort: L.

---

### upstream#1519 — Preserving comments when saving

**Verdict:** `needs-investigation`

**Ask summary (1-2 sentences):** User reports comments disappear when modifying `run.text` in a paragraph that contains a comment. Thread incomplete — OP was to provide a repro.

**Evidence:**
- Fork has full Phase-A/D comments module (`src/docx/comments.py`, `src/docx/oxml/comments.py`) including reply support.
- Whether `run.text` assignment removes `w:commentRangeStart/End` markers in the paragraph is not obvious from a quick grep; needs a reproducer.

**TODO (if applicable):** Write a regression test: create doc with a comment spanning a run, assign to `run.text`, re-save, verify `w:commentRangeStart/End` and `w:commentReference` are preserved. Fix if broken. Effort: M.

---

### upstream#1517 — doc.add_heading('Formulário...', level=1) (incomplete issue body)

**Verdict:** `out-of-scope`

**Ask summary (1-2 sentences):** Issue body is a dump of Portuguese user code with no question, error, or ask. Title is just a line of their script. Appears to be a mistaken post.

**Evidence:**
- Zero comments, no stated problem, no traceback.
- `doc.add_heading` works; emoji in headings works.

**TODO (if applicable):** none.

---

### upstream#1516 — Slow table parsing for huge tables

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** Parsing a 9000×10 table via `Table.rows` / `cell.text` is too slow due to conservative re-parsing on every access. Maintainer suggests a `ReadOnlyTable` / `TableSnapshot` API.

**Evidence:**
- `src/docx/table.py` — `Table.rows`, `_Row.cells`, `_Cell.text` walk the XML tree repeatedly; no snapshot/iterator caching.
- No `ReadOnlyTable` or `iter_rows` fast path in the fork.

**TODO (if applicable):** Add `Table.iter_rows_fast()` (or a `TableSnapshot`) that walks `w:tr`/`w:tc` once and returns plain strings, bypassing mutation-safety. Effort: M.

---

### upstream#1515 — Feature Request: Add Hyperlinks With Multiple Runs

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** Want to insert a hyperlink containing multiple formatted runs (e.g. "Google " bold + "Search" italic inside one `<w:hyperlink>`).

**Evidence:**
- `src/docx/text/paragraph.py:162` `add_hyperlink(url, text, style, anchor)` only supports a single `text`/`style` — one run.
- `src/docx/text/hyperlink.py` `Hyperlink` class has a `runs` getter but no `add_run()` / `clear_runs()` writer.

**TODO (if applicable):** Add `Hyperlink.add_run(text, style)` writer (mirror `Paragraph.add_run`) so callers can build multi-run hyperlinks. Effort: S.

---

### upstream#1512 — Preserving table content/format during replace

**Verdict:** `out-of-scope`

**Ask summary (1-2 sentences):** User's secondary code path rewrites `run.text` in ways that destroy cell formatting; issue has no actionable library bug, it's user code comparing two of their own snippets.

**Evidence:**
- Upstream got no response from maintainers. No reproducible library failure described; it's a user-code question about preserving formatting when they manipulate formats.
- Search/replace with formatting preservation already exists: `src/docx/search.py` + `Document.replace_regex`.

**TODO (if applicable):** none — answer is "use Document.replace_regex / search module".

---

### upstream#1510 — How do I add style (border) to images?

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** Want to add picture outline / border (like Word's "Format Picture → Line") — solid color, width, transparency.

**Evidence:**
- `src/docx/oxml/shape.py:652` mentions `pic:spPr` and `a:ln` in templates (line 659) but there is no public API for picture borders; the user's raw OxmlElement hack breaks the file because of namespace/URI issues.
- No `InlineShape.outline` / `.border` property.

**TODO (if applicable):** Add `InlineShape.outline` / `FloatingImage.outline` with width/rgb/transparency properties writing `pic:spPr/a:ln`. Effort: M.

---

### upstream#1504 — Support reading pptx chart in CT_GraphicalObjectData (c:chart)

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** Request: add `c:chart` as ZeroOrOne on CT_GraphicalObjectData so python-docx can recognise chart drawings.

**Evidence:**
- `src/docx/oxml/shape.py:363-368` — the public CT_GraphicalObjectData still declares only `pic:pic` + `uri`, but higher up (line 442) the chart code writes `c:chart` directly and `src/docx/drawing/__init__.py:76` reads `wp:inline/a:graphic/a:graphicData/c:chart`.
- Charts are read via `Document.charts` (HISTORY.rst lists "Charts read + add_chart() (#111)").

**TODO (if applicable):** Optional — declare `cChart = ZeroOrOne("c:chart")` on CT_GraphicalObjectData for a cleaner API (today it works via xpath). Effort: S.

---

### upstream#1503 — Merged-cell parsing issues

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** User reports that a table with many merge/split operations reads as 7 columns even though visually it's 3; asks for correct parsing.

**Evidence:**
- `src/docx/table.py:659` `_Cell.is_merge_origin` (tri-state), `Cell.merge_origin` (#145), and gridSpan/vMerge awareness cover merged cells.
- `Table.row_cells` / visible-cell iteration respects vMerge="continue" (line 505).
- The "7 columns but visually 3" reflects the actual `w:tblGrid` — that is the file's truth. Correct answer is: use visible-cell traversal, which the fork already supports.

**TODO (if applicable):** Possibly document "visible vs grid columns" behavior more explicitly. Effort: S (docs).

---

### upstream#1502 — Table direction RTL not supported

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** Setting table to RTL so columns flow right-to-left.

**Evidence:**
- `src/docx/table.py:482` `Table.table_direction` getter/setter using `WD_TABLE_DIRECTION` enum; writes `w:tblPr/w:bidiVisual`.
- `src/docx/enum/table.py:354` defines `WD_TABLE_DIRECTION`.
- A comment on the issue already confirms this works in current code.

**TODO (if applicable):** none.

---

### upstream#1500 — Deleting a column in a simple table breaks alignment

**Verdict:** `new-feature-needed`

**Ask summary (1-2 sentences):** User removed a column with raw lxml (`row._tr.remove(row.cells[2]._tc)`); table becomes malformed. Maintainer said there is no published `delete_column` API.

**Evidence:**
- `src/docx/table.py` has no `delete_column` / `remove_column` method.
- `oxml/table.py` has no helper either.

**TODO (if applicable):** Add `Table.delete_column(index)` and `_Column.delete()` that correctly remove the `w:gridCol`, each row's corresponding `w:tc`, and update grid spans. Effort: M.

---

### upstream#1499 — lxml version is too low

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** Request to bump `lxml >= 3.1.0` floor to something modern (lxml 5).

**Evidence:**
- `pyproject.toml` line 28: `"lxml>=4.9.1"` — already far above 3.1.0.
- Upstream maintainer agreed in thread, and our fork's pyproject already reflects a modern floor.

**TODO (if applicable):** Optional bump further to `lxml>=5` to match thread consensus. Effort: S.

---

### upstream#1497 — Remove Python 3.7/3.8 from project classifiers

**Verdict:** `resolved-in-loadfix`

**Ask summary (1-2 sentences):** pyproject classifiers list Python 3.7/3.8 even though `requires-python = ">=3.9"`.

**Evidence:**
- `pyproject.toml` classifiers (lines 8-22) list only `3`, `3.9`, `3.10`, `3.11`, `3.12`, `3.13`. 3.7/3.8 removed.

**TODO (if applicable):** none.

---

## Batch summary

- `resolved-in-loadfix`: 10 (#1549, #1547, #1546, #1541, #1540, #1504, #1503, #1502, #1499, #1497)
- `new-feature-needed`: 9 (#1544, #1533, #1532, #1528, #1520, #1516, #1515, #1510, #1500)
- `new-bug-needed`: 1 (#1539)
- `needs-investigation`: 3 (#1542, #1535, #1519)
- `out-of-scope`: 2 (#1517, #1512)

Total: 25.

Patterns noticed:
- Table API gaps keep recurring (column deletion #1500, perf #1516, merged-cell clarity #1503) — a future "tables-advanced" wave would be productive.
- Image / picture customisation is under-served (#1510 outline, #1547 wrap, #1544 pathlib).
- Several issues (#1542, #1533) are timezone-adjacent — one follow-up could standardise tz handling for both core_properties and comments.
- Two `pyproject.toml`/deps housekeeping issues (#1497, #1499) are already fixed in the fork.

### upstream#1494 — Zero-DPI image causes division by zero error
**Verdict:** new-bug-needed
**Ask summary:** A JPEG that reports a DPI of 0 triggers `ZeroDivisionError` in `Image.width/height`; should fall back to 72 DPI like "missing DPI" case.
**Evidence:** `src/docx/image/image.py:108` still does `Inches(px_width / horz_dpi)` with no zero guard; no fallback in `src/docx/image/jpeg.py` App0/App1 parsing.
**TODO (if applicable):** Guard `horz_dpi`/`vert_dpi` in `Image`/`_ImageHeaderBase` so zero/invalid values resolve to 72. S

### upstream#1493 — Error when adding a .jpg picture to a file
**Verdict:** needs-investigation
**Ask summary:** User reports a traceback from `Document.add_picture` on a specific JPG but no image attached and the traceback body was truncated; could be duplicate of #1494.
**Evidence:** Body has no stack trace; two users report reproduction but no shareable image. Commenter "worked for me" suggests environmental or image-specific.
**TODO (if applicable):** n/a until reproducer is supplied.

### upstream#1492 — Project Status?
**Verdict:** out-of-scope
**Ask summary:** Meta question about upstream project maintenance; not a feature/bug.
**Evidence:** Not a library issue; loadfix fork exists precisely because upstream is slow-moving.
**TODO (if applicable):** n/a.

### upstream#1489 — Cells with field values / drop down menu not recognised as text
**Verdict:** resolved-in-loadfix
**Ask summary:** `_Cell.text` returns empty for cells whose content is in `w:sdt` (content-control dropdown) or a complex field.
**Evidence:** `src/docx/oxml/text/paragraph.py:274` iterates `w:r | w:hyperlink | w:fldSimple | w:sdt`, so SDT and simple-field text is included in `paragraph.text`, which `_Cell.text` joins.
**TODO (if applicable):** n/a (verify with a doc containing only complex-field text if caller reports remainder).

### upstream#1484 — character format application issue
**Verdict:** needs-investigation
**Ask summary:** Applying a character style to a run reportedly doesn't take effect unless the user also iterates runs and re-applies.
**Evidence:** Upstream maintainer asked for an MRE; none supplied. No obvious regression in `src/docx/text/run.py` or `font.py`.
**TODO (if applicable):** n/a until minimal repro exists.

### upstream#1483 — how can i change rtl in python
**Verdict:** out-of-scope
**Ask summary:** User asks about a `tkinter`/`customtkinter` text widget RTL, not python-docx.
**Evidence:** Mentions `tag_configure("right", justify="right")` (Tk API); unrelated to docx. Loadfix does provide RTL (HISTORY "RTL / bidi on Paragraph and Run (#127)").
**TODO (if applicable):** n/a.

### upstream#1482 — recursively resolve core / custom properties before extracting text
**Verdict:** new-feature-needed
**Ask summary:** When a paragraph references a property field (e.g. DOCPROPERTY), `paragraph.text` should resolve the property value instead of the cached text.
**Evidence:** Loadfix has custom properties (`src/docx/custom_properties.py`) and complex fields (`src/docx/fields.py`), but `CT_P.text` does not resolve `DOCPROPERTY`/`PAGE`/`NUMPAGES` field codes — only returns displayed `w:r` children.
**TODO (if applicable):** Add resolver that substitutes `DOCPROPERTY` / core-property field results into rendered text. M

### upstream#1481 — Potential Security Improvements (Scorecard)
**Verdict:** out-of-scope
**Ask summary:** Requests repository-policy actions (branch protection, SAST, Dependabot, SECURITY.md).
**Evidence:** Repo/CI configuration concerns, not library code. Loadfix already has AI-agent CI pipeline; SECURITY policy is a maintainer decision.
**TODO (if applicable):** n/a.

### upstream#1479 — Adding an HTML iframe with an interactive Google map as a document object
**Verdict:** out-of-scope
**Ask summary:** User wants to embed an interactive HTML iframe inside a .docx.
**Evidence:** OOXML / Word do not render iframes; this is a Word-platform limitation, not a library gap.
**TODO (if applicable):** n/a.

### upstream#1477 — Clarification on Previous Issue #1476
**Verdict:** out-of-scope
**Ask summary:** User apology/clarification note; no technical ask.
**Evidence:** Author acknowledges original issue was based on AI misunderstanding.
**TODO (if applicable):** n/a.

### upstream#1475 — Non-integer font sizes
**Verdict:** new-bug-needed
**Ask summary:** `run.font.size` crashes on runs whose `w:sz/@w:val` is a non-integer decimal ("36.5625…") that Word tolerates.
**Evidence:** `src/docx/oxml/simpletypes.py:364` `ST_HpsMeasure.convert_from_xml` still does `Pt(int(str_value) / 2.0)`; no float branch.
**TODO (if applicable):** Accept float half-points: `Pt(float(str_value) / 2.0)` (and mirror in any related sz types). S

### upstream#1474 — Watermark image is moved
**Verdict:** needs-investigation
**Ask summary:** User says a watermark image placed via python-docx shifts position on reopen, regardless of settings.
**Evidence:** Loadfix added watermark support (HISTORY D.23; `src/docx/watermark.py`, `src/docx/oxml/watermark.py`). Issue predates that feature but could still apply. Need reproducer against current loadfix watermark API.
**TODO (if applicable):** Verify watermark image anchoring is preserved after save/round-trip with current watermark module. S

### upstream#1473 — WD_PARAGRAPH_ALIGNMENT EXTENSION (start / end)
**Verdict:** new-bug-needed
**Ask summary:** Enum mapping raises `ValueError: WD_PARAGRAPH_ALIGNMENT has no XML mapping for 'start'` when reading paragraphs whose `w:jc/@w:val` is `start`/`end` (valid per OOXML Transitional).
**Evidence:** `src/docx/enum/text.py:10-64` lists only `left|center|right|both|distribute|*Kashida|thaiDistribute`; no `start`/`end` member, so any doc using those values raises on read.
**TODO (if applicable):** Add `START` and `END` enum members mapping to `start`/`end` XML values (likely aliasing LEFT/RIGHT). S

### upstream#1472 — Extraction of numbering and multi-level sign
**Verdict:** resolved-in-loadfix
**Ask summary:** User wants an API to walk the numbering part and expose `numId`/`abstractNumId`/`lvlText`/`numFmt` per level.
**Evidence:** HISTORY "D.9 Numbering style control (#22)"; `src/docx/numbering.py` exposes `Numbering`, `NumberingDefinition`, `Level.text`, `Level.number_format`, `Level.start`, `Paragraph.list_format`, `Paragraph.numbering_format`.
**TODO (if applicable):** n/a (rendered-number computation is separately requested in #1454).

### upstream#1469 — Python-docx vertically align rectangle and paragraph
**Verdict:** out-of-scope
**Ask summary:** User asks for help aligning a matplotlib-generated rectangle image next to a paragraph — layout/usage question, not a library API gap.
**Evidence:** Concerns matplotlib figure positioning before the image is inserted; no docx code change needed.
**TODO (if applicable):** n/a.

### upstream#1468 — TOC of List of Figures and Tables are not clickable in pdf
**Verdict:** out-of-scope
**Ask summary:** TOC-of-figures hyperlinks disappear after `unoconv` PDF conversion.
**Evidence:** Problem is in the PDF converter (unoconv/LibreOffice), not in python-docx TOC emission. Loadfix TOC feature (`Document.add_table_of_contents`) produces the same XML Word itself writes.
**TODO (if applicable):** n/a.

### upstream#1466 — Unable to parse 3-digit RGB string
**Verdict:** new-feature-needed
**Ask summary:** `RGBColor.from_string` only accepts 6-hex-digit strings; users want 3-digit shorthand (CSS style) accepted.
**Evidence:** `src/docx/shared.py:133` still slices `[:2][2:4][4:]` with no length guard or shorthand expansion.
**TODO (if applicable):** Extend `RGBColor.from_string` to accept 3-hex-digit form by doubling each nibble; raise clear `ValueError` otherwise. S

### upstream#1465 — Streaming Support for Incremental Document Uploads
**Verdict:** out-of-scope
**Ask summary:** Requests piecewise/streaming document writing to S3.
**Evidence:** .docx is a ZIP of XML parts; content must be fully composed before final zip. Commenter on thread notes same. Architectural limitation of the format.
**TODO (if applicable):** n/a.

### upstream#1464 — Option to Remove Metadata from New Documents
**Verdict:** new-feature-needed
**Ask summary:** Caller wants to omit core/app metadata on new documents so file-part reassembly downstream works.
**Evidence:** `src/docx/coreprops.py` / `Package._core_properties_part` always initialise core-properties; no option to suppress them.
**TODO (if applicable):** Add `Document(include_metadata=False)` or post-hoc `Document.core_properties.clear_all()` to blank author/created/modified. S

### upstream#1463 — Feature request: Crop images when adding
**Verdict:** new-feature-needed
**Ask summary:** Allow `add_picture` (or post-insertion API) to set `a:srcRect` crop box.
**Evidence:** `src/docx/oxml/shape.py:351` lists `a:srcRect` as a successor but there is no proxy/setter on `InlineShape` / `FloatingImage`; no `crop_left/top/right/bottom` properties.
**TODO (if applicable):** Add `Picture.crop` setter (four ST_Percentage values) writing `a:srcRect`. M

### upstream#1462 — retrieve document structure
**Verdict:** resolved-in-loadfix
**Ask summary:** User wants to correlate headings, subsections, tables (document outline).
**Evidence:** Standard API already exposes paragraphs + styles (`Paragraph.style.name` — heading levels) and `Document.iter_inner_content()` gives ordered elements; loadfix also ships `src/docx/accessibility.py` (HISTORY "heading-structure accessibility validator (#159)") for structural walk.
**TODO (if applicable):** n/a (usage question with available API).

### upstream#1461 — farshidasadi68:patch-1
**Verdict:** out-of-scope
**Ask summary:** Empty / stray issue cross-posted from a PR comment; no actionable content.
**Evidence:** Body is just the PR branch name; attachment is a zip with no described ask.
**TODO (if applicable):** n/a.

### upstream#1458 — ValueError while access cell._tc.bottom
**Verdict:** new-bug-needed
**Ask summary:** `cell._tc.bottom` raises `ValueError: no tc element at grid_offset=3` for tables with gridBefore/omitted cells; commenter posted a working fix.
**Evidence:** `src/docx/oxml/table.py:360-379` `tc_at_grid_offset` still uses strict equality on `remaining_offset == 0`; upstream PR #1526 (cited in thread) has not landed in loadfix either.
**TODO (if applicable):** Tolerate gridBefore / omitted cells in `tc_at_grid_offset` (match range, not equality) per commenter's patch. S

### upstream#1457 — Copying Images from document
**Verdict:** new-feature-needed
**Ask summary:** User wants a supported way to copy body content + image relationships between documents; currently has to manually rewrite rIds.
**Evidence:** No `copy_body` / `copy_to` helper in `src/docx/`; no shared image-relationship cloning utility. Loadfix has part-level relationship helpers but no body-copy primitive.
**TODO (if applicable):** Add `Document.append_body(other_doc)` that also clones image/media parts and rewrites rIds. L

### upstream#1454 — Correctly reading enumeration and list numbers/letters
**Verdict:** new-feature-needed
**Ask summary:** Provide the rendered list number per paragraph (e.g. "1.2.3", "a)") computed from numbering definitions + document traversal state.
**Evidence:** Loadfix has numbering-definition reading (`src/docx/numbering.py` `Level.text`/`number_format`) but no stateful renderer that walks paragraphs and produces the displayed label.
**TODO (if applicable):** Add `Paragraph.list_label` / `Document.list_labels()` that traverses body in order and formats per-level counters via `lvlText`. L

## Batch summary

- resolved-in-loadfix: 3 (#1489, #1472, #1462)
- new-feature-needed: 6 (#1482, #1466, #1464, #1463, #1457, #1454)
- new-bug-needed: 4 (#1494, #1475, #1473, #1458)
- needs-investigation: 3 (#1493, #1484, #1474)
- out-of-scope: 9 (#1492, #1483, #1481, #1479, #1477, #1469, #1468, #1465, #1461)

Total: 25.

### upstream#1453 — How can I use the library to get the word count of a .doc/.docx files as reported in MS-Word ?

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** User wants accurate word count matching MS Word output from python-docx.

**Evidence:**
- `src/docx/statistics.py` defines `DocumentStatistics` with `words`, `characters`, `characters_no_spaces`, `paragraphs`; wired via `Document.statistics` (HISTORY #161).

**TODO (if applicable):** n/a


### upstream#1449 — format page in docx on page where specific paragraph is placed

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** User wants to detect which physical "page" contains a given paragraph to apply formatting. This requires a rendering/pagination engine not present in any OOXML library (layout is computed by Word).

**Evidence:**
- No pagination/layout engine in `src/docx/`; loadfix only exposes `rendered_page_breaks` (authored break markers), not page indexing.

**TODO (if applicable):** n/a


### upstream#1448 — May I ask how to obtain domain codes from Word and docx files

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** Terse question about "domain codes" (likely field codes, given machine translation). If fields: already covered by loadfix Phase C fields API.

**Evidence:**
- `src/docx/fields.py` provides complex & simple field code access (HISTORY Phase C #10).

**TODO (if applicable):** Clarify intent; likely resolved-in-loadfix via `Document.fields` / `instr_text`. S.


### upstream#1447 — Distribution packages miss `RECORD` file

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** User claims wheel lacks `RECORD`. Upstream reply verified wheels do include it (setuptools build-backend generates it automatically).

**Evidence:**
- `pyproject.toml` uses `setuptools.build_meta`; no override of RECORD generation.

**TODO (if applicable):** n/a


### upstream#1445 — The element property of a paragraph is private can we access it as a public property

**Verdict:** new-feature-needed

**Ask summary (1-2 sentences):** Expose a public `Paragraph.element` property aliasing `_element`, for IDE friendliness.

**Evidence:**
- `src/docx/text/paragraph.py:54` sets `self._p = self._element = p` but no public `element` attribute; `Run` at `src/docx/text/run.py:39` already does `self.element = r`.

**TODO (if applicable):** Add `Paragraph.element`, `Table.element`, `_Cell.element` public read-only aliases. S.


### upstream#1444 — dumping a document

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Feature request: produce python-docx code that reconstructs the document (reverse generator / "scaffolding").

**Evidence:**
- No such utility in loadfix; tangential to library mission.

**TODO (if applicable):** n/a


### upstream#1443 — Enumeration / Lists: Missing indention

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** Built-in `List Bullet` / `List Number` styles produced by python-docx's default template lack indentation unlike Word's own list styles.

**Evidence:**
- Default template at `src/docx/templates/default.docx` / `default-styles.xml` supplies these style defs; not modified by loadfix. Known template gap.

**TODO (if applicable):** Review default-styles.xml list indents against Word's output; M.


### upstream#1441 — Kivy & Kivymd

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** User asks whether python-docx works with Kivy/KivyMD. No docx-related action.

**Evidence:**
- Pure-Python library; no UI framework coupling.

**TODO (if applicable):** n/a


### upstream#1439 — WD_BUILTIN_STYLE gives no style with name

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** `styles[WD_STYLE.BODY_TEXT]` raises `KeyError` because lookup stringifies the enum member ("BODY_TEXT (-67)") instead of translating to the UI style name.

**Evidence:**
- `src/docx/styles/styles.py:30-46` `__getitem__` only accepts str names/ids; `WD_BUILTIN_STYLE` is a `BaseEnum` value at `src/docx/enum/style.py:6`.

**TODO (if applicable):** Accept `WD_BUILTIN_STYLE` member in `Styles.__getitem__` (map enum -> UI name). S.


### upstream#1438 — How to insert a table properly with alignment and formatting

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** User wants a "keep whole table on one page" behaviour conditional on available space. Physical pagination is not computable without a layout engine; only per-row `cantSplit` is modifiable.

**Evidence:**
- `src/docx/table.py` exposes row `cant_split`/`allow_break_across_pages` (HISTORY D.16 #51) but no page-space detection.

**TODO (if applicable):** n/a


### upstream#1435 — Visio Format file handle

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Support add_picture for Visio (.vsd/.vsdx) embeddings. Visio files aren't images; OOXML embeds them via `oleObject` with a fallback image.

**Evidence:**
- `src/docx/image/__init__.py` SIGNATURES list has no Visio entry; OLE read-only exists (`embedded_objects.py`) but no Visio-specific insert API.

**TODO (if applicable):** Design write-side OLE embed API; L.


### upstream#1434 — Issue Editing Complex Word Tables with Merged Cells

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** `Table._cells` raises IndexError on orphan vMerge=continue cells in row 1 (malformed tables).

**Evidence:**
- `src/docx/table.py:495-518` guards `len(cells) >= col_count` and falls back to fresh `_Cell(tc, self)` when an orphan continuation is encountered.

**TODO (if applicable):** n/a


### upstream#1433 — can't get bottom

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** PyCharm debugger resolver error accessing `tc.bottom` on a vMerge cell. Likely PyCharm-specific (getattr probing raises) rather than API bug; loadfix has unchanged `CT_Tc.bottom` semantics.

**Evidence:**
- `grep -n "def bottom" src/docx/oxml/table.py` — property exists; upstream provides `_Cell.bottom` too; no reproducer beyond IDE.

**TODO (if applicable):** Attempt repro with plain Python against reported merge pattern; S.


### upstream#1431 — How to accurately estimate the size of an object from python-docx

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** User wants accurate RAM sizing for proxy objects; pympler hits lxml C-level `_OxmlElementBase.__dict__` descriptor. Not a python-docx bug.

**Evidence:**
- Error originates in lxml/pympler interaction, not python-docx code.

**TODO (if applicable):** n/a


### upstream#1430 — JPEG image not handled.

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** Valid-looking JPEG rejected by `_ImageHeaderFactory`; likely missing JFIF/Exif APP0/APP1 marker so header sniffer fails. Upstream comment notes stripped Exif/JFIF.

**Evidence:**
- `src/docx/image/jpeg.py:107-120` requires APP0 (JFIF) or APP1 (Exif) sequences; other marker layouts fall through.

**TODO (if applicable):** Relax JPEG sniffer to accept any valid SOI+SOF sequence; S.


### upstream#1429 — MS office DOCX/PPTX to PDF

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Help request for building a docx→PDF converter; suggests using LibreOffice. No library change.

**Evidence:**
- No rendering engine in scope.

**TODO (if applicable):** n/a


### upstream#1428 — Memory leak when using docx.Document to parse large word file

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** User reports 2.6GB RSS after loading a docx; comment traces the retention to lxml `etree.fromstring`, not to python-docx proxies. Duplicate of #1364.

**Evidence:**
- Root cause is lxml retention; loadfix has no memory-freeing override in `src/docx/oxml/parser.py`.

**TODO (if applicable):** Investigate whether dropping `Package._rels` references / explicit `_element.clear()` on close helps; M.


### upstream#1427 — paragraphs.insert doesn't work

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** User calling `list.insert` on the `.paragraphs` sequence expected paragraph to move in-document. Loadfix supplies proper insertion APIs.

**Evidence:**
- `src/docx/text/paragraph.py:743` `insert_paragraph_before` plus `insert_paragraph_after` (HISTORY D.13 #26).

**TODO (if applicable):** n/a (docs could mention, but API exists).


### upstream#1426 — Youtube tutorials for Python-Docx

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Community discussion inviting feedback on tutorial content. No code change.

**Evidence:**
- n/a

**TODO (if applicable):** n/a


### upstream#1425 — how to correctly delete an image with its relationships?

**Verdict:** new-feature-needed

**Ask summary (1-2 sentences):** Provide a high-level API for deleting an inline/floating image that also prunes the orphaned `r:embed` relationship and (optionally) the image part.

**Evidence:**
- `InlineShape` in `src/docx/shape.py` has no `delete()`; `Run.delete` at `src/docx/text/run.py:186` removes the run but leaves relationships. HISTORY mentions no image-delete helper.

**TODO (if applicable):** Add `InlineShape.delete()` / `FloatingImage.delete()` that drop `w:drawing`, prune unused `rId`, and optionally part. M.


### upstream#1422 — Set line spacing does not display properly in word

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Visual discrepancy attributable to Word's "Don't add space between paragraphs of the same style" compat setting (user self-resolved in comment).

**Evidence:**
- `src/docx/text/parfmt.py:236-270` writes standard `w:spacing/@w:line` + `@w:lineRule`; correct per spec.

**TODO (if applicable):** n/a


### upstream#1418 — Setting placeholder text in list of tables via field code 'TOC \\h \\z \\c "Table"'

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** User wants to customize the "No table of figures entries found." placeholder via field-code switches; Word uses the `\f` or the literal separator-run text to drive this.

**Evidence:**
- `src/docx/fields.py` TOC field support exists but custom "empty placeholder" result text write-path not obvious; API `Field.set_result_text` exists.

**TODO (if applicable):** Document that placeholder is just the field result text; user can set via `field.set_result_text(...)`. S.


### upstream#1417 — support for Text Form Fields and Checkboxes

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** Request for legacy Text Input / Checkbox / Dropdown form field support plus bookmark-driven editing.

**Evidence:**
- `src/docx/form_fields.py` (`FormField`, `TextInputFormField`, `CheckboxFormField`, `DropdownFormField`, `new_*_form_field_ffData` builders); HISTORY "Add legacy form fields (#123)".

**TODO (if applicable):** n/a


### upstream#1416 — easily replacing / removing&adding image or text in header

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** User wants to find/replace template placeholders (text and image) including within headers, tolerant of run splits.

**Evidence:**
- `src/docx/document.py:759,815,930,979` `replace_all`/`replace_regex`/`search_all`; `src/docx/search.py:360` iterates headers, footers, footnotes, endnotes, comments. Image-replace is the one half not covered (see #1425).

**TODO (if applicable):** Text side resolved; image placeholder replacement depends on #1425 delete API. n/a here.


### upstream#1414 — How do I add an xml object by coordinates?

**Verdict:** new-feature-needed

**Ask summary (1-2 sentences):** User wants to add a floating shape (transparent rectangle) at explicit coordinates. Loadfix has `Paragraph.add_shape` but it only creates inline shapes; floating images exist, floating shapes do not.

**Evidence:**
- `src/docx/text/paragraph.py:396` `add_shape` uses `new_inline_shape_drawing` only; `src/docx/shape.py:140` FloatingImage is image-specific; no anchor-based shape builder in `src/docx/oxml/shape.py`.

**TODO (if applicable):** Add `Paragraph.add_floating_shape(shape_type, ..., h/v anchor, offset)`; M.


## Batch summary

- resolved-in-loadfix: 5 (#1453, #1434, #1427, #1417, #1416)
- new-feature-needed: 3 (#1445, #1425, #1414)
- new-bug-needed: 1 (#1439)
- needs-investigation: 6 (#1448, #1443, #1433, #1430, #1428, #1418)
- out-of-scope: 10 (#1449, #1447, #1444, #1441, #1438, #1435, #1431, #1429, #1426, #1422)

### upstream#1410 — feat: distinguish "no such file" from "not a ZIP"
**Verdict:** new-feature-needed
**Ask summary:** Raise distinct exceptions (or clearer messages) for missing path vs. non-ZIP input in `Document()` for better diagnostics.
**Evidence:** `src/docx/api.py` / `src/docx/opc/package.py` still raise generic `PackageNotFoundError`; no `isfile`/`FileNotFoundError` branch.
**TODO (if applicable):** Add pre-open `os.path.isfile` check and distinct exception subclass. S

### upstream#1409 — no read apis
**Verdict:** out-of-scope
**Ask summary:** User cannot find read examples; request is a docs/usage question, not a code gap.
**Evidence:** Extensive read APIs exist (Document.paragraphs, tables, search_all, etc.).
**TODO (if applicable):** n/a.

### upstream#1406 — Highlight certain words of a paragraph to bold
**Verdict:** out-of-scope
**Ask summary:** Usage question about bolding substrings within a paragraph; solvable with existing multi-run Paragraph API.
**Evidence:** `paragraph.add_run(...).bold = True` already supported; no library gap.
**TODO (if applicable):** n/a.

### upstream#1404 — OSS-Fuzz Integration
**Verdict:** out-of-scope
**Ask summary:** External contributor offers to add OSS-Fuzz harness + build scripts for continuous fuzzing.
**Evidence:** No fuzzing infra in repo; decision is maintainer/project-policy, not a loadfix feature gap.
**TODO (if applicable):** n/a (consider if fork wants it).

### upstream#1403 — Auto refresh Table of Contents using docx
**Verdict:** new-feature-needed
**Ask summary:** Users want to mark/refresh a TOC so Word recomputes it on open, without win32com.
**Evidence:** `src/docx/toc.py` has `add_table_of_contents`; no "dirty" or `updateFields` API surfaced on `Field`/`TOC`.
**TODO (if applicable):** Add `TOC.mark_dirty()` / `Settings.update_fields_on_open`. S

### upstream#1401 — How to add internal top and bottom table cell spacings?
**Verdict:** new-feature-needed
**Ask summary:** Expose table-level default `w:tblCellMar` (top/bottom/left/right) on `Table`.
**Evidence:** `src/docx/oxml/table.py` lists `w:tblCellMar` in `_tag_seq` but has no ZeroOrOne/proxy; `Table.margins` is not defined (only `_Cell.margins`).
**TODO (if applicable):** Add `Table.cell_margins` proxy mirroring `_Cell.margins`. S

### upstream#1399 — Inline support for SVG file stream
**Verdict:** resolved-in-loadfix
**Ask summary:** Request for SVG image support via `add_picture` streams.
**Evidence:** HISTORY D.22 "SVG image support (#76)"; SVG handling in `src/docx/image/` and `src/docx/parts/story.py`.
**TODO (if applicable):** n/a.

### upstream#1398 — 打开空的docx文档时报错 (error opening empty docx)
**Verdict:** resolved-in-loadfix
**Ask summary:** Document() fails on empty/malformed docx bytes; users want graceful handling.
**Evidence:** `recover=True` mode added (`src/docx/api.py:20`, `Document.recovery_warnings`); HISTORY "recover=True mode for malformed .docx (#151)".
**TODO (if applicable):** n/a.

### upstream#1397 — Can not read an empty docx
**Verdict:** resolved-in-loadfix
**Ask summary:** Duplicate of #1398 — request to tolerate empty/corrupt docx.
**Evidence:** Same as #1398 — `recover=True` path in `src/docx/api.py`.
**TODO (if applicable):** n/a.

### upstream#1396 — Chinese fonts only non-Chinese parts valid
**Verdict:** resolved-in-loadfix
**Ask summary:** Setting `run.font.name` doesn't apply to CJK glyphs because only `w:ascii`/`w:hAnsi` are set, not `w:eastAsia`.
**Evidence:** `src/docx/text/font.py:581` adds `name_east_asia` and `name_far_east` setters writing `rFonts/@w:eastAsia`.
**TODO (if applicable):** n/a (API available; workaround documented).

### upstream#1391 — [Feature] Support EMF image
**Verdict:** new-feature-needed
**Ask summary:** Add EMF (and WMF) inline image support similar to python-docx-ng.
**Evidence:** `src/docx/opc/spec.py:10` registers `emf` content-type mapping, but no `Emf` header class in `src/docx/image/` (grep returns no module).
**TODO (if applicable):** Add EMF/WMF image header parser + `Image.from_*` dispatch. M

### upstream#1390 — Failed to read entire text of a cell in a table
**Verdict:** needs-investigation
**Ask summary:** `cell.text` drops substrings (e.g. "0.5m") in specific CJK docs; likely a content-specific parsing issue (possibly `mc:AlternateContent`, fields, or merged cells).
**Evidence:** `_Cell.text` at `src/docx/table.py:748`; no obvious regression fix landed in loadfix. Sample file not fully attached.
**TODO (if applicable):** Reproduce with attached test.docx; investigate field/AlternateContent interaction. M

### upstream#1389 — doc.paragraphs not including `<mc:AlternateContent>` content
**Verdict:** needs-investigation
**Ask summary:** Text inside `mc:AlternateContent/mc:Choice` (e.g. wps text boxes) is invisible to iteration APIs.
**Evidence:** No `AlternateContent` / `mc:Fallback` handling in `src/docx/`; `iter_inner_content` only yields CT_P/CT_Tbl. Related to D.27 shapes/text-box (#75) which added DrawingML shape access — needs confirmation whether mc:Choice path is covered.
**TODO (if applicable):** Verify D.27 covers mc:AlternateContent/wps:txbx text-box traversal. M

### upstream#1376 — "FollowedHyperlink" cannot be found in some documents
**Verdict:** needs-investigation
**Ask summary:** Built-in `FollowedHyperlink` style is sometimes absent; creating it yields `FollowedHyperlink1` rather than overriding.
**Evidence:** No related fix in loadfix HISTORY; style-creation semantics unchanged (`src/docx/styles/`).
**TODO (if applicable):** Add helper to materialize latent built-in styles (e.g. `styles.add_latent_style`). M

### upstream#1375 — Duplicate document styles
**Verdict:** new-feature-needed
**Ask summary:** Provide a way to copy styles (with properties) from one document to another; current deep-copy loses formatting.
**Evidence:** No `Styles.import_from` / style-clone API in `src/docx/styles/`.
**TODO (if applicable):** Add `Styles.import_from(source_doc, names=None)` that clones `w:style` elements. M

### upstream#1372 — 怎么样获取word中生成的序号 (get list-generated numbers)
**Verdict:** new-feature-needed
**Ask summary:** Retrieve the auto-generated list numbers (e.g. "1.1") as part of paragraph text.
**Evidence:** `src/docx/numbering.py` exposes Numbering/Level/start but no `Paragraph.list_number_text` renderer.
**TODO (if applicable):** Add read-only list label renderer on Paragraph (sequence-aware). L

### upstream#1370 — Unable to get text of MergeFields (sometimes)
**Verdict:** needs-investigation
**Ask summary:** Replacing placeholders inside `MERGEFIELD` instructions sporadically fails because placeholder spans multiple runs/field ranges.
**Evidence:** Phase C added field read + D.10 search/replace with formatting preservation (`replace_all`, `replace_regex`); unclear whether MERGEFIELD `instrText` placeholders are reached.
**TODO (if applicable):** Verify `replace_all` traverses `w:instrText` / field ranges for MERGEFIELDs. S

### upstream#1369 — TypeError: int() argument must be ... not 'NoneType'
**Verdict:** needs-investigation
**Ask summary:** Crash during docx→html conversion due to `int(None)` on some attribute; requester wants a None-safe default.
**Evidence:** Traceback/screenshot insufficient to localize in loadfix; no clear match via grep.
**TODO (if applicable):** Request reproduction; add None-guard once attribute/line identified. S

### upstream#1368 — Clean way to create a multi-column document
**Verdict:** resolved-in-loadfix
**Ask summary:** Public API to set multi-column section layout (currently requires touching `_sectPr`).
**Evidence:** `Section.columns` proxy (`src/docx/section.py:60`, class `SectionColumns` L919) + HISTORY D.19 "Multi-column section layout (#60)".
**TODO (if applicable):** n/a.

### upstream#1367 — table.column_cells(0) mixes cells across rows
**Verdict:** needs-investigation
**Ask summary:** For a specific docx, `column_cells(0)` returns a second-row cell inside first-row result — likely vertical-merge handling bug.
**Evidence:** No related fix noted in HISTORY; `Table.column_cells` unchanged. Needs test with attached `2.docx`.
**TODO (if applicable):** Reproduce and fix vMerge accounting in `column_cells`. M

### upstream#1366 — A row in a table cannot span multiple pages
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose `w:cantSplit` / allow row to break across pages.
**Evidence:** HISTORY D.16 "Row.allow_break_across_pages (#51)".
**TODO (if applicable):** n/a.

### upstream#1365 — Reading bullet point values (e.g. "1.1")
**Verdict:** new-feature-needed
**Ask summary:** Read the rendered list label ("1.1 Sample text") alongside paragraph text.
**Evidence:** Duplicate ask of #1372; no list-label renderer in `src/docx/numbering.py` / `Paragraph`.
**TODO (if applicable):** Add `Paragraph.list_label` / renderer (shared with #1372). L

### upstream#1359 — Completeness of example code
**Verdict:** out-of-scope
**Ask summary:** User asking another user to flesh out a field-insertion code snippet; a support question, not a library request.
**Evidence:** n/a.
**TODO (if applicable):** n/a.

### upstream#1357 — How can I convert docx to XML
**Verdict:** out-of-scope
**Ask summary:** User wants to translate `w:t` tags by round-tripping docx→xml→docx; usage/how-to, not a library gap.
**Evidence:** Existing `element.xml` / save round-trip already supports this.
**TODO (if applicable):** n/a.

### upstream#1354 — Insert picture issue (header logo)
**Verdict:** out-of-scope
**Ask summary:** User self-resolved; thread is a request for their solution, not a library bug report.
**Evidence:** Issue body indicates resolution outside library.
**TODO (if applicable):** n/a.

## Batch summary
- resolved-in-loadfix: 6 (#1399, #1398, #1397, #1396, #1368, #1366)
- new-feature-needed: 7 (#1410, #1403, #1401, #1391, #1375, #1372, #1365)
- new-bug-needed: 0
- needs-investigation: 6 (#1390, #1389, #1376, #1370, #1369, #1367)
- out-of-scope: 6 (#1409, #1406, #1404, #1359, #1357, #1354)
- Total: 25

### upstream#1352 — Docu: Reference to root package missing in object.inv file
**Verdict:** out-of-scope
**Ask summary:** Sphinx `object.inv` intersphinx file omits the root `docx` package reference, breaking pydoctor cross-linking.
**Evidence:** Docs build tooling issue; loadfix has not touched Sphinx `conf.py` intersphinx configuration.
**TODO (if applicable):** Add `py:module` directive for `docx` root in docs/conf.py — S.

### upstream#1349 — There is no item named %r in the archive
**Verdict:** needs-investigation
**Ask summary:** Rendering a docx containing a link to an internal bookmark raises `KeyError: "There is no item named %r in the archive"`. No repro/traceback provided.
**Evidence:** loadfix has bookmarks (src/docx/bookmarks.py, Phase C) and hyperlink API but no specific handler for this archive-lookup KeyError; message signature not found in src/docx/.
**TODO (if applicable):** Reproduce with sample docx; trace OPC archive access on hyperlink/bookmark load — M.

### upstream#1348 — Delete break section with next page?
**Verdict:** new-feature-needed
**Ask summary:** User wants to remove a "section break (next page)" programmatically.
**Evidence:** `src/docx/section.py` has `remove_page_borders`, `remove_watermark`, etc. but no section-delete / merge-into-previous API; no `sections.remove` / `del sections[i]`.
**TODO (if applicable):** Add `Section.delete()` / `sections.pop()` merging sectPr into next section — M.

### upstream#1347 — Track change
**Verdict:** resolved-in-loadfix
**Ask summary:** Wants to programmatically produce Word "Track Changes" marking inserted/deleted text when merging a text file.
**Evidence:** Phase B in HISTORY.rst covers read + accept/reject of tracked changes; src/docx/tracked_changes.py exists. Authoring tracked insertions/deletions may still need extension.
**TODO (if applicable):** Confirm Phase B covers write/authoring of `w:ins`/`w:del`; if read-only, add tracked-change authoring API — M.

### upstream#1344 — Split docx file at all Headings and keep styles?
**Verdict:** out-of-scope
**Ask summary:** Usage question — how to split a doc on Heading 1 while preserving styles (user's snippet creates blank docs).
**Evidence:** Styles carry via the template; this is user-code territory. No "split_document" helper is in loadfix nor expected.
**TODO (if applicable):** n/a (documentation / recipe only).

### upstream#1341 — How to change default Cambria Math font for math equations
**Verdict:** needs-investigation
**Ask summary:** User wants to change math-equation font (currently always Cambria Math) via python-docx.
**Evidence:** src/docx/equations.py and src/docx/oxml/math.py exist (minimal create API per HISTORY); no API for math-font override found.
**TODO (if applicable):** Expose `w:mathPr/w:mathFont` on settings or equation to override default — M.

### upstream#1339 — allows to copy data_frame styling
**Verdict:** out-of-scope
**Ask summary:** Feature request: ingest pandas `DataFrame.style` HTML and render as styled docx table.
**Evidence:** No HTML-to-docx or altChunk support in loadfix; outside fork mission (pandas integration).
**TODO (if applicable):** n/a.

### upstream#1338 — Inaccurately extracts underlined words from docx file
**Verdict:** needs-investigation
**Ask summary:** `run.font.underline` returns True for some runs that visually appear un-underlined (likely style inheritance / `w:u val="none"` interplay).
**Evidence:** font.underline in src/docx/text/font.py uses standard ZeroOrOne mechanism; no special handling for style-inherited underline overrides.
**TODO (if applicable):** Verify `w:u val="none"` handling vs. style inheritance; consider `effective_underline` — M.

### upstream#1334 — Incorrect column count estimation on some tables
**Verdict:** needs-investigation
**Ask summary:** `_Table._column_count` diverges from `len(_Row.cells)` in some tables (likely gridSpan / tblGrid mismatch).
**Evidence:** `_column_count` in src/docx/table.py uses tblGrid; known upstream class of bugs. No explicit loadfix fix noted.
**TODO (if applicable):** Investigate row vs tblGrid reconciliation; patch `_column_count` for gridSpan-heavy tables — M.

### upstream#1332 — How to get page number?
**Verdict:** out-of-scope
**Ask summary:** Wants runtime page number when iterating a table — requires a layout engine.
**Evidence:** OOXML does not store persistent pagination; python-docx does not lay out. loadfix adds `RenderedPageBreak` support but cannot compute arbitrary page numbers.
**TODO (if applicable):** n/a.

### upstream#1331 — CANNOT Export Cropping Images
**Verdict:** new-feature-needed
**Ask summary:** When extracting images, Word's crop (`a:srcRect`) is ignored — raw (un-cropped) image is returned.
**Evidence:** `a:srcRect` is referenced only as a successor sibling in src/docx/oxml/shape.py; no cropping accessor or crop-aware export.
**TODO (if applicable):** Expose `InlineShape.crop` (srcRect l/t/r/b) and/or apply crop when exporting bytes — M.

### upstream#1330 — One character < is failing to write in the document
**Verdict:** needs-investigation
**Ask summary:** Adding a run with text `"<"` appears to drop the char in output (possibly parser treating as start-of-tag).
**Evidence:** No special `<`-escaping logic in src/docx/text/run.py; lxml normally escapes text nodes, so this may be reproduction-specific.
**TODO (if applicable):** Add regression test using `run.add_text("<")` and verify serialized XML contains `&lt;` — S.

### upstream#1327 — Missing content while reading paragraph
**Verdict:** needs-investigation
**Ask summary:** Some paragraph content is missing when reading via `paragraph.text` — likely SDT / content-control / hyperlink / field wrapped content not walked.
**Evidence:** loadfix has content_controls.py, fields.py, hyperlinks, plus Paragraph.iter_inner_content (1.0.0); but `.text` may still skip some wrappers.
**TODO (if applicable):** Audit Paragraph.text traversal for w:sdt / w:smartTag / nested content — M.

### upstream#1324 — Table normal style doesn't correspond to Microsoft Word's table normal style
**Verdict:** needs-investigation
**Ask summary:** "Table Normal" style applied via API produces a borderless table, while Word's Table Normal shows borders.
**Evidence:** Style resolution is in src/docx/styles/; loadfix adds Table.style_flags/borders (#144, #102) but no documented fix for this style-lookup discrepancy.
**TODO (if applicable):** Verify Table Normal style chain and default borders — M.

### upstream#1320 — Convert jupyter notebook into docx
**Verdict:** out-of-scope
**Ask summary:** Wants nbconvert-style ipynb→docx with styles preserved.
**Evidence:** No notebook conversion; outside scope (belongs to nbconvert / pandoc).
**TODO (if applicable):** n/a.

### upstream#1317 — altChunk doesn't insert in header's paragraphs
**Verdict:** new-feature-needed
**Ask summary:** Inserting `w:altChunk` into a header paragraph is ignored by Word on open.
**Evidence:** `altChunk` not found in src/docx/; not supported in loadfix. Word itself rejects altChunk in headers (schema restriction), so this may be resolve-as-wontfix.
**TODO (if applicable):** Document limitation or add main-body-only altChunk helper — M.

### upstream#1316 — Image opacity
**Verdict:** new-feature-needed
**Ask summary:** Provide API to set image opacity (`a:alphaModFix`) on inline/floating images.
**Evidence:** "alphaModFix" / "opacity" not found in src/docx/; Phase D.17 adds floating images but no opacity control.
**TODO (if applicable):** Add `InlineShape.opacity` / `FloatingImage.opacity` setter producing `a:alphaModFix` — S.

### upstream#1315 — Dealing with complex script in word document
**Verdict:** resolved-in-loadfix
**Ask summary:** Accessing font attributes of complex-script (Arabic/RTL) runs returns None — need `rFonts_cs`, bidi, complex_script.
**Evidence:** src/docx/text/font.py exposes `complex_script` setter/getter (l.232), `bidi_language` (l.425), `rFonts_cs` (l.573), `rFonts_eastAsia` — per HISTORY entries #160, #128, #127.
**TODO (if applicable):** n/a — covered by #127 / #128 / #160.

### upstream#1314 — How to quit the lock aspect relation of an image?
**Verdict:** new-feature-needed
**Ask summary:** Need API to clear `noChangeAspect` on picture so width/height can be set independently.
**Evidence:** `noChangeAspect` referenced in src/docx/oxml/shape.py but no public setter on Shape/Image proxies.
**TODO (if applicable):** Add `InlineShape.lock_aspect_ratio` and equivalent on floating image — S.

### upstream#1313 — NumberingPart.new NotImplementedError
**Verdict:** resolved-in-loadfix
**Ask summary:** Accessing `doc.part.numbering_part` raises NotImplementedError in stock python-docx.
**Evidence:** src/docx/parts/numbering.py defines `NumberingPart.new()` and `NumberingPart.default(package)` returning a valid part. HISTORY Phase D.9 (#22).
**TODO (if applicable):** n/a.

### upstream#1312 — How to detect merged cells?
**Verdict:** resolved-in-loadfix
**Ask summary:** Need a reliable API to detect merged cells (horizontal + vertical) rather than comparing cell text.
**Evidence:** src/docx/table.py provides `Cell.is_merge_origin` (l.659) and `Cell.merge_origin` (l.687) walking vMerge/gridSpan — HISTORY #145.
**TODO (if applicable):** n/a.

### upstream#1308 — Enable copying of a run's font
**Verdict:** new-feature-needed
**Ask summary:** Provide a `Font.copy_to(other)` / `copy_format` helper so replacing text preserves rPr formatting.
**Evidence:** No `copy_format` / `copy_font` API in src/docx/text/ (grep returns no hits). `Run.split` exists (#94) but no font-copy.
**TODO (if applicable):** Add `Font.copy_to(target_font)` clone of `w:rPr` — S.

### upstream#1307 — Missing embedded font
**Verdict:** new-feature-needed
**Ask summary:** Embedded fonts (font-embedding fontTable parts) are dropped when python-docx saves a modified file.
**Evidence:** src/docx/font_table.py exposes `is_embedded` (read) but `FontTablePart` / parts-relation handling for the binary embedded-font parts is not obviously preserved on save.
**TODO (if applicable):** Ensure `/word/fonts/*.odttf` parts & rels survive save round-trip — M.

### upstream#1303 — How to copy the tables from docx to a new docx file
**Verdict:** out-of-scope
**Ask summary:** Usage question about copying a `Table` element between documents preserving styles.
**Evidence:** No generic cross-document copy helper; requires style/numbering migration which is larger effort.
**TODO (if applicable):** Recipe / optional `Table.copy_to(other_doc)` — L.

### upstream#1300 — Test failure on Python 3.12 due to datetime.datetime.utcnow() deprecation
**Verdict:** resolved-in-loadfix
**Ask summary:** Upstream tests fail on 3.12 due to `utcnow()` DeprecationWarning surfaced as error.
**Evidence:** src/docx/opc/parts/coreprops.py:34 uses `dt.datetime.now(dt.timezone.utc)`; tests in tests/opc/parts/test_coreprops.py and tests/test_comments.py also use timezone-aware now.
**TODO (if applicable):** n/a.

## Batch summary
- resolved-in-loadfix: 5 (#1347, #1315, #1313, #1312, #1300)
- new-feature-needed: 6 (#1348, #1331, #1317, #1316, #1314, #1308, #1307) [7]
- needs-investigation: 7 (#1349, #1341, #1338, #1334, #1330, #1327, #1324)
- out-of-scope: 5 (#1352, #1344, #1339, #1332, #1320, #1303) [6]
- new-bug-needed: 0

Totals (authoritative): resolved-in-loadfix 5, new-feature-needed 7, needs-investigation 7, out-of-scope 6, new-bug-needed 0 = 25.

### upstream#1293 — Page background-color
**Verdict:** resolved-in-loadfix
**Ask summary:** Support setting the document page background color.
**Evidence:** `src/docx/document.py:310-334` — `Document.background_color` RGBColor getter/setter (HISTORY `#118`).
**TODO (if applicable):** n/a

### upstream#1285 — Table cells should support recursively add_table with style
**Verdict:** new-feature-needed
**Ask summary:** Add an optional `style` kwarg to `_Cell.add_table(rows, cols)` so cell tables can be styled on creation like `Document.add_table`.
**Evidence:** `src/docx/table.py:562-576` — `_Cell.add_table` still lacks `style` param.
**TODO (if applicable):** Add `style=None` param to `_Cell.add_table` mirroring `Document.add_table`. S

### upstream#1280 — can not get TOC of docx with paragraphs
**Verdict:** new-feature-needed
**Ask summary:** TOC paragraphs wrapped in `w:sdt` aren't visible via `document.paragraphs`; user wants flattening / SDT-aware iteration.
**Evidence:** `src/docx/oxml/document.py:113-119` `inner_content_elements` xpath only picks `./w:p | ./w:tbl` (no SDT descent); `blkcntnr.paragraphs` just iterates `p_lst`.
**TODO (if applicable):** Optional `include_sdt`/flatten mode on body iteration to surface SDT-wrapped paragraphs. M

### upstream#1277 — How to get filename of extracted OLE objects
**Verdict:** resolved-in-loadfix
**Ask summary:** Retrieve the filename / part identity of embedded OLE objects.
**Evidence:** `src/docx/embedded_objects.py:61-70` — `EmbeddedObject.embedded_partname` exposes OPC partname (HISTORY `#140`).
**TODO (if applicable):** n/a

### upstream#1268 — How to add a ZIP attachment to Word
**Verdict:** out-of-scope
**Ask summary:** Support inserting an arbitrary ZIP attachment as an embedded package.
**Evidence:** no match — `src/docx/embedded_objects.py` is read-only; no `add_embedded_object`/`add_package` API.
**TODO (if applicable):** Write-side embedded-package attachment API. L

### upstream#1252 — Documenting overwrite behaviour of Document.save
**Verdict:** needs-investigation
**Ask summary:** Document that `Document.save()` overwrites the target .docx without warning.
**Evidence:** docs/user/documents.rst does not call out overwrite behavior explicitly.
**TODO (if applicable):** Add note in `docs/user/documents.rst` and `Document.save` docstring. S

### upstream#1250 — Inconsistent document icons
**Verdict:** out-of-scope
**Ask summary:** Generated doc's chart icon differs from one created in Office; cosmetic/thumbnail rendering behavior.
**Evidence:** no match — OS/Office thumbnail generation, not a python-docx feature.
**TODO (if applicable):** n/a

### upstream#1238 — Get and set the shading color on run and paragraphs
**Verdict:** needs-investigation
**Ask summary:** Expose shading (`w:shd/@w:fill`) reads/writes on both runs and paragraphs.
**Evidence:** `src/docx/text/font.py:700-735` implements `Font.shading_color` (run). No paragraph-level proxy property for `pPr/w:shd` (search of `src/docx/text/parfmt.py` / `paragraph.py` for "shading" empty).
**TODO (if applicable):** Add `ParagraphFormat.shading_color` mirroring `Font.shading_color`. S

### upstream#1235 — Accessing characters in equations
**Verdict:** new-feature-needed
**Ask summary:** Edit characters / replace text inside OMML equations while preserving formatting.
**Evidence:** `src/docx/equations.py:64-135` provides read (`omml_xml`, `text`) and minimal create; no text/mutate API for children.
**TODO (if applicable):** Add equation-node text traversal/edit helpers on `Equation`. M

### upstream#1233 — Find texts with Normal style and size != 11
**Verdict:** out-of-scope
**Ask summary:** Usage question about filtering runs where style is Normal and font size differs from 11pt; no API gap.
**Evidence:** `Run.font.size` and `Paragraph.style.name` already available; behavior is user's own logic bug (inheritance).
**TODO (if applicable):** n/a

### upstream#1231 — Custom Font Inclusion
**Verdict:** new-feature-needed
**Ask summary:** Ability to embed (package) TTF fonts in the output docx so readers without the font can still render correctly.
**Evidence:** `src/docx/font_table.py` and `FontTablePart` expose read of existing font references; no API to embed `embeddedRegular`/obfuscated font parts or add font parts.
**TODO (if applicable):** Add `FontTable.add_embedded_font(path)` wiring embedded font parts + relationships. L

### upstream#1228 — Automerging tables issue 2
**Verdict:** out-of-scope
**Ask summary:** Word auto-merges adjacent tables unless separated by a paragraph; user wants a smaller separator.
**Evidence:** This is a Word rendering behavior; library already supports inserting a paragraph between tables (workaround is the canonical OOXML solution).
**TODO (if applicable):** n/a

### upstream#1227 — WD_TABLE_DIRECTION.RTL not working
**Verdict:** needs-investigation
**Ask summary:** Setting `table.direction = WD_TABLE_DIRECTION.RTL` has no effect in Word.
**Evidence:** `src/docx/table.py:482-492` exposes `Table.table_direction` (new name) that writes `w:tblPr/w:bidiVisual`; `Table.direction` (old upstream name) not confirmed. Worth testing upstream-like reproducer against loadfix.
**TODO (if applicable):** Verify `table_direction=RTL` round-trips and renders; add `direction` alias if needed. S

### upstream#1209 — Buffered operations for bulk table read/write
**Verdict:** new-feature-needed
**Ask summary:** Provide an "official" cached cell access to avoid `_cells` being rebuilt on every `Table.cell(r, c)` call.
**Evidence:** `src/docx/table.py:372-378` still calls `self._cells[...]` each time; no `lazyproperty`/cached variant.
**TODO (if applicable):** Add public `Table.cells` (cached/buffered) or cache `_cells` with invalidation. M

### upstream#1208 — Recursive call in Table Cell merging causes RecursionError
**Verdict:** new-bug-needed
**Ask summary:** `CT_Tc._grow_to` recurses row-by-row; large tables hit Python recursion limit.
**Evidence:** `src/docx/oxml/table.py:982-1003` still recursive (`tc_below._grow_to(...)` tail call).
**TODO (if applicable):** Convert `_grow_to` to iterative loop. S

### upstream#1201 — Add paragraph in a specific section location
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert a paragraph at an arbitrary location within a section.
**Evidence:** Phase D.13 — `Paragraph.insert_paragraph_before`/`insert_paragraph_after` (`src/docx/text/paragraph.py:743-780`, HISTORY `#26`); `Section.iter_inner_content` (`src/docx/section.py:235`) gives ordered section content to target a position.
**TODO (if applicable):** n/a

### upstream#1197 — Read info from select/combo box
**Verdict:** resolved-in-loadfix
**Ask summary:** Read the selected value of a combo-box / dropdown SDT form field.
**Evidence:** `src/docx/content_controls.py` (D.14 `#27`) registers `ContentControlType.COMBO_BOX` / `DROPDOWN` and exposes `ContentControl` proxy; `Document.content_controls` iteration (`document.py:1162`).
**TODO (if applicable):** n/a

### upstream#1189 — Add new rows with same style/format as a template row
**Verdict:** new-feature-needed
**Ask summary:** A helper to append rows that inherit a specific existing row's cell/paragraph formatting.
**Evidence:** no match — `Table.add_row` does not clone formatting from another row.
**TODO (if applicable):** Add `Table.add_row(source_row=...)` / `_Row.copy()` clone helper. M

### upstream#1178 — ElementProxy is missing in the docs
**Verdict:** needs-investigation
**Ask summary:** `ElementProxy` class is referenced but not exposed in the Sphinx API documentation.
**Evidence:** `docs/user/api-concepts.rst:60` mentions it; no dedicated `autoclass` in `docs/api/`.
**TODO (if applicable):** Add `.. autoclass:: docx.shared.ElementProxy` in an appropriate `docs/api/` page. S

### upstream#1177 — Deactivate spell/grammatical checking for a document
**Verdict:** new-feature-needed
**Ask summary:** Toggle `w:hideSpellingErrors` / `w:hideGrammaticalErrors` (and default language) via `Settings`.
**Evidence:** `src/docx/oxml/settings.py:366-369` declares the tags in the `_tag_seq` but no `ZeroOrOne` wiring or `Settings` proxy property (`grep -n hide_spell` src/docx/settings.py empty).
**TODO (if applicable):** Wire `w:hideSpellingErrors`/`w:hideGrammaticalErrors` CT_OnOff children + `Settings.hide_spelling_errors`/`hide_grammatical_errors`. S

### upstream#1176 — Update hyperlink with anchor
**Verdict:** new-feature-needed
**Ask summary:** Update/remove the anchor (`#bookmark`) portion of a hyperlink's target.
**Evidence:** `src/docx/text/hyperlink.py:59-80` exposes `address` / `fragment` as read-only properties; no setter to change the target's anchor (nor the `rels` Target for external `_Target`).
**TODO (if applicable):** Add setters for `Hyperlink.address` and `Hyperlink.fragment` (updating `w:anchor` and the relationship target). S

### upstream#1164 — Image size problem (resets after moving)
**Verdict:** needs-investigation
**Ask summary:** Image inserted with explicit width/height appears wrong until user moves it in Word.
**Evidence:** no match — likely an EMU/aspect-ratio or inline-shape sizing bug; need a repro against loadfix `inline_shapes`/`FloatingImage`.
**TODO (if applicable):** Repro with current `add_picture(width=, height=)`; confirm `<wp:extent>` vs `<a:ext>` consistency. M

### upstream#1161 — TableStyle overwritten by document style
**Verdict:** needs-investigation
**Ask summary:** Modifying `Normal` style appears to suppress a custom `TABLE` style's font settings.
**Evidence:** no explicit fix in HISTORY; style inheritance interaction not covered by Phase C/D entries.
**TODO (if applicable):** Investigate whether custom table style writes `w:tblStylePr/w:rPr` correctly so it wins over Normal. M

### upstream#1160 — Improve documentation about tables
**Verdict:** needs-investigation
**Ask summary:** Docs should show how to format text (alignment, bold) inside table cells, e.g. for a header row.
**Evidence:** `docs/user/tables.rst` and `tables-advanced.rst` exist but may not cover per-cell run formatting examples.
**TODO (if applicable):** Add "Formatting cell content" section with header-row bold/centered example. S

### upstream#1159 — Table autofit default value
**Verdict:** needs-investigation
**Ask summary:** Clarify in `Table.autofit` docstring what the default value is (True?).
**Evidence:** `src/docx/table.py:203-218` — docstring on `autofit` does not state the default.
**TODO (if applicable):** Extend `Table.autofit` docstring to note default (True when `w:tblLayout` absent). S

## Batch summary

- resolved-in-loadfix: 4 (1293, 1277, 1201, 1197)
- new-feature-needed: 8 (1285, 1280, 1235, 1231, 1209, 1189, 1177, 1176)
- new-bug-needed: 1 (1208)
- needs-investigation: 8 (1252, 1238, 1227, 1178, 1164, 1161, 1160, 1159)
- out-of-scope: 4 (1268, 1250, 1228, 1233)
- total: 25

### upstream#1156 — feat: insert hyperlink in paragraph
**Verdict:** resolved-in-loadfix
**Ask summary:** Ability to insert hyperlinks pointing to other documents from within a paragraph/table cell.
**Evidence:** `src/docx/text/paragraph.py:162` `add_hyperlink`; HISTORY D.1 (#97).
**TODO (if applicable):** n/a

### upstream#1154 — Project status in README.md
**Verdict:** out-of-scope
**Ask summary:** Asks upstream maintainer about project status and forks.
**Evidence:** Meta question about upstream repo.
**TODO (if applicable):** n/a

### upstream#1150 — Applying Ligature to Font
**Verdict:** new-feature-needed
**Ask summary:** Set OpenType ligatures (`w14:ligatures`) on a run.
**Evidence:** No `ligatures` in `src/docx/oxml/` or `text/font.py`.
**TODO (if applicable):** Add `Font.ligatures` mapping `w:rPr/w14:ligatures/@w14:val`. S

### upstream#1147 — Replace text in paragraph while retaining inline shape
**Verdict:** resolved-in-loadfix
**Ask summary:** Replace paragraph text without deleting inline shapes/images.
**Evidence:** `src/docx/search.py` run-level replace preserves non-text runs; HISTORY D.10 (#91), `Document.replace_regex` docs.
**TODO (if applicable):** n/a

### upstream#1144 — add Table.indent
**Verdict:** new-feature-needed
**Ask summary:** Expose `w:tblPr/w:tblInd` (left indent for tables).
**Evidence:** No `tblInd` or `Table.indent` in `src/docx/table.py` / `oxml/table.py`.
**TODO (if applicable):** Add `Table.indent` property. S

### upstream#1141 — Feature - Edit Chart Data
**Verdict:** new-feature-needed
**Ask summary:** Programmatically edit numeric/category data in embedded charts.
**Evidence:** `src/docx/chart.py` provides read + `add_chart` only (HISTORY "Charts read + add_chart"); no `replace_data`/ChartData writer.
**TODO (if applicable):** Add `Chart.replace_data` / ChartData rewrite (mirror python-pptx). L

### upstream#1139 — Wish there is a way to set up Asia text font
**Verdict:** resolved-in-loadfix
**Ask summary:** Set East-Asian font name (`rFonts/@w:eastAsia`) on a run.
**Evidence:** `src/docx/text/font.py:596` `name_far_east` / `name_east_asian`.
**TODO (if applicable):** n/a

### upstream#1134 — Capture Bullet points
**Verdict:** resolved-in-loadfix
**Ask summary:** Read bullet / numbering info (ilvl, numId, abstractNum format).
**Evidence:** HISTORY D.9 (#22) numbering style control; `src/docx/numbering.py`; `paragraph.list_level`, `paragraph.numbering_definition` in `text/paragraph.py`.
**TODO (if applicable):** n/a

### upstream#1130 — How to attach a file into DOCX files?
**Verdict:** new-feature-needed
**Ask summary:** Embed arbitrary binary attachments (e.g. PDFs, Office files) in a .docx.
**Evidence:** `embedded_objects.py` is read-only; no `add_ole_object` / `add_embedded_file` writer.
**TODO (if applicable):** Add write API for embedded OLE / altChunk attachments. L

### upstream#1127 — Feature - Embed binary objects?
**Verdict:** new-feature-needed
**Ask summary:** Create/embed OLE binary objects (PDF, Excel, etc.) in a document.
**Evidence:** Same as #1130 — read-only EmbeddedObject class only.
**TODO (if applicable):** Add OLE embed write API (pairs with #1130). L

### upstream#1124 — Unparsable resolution_tag IFD-Entries
**Verdict:** new-bug-needed
**Ask summary:** JPEG with non-rational EXIF resolution tag crashes `_TiffParser._dpi` with TypeError.
**Evidence:** `src/docx/image/tiff.py:107` multiplies `dots_per_unit` by int/float; no type-guard.
**TODO (if applicable):** Guard `_dpi` against non-numeric `dots_per_unit`; fall back to 72 dpi. S

### upstream#1123 — paragraph.text is missing text (tracked changes)
**Verdict:** resolved-in-loadfix
**Ask summary:** `paragraph.text` should include inserted/deleted tracked-change text.
**Evidence:** HISTORY Phase B (#53, #7); `tracked_changes.py`, `paragraph.revision_marks_text()` at `text/paragraph.py:1185`.
**TODO (if applicable):** n/a

### upstream#1121 — header images share id-space with document
**Verdict:** new-bug-needed
**Ask summary:** `docPr/@id` collisions between body and header drawings corrupt docs.
**Evidence:** `oxml/drawing.py:201-246` hardcodes `shape_id`; no document-wide id registry.
**TODO (if applicable):** Allocate `docPr/@id` from a document-scoped counter across all stories. M

### upstream#1120 — Tables extraction from DOCX (hidden columns)
**Verdict:** needs-investigation
**Ask summary:** Ignore hidden columns / hidden text when extracting tables from PDF-converted docx.
**Evidence:** `Font.hidden` read/write exists (`text/font.py:296`); no explicit "skip hidden column" table helper.
**TODO (if applicable):** Consider `Table.visible_cells` / skip `w:vanish` helper. S

### upstream#1119 — How to add content after a text box
**Verdict:** needs-investigation
**Ask summary:** Insert paragraphs before/after a text box, not just edit its content.
**Evidence:** `drawing/__init__.py` exposes txbxContent read; HISTORY D.13 covers paragraph insert but not textbox anchoring.
**TODO (if applicable):** Add `Drawing.insert_paragraph_before/after` sugar. S

### upstream#1112 — How can I draw rounded rectangle inside the document?
**Verdict:** new-feature-needed
**Ask summary:** Create DrawingML preset shapes (e.g. roundRect) programmatically.
**Evidence:** HISTORY D.27 adds read of DrawingML shapes; `oxml/drawing.py:200` builds rectangles via low-level helpers but no public `add_shape` API.
**TODO (if applicable):** Add `Document.add_shape(WD_SHAPE, ...)` writer. M

### upstream#1111 — document.save() with colon in filename silently fails
**Verdict:** new-bug-needed
**Ask summary:** On Windows, `Document.save('Test:Test.docx')` silently writes an empty file with no error.
**Evidence:** `opc/package.py` save path doesn't validate filename; no sanitization.
**TODO (if applicable):** Raise OSError/ValueError on Windows-invalid filename chars in `Document.save`. S

### upstream#1110 — edit other tab? (checkbox form controls)
**Verdict:** resolved-in-loadfix
**Ask summary:** Toggle a Word checkbox / form control.
**Evidence:** `content_controls.py:30` CHECKBOX type; `form_fields.py:343` CheckboxFormField; HISTORY D.14 (#27) + legacy form fields (#123).
**TODO (if applicable):** n/a

### upstream#1109 — Introduce feature to add section breaks
**Verdict:** resolved-in-loadfix
**Ask summary:** Add section breaks to a document programmatically.
**Evidence:** `document.py:257` `Document.add_section` (already upstream); no loadfix gap.
**TODO (if applicable):** n/a

### upstream#1108 — bold don't working
**Verdict:** out-of-scope
**Ask summary:** User bug in their own code (`.isnumeric` called as attribute instead of method).
**Evidence:** Not a library defect.
**TODO (if applicable):** n/a

### upstream#1106 — Unable to get text from shape and mathematical formula
**Verdict:** resolved-in-loadfix
**Ask summary:** Extract text inside shapes/textboxes and math formulas.
**Evidence:** `equations.py:74` `Equation.text`; `drawing/__init__.py:114` `txbxContent.text`; HISTORY "Equation read" (#113) + D.27 text-box (#75).
**TODO (if applicable):** n/a

### upstream#1105 — "image part with relationship rID8 was not found" error
**Verdict:** resolved-in-loadfix
**Ask summary:** Open docx with dangling image relationships without crashing.
**Evidence:** HISTORY "Add recover=True mode for malformed .docx" (#151); `opc/package.py:130` `open(..., recover=True)`.
**TODO (if applicable):** n/a

### upstream#1103 — About Attachments
**Verdict:** new-feature-needed
**Ask summary:** Extract attached files (embedded OLE/altChunk) from a docx.
**Evidence:** `embedded_objects.py` read-only supports OLE blob bytes; altChunk not covered.
**TODO (if applicable):** Add altChunk read + `Document.attachments` helper. M

### upstream#1102 — add_column is corrupting my file
**Verdict:** new-bug-needed
**Ask summary:** `_Columns.add_column(width)` only adds `w:gridCol` but not a `w:tc` per row, producing an unreadable doc.
**Evidence:** `src/docx/table.py:76` `add_column` adds only `gridCol`; see also 1431 `_Column`.
**TODO (if applicable):** Make `add_column` also insert matching `w:tc` in every row. S

### upstream#1099 — Python can't find Document module
**Verdict:** out-of-scope
**Ask summary:** User install/tooling issue (`pip install python-docx-1` — wrong package name).
**Evidence:** Not a library defect.
**TODO (if applicable):** n/a

## Batch summary
- resolved-in-loadfix: 9 (1156, 1147, 1139, 1134, 1123, 1110, 1109, 1106, 1105)
- new-feature-needed: 7 (1150, 1144, 1141, 1130, 1127, 1112, 1103)
- new-bug-needed: 4 (1124, 1121, 1111, 1102)
- needs-investigation: 2 (1120, 1119)
- out-of-scope: 3 (1154, 1108, 1099)
- Total: 25

### upstream#1094 — Hyphen/Dashes are removed from tables
**Verdict:** needs-investigation
**Ask summary:** User claims hyphens in table text content are stripped, citing the `Table.style` docstring that says "hyphen must be removed".
**Evidence:** The docstring at `src/docx/table.py:421` refers only to UI→internal style-name translation (see `BabelFish` in `src/docx/styles/__init__.py`); there is no code that strips hyphens from cell text. Likely user misreading, no reproducer attached.
**TODO (if applicable):** Clarify docstring wording to avoid confusion — S.

### upstream#1091 — font size for merged cells
**Verdict:** needs-investigation
**Ask summary:** Setting font.size on a merged table cell doesn't take effect.
**Evidence:** `src/docx/oxml/table.py` exposes vMerge/gridSpan handling and `Cell.merge_origin`/`is_merge_origin` exist; but no specific handling for font-size propagation on merged cells.
**TODO (if applicable):** Reproduce and document that user must target `merge_origin.paragraphs[0].runs[0]` — S.

### upstream#1089 — Font.name not applying to Headings
**Verdict:** needs-investigation
**Ask summary:** `styles['Normal'].font.name = 'Times New Roman'` does not change heading font (headings use theme fonts like Calibri Light).
**Evidence:** Style font handling via `BabelFish` name mapping intact; headings inherit from theme via `w:rFonts` asciiTheme/hAnsiTheme which override Normal. Not a bug — FAQ candidate.
**TODO (if applicable):** Add docs note on theme-font override in heading styles — S.

### upstream#1087 — feature: add endnote
**Verdict:** resolved-in-loadfix
**Ask summary:** Request API to add endnotes / cross-references.
**Evidence:** `src/docx/endnotes.py` Endnotes.add(); HISTORY "Phase A — Footnotes and endnotes"; REF/PAGEREF under Phase C (#115). Cross-refs via fields.

### upstream#1086 — Max Memory Error - 10MB Max (AttValue too long)
**Verdict:** new-bug-needed
**Ask summary:** Large embedded images/attributes (>10MB AttValue) fail to parse; can the limit be raised?
**Evidence:** `src/docx/oxml/parser.py:22` sets `huge_tree=False`; no opt-in for huge_tree.
**TODO (if applicable):** Add `Document(..., huge_tree=True)` opt-in to parser config — S.

### upstream#1084 — extract number of pages
**Verdict:** new-feature-needed
**Ask summary:** Return number of pages in the document.
**Evidence:** `src/docx/statistics.py` counts paragraphs/words/chars only; no pages count. app.xml `Pages` not exposed.
**TODO (if applicable):** Add `DocumentStatistics.pages` (read cached value from app.xml ExtendedProperties) — M.

### upstream#1083 — Add styles between documents
**Verdict:** new-feature-needed
**Ask summary:** Copy a style definition from one document into another.
**Evidence:** `src/docx/styles/styles.py:55` `add_style()` only creates a blank style of a given type; no cross-document copy helper.
**TODO (if applicable):** Add `Styles.import_style(source_style)` to deep-copy style XML into target — M.

### upstream#1082 — Trouble recognising tables
**Verdict:** out-of-scope
**Ask summary:** Tables produced by python-pdf2docx are missed; user asks python-docx to detect them anyway.
**Evidence:** Issue lies in upstream pdf2docx converter producing non-table markup; python-docx correctly reflects what's in the XML.

### upstream#1081 — Bold not working for Arabic/Persian
**Verdict:** resolved-in-loadfix
**Ask summary:** `run.bold = True` does not make Arabic text bold in Word/LibreOffice (they require complex-script `w:bCs`).
**Evidence:** `src/docx/text/font.py:245` exposes `cs_bold` (and `cs_italic`, `rtl`, `bidi_language`); HISTORY "Add RTL / bidi on Paragraph and Run (#127)" and Font.language additions (#160).

### upstream#1080 — strange bug about cv2 and docx
**Verdict:** out-of-scope
**Ask summary:** Empty body; likely import-order issue between cv2 and docx.
**Evidence:** No reproducer, no python-docx code implicated.

### upstream#1079 — Print user-friendly message on section IndexError
**Verdict:** out-of-scope
**Ask summary:** Tutorial-level question on catching IndexError from `sections[1]`. Their try/except typos `WordFile.section[0]` (missing 's').
**Evidence:** `src/docx/section.py:30` raises normal IndexError; standard Python try/except applies.

### upstream#1077 — python-docx on Macbook M1 (lxml _xmlFree)
**Verdict:** out-of-scope
**Ask summary:** `ImportError: symbol not found '_xmlFree'` on Apple Silicon — an lxml install issue.
**Evidence:** Error originates in `lxml/etree.cpython-310-darwin.so`, unrelated to docx code.

### upstream#1075 — _Cell class and add text by cell number
**Verdict:** out-of-scope
**Ask summary:** Tutorial question — add text to a cell using linear index instead of (row, col).
**Evidence:** Can be done via `table._cells[i]` today; no API change warranted.

### upstream#1074 — #455 is not fixed yet
**Verdict:** resolved-in-loadfix
**Ask summary:** Claims next_id increment fix from #455 regressed.
**Evidence:** `src/docx/parts/story.py:131` `next_id` uses `max(used_ids) + 1` (doesn't fill gaps) — HISTORY 0.8.7 records "#455 increment next_id, don't fill gaps". Fork retains correct implementation.

### upstream#1071 — paragraph_format.AddSpaceBetweenFarEastAndDigit
**Verdict:** new-feature-needed
**Ask summary:** Expose the "auto-space between East-Asian and digits" paragraph flag (Word option `w:autoSpaceDN`).
**Evidence:** `src/docx/oxml/text/parfmt.py:195` lists `w:autoSpaceDE`/`w:autoSpaceDN` in tag sequence but no proxy property on `ParagraphFormat`.
**TODO (if applicable):** Add `ParagraphFormat.auto_space_de` / `auto_space_dn` boolean properties — S.

### upstream#1070 — Run._element copy mechanism
**Verdict:** out-of-scope
**Ask summary:** Confusion: `run.font.name = ...` followed by `run._element.rPr.rFonts.set(...)` inside a function raises because font.name assignment triggers rFonts creation that the local ref missed.
**Evidence:** Font.name setter in `src/docx/text/font.py` works through element; user is hitting a stale reference pattern, not a library defect.

### upstream#1069 — Extract text from header2.xml
**Verdict:** resolved-in-loadfix
**Ask summary:** How to read paragraphs/tables inside a specific header XML (even-page, first-page).
**Evidence:** `src/docx/section.py:147` `even_page_header`, `:167` `first_page_header`, `:214` `header` all return `_Header` (BlockItemContainer) with paragraphs/tables access. HISTORY "Add Section odd/even page header-footer (#149)".

### upstream#1068 — Upload wheels to PyPI
**Verdict:** out-of-scope
**Ask summary:** Request wheel distributions on PyPI for python-openxml/python-docx.
**Evidence:** Release/distribution concern of upstream project; loadfix fork packaging not impacted.

### upstream#1065 — OxmlElement discriminating against valid tags (w:keepNext)
**Verdict:** resolved-in-loadfix
**Ask summary:** `OxmlElement('w:keepNext')` appended to `w:tr` vanishes on save, unlike `w:cantSplit`.
**Evidence:** `src/docx/oxml/__init__.py:575` registers `w:keepNext` as `CT_OnOff`; note keepNext is a paragraph-property child of `w:pPr`, not `w:tr` — user was inserting in wrong parent. Serializer keeps registered elements. No bug.

### upstream#1060 — Cannot generate word directory structure (TOC)
**Verdict:** resolved-in-loadfix
**Ask summary:** Ask for a method to generate a table-of-contents / directory.
**Evidence:** `src/docx/toc.py` plus `Document.add_table_of_contents`, `Paragraph.insert_table_of_contents_before/after` (HISTORY `#116`).

### upstream#1059 — Demo problem regression (lists/table style)
**Verdict:** needs-investigation
**Ask summary:** The bundled example script's numbered/bulleted lists and table style don't render after 0.8.4.
**Evidence:** No obvious fix in HISTORY 1.3.0.dev0; fork retains default template; bisect points to upstream 0.8.5.
**TODO (if applicable):** Re-run attached demo against fork default template; if broken, audit `src/docx/templates/default.docx` — M.

### upstream#1057 — add_paragraph auto-escape behaviour
**Verdict:** out-of-scope
**Ask summary:** Interaction with python-docx-template escaping; user sees escaped `<,>,&` in rendered output.
**Evidence:** `Paragraph.add_run` stores raw text; lxml handles escape on serialize automatically. Issue is in docxtpl template rendering, not docx.

### upstream#1056 — Preserve formatting
**Verdict:** resolved-in-loadfix
**Ask summary:** Replace text in specific columns while preserving runs' formatting (bold/italic/colour).
**Evidence:** HISTORY "D.10 Search and replace with formatting preservation (#91)", `src/docx/search.py`, `Document.replace_regex` etc.

### upstream#1055 — Draw red dividing line / remove heading-0 line
**Verdict:** needs-investigation
**Ask summary:** (1) Insert a colored horizontal rule; (2) suppress the underline/line appearing under `add_heading(level=0)`.
**Evidence:** Paragraph borders added via HISTORY "D.7 Paragraph borders (#109)" — red border on paragraph can fake a rule; heading-0 line comes from Title style's `w:pBdr`.
**TODO (if applicable):** Document pattern for horizontal-rule paragraph + Title border override — S.

### upstream#1054 — Convert Document object directly to PDF
**Verdict:** out-of-scope
**Ask summary:** Pass a `Document` object to `docx2pdf.convert()` without writing to disk.
**Evidence:** `docx2pdf` is a separate package requiring a path/filename; python-docx can't change its API.

## Batch summary
- resolved-in-loadfix: 7 (#1087, #1081, #1074, #1069, #1065, #1060, #1056)
- new-feature-needed: 3 (#1084, #1083, #1071)
- new-bug-needed: 1 (#1086)
- needs-investigation: 5 (#1094, #1091, #1089, #1059, #1055)
- out-of-scope: 9 (#1082, #1080, #1079, #1077, #1075, #1070, #1068, #1057, #1054)

### upstream#1053 — Read equation as linear(latex) or images from word document
**Verdict:** resolved-in-loadfix
**Ask summary:** Read math equations (as LaTeX/linear) and images from a .docx.
**Evidence:** `src/docx/equations.py` (Equation read API); images already supported upstream.
**TODO (if applicable):** n/a

### upstream#1052 — Seach string and insert page break
**Verdict:** resolved-in-loadfix
**Ask summary:** Find a string and insert a page break at that location.
**Evidence:** `src/docx/search.py` (Phase D.10 search/replace with formatting preservation); `Run.add_break` exists upstream.
**TODO (if applicable):** n/a

### upstream#1050 — Saving/Exporting documents to jpg/jpeg/gif etc
**Verdict:** out-of-scope
**Ask summary:** Render/export a .docx as a raster image.
**Evidence:** no match; requires a Word-rendering engine, outside python-docx.
**TODO (if applicable):** n/a

### upstream#1049 — Calling an existing Word macro
**Verdict:** out-of-scope
**Ask summary:** Execute (run) a VBA macro stored in a document from Python.
**Evidence:** loadfix supports .docm round-trip (`D.24`) but VBA execution requires Word/COM.
**TODO (if applicable):** n/a

### upstream#1048 — alternative text (alt_text) on Table (tblCaption, tblDescription)
**Verdict:** new-feature-needed
**Ask summary:** Expose Table.alt_text (w:tblCaption) and alt_description (w:tblDescription).
**Evidence:** `src/docx/oxml/table.py:689-690` lists tblCaption/tblDescription in _tag_seq but no proxy accessors in `src/docx/table.py`.
**TODO (if applicable):** Add Table.alt_text / Table.alt_description read-write properties (CT_TblPr). Size S.

### upstream#1047 — Detect page break within tables
**Verdict:** out-of-scope
**Ask summary:** Detect page-break boundaries inside a large table for chunked rendering.
**Evidence:** page layout is determined by Word at render time; no DOM info.
**TODO (if applicable):** n/a

### upstream#1046 — Background color of row/cell
**Verdict:** resolved-in-loadfix
**Ask summary:** Set shading/background color on a row or cell.
**Evidence:** Phase D.6 (cell shading) — `src/docx/table.py` shading; row shading via oxml.
**TODO (if applicable):** n/a

### upstream#1045 — Page numbers and internal hyperlinks in TOC
**Verdict:** resolved-in-loadfix
**Ask summary:** Build TOC with page numbers and internal hyperlinks.
**Evidence:** `src/docx/toc.py` (Document.add_table_of_contents); Word TOC fields render page numbers and hyperlinks natively.
**TODO (if applicable):** n/a

### upstream#1042 — Bit-identical docx files (stable zip timestamps)
**Verdict:** new-feature-needed
**Ask summary:** Make _ZipPkgWriter write deterministic timestamps so repeated saves are bit-identical.
**Evidence:** `src/docx/opc/phys_pkg.py:175-178` still uses `writestr(membername, blob)` with no ZipInfo.
**TODO (if applicable):** Add opt-in deterministic mode (ZipInfo date_time=(1980,1,1,0,0,0)) on Document.save. Size S.

### upstream#1041 — How to add a hyperlink to a table
**Verdict:** resolved-in-loadfix
**Ask summary:** User-facing recipe for adding hyperlinks in a table cell.
**Evidence:** Phase D.1 hyperlink creation — `src/docx/text/paragraph.py:162 add_hyperlink()` works in cell paragraphs.
**TODO (if applicable):** n/a

### upstream#1039 — Footnote style not working ("Footnote Reference" KeyError)
**Verdict:** resolved-in-loadfix
**Ask summary:** `document.styles['Footnote Reference']` missing in default template.
**Evidence:** Phase A footnotes ensures FootnoteReference style; `src/docx/oxml/footnotes.py:50` writes `w:rStyle w:val="FootnoteReference"`. Style present in enum `FOOTNOTE_REFERENCE`.
**TODO (if applicable):** n/a

### upstream#1037 — Setting core_properties.last_modified_by duplicates element
**Verdict:** needs-investigation
**Ask summary:** After setting `last_modified_by`, core.xml contains two `cp:lastModifiedBy` entries; document invalid.
**Evidence:** `src/docx/oxml/coreprops.py:152` uses ZeroOrOne (should not duplicate) but this is a parse-time condition from the original file; no test covers the duplicate case.
**TODO (if applicable):** Add regression test that opens a core.xml with duplicate lastModifiedBy and verifies single-element result. Size S.

### upstream#1035 — Retrieve shapes (triangle/square/line/textbox) from .docx
**Verdict:** resolved-in-loadfix
**Ask summary:** Read non-image shapes (text boxes, autoshapes) from a document.
**Evidence:** Phase D.27 DrawingML shapes — `src/docx/drawing/__init__.py` exposes shapes/text_box/group_shapes.
**TODO (if applicable):** n/a

### upstream#1034 — Problem in pyinstaller PY TO EXE
**Verdict:** out-of-scope
**Ask summary:** PyInstaller packaging issue, no details.
**Evidence:** pkg issue unrelated to python-docx source.
**TODO (if applicable):** n/a

### upstream#1033 — Change the size of List Bullet
**Verdict:** out-of-scope
**Ask summary:** Change font size of "List Bullet" style output.
**Evidence:** already supported via `document.styles['List Bullet'].font.size` / per-run font.size.
**TODO (if applicable):** n/a

### upstream#1030 — Table.width getter/setter
**Verdict:** resolved-in-loadfix
**Ask summary:** Set total table width and have columns scale.
**Evidence:** `src/docx/table.py:286 preferred_width` (Phase D.26 table autofit/column width), plus `table.autofit`.
**TODO (if applicable):** n/a

### upstream#1027 — ParagraphFormat.font (rPr at pPr level)
**Verdict:** out-of-scope
**Ask summary:** Expose a Font proxy on ParagraphFormat for `pPr/rPr`.
**Evidence:** per scanny/HtheChemist in thread, rPr only affects paragraph-mark glyph, so low value; not implemented.
**TODO (if applicable):** n/a

### upstream#1025 — Adding paragraph not recorded as tracked insertion
**Verdict:** new-feature-needed
**Ask summary:** When track-changes mode is set in the doc, inserts via `add_paragraph` should be wrapped in `w:ins`.
**Evidence:** Phase B handles read/accept/reject only; no create API wrapping new runs in `w:ins`. See `src/docx/tracked_changes.py`.
**TODO (if applicable):** Add optional `author=/date=` to add_paragraph/add_run that emits `w:ins`. Size M.

### upstream#1023 — Run.add_ole_object()
**Verdict:** new-feature-needed
**Ask summary:** Add ability to embed OLE object (e.g. PDF, workbook) as w:object in a run.
**Evidence:** `src/docx/embedded_objects.py` is read-only ("Creation and modification are intentionally not supported"). Parts infrastructure exists in `src/docx/parts/embedded_object.py`.
**TODO (if applicable):** Add Run.add_ole_object(fn, icon, prog_id) writing o:OLEObject + embedding part. Size L.

### upstream#1021 — Hyperlinks not being highlighted
**Verdict:** needs-investigation
**Ask summary:** When applying highlight to a paragraph that contains hyperlinks, the hyperlink runs stay unhighlighted.
**Evidence:** `Font.highlight_color` exists (`src/docx/text/font.py:309`); iterator over paragraph runs likely skips `w:hyperlink/w:r`.
**TODO (if applicable):** Verify highlight iterator visits hyperlink runs; add test + fix if gap. Size S.

### upstream#1019 — Copyright of test images under features/steps/test_files/
**Verdict:** out-of-scope
**Ask summary:** Replace Lena and other unclear-licence test images for Debian packaging.
**Evidence:** concerns upstream repo contents; loadfix inherits same images.
**TODO (if applicable):** n/a (cosmetic repo hygiene — optionally swap for CC0 images; Size S).

### upstream#1015 — Not all tables extracted (nested tables)
**Verdict:** resolved-in-loadfix
**Ask summary:** `document.tables` only returns top-level tables; nested ones missed.
**Evidence:** Documented limitation in `src/docx/document.py:1087-1096`; cell iteration supports nested access via `cell.tables`. Recursion helper is standard recipe.
**TODO (if applicable):** n/a (documentation-level — optional convenience iterator).

### upstream#1012 — Add auto-numbering to headings (1, 1.1, 1.1.1)
**Verdict:** resolved-in-loadfix
**Ask summary:** Apply multilevel numbering to Heading styles.
**Evidence:** Phase D.9 numbering style control — `src/docx/numbering.py` `NumberingDefinition.apply_to(paragraph)`.
**TODO (if applicable):** n/a

### upstream#1011 — Support for adding/updating math equations
**Verdict:** resolved-in-loadfix
**Ask summary:** Add/update OMML equations programmatically.
**Evidence:** "Equation read + minimal create API (#113)" in HISTORY.rst; `src/docx/equations.py`.
**TODO (if applicable):** n/a

### upstream#1009 — Add footnote in cell (superscript number)
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert a footnote reference inside a table cell, rendered as superscript.
**Evidence:** Phase A `Document.add_footnote` wraps `w:footnoteReference` with `w:rStyle="FootnoteReference"` — which is a superscript style per default styles.xml.
**TODO (if applicable):** n/a

## Batch summary
- resolved-in-loadfix: 12 (#1053, #1052, #1046, #1045, #1041, #1039, #1035, #1030, #1015, #1012, #1011, #1009)
- new-feature-needed: 4 (#1048, #1042, #1025, #1023)
- new-bug-needed: 0
- needs-investigation: 2 (#1037, #1021)
- out-of-scope: 7 (#1050, #1049, #1047, #1034, #1033, #1027, #1019)
- Total: 25

### upstream#1007 — NO numbering num to abs_num map
**Verdict:** resolved-in-loadfix
**Ask summary:** Provide API to get abstractNum (with numFmt / lvl info) from a num id.
**Evidence:** `src/docx/numbering.py` exposes AbstractNum, abstract_num_id, num_id mapping; `Level.number_format`/`numFmt_val` at lines 193/310.
**TODO (if applicable):** n/a

### upstream#1002 — Adding/Loading hyperlink pictures (URLs) from web/external
**Verdict:** new-feature-needed
**Ask summary:** Allow inserting images referenced by URL (a:blip@r:link) that Word loads on open rather than embedding.
**Evidence:** `src/docx/shape.py:306` reads `r:link` but there is no API to create a linked (external) picture.
**TODO (if applicable):** Add `add_picture(url=..., linked=True)` using r:link blip relationship. M

### upstream#999 — package hangs trying to find default style
**Verdict:** needs-investigation
**Ask summary:** Reports a hang in `Styles._iter_styles` xpath when looking up the default style for a paragraph.
**Evidence:** `src/docx/styles/styles.py:70 default_for` still uses same `_element.default_for`; no reported fix for perf/hang in history.
**TODO (if applicable):** Reproduce with pathological styles.xml and add guard or fallback. M

### upstream#998 — Check boxes in the docx table
**Verdict:** resolved-in-loadfix
**Ask summary:** Cell.text drops checkbox form fields; user wants checked/unchecked state when extracting text.
**Evidence:** `src/docx/form_fields.py` provides CheckboxFormField with `checked`; content_controls.py handles w14:checkbox SDT.
**TODO (if applicable):** n/a

### upstream#997 — Detecting lists inside a table
**Verdict:** resolved-in-loadfix
**Ask summary:** Detect bulleted/numbered lists inside table cells.
**Evidence:** Paragraphs inside cells expose numPr via numbering API (`src/docx/numbering.py` + `Paragraph.style`); numbering phase D.9 (#22).
**TODO (if applicable):** n/a

### upstream#995 — Wrong image size created when width is used!
**Verdict:** out-of-scope
**Ask summary:** Complaint that `add_picture(width=Inches(1))` renders at 120% — really a Word DPI/EMU interpretation question.
**Evidence:** No bug in lib; image EMU dims correct. Documentation note at best.
**TODO (if applicable):** n/a

### upstream#993 — Bad code #2 in Quickstart (style lookup warning)
**Verdict:** out-of-scope
**Ask summary:** Upstream docs use deprecated style_id ('LightShading-Accent1') causing UserWarning.
**Evidence:** Warning logic lives in `src/docx/styles/styles.py:96`; quickstart doc is upstream's issue.
**TODO (if applicable):** Update loadfix quickstart if still using deprecated id. S

### upstream#992 — 'List out of range' error: vMerge first cell + unaligned table border
**Verdict:** resolved-in-loadfix
**Ask summary:** Table._cells crashes when first row has vMerge=continue in an irregular/merged grid.
**Evidence:** `src/docx/table.py:510` guards `if len(cells) >= col_count` before `cells[-col_count]`.
**TODO (if applicable):** n/a

### upstream#990 — AttributeError: 'tuple' object has no attribute 'qty'
**Verdict:** out-of-scope
**Ask summary:** Upstream quickstart example relies on a namedtuple the user didn't define.
**Evidence:** Docs-only; no library change needed.
**TODO (if applicable):** n/a

### upstream#989 — Unwanted duplicate pages under specific conditions
**Verdict:** needs-investigation
**Ask summary:** Table + blank paragraph layout triggers content duplication across pages on render in Word.
**Evidence:** No reproducible code; likely Word rendering quirk not library bug.
**TODO (if applicable):** Attempt reproduction with minimal script. S

### upstream#988 — Single character modify style
**Verdict:** resolved-in-loadfix
**Ask summary:** User wants to re-style a single character inside a run sharing formatting with its neighbours.
**Evidence:** `Run.split(offset)` at `src/docx/text/run.py:268`; also search/replace phase D.10 (#91).
**TODO (if applicable):** n/a

### upstream#987 — Numbering in table
**Verdict:** resolved-in-loadfix
**Ask summary:** New rows added via `table.add_row()` don't continue an existing numbered list in first column.
**Evidence:** Numbering style control phase D.9 (#22) lets paragraphs carry numPr; user applies paragraph to cell.
**TODO (if applicable):** n/a

### upstream#983 — Making executable (pyinstaller)
**Verdict:** out-of-scope
**Ask summary:** PyInstaller misses template files; packaging question not library.
**Evidence:** `src/docx/templates/` bundled via package_data; downstream packaging issue.
**TODO (if applicable):** n/a

### upstream#981 — run.add_picture() won't insert image
**Verdict:** needs-investigation
**Ask summary:** `paragraph.add_run().add_picture(...)` runs without error but image absent on open.
**Evidence:** `src/docx/text/run.py:62 add_picture` exists; likely run is empty/unsaved or existing content pushes image off. Needs repro.
**TODO (if applicable):** Add regression test for `Run.add_picture` on existing paragraph. S

### upstream#980 — isolate_run()
**Verdict:** resolved-in-loadfix
**Ask summary:** Helper snippet from maintainer to split runs to isolate a character range.
**Evidence:** `Run.split(offset)` at `src/docx/text/run.py:268` implements this functionality.
**TODO (if applicable):** n/a

### upstream#979 — Table repeat cell text if page changes
**Verdict:** out-of-scope
**Ask summary:** Wants cell text to auto-append "(continue)" on new page when cell splits.
**Evidence:** Word layout feature; OOXML cannot express runtime duplication. No fix possible.
**TODO (if applicable):** n/a

### upstream#976 — Setting character spacing = expanded
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose `w:spacing` (expanded/condensed) on Font.
**Evidence:** `Font.character_spacing` at `src/docx/text/font.py:41` reads/writes `w:spacing/@w:val`.
**TODO (if applicable):** n/a

### upstream#973 — font.rtl = True disturbs font.size
**Verdict:** needs-investigation
**Ask summary:** With rtl=True, font.size set on style is ignored; likely need to also set szCs (complex-script size).
**Evidence:** `Font.size` writes `w:sz` only; `w:szCs` may be needed when bidi. See rtl/bidi support (#127, #160).
**TODO (if applicable):** Ensure Font.size also sets szCs or add `complex_script_size`. S

### upstream#971 — Distinguish between bulleted list and numbered list
**Verdict:** resolved-in-loadfix
**Ask summary:** Differentiate bullet vs numbered list paragraphs when both use ListParagraph style.
**Evidence:** `src/docx/numbering.py` exposes Level.numFmt including "bullet"/"decimal" via abstractNum lookup.
**TODO (if applicable):** n/a

### upstream#966 — How can I remove last page from a word document
**Verdict:** out-of-scope
**Ask summary:** Remove the last rendered "page" — Word has no static page concept in OOXML.
**Evidence:** Pages are reflow output; only sections exist. No API possible.
**TODO (if applicable):** n/a

### upstream#965 — Content Control tags not updating in word after editing w:t
**Verdict:** needs-investigation
**Ask summary:** Editing SDT w:t text has no effect because dataBinding overrides content from customXml at open.
**Evidence:** `src/docx/content_controls.py` + `oxml/content_controls.py:252 dataBinding` read. No helper to update the bound customXml part.
**TODO (if applicable):** Add ContentControl.write-through to customXml target or document workaround. M

### upstream#959 — Encoding special characters
**Verdict:** out-of-scope
**Ask summary:** Polish diacritics saved as mojibake; almost certainly caller's source encoding, not lib.
**Evidence:** lxml serialises unicode correctly; library has no double-encode path.
**TODO (if applicable):** n/a

### upstream#954 — Missing Formatting of content in Table
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose run-level formatting inside cells (not just cell.text).
**Evidence:** Iterating `cell.paragraphs[i].runs[j].font` already provides full formatting; D.6/D.20 shading added.
**TODO (if applicable):** n/a

### upstream#950 — How can I add elements in a paragraph? but not in a run.
**Verdict:** resolved-in-loadfix
**Ask summary:** User needs to insert `w:bookmarkStart` as a direct child of `w:p`, not inside a run.
**Evidence:** `Paragraph.add_bookmark` at `src/docx/text/paragraph.py:56` (Phase C, #52).
**TODO (if applicable):** n/a

### upstream#948 — symbol font not supported
**Verdict:** resolved-in-loadfix
**Ask summary:** `w:sym` elements are ignored when reading run text, preventing round-trip.
**Evidence:** `Run.symbols` iterator + `CT_Sym` at `src/docx/text/run.py:122`, `src/docx/oxml/text/run.py:357` (phase-other #114). Note: `Run.text` may still not inline sym glyphs.
**TODO (if applicable):** Optionally include symbol char in `Run.text` representation. S

## Batch summary
- resolved-in-loadfix: 11
- new-feature-needed: 1
- new-bug-needed: 0
- needs-investigation: 6
- out-of-scope: 7

### upstream#947 — How to add run in paragraph using python-docx
**Verdict:** out-of-scope
**Ask summary:** User support question on copying runs between documents. Not a defect or feature request.
**Evidence:** no match — usage question.

### upstream#945 — How can I extracted highlighted text?
**Verdict:** resolved-in-loadfix
**Ask summary:** User wants to read a run's highlight color, not just set it.
**Evidence:** `src/docx/text/font.py:309` exposes `Font.highlight_color` getter (returns `WD_COLOR_INDEX | None`).

### upstream#940 — _next_numId takes too long on big documents
**Verdict:** new-bug-needed
**Ask summary:** O(n^2) linear scan in `CT_Numbering._next_numId` becomes a major bottleneck for documents with thousands of tables/numbered lists.
**Evidence:** `src/docx/oxml/numbering.py:297` still uses the original range-scan algorithm.
**TODO:** Optimize `_next_numId` and `_next_abstractNumId` via max(+1) fast path or gap-cache while preserving gap-fill semantics; add perf regression test. S.

### upstream#939 — Split cells in tables are read wrong
**Verdict:** needs-investigation
**Ask summary:** After splitting a previously-merged cell in Word, python-docx reports wrong cell layout (cells "leak" into neighbouring rows).
**Evidence:** vMerge/gridSpan logic in `src/docx/table.py:662` present but not verified against this fixture; no related loadfix commit.
**TODO:** Reproduce with the attached demo-word.docx, add failing test, fix cell-grid reconstruction for split-after-merge case. M.

### upstream#935 — How set font spacing?
**Verdict:** resolved-in-loadfix
**Ask summary:** User asks how to change inter-character spacing (kerning-style, w:spacing val).
**Evidence:** `src/docx/text/font.py:41` `Font.character_spacing` getter/setter; documented at `docs/user/text-advanced.rst:95`.

### upstream#932 — Missing text fragments when SmartTags are used
**Verdict:** new-bug-needed
**Ask summary:** Runs nested inside `<w:smartTag>` elements are skipped when iterating paragraph text, losing content like "Diyarbakır".
**Evidence:** No `w:smartTag` handling in `src/docx/oxml/text/paragraph.py` or run iteration; only settings-level `w:smartTagType` is referenced.
**TODO:** Treat `w:smartTag` (and nested `w:customXml`/`w:sdt`) as transparent containers in run/text iteration; add fixture-based test. M.

### upstream#930 — Is there a way to read the revisions?
**Verdict:** resolved-in-loadfix
**Ask summary:** User wants to read tracked-change (revision) content from paragraphs.
**Evidence:** Phase B tracked-changes API — `src/docx/tracked_changes.py`, HISTORY "Add read of tracked insertions and deletions (#53)".

### upstream#929 — ata frame cell by cell inpython
**Verdict:** out-of-scope
**Ask summary:** Empty body / garbled title; no actionable content.
**Evidence:** no match.

### upstream#928 — Document presentation error (WD_ALIGN_PARAGRAPH)
**Verdict:** resolved-in-loadfix
**Ask summary:** Docs showed import of `WD_PARAGRAPH_ALIGNMENT` but canonical name is `WD_ALIGN_PARAGRAPH`.
**Evidence:** `src/docx/enum/text.py:67` exports both names as aliases; docstring demonstrates `WD_ALIGN_PARAGRAPH` import.

### upstream#926 — feature: black format whole codebase
**Verdict:** out-of-scope
**Ask summary:** Request to bulk-reformat codebase with black.
**Evidence:** `pyproject.toml` uses ruff (line-length 100). Style choice for fork maintainers; not a user-facing feature.

### upstream#925 — Extracting hyperlinks from table cell
**Verdict:** resolved-in-loadfix
**Ask summary:** Hyperlink text inside cell paragraphs was missing; users want hyperlinks surfaced.
**Evidence:** `Paragraph.hyperlinks` and `Paragraph.iter_inner_content` in `src/docx/text/paragraph.py:708,919`; already upstream in 1.0.0 and carried into fork.

### upstream#924 — How to generate table of contents automatically
**Verdict:** resolved-in-loadfix
**Ask summary:** Programmatically insert a TOC (and PAGEREF fields) when building a docx.
**Evidence:** `Document.add_table_of_contents` at `src/docx/document.py:267`; `src/docx/toc.py`; HISTORY `#116` + Phase C REF/PAGEREF (#115).

### upstream#922 — can't get the font size
**Verdict:** out-of-scope
**Ask summary:** Support question; `font.size` returns None when size is inherited from style. Expected behaviour.
**Evidence:** no match — inheritance semantics already documented.

### upstream#921 — Add table caption and description (alt_text)
**Verdict:** new-feature-needed
**Ask summary:** Expose `w:tblCaption` and `w:tblDescription` as Table API (accessibility alt-text/title for tables).
**Evidence:** Tags listed in `src/docx/oxml/table.py:689-690` tag sequence but no `ZeroOrOne` declaration and no `Table.alt_text` / `Table.caption` proxy; contrast with InlineShape alt_text (HISTORY #158).
**TODO:** Add `Table.caption` and `Table.description` (alt_text) read/write properties with ZeroOrOne CT_String children; mirror InlineShape alt_text precedent. S.

### upstream#919 — How to add a hyperlink to a document?
**Verdict:** resolved-in-loadfix
**Ask summary:** Provide an API to append hyperlinks when creating a document.
**Evidence:** `Paragraph.add_hyperlink` at `src/docx/text/paragraph.py:162`; Phase D.1 `#97`.

### upstream#918 — Extracting text and images and write to new docx
**Verdict:** out-of-scope
**Ask summary:** Generic how-to for copy/transform workflow; no concrete python-docx bug.
**Evidence:** no match.

### upstream#917 — table with multiple rows in table header
**Verdict:** resolved-in-loadfix
**Ask summary:** Allow more than one row to be marked as the table-header (repeated on each page) via `w:tblHeader`.
**Evidence:** `src/docx/oxml/table.py:1389` defines `tblHeader` ZeroOrOne on trPr and `src/docx/oxml/table.py:348` exposes `_Row.is_header` setter; HISTORY "Add _Row.is_header (#93)".

### upstream#916 — creating linked image without stored with document
**Verdict:** new-feature-needed
**Ask summary:** Support inserting a picture linked to an external file (`r:link` / `a:blip r:link`) rather than embedded.
**Evidence:** Read-side handling of `r:link` exists at `src/docx/shape.py:306` and `src/docx/oxml/shape.py:343`, but no `add_linked_picture` / `link_to_file` API on `Run.add_picture` / `Document.add_picture` (`src/docx/document.py:238`).
**TODO:** Add `add_picture(..., link=True, save_with_document=False)` creating an external relationship with `TargetMode=External` and `r:link` on blip; cover both "link-only" and "link + embed". M.

### upstream#914 — how to get ole object type in word by python
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose the OLE progId / object type of an embedded OLE object.
**Evidence:** `src/docx/embedded_objects.py:84` exposes ProgID and related attrs; HISTORY "Add read-only embedded OLE objects (#140)".

### upstream#913 — how to extract txt or zip file from OLE in docx?
**Verdict:** out-of-scope
**Ask summary:** Q about using `olefile` to dig contents out of CFBF OLE streams; unrelated to python-docx API.
**Evidence:** no match — external-library question.

### upstream#911 — Add more core properties (Company, Manager)
**Verdict:** new-feature-needed
**Ask summary:** Expose extended (app.xml) properties such as `Company` and `Manager`, which are not in core (Dublin-Core) properties.
**Evidence:** Placeholders exist in `src/docx/templates/default-docx-template/docProps/app.xml` but no `ExtendedProperties` proxy or app.xml part wrapper; `custom_properties.py` covers only `docProps/custom.xml`.
**TODO:** Add `Document.extended_properties` (ExtendedPropertiesPart wrapping `docProps/app.xml`) with Company, Manager, Application, etc. read/write. M.

### upstream#909 — How to Add Line Numbers?
**Verdict:** resolved-in-loadfix
**Ask summary:** Add line numbering in the gutter for a section (`w:lnNumType`).
**Evidence:** `src/docx/oxml/section.py:310` defines `lnNumType`; HISTORY "Add Section.line_numbering (#122)".

### upstream#902 — Unable to read docx containing pictures linking to internal bookmarks
**Verdict:** new-bug-needed
**Ask summary:** Opening a docx fails with `KeyError: "There is no item named 'word/#MyBookmark'"` because an internal hyperlink relationship whose target is a `#bookmark` anchor is treated as a package part.
**Evidence:** `src/docx/opc/rel.py:57` / `pkgreader.py` walk treats relationships as internal parts; no anchor-only / fragment-only handling observed; `recover=True` (api.py:34) may not catch this.
**TODO:** Detect relationships whose `Target` is purely a `#fragment` (internal bookmark anchor) and skip package-reader lookup; add test with fixture `PictureBookmarks.docx`. S.

### upstream#895 — feature: floating images
**Verdict:** resolved-in-loadfix
**Ask summary:** Support reading/writing floating (wp:anchor) images — behind/in-front-of text.
**Evidence:** Phase D.17 in HISTORY — "Add Floating images with wp:anchor positioning (#30)"; `src/docx/drawing/` covers this.

### upstream#892 — Enhancement: OPC OOXML (Flat OPC) support
**Verdict:** new-feature-needed
**Ask summary:** Load/save Flat-OPC XML packages (`<pkg:package>` single-file XML representation emitted by Office-JS `getOoxml`).
**Evidence:** `grep -rn "xmlPackage\|FlatOPC\|pkg:package"` returns no hits in `src/docx/opc/`; only zip-based package reader.
**TODO:** Add Flat-OPC reader/writer (`Document(flat_opc=True)` and `Package.save_as_flat_opc`) wrapping a synthetic `PhysPkgReader` that walks `<pkg:part>` nodes. L.

## Batch summary
- resolved-in-loadfix: 11  (945, 935, 930, 928, 925, 924, 919, 917, 914, 909, 895)
- new-feature-needed:  4   (921, 916, 911, 892)
- new-bug-needed:      3   (940, 932, 902)
- needs-investigation: 1   (939)
- out-of-scope:        6   (947, 929, 926, 922, 918, 913)
- Total:               25

### upstream#888 — Applying the "next_paragraph_style" attribute
**Verdict:** new-feature-needed
**Ask summary:** User requests that python-docx either apply `style.next_paragraph_style` automatically when adding a paragraph, or at least document that this setting is not honoured by python-docx (only by Word).
**Evidence:** `src/docx/styles/style.py` exposes `next_paragraph_style` but `Document.add_paragraph` / `BlockItemContainer.add_paragraph` do not consult it. No mention in HISTORY.rst 1.3.0.dev0 or in `docs/user/styles-*.rst`.
**TODO (if applicable):** Document limitation and optionally auto-apply `next_paragraph_style` in `add_paragraph`. S

### upstream#861 — Table is not a rectangle
**Verdict:** resolved-in-loadfix
**Ask summary:** Non-rectangular tables (rows using `w:gridBefore` / `w:gridAfter` / `w:wBefore` / `w:wAfter`) mis-place cells when enumerated through `row.cells`.
**Evidence:** `src/docx/table.py:1634` adds `_Row.grid_cols_before` / `grid_cols_after` (versionadded 1.3.0.dev0) and the `Row.cells` docstring explicitly instructs callers to use them for matrix reconstruction.
**TODO (if applicable):** none.

### upstream#852 — No. of columns are wrong in some table objects
**Verdict:** resolved-in-loadfix
**Ask summary:** Same root cause as #861 — 3 `gridCol` declared but row has only 2 `w:tc` because a trailing `gridAfter` skips a column.
**Evidence:** `_Row.grid_cols_after` (src/docx/table.py:1634) and underlying `CT_TrPr.gridAfter` (src/docx/oxml/table.py:1383). Matches upstream fix #531/#1146 ("Index error on table with misaligned borders") already present.
**TODO (if applicable):** none.

### upstream#845 — feature: add watermark
**Verdict:** resolved-in-loadfix
**Ask summary:** Add the ability to insert a watermark (text or image) into a document.
**Evidence:** `src/docx/watermark.py`, docs `docs/user/watermarks.rst`, HISTORY.rst D.23 "Watermark support (text and image) (#36)", commit `0036485`.
**TODO (if applicable):** none.

### upstream#827 — Change reading direction to Right-To-Left not working
**Verdict:** resolved-in-loadfix
**Ask summary:** Ability to set RTL / bidi reading direction for sections, paragraphs, runs (for Arabic).
**Evidence:** `Section.bidi` / `Section.rtl_gutter` (src/docx/section.py:526), Section.text_direction (HISTORY #148), paragraph/run RTL via parfmt / run (HISTORY "Add RTL / bidi on Paragraph and Run (#127)"), `Table.bidi_visual` (src/docx/table.py:488).
**TODO (if applicable):** none.

### upstream#811 — Modification to enable producing consistent binary output
**Verdict:** needs-investigation
**Ask summary:** Request deterministic / reproducible binary output from python-docx so two runs with identical input yield identical `.docx` bytes (linked to upstream PR #810).
**Evidence:** `src/docx/ids.py` provides deterministic stable_id derivation, but no end-to-end reproducible-write option found (no freeze of `dcterms:created`/`modified`, zip mtimes, or rsid generation) in `src/docx/opc/` or `OpcPackage.save`.
**TODO (if applicable):** Add opt-in `Document.save(reproducible=True)` that zeros zip mtimes, clears rsids/`w:rsid*`, and pins core-props timestamps. M

### upstream#790 — How can I add xml tags or set new values in document.xml
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert `<w:pPr><w:bidi/><w:spacing .../></w:pPr>` for every paragraph via the API rather than unzipping.
**Evidence:** `ParagraphFormat` exposes `right_to_left`/`bidi` (src/docx/text/parfmt.py:336) and `space_before`/`space_after`/`line_spacing`. Works for every paragraph via `paragraph_format` proxy.
**TODO (if applicable):** none.

### upstream#771 — Problem extracting table when border is not aligned (duplicate columns)
**Verdict:** resolved-in-loadfix
**Ask summary:** `row.cells` returns a duplicated trailing cell for tables with misaligned `gridSpan`/`gridAfter`, causing column misalignment in extraction.
**Evidence:** Same machinery as #861/#852. HISTORY 1.1.1 note references upstream fix "#531, #1146 Index error on table with misaligned borders". `iter_tc_cells` and `grid_cols_after` handle this (src/docx/table.py).
**TODO (if applicable):** none.

### upstream#770 — _Columns slicing returns single _Column
**Verdict:** new-bug-needed
**Ask summary:** `table.columns[1:]` returns a single `_Column` instead of a list, unlike `_Rows` which correctly returns a list for slice indexing.
**Evidence:** `src/docx/table.py:1507` — `_Columns.__getitem__` delegates blindly to `self._gridCol_lst[idx]`; no slice overload (compare `_Rows.__getitem__` at line 1731 which handles `slice`).
**TODO (if applicable):** Add slice support to `_Columns.__getitem__` returning `list[_Column]`, matching `_Rows`. S

### upstream#747 — How to access the floating shape?
**Verdict:** resolved-in-loadfix
**Ask summary:** Access floating (anchored) shapes in addition to inline shapes.
**Evidence:** `src/docx/text/paragraph.py:442` `add_floating_image`, `src/docx/oxml/shape.py:89` `CT_Anchor`, HISTORY D.17 "Floating images with wp:anchor positioning", commit `f51e7a9`. `FloatingImage` proxy exposed (src/docx/shape.py:140).
**TODO (if applicable):** none.

### upstream#745 — Add shape in document
**Verdict:** resolved-in-loadfix
**Ask summary:** Add arbitrary `<wp:inline>` / drawing shapes to a document.
**Evidence:** `Paragraph.add_shape` (src/docx/text/paragraph.py:396), HISTORY D.27 "DrawingML shapes and text-box content access (#75)", plus inline picture / floating picture APIs. Caption helpers also available.
**TODO (if applicable):** none.

### upstream#744 — Check if a paragraph or table is separated by paging
**Verdict:** new-feature-needed
**Ask summary:** User wants to know, after layout, whether a given paragraph or table is split across pages so they can annotate headings with "- ext" on each continuation page.
**Evidence:** python-docx has no layout engine. `rendered_page_breaks` (src/docx/text/paragraph.py) surfaces Word-cached `w:lastRenderedPageBreak`, but no helper for "is table split across pages". No API in `src/docx/table.py`.
**TODO (if applicable):** Add `Paragraph.page_breaks_inside` / `Table.spans_page_break` readers leveraging `w:lastRenderedPageBreak` markers in cell paragraphs. M

### upstream#743 — Embedding file into docx
**Verdict:** new-feature-needed
**Ask summary:** Programmatically embed one or more files (e.g. text files) into a `.docx`.
**Evidence:** `src/docx/embedded_objects.py` exposes read-only OLE object access ("Creation and modification are intentionally not supported"). No `add_embedded_object` / `add_ole_object` API found.
**TODO (if applicable):** Add writable `Paragraph.add_embedded_object(file, progid, icon=...)` producing an `<o:OLEObject>` and embeddings part. L

### upstream#740 — feature: Run.delete()
**Verdict:** resolved-in-loadfix
**Ask summary:** Provide a `Run.delete()` method so empty runs can be removed.
**Evidence:** `src/docx/text/run.py:186` implements `Run.delete` (versionadded 1.3.0.dev0); HISTORY line 92 "Add Paragraph.delete / Run.delete / Table.delete (#50)".
**TODO (if applicable):** none.

### upstream#738 — How to get the content of "List Paragraph"
**Verdict:** out-of-scope
**Ask summary:** Usage question — user cannot read `.text` of a paragraph styled "List Paragraph".
**Evidence:** `Paragraph.text` already returns all run text regardless of style; this is a user/support question, not a library gap. No code change warranted.
**TODO (if applicable):** none.

### upstream#736 — Stack multiple table
**Verdict:** out-of-scope
**Ask summary:** User wants to minimise paragraphs / cell margins when stacking tables (layout/formatting question).
**Evidence:** `_Cell.margins` (HISTORY #143) and paragraph spacing APIs already expose the necessary controls; no missing feature.
**TODO (if applicable):** none.

### upstream#733 — How to split a merged cell
**Verdict:** new-feature-needed
**Ask summary:** Split (unmerge) a previously-merged table cell back into its component cells.
**Evidence:** `_Cell.merge` exists, and `is_merge_origin` / `merge_origin` are read-only (HISTORY #145). No `_Cell.split` / `unmerge` method in `src/docx/table.py`.
**TODO (if applicable):** Add `_Cell.split()` that clears `gridSpan` / `vMerge` and reinstates the removed `<w:tc>` placeholders. M

### upstream#730 — Create and apply a style for heading
**Verdict:** out-of-scope
**Ask summary:** Support/usage request on how to create a custom style and apply it to all headings.
**Evidence:** `Styles.add_style` and `paragraph.style = ...` already cover this; `docs/user/styles-using.rst` exists. No gap.
**TODO (if applicable):** none.

### upstream#729 — Extract a heading level i.e. 1-9
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose the outline level integer (1..9) for a heading paragraph instead of requiring regex on style name.
**Evidence:** `src/docx/toc.py:58` `_paragraph_heading_level` and `src/docx/accessibility.py:48` `_heading_level` both implement this. These are private helpers; upstream request is for a public accessor.
**TODO (if applicable):** Surface a public `Paragraph.heading_level` property delegating to the existing helper. S

### upstream#727 — Setting the document language has no effect
**Verdict:** new-feature-needed
**Ask summary:** `core_properties.language = "hu-HU"` writes core.xml but Word/LibreOffice use `w:themeFontLang` / per-run `w:lang` instead, so spell-check language is unchanged.
**Evidence:** `src/docx/oxml/settings.py:432` lists `w:themeFontLang` in tag_seq but no accessor in `Settings` (src/docx/settings.py). `Font.language` / `east_asian_language` / `bidi_language` exist per-run (HISTORY #160) but no document-wide setter.
**TODO (if applicable):** Add `Settings.theme_font_language` (val/eastAsia/bidi) that writes `w:themeFontLang`, and optionally a one-shot `Document.set_language()` convenience. M

### upstream#726 — Adding Hyperlinks between pages of same document
**Verdict:** resolved-in-loadfix
**Ask summary:** Support internal (anchor/bookmark) hyperlinks inside a document.
**Evidence:** `Paragraph.add_hyperlink` accepts `anchor=` parameter (src/docx/text/paragraph.py:174, src/docx/oxml/text/paragraph.py:46), bookmarks API exists (src/docx/bookmarks.py), HISTORY D.1 "Hyperlink creation API (#97)" and Phase C bookmarks.
**TODO (if applicable):** none.

### upstream#723 — Add Table of Contents, List of Figures and List of Tables
**Verdict:** needs-investigation
**Ask summary:** Built-in support for Table of Contents, List of Figures, and List of Tables that works cross-platform.
**Evidence:** `Document.add_table_of_contents` (src/docx/document.py:267) and `src/docx/toc.py` cover TOC including cached result preview. Caption helpers exist (HISTORY #141). No `add_list_of_figures` / `add_list_of_tables` (TOC of SEQ fields) helpers.
**TODO (if applicable):** Add `Document.add_list_of_figures` / `add_list_of_tables` builders emitting `TOC \c "Figure"` / `"Table"` field instructions. M

### upstream#722 — Impossible to group figure with its caption
**Verdict:** resolved-in-loadfix
**Ask summary:** User needs floating pictures (so figure+caption can be grouped). Asks for `FloatingShape` / `add_floating_picture`.
**Evidence:** `Paragraph.add_floating_image` (src/docx/text/paragraph.py:442), `FloatingImage` proxy (src/docx/shape.py:140), HISTORY D.17. Caption helpers exist (HISTORY #141).
**TODO (if applicable):** none.

### upstream#721 — How can I use it in uwsgi?
**Verdict:** out-of-scope
**Ask summary:** User writes `.doc` (not `.docx`) via uwsgi and sees garbled output — deployment/usage question.
**Evidence:** python-docx only writes `.docx` (Office Open XML); saving to `.doc` produces a `.docx`-payload file with the wrong extension. Not a library bug.
**TODO (if applicable):** none.

### upstream#717 — New Image Format: WEBP
**Verdict:** new-feature-needed
**Ask summary:** Accept `.webp` images via `add_picture` (currently raises `UnrecognizedImageError`).
**Evidence:** `src/docx/image/` contains bmp/gif/jpeg/png/svg/tiff handlers but no `webp.py`. HISTORY D.22 adds SVG only.
**TODO (if applicable):** Add `WebP(Image)` header parser + registration in `src/docx/image/image.py` `_ImageHeaderFactory`. S

## Batch summary

- resolved-in-loadfix: 12 (#861, #852, #845, #827, #790, #771, #747, #745, #740, #729 partial->S todo, #726, #722)
- new-feature-needed: 8 (#888, #770 bug-like, #744, #743, #733, #727, #717, plus #729 public surface)
- new-bug-needed: 1 (#770)
- needs-investigation: 2 (#811, #723)
- out-of-scope: 4 (#738, #736, #730, #721)

Note: #729 counted as resolved (private helper exists) with optional S TODO to expose it publicly; #770 counted as new-bug-needed only.

### upstream#667 — hidden inline images
**Verdict:** needs-investigation
**Ask summary:** User reports inline picture inserted via `run.add_picture` is present in media/ but not rendered; asks for doc/style guidance.
**Evidence:** no match — no dedicated "hidden image" guard or docs located; image insertion code unchanged semantics.
**TODO (if applicable):** Reproduce hidden-image scenario; document common causes (anchor vs inline, bookmark run placement). S

### upstream#668 — Replacing a header in one docx with the header from another docx
**Verdict:** new-feature-needed
**Ask summary:** Request ability to assign/copy a header from another document to a section.
**Evidence:** no match — no copy-header API; `src/docx/parts/hdrftr.py` exists but no cross-document clone.
**TODO (if applicable):** Add `Section.copy_header_from(other_section)` and footer twin. M

### upstream#670 — Repeated output of same table cell on iter_block_items
**Verdict:** needs-investigation
**Ask summary:** User-side `iter_block_items` generator re-yields cell contents for nested tables.
**Evidence:** `src/docx/oxml/table.py:902 iter_block_items` exists; upstream issue is about user-authored recursive helper, not a fork bug.
**TODO (if applicable):** Add doc snippet / recipe for proper nested-table iteration. S

### upstream#672 — v0.8.10 header/footer not exposed on new docs
**Verdict:** resolved-in-loadfix
**Ask summary:** Claim that `sections` returns None and headers/footers unusable on new documents.
**Evidence:** Upstream 0.8.8 added headers/footers; loadfix inherits 1.2.0+ baseline with full `Section.header/footer`; HISTORY.rst 0.8.8.
**TODO (if applicable):** n/a

### upstream#674 — How to center the picture?
**Verdict:** out-of-scope
**Ask summary:** Usage question about centering an inline picture.
**Evidence:** Standard `paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER` already supported.
**TODO (if applicable):** n/a

### upstream#676 — Captions that behave like captions
**Verdict:** resolved-in-loadfix
**Ask summary:** Request real caption behavior: SEQ numbering, bound to figure, cross-referenceable.
**Evidence:** `src/docx/captions.py` + HISTORY "Add caption helpers (#141)"; Document.add_caption plus Paragraph.add_caption_before/after.
**TODO (if applicable):** n/a

### upstream#677 — Getting the heading number of a Word section heading
**Verdict:** resolved-in-loadfix
**Ask summary:** Request to retrieve the numbering text ("1.2.3") for a heading paragraph.
**Evidence:** `src/docx/numbering.py:309 number_format` + `number`/`text` on numbering level; Phase D.9 (#22) in HISTORY.
**TODO (if applicable):** n/a

### upstream#678 — fix: accommodate NULL relationship (by skipping)
**Verdict:** needs-investigation
**Ask summary:** Open-document KeyError on relationship to a "NULL" target; shortlist label indicates proposed tolerant behavior.
**Evidence:** no match in src/docx/opc for "NULL relationship" skip; recover=True (#151) is generic but not targeted.
**TODO (if applicable):** Skip/log relationships whose target is NULL/missing in package loader. M

### upstream#680 — How to change settings.xml (autoHyphenation)
**Verdict:** new-feature-needed
**Ask summary:** Request property to set `w:autoHyphenation` in settings.
**Evidence:** `src/docx/oxml/settings.py:386` lists the tag in successors list only; no proxy getter/setter on Settings.
**TODO (if applicable):** Add `Settings.auto_hyphenation` bool property (plus related hyphenation flags). S

### upstream#682 — Text with color or underlined cannot be read
**Verdict:** out-of-scope
**Ask summary:** User reports paragraph.text omits colored/underlined runs.
**Evidence:** no match — `Paragraph.text` concatenates all run text regardless of formatting; likely user-code issue.
**TODO (if applicable):** n/a

### upstream#684 — install error: doesn't exist or not a regular file
**Verdict:** resolved-in-loadfix
**Ask summary:** 0.8.10 source install fails copying default-docx-template.
**Evidence:** HISTORY 0.8.10 and 0.8.9 notes cover the build fix; loadfix now uses pyproject.toml + src/ layout with wheel packaging of templates.
**TODO (if applicable):** n/a

### upstream#689 — Table cells with form fields parsed incorrectly
**Verdict:** needs-investigation
**Ask summary:** Cells containing legacy form fields yield empty text on iteration.
**Evidence:** `src/docx/form_fields.py` adds read/write legacy FF; but Cell.text / Paragraph.text still concatenates runs — need to confirm ff result text is included.
**TODO (if applicable):** Verify Cell.text includes form-field result text; add regression test. S

### upstream#693 — Unable to read Strict Open XML Docx file
**Verdict:** new-feature-needed
**Ask summary:** Request loader support for Strict Open XML (ISO/IEC 29500 Strict) files.
**Evidence:** no match — no namespace remapping or Strict→Transitional transform.
**TODO (if applicable):** Add Strict→Transitional namespace rewrite on open. L

### upstream#694 — Can't set table style 'Table Grid'
**Verdict:** needs-investigation
**Ask summary:** Assigning `table.style = 'Table Grid'` raises KeyError when style is latent in styles.xml.
**Evidence:** no match for automatic latent-style materialization when assigning to table.style.
**TODO (if applicable):** On style assignment, promote latent style to actual style entry if needed. M

### upstream#696 — Detect dropdown and get the values
**Verdict:** resolved-in-loadfix
**Ask summary:** Read dropdown form-field values from a docx.
**Evidence:** `src/docx/form_fields.py:50 DROPDOWN`, `.dropdown` view exposes entries; HISTORY legacy form fields (#123).
**TODO (if applicable):** n/a

### upstream#698 — Can I update my docx contents by code?
**Verdict:** out-of-scope
**Ask summary:** General usage question about editing docx content with python.
**Evidence:** n/a
**TODO (if applicable):** n/a

### upstream#700 — feature-request: ability to set text language
**Verdict:** resolved-in-loadfix
**Ask summary:** Allow setting language on runs / paragraphs for spell check.
**Evidence:** `src/docx/text/font.py:355 language`, `:390 east_asian_language`, `:425 bidi_language`; HISTORY Font.language (#160).
**TODO (if applicable):** n/a

### upstream#701 — Question about style transfer between documents
**Verdict:** new-feature-needed
**Ask summary:** How to copy styles between two documents and list only used styles.
**Evidence:** no match for `copy_style`/`copy_styles`; no "used styles" filter helper.
**TODO (if applicable):** Add `Styles.copy_from(other_styles)` and `Styles.in_use` helper. M

### upstream#702 — add_table set row background / word color / outline color
**Verdict:** resolved-in-loadfix
**Ask summary:** Need API for cell/row background, font color, cell border color.
**Evidence:** Phase D.6 cell shading (#63) and Table.borders / _Cell.borders (#102) in HISTORY; `Font.color` existed.
**TODO (if applicable):** n/a

### upstream#703 — Pandas DataFrame to Word Table takes a lot of time
**Verdict:** needs-investigation
**Ask summary:** Performance issue: large table insertion very slow.
**Evidence:** no dedicated bulk-insert path found; `add_table` creates rows/cells via lxml insert.
**TODO (if applicable):** Add bulk/batch table creation API or doc recipe using lxml direct XML. M

### upstream#707 — How can I set paragraph_format.element.xml?
**Verdict:** out-of-scope
**Ask summary:** User wants to extract `paragraph_format.element.xml` and apply to another doc.
**Evidence:** Possible via internal XML manipulation; not a public-API request.
**TODO (if applicable):** n/a

### upstream#709 — How can I insert xml with style for another docx file?
**Verdict:** needs-investigation
**Ask summary:** Cross-document paragraph move; related to style/numbering reference transfer.
**Evidence:** Overlaps with #701 (style copy) and #668 (header copy); no generic cross-doc block clone API.
**TODO (if applicable):** Investigate `Document.import_block(other_block)` that rewires style/numbering refs. L

### upstream#713 — How to add attachment (embedded xlsx/docx)
**Verdict:** needs-investigation
**Ask summary:** Request to embed another office file as attachment (OLE package).
**Evidence:** `src/docx/embedded_objects.py` is read-only per HISTORY ("read-only embedded OLE objects (#140)").
**TODO (if applicable):** Add `add_embedded_object(path)` write API for OLE packages. L

### upstream#714 — Read the endnotes.xml
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose endnotes content on Document.
**Evidence:** Phase A — `Document.endnotes` (#17, #96); `src/docx/parts/endnotes.py`, `src/docx/endnotes.py`.
**TODO (if applicable):** n/a

### upstream#715 — How to read charts in docx
**Verdict:** resolved-in-loadfix
**Ask summary:** Read chart content from docx.
**Evidence:** `src/docx/chart.py` Chart/ChartSeries; `src/docx/parts/chart.py`; HISTORY "Charts read + add_chart() (#111)".
**TODO (if applicable):** n/a

## Batch summary
- resolved-in-loadfix: 9 (#672, #676, #677, #684, #696, #700, #702, #714, #715)
- new-feature-needed: 4 (#668, #680, #693, #701)
- needs-investigation: 8 (#667, #670, #678, #689, #694, #703, #709, #713)
- out-of-scope: 4 (#674, #682, #698, #707)
- new-bug-needed: 0
- Total: 25

### upstream#665 — AttributeError: 'ColorFormat' object has no attribute 'brightness' when using a theme color
**Verdict:** new-feature-needed
**Ask summary:** Documentation advertises `ColorFormat.brightness` (theme-color tint/shade adjust) but it's not implemented, raising AttributeError.
**Evidence:** src/docx/dml/color.py defines only `rgb`, `theme_color`, `type`, `_color`; no `brightness` property. grep "brightness" src/docx/dml/ returns no matches.
**TODO (if applicable):** Add ColorFormat.brightness getter/setter writing w:rPr/w:color/@w14:luminance-style shade/tint lumMod/lumOff — S.

### upstream#663 — How to delete table
**Verdict:** resolved-in-loadfix
**Ask summary:** User wants to delete a table in a document.
**Evidence:** src/docx/table.py:62 `Table.delete()` exists; HISTORY.rst lists "Add Paragraph.delete / Run.delete / Table.delete (#50)".

### upstream#662 — Add text on image
**Verdict:** out-of-scope
**Ask summary:** Overlaying text on top of an image (textbox over picture / z-order).
**Evidence:** Requires DrawingML anchor stacking & text-box positioning beyond python-docx's remit; no related module. HISTORY mentions read-only text-box content only (D.27).

### upstream#661 — How to clear the document background?
**Verdict:** resolved-in-loadfix
**Ask summary:** Remove `<w:background>` page-color element.
**Evidence:** src/docx/document.py:310-328 exposes `Document.background_color` getter/setter; setting to None removes it. HISTORY: "Add Document.background_color (#118)".

### upstream#657 — How to install python-docx at Ubuntu?
**Verdict:** out-of-scope
**Ask summary:** Support question about installing the package on Ubuntu without pip.
**Evidence:** No code issue; packaging/install guidance outside the fork's OOXML scope.

### upstream#651 — Add svg picture
**Verdict:** resolved-in-loadfix
**Ask summary:** Support inserting SVG images via `add_picture`.
**Evidence:** src/docx/image/svg.py exists; image.py:183-187 detects SVG streams. HISTORY: "D.22 SVG image support (#76)".

### upstream#650 — How to use Python iteration to read paragraphs, tables and pictures in word?
**Verdict:** resolved-in-loadfix
**Ask summary:** Iterate Document content in order including inline images.
**Evidence:** `BlockItemContainer.iter_inner_content` (blkcntnr.py:77), `Run.iter_inner_content` yields Drawing (run.py:233). HISTORY lists iter_inner_content additions.

### upstream#647 — Is it possible to do a mailto:email@domain.com
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert a mailto hyperlink.
**Evidence:** src/docx/text/paragraph.py:162 `add_hyperlink(url=...)` uses RT.HYPERLINK external rel; any URL scheme (incl. mailto:) works. HISTORY: "D.1 Hyperlink creation API (#97)".

### upstream#645 — I want to know how to insert pictures in the table or cell
**Verdict:** resolved-in-loadfix
**Ask summary:** Add picture into a table cell.
**Evidence:** `_Cell` inherits BlockItemContainer; use `cell.paragraphs[0].add_run().add_picture()` — supported since upstream. No fork change needed.

### upstream#644 — Looking to help
**Verdict:** out-of-scope
**Ask summary:** Contributor offering help, not a bug/feature.
**Evidence:** No code ask.

### upstream#639 — Number List rendering
**Verdict:** needs-investigation
**Ask summary:** Numbered list numbers continue incorrectly across paragraphs when rendered.
**Evidence:** No attached docx; ambiguous. Fork has Paragraph.restart_numbering and Numbering.add_style (HISTORY D.9) which addresses restart scenarios. Need user repro.

### upstream#626 — How to identify section/table index in footer?
**Verdict:** resolved-in-loadfix
**Ask summary:** Access tables inside footer; `footer.tables` sometimes empty.
**Evidence:** `_BaseHeaderFooter(BlockItemContainer)` in section.py:1355 inherits `tables` (blkcntnr.py:93). Primary headers/footers and inheritance are handled.

### upstream#625 — Handling poorly formed XML
**Verdict:** resolved-in-loadfix
**Ask summary:** Load docx where `<w:br/>` appears inside a single `<w:t>`, which truncates text.
**Evidence:** HISTORY: "Add recover=True mode for malformed .docx (#151)"; src/docx/api.py:20 adds `recover` flag using lxml recovering parser. Covers broad malformed-XML cases.

### upstream#623 — Extracting custom fields as text, from Word table
**Verdict:** resolved-in-loadfix
**Ask summary:** Read evaluated/cached text of custom fields (incl. linked Excel) inside tables.
**Evidence:** src/docx/fields.py exposes Field with cached result via fldChar begin/separate/end; HISTORY Phase C: "Add simple and complex field codes (#10)".

### upstream#621 — Custom namespaces for XPath
**Verdict:** resolved-in-loadfix
**Ask summary:** `element.xpath(expr, namespaces=...)` broken after nsmap override; can't access w14 elements.
**Evidence:** src/docx/oxml/xmlchemy.py:688 override accepts `**kwargs` and merges; ns.py:27 includes w14. User-supplied namespaces pass through via kwargs.

### upstream#617 — Identifying Indents Not Working
**Verdict:** resolved-in-loadfix
**Ask summary:** `paragraph.style.paragraph_format.left_indent` returns None for numbered-list paragraphs.
**Evidence:** Fork adds Paragraph.numbering_format returning Level with `indent` (numbering.py:340), exposing list-level indent separate from style.

### upstream#616 — i wanna save my docx in specific directory
**Verdict:** out-of-scope
**Ask summary:** User confusion about `document.save(path)` — standard API already accepts a path.
**Evidence:** No code issue.

### upstream#615 — how to add a inlineshape object as image directly?
**Verdict:** out-of-scope
**Ask summary:** Empty body; unclear request to add an InlineShape object directly.
**Evidence:** No actionable ask.

### upstream#614 — Reading list numbers along with text
**Verdict:** new-feature-needed
**Ask summary:** Read rendered list numbers ("1.", "2.1") alongside paragraph text.
**Evidence:** Fork exposes `Paragraph.numbering_format.number_format` + `.text` (pattern like "%1.") but does not compute the effective rendered label per-paragraph (counter resolution).
**TODO (if applicable):** Add Paragraph.list_number / rendered_number that walks numbering instance state to produce "1.2" strings — M.

### upstream#612 — How to add_table with table object parameter
**Verdict:** new-feature-needed
**Ask summary:** Clone/append an existing Table (from another doc) into a new Document.
**Evidence:** No `copy_table` / `Document.add_table(table=...)` in table.py or document.py. Cross-part image/numbering rewiring not implemented.
**TODO (if applicable):** Add Document.add_table_copy(table) copying w:tbl XML plus rewiring rIds (images, styles, numbering) — L.

### upstream#610 — Hyperlinks
**Verdict:** new-feature-needed
**Ask summary:** Wrap an existing substring inside an existing paragraph with a hyperlink without restructuring runs.
**Evidence:** add_hyperlink only appends to end of paragraph (paragraph.py:162). No `Run.wrap_hyperlink` / inline insertion. Related to search/replace; D.10 replaces formatting but not hyperlink creation at arbitrary position.
**TODO (if applicable):** Add Run.make_hyperlink / Paragraph.insert_hyperlink_at(run, url) that splits runs around anchor text — M.

### upstream#609 — Need support for python 2.6
**Verdict:** out-of-scope
**Ask summary:** Restore Python 2.6 compatibility.
**Evidence:** HISTORY 1.0.0: "Remove Python 2 support. Supported versions are 3.7+"; fork targets 3.9+.

### upstream#607 — Open Document in Google Docs, unable to convert Document
**Verdict:** needs-investigation
**Ask summary:** Generated .docx fails Google Docs conversion.
**Evidence:** No repro file; root cause unclear (likely missing/invalid styles or sectPr). May be mitigated by existing HISTORY items (DocumentProtection D.3, theme D.22). Need sample file.

### upstream#606 — About the new version 0.8.9
**Verdict:** needs-investigation
**Ask summary:** `document.part.related_parts[rId]._blob` returns None after headers/footers change.
**Evidence:** Internal API change from 0.8.9; fork adds headers-footer and many parts but didn't document `_blob` access pattern. Private attr; need confirmation whether it affects public API.

### upstream#605 — Extracting information from docx file(text including inline images)
**Verdict:** resolved-in-loadfix
**Ask summary:** Extract text plus inline images from docx in order.
**Evidence:** `Run.iter_inner_content` yields str | Drawing | RenderedPageBreak (text/run.py:233); paragraphs expose images via inline_shapes / drawings.

## Batch summary
- resolved-in-loadfix: 12 (#663, #661, #651, #650, #647, #645, #626, #625, #623, #621, #617, #605)
- new-feature-needed: 4 (#665, #614, #612, #610)
- needs-investigation: 3 (#639, #607, #606)
- out-of-scope: 6 (#662, #657, #644, #616, #615, #609)
- new-bug-needed: 0
- Total: 25

### upstream#548 — docx.oxml.exceptions.InvalidXmlError: required `<w:tblGrid>` child element not present
**Verdict:** needs-investigation
**Ask summary:** Malformed table (missing `w:tblGrid`) raises InvalidXmlError when opening a .docx.
**Evidence:** Recover mode exists (`src/docx/opc/package.py:130`) but it targets XML parse errors, not missing required children. No targeted fix.
**TODO (if applicable):** Harden CT_Tbl to synthesize missing `w:tblGrid` under recover=True (S).

### upstream#549 — 3d mapping of column chart column color based on Range for Word Tables
**Verdict:** out-of-scope
**Ask summary:** User wants to use matplotlib 3D axes with data from Word tables; needs Excel/plot help.
**Evidence:** Not a python-docx concern; no match in fork.
**TODO (if applicable):** None.

### upstream#550 — 3d column graph for a word table coloring the column based on range or cell bg color
**Verdict:** out-of-scope
**Ask summary:** Duplicate of #549 — matplotlib 3D rendering using Word table cell bg color.
**Evidence:** Not a python-docx concern; no match.
**TODO (if applicable):** None.

### upstream#554 — Fetch elements numbering values
**Verdict:** new-feature-needed
**Ask summary:** Obtain the computed list/heading number text (e.g. "1." for "1. Summary"), not just `p.text`.
**Evidence:** Fork exposes raw `numId`/`ilvl` (`src/docx/numbering.py`, `text/paragraph.py:942`) but no resolved-number renderer.
**TODO (if applicable):** Add `Paragraph.list_number_text` that walks numbering.xml to compute the displayed number (M).

### upstream#558 — Appending one document to another
**Verdict:** new-feature-needed
**Ask summary:** Merge/append one .docx body into another, preserving styles/watermarks.
**Evidence:** No `append_document` / `merge` API in `src/docx/document.py`.
**TODO (if applicable):** Add `Document.append_document(other)` (body, styles, numbering, images, relationships) (L).

### upstream#560 — Extract Chart Object From Doc
**Verdict:** resolved-in-loadfix
**Ask summary:** Access chart objects embedded in a document.
**Evidence:** `Document.charts` property and `Chart` class (`src/docx/chart.py`, `document.py:346`); HISTORY "Charts read + add_chart() (#111)".
**TODO (if applicable):** None.

### upstream#563 — add_table() adds a table even when raising exception
**Verdict:** new-bug-needed
**Ask summary:** Failed `add_table()` (missing style) leaves a stray table element in the document.
**Evidence:** `Document.add_table` in `src/docx/document.py` builds then styles; no transactional rollback visible.
**TODO (if applicable):** Make `add_table` validate style before inserting `w:tbl`, or remove on failure (S).

### upstream#567 — How I can insert a chart to document in python-docx?
**Verdict:** resolved-in-loadfix
**Ask summary:** Add charts with X/Y data to a document.
**Evidence:** `Document.add_chart()` at `src/docx/document.py:186`; HISTORY "Charts read + add_chart() (#111)".
**TODO (if applicable):** None.

### upstream#569 — left_indent not getting populated when reading existing document
**Verdict:** needs-investigation
**Ask summary:** `paragraph_format.left_indent` returns None for visibly indented paragraphs (style-inherited).
**Evidence:** `left_indent` in `src/docx/text/parfmt.py:217` reads local `w:ind` only; no style-resolution helper. Upstream design, may need a resolved-format accessor.
**TODO (if applicable):** Add `paragraph_format.resolved_left_indent` (walks style chain) (M).

### upstream#570 — Error with paragraph.style = 'List Bullet'!
**Verdict:** needs-investigation
**Ask summary:** Setting style `'List Bullet'` in a cell of a loaded document raises KeyError (style missing from that doc).
**Evidence:** Style resolution raises KeyError when style absent; no auto-add for built-in styles in fork.
**TODO (if applicable):** Auto-load built-in style from Word defaults when first referenced (M).

### upstream#572 — How to edit/read docx's some properties such as 'Company','Manager'?
**Verdict:** new-feature-needed
**Ask summary:** Read/write extended (app.xml) properties like Company, Manager.
**Evidence:** Fork has `CoreProperties` and `CustomProperties` but no ExtendedProperties wrapper; app.xml only in template rels.
**TODO (if applicable):** Add `Document.extended_properties` exposing Company/Manager/Application etc. (M).

### upstream#573 — old settings.xml
**Verdict:** needs-investigation
**Ask summary:** Merging docx via python-docx writes a minimal settings.xml that breaks text-box (mc:AlternateContent) offsets.
**Evidence:** Default template settings.xml exists; no textbox-offset-specific settings preservation logic.
**TODO (if applicable):** Preserve original settings.xml (incl. compatSettings) on round-trip or when merging (M).

### upstream#578 — I want to know how to insert pictures in the form of word.
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert a picture into a Word table cell.
**Evidence:** `_Cell.add_paragraph().add_run().add_picture()` via `Run.add_picture` (`src/docx/text/run.py:62`).
**TODO (if applicable):** None (doc-only; consider adding cell-picture example).

### upstream#583 — How to get table next/before a paragraph
**Verdict:** new-feature-needed
**Ask summary:** Navigate from a paragraph/table to adjacent block items (next/previous sibling).
**Evidence:** No `next_sibling`/`previous_sibling` helper on Paragraph/Table in `src/docx/text/paragraph.py` or `table.py`.
**TODO (if applicable):** Add `.next_block` / `.previous_block` on block items (S).

### upstream#584 — How to determine table head alignment while reading docx file
**Verdict:** resolved-in-loadfix
**Ask summary:** Detect whether header cells are top/left aligned when reading table data.
**Evidence:** `_Cell` exposes vertical alignment via `tcPr.vAlign_val` (`src/docx/table.py:866`); horizontal via paragraph alignment.
**TODO (if applicable):** None.

### upstream#585 — how to change some words style like font size and font style like cambria or time news man?
**Verdict:** out-of-scope
**Ask summary:** Usage question (body is empty); wants to change run font.
**Evidence:** `Run.font.name/size` already exist upstream.
**TODO (if applicable):** None.

### upstream#586 — Creating table after certain Inches in the page.
**Verdict:** needs-investigation
**Ask summary:** Position a table N inches from the left page margin (left indent on table).
**Evidence:** Fork has autofit/column width (HISTORY D.26) but table-level left indent (`w:tblInd`) not obviously exposed.
**TODO (if applicable):** Add `Table.left_indent` mapping `w:tblInd` (S).

### upstream#588 — How to parse the position information of characters in docx?
**Verdict:** out-of-scope
**Ask summary:** Wants absolute page/x-y position of individual characters.
**Evidence:** Word does not store character positions; requires rendering engine. Out of scope.
**TODO (if applicable):** None.

### upstream#590 — How to return the heading number
**Verdict:** new-feature-needed
**Ask summary:** Return the computed list-numbering value (e.g. "1.2.3") for a heading paragraph.
**Evidence:** Same gap as #554 — raw numPr exposed but no resolver in `src/docx/numbering.py`.
**TODO (if applicable):** Covered by #554 task `Paragraph.list_number_text` (M).

### upstream#591 — How to open a document with password?
**Verdict:** out-of-scope
**Ask summary:** Open/save encrypted .docx with a password.
**Evidence:** No crypto in fork; requires `msoffcrypto-tool` or similar; classic out-of-scope.
**TODO (if applicable):** None.

### upstream#594 — can't copy 'docx/templates/default-docx-template' on pip install python-docx
**Verdict:** resolved-in-loadfix
**Ask summary:** `default-docx-template` not installed by pip; PackageNotFoundError at import.
**Evidence:** `src/docx/templates/default-docx-template/` and `default.docx` ship in src-layout.
**TODO (if applicable):** None.

### upstream#596 — How to read unicode characters
**Verdict:** out-of-scope
**Ask summary:** Determine encoding + transliterate to ASCII (Unidecode).
**Evidence:** Word text is always Unicode per OOXML; no encoding detection needed; user-level concern.
**TODO (if applicable):** None.

### upstream#600 — How to get automatic caption numbering?
**Verdict:** resolved-in-loadfix
**Ask summary:** Generate sequential caption numbers like "Figure 1".
**Evidence:** `src/docx/captions.py` and `Document.add_caption`; HISTORY "Add caption helpers (#141)".
**TODO (if applicable):** None.

### upstream#602 — Protection classification
**Verdict:** needs-investigation
**Ask summary:** Set a sensitivity classification (Public/Internal/Confidential/Restricted) on a document.
**Evidence:** Fork supports DocumentProtection (`src/docx/settings.py:166`) but not MS Information Protection sensitivity labels (stored in CustomXML `MSIP_Label_*`).
**TODO (if applicable):** Add helper `Document.sensitivity_label` via custom_xml (M).

### upstream#604 — Better handle missing items when extracting (with Window Subsystem for Linux)
**Verdict:** needs-investigation
**Ask summary:** Opening certain .docx with images on WSL raises KeyError; asks for better error handling.
**Evidence:** `recover=True` mode (`src/docx/opc/package.py:130`) addresses malformed XML; unclear if it catches missing-part KeyErrors.
**TODO (if applicable):** Extend recover mode to warn-and-skip on missing part references instead of KeyError (S).

## Batch summary
- resolved-in-loadfix: 6 (#560, #567, #578, #584, #594, #600)
- new-feature-needed: 5 (#554, #558, #572, #583, #590)
- new-bug-needed: 1 (#563)
- needs-investigation: 7 (#548, #569, #570, #573, #586, #602, #604)
- out-of-scope: 6 (#549, #550, #585, #588, #591, #596)
- Total: 25

### upstream#545 — Mimetype detected by libmagic is inaccurate
**Verdict:** needs-investigation
**Ask summary:** libmagic reports `application/octet-stream` for files created by python-docx rather than the OOXML MIME type.
**Evidence:** no match; issue is about zip magic ordering in OPC packaging, not a loadfix phase.
**TODO (if applicable):** Investigate writing `[Content_Types].xml` / `mimetype` first in the zip and ensure a uncompressed mimetype stream (S).

### upstream#543 — how to combine all other document-files to a document?
**Verdict:** new-feature-needed
**Ask summary:** User wants to merge/append another .docx file into the current document.
**Evidence:** no match; `grep -rn "merge_documents\|append_document"` empty.
**TODO (if applicable):** Add `Document.append_document(other)` that imports body, styles, numbering, images (L).

### upstream#542 — Table of Contents header
**Verdict:** needs-investigation
**Ask summary:** User reports `WD_BUILTIN_STYLE.INDEX_HEADING` raises on `add_paragraph(style=...)` when used above a TOC.
**Evidence:** fork has `toc.py` and `add_table_of_contents` (#116); not clear if TOC Heading built-in style is handled.
**TODO (if applicable):** Verify `WD_BUILTIN_STYLE.TOC_HEADING` maps correctly and document usage with `add_table_of_contents` (S).

### upstream#540 — Problem with rotation when adding images
**Verdict:** needs-investigation
**Ask summary:** Portrait images auto-rotate to landscape when added via `add_picture`.
**Evidence:** no match; likely EXIF orientation not honored in `src/docx/image/`.
**TODO (if applicable):** Honor EXIF orientation tag when computing image dimensions / rendering (M).

### upstream#538 — would it support textBox later?
**Verdict:** resolved-in-loadfix
**Ask summary:** User asks whether text-box support is planned.
**Evidence:** `src/docx/drawing/__init__.py:226` reads `wps:txbx/w:txbxContent`; HISTORY D.27 "DrawingML shapes and text-box content access (#75)".
**TODO (if applicable):** n/a (read-only; creation may be separate issue).

### upstream#536 — Defining bullets
**Verdict:** resolved-in-loadfix
**Ask summary:** How to define a new bullet/numbering list.
**Evidence:** `src/docx/numbering.py` with `add_numbering_definition` / `add_abstractNum`; HISTORY D.9 (#22).
**TODO (if applicable):** n/a.

### upstream#532 — Is it possible to Add a cell to a particular row?
**Verdict:** new-feature-needed
**Ask summary:** Expose API for adding a cell to an existing row.
**Evidence:** no match for `Row.add_cell` / `insert_cell` in `table.py`.
**TODO (if applicable):** Add `Row.add_cell()` / `Row.insert_cell(index)` that inserts a `w:tc` element (S).

### upstream#526 — LGPL License ?
**Verdict:** out-of-scope
**Ask summary:** Question about licensing (project is MIT).
**Evidence:** no code change.
**TODO (if applicable):** n/a.

### upstream#524 — Writing to a Textbox.
**Verdict:** new-feature-needed
**Ask summary:** Create and write formatted text into a text box, control fill/border/position.
**Evidence:** fork only reads text-box content (`drawing/__init__.py:226`); no creation API.
**TODO (if applicable):** Add `Document.add_text_box()` / `Run.add_text_box()` creating `wps:wsp/wps:txbx` with formatting (L).

### upstream#520 — python-docx installed but import fails
**Verdict:** out-of-scope
**Ask summary:** User's pip installed into Python 2.7 site-packages instead of 3.7.
**Evidence:** user-environment issue.
**TODO (if applicable):** n/a.

### upstream#519 — Splitting a run
**Verdict:** resolved-in-loadfix
**Ask summary:** Need utilities to split a run to apply formatting to a substring.
**Evidence:** `src/docx/text/run.py:269` `split_run(offset)`; `search_regex`/`replace_regex` in `search.py` (D.10 #91, #153, #154).
**TODO (if applicable):** n/a.

### upstream#518 — Remove series of images and replace with special characters
**Verdict:** new-feature-needed
**Ask summary:** Delete images in a document and insert placeholder characters in their place.
**Evidence:** no explicit image-removal helper; `Run.clear()` exists but no `remove_drawing` helper.
**TODO (if applicable):** Add `Run.remove_drawings()` / `Paragraph.remove_inline_images()` helpers (S).

### upstream#517 — Cannot add shape to a document
**Verdict:** new-feature-needed
**Ask summary:** Add a rectangle shape with border and inner text.
**Evidence:** fork reads DrawingML shapes (D.27) but has no shape-creation API.
**TODO (if applicable):** Add `Document.add_shape(type, width, height, ...)` creating `wps:wsp` (L).

### upstream#515 — add_picture ZeroDivisionError: division by zero
**Verdict:** new-bug-needed
**Ask summary:** JPEG with zero EXIF `px_per_inch` leads to division-by-zero in `add_picture`.
**Evidence:** `src/docx/image/` parses DPI; divide occurs when computing EMU dimensions.
**TODO (if applicable):** Clamp horz/vert dpi to default (72) when parsed value is 0 (S).

### upstream#514 — IndexError when adding table (no sectPr)
**Verdict:** new-bug-needed
**Ask summary:** `add_table` crashes on documents whose body has no `w:sectPr`.
**Evidence:** `document.py:_block_width` still does `self.sections[-1]`; no fallback in loadfix.
**TODO (if applicable):** Fall back to `EMU(0)` or page-size default when `sectPr_lst` is empty (S).

### upstream#513 — No option to set cell margins of table
**Verdict:** resolved-in-loadfix
**Ask summary:** Control per-cell margins (`w:tcMar`).
**Evidence:** `src/docx/table.py:784` `CellMargins` proxy with `w:tcMar` getter/setter; not explicitly in HISTORY but implemented.
**TODO (if applicable):** n/a.

### upstream#510 — Problem with RTL property
**Verdict:** needs-investigation
**Ask summary:** Setting `font.rtl = True` stops other run properties (size) from taking effect.
**Evidence:** `font.py:660` rtl getter/setter; unclear if ordering of rPr children is still XSD-correct.
**TODO (if applicable):** Verify `w:rtl` successors tuple places it after `w:sz` etc. per CT_RPr XSD (S).

### upstream#508 — copy styles between Word documents — add_style(Style)?
**Verdict:** new-feature-needed
**Ask summary:** Copy whole style objects from one document to another.
**Evidence:** `styles.py:add_style(name, style_type, builtin)` takes name only; no deep-copy importer.
**TODO (if applicable):** Add `Styles.import_style(style_from_other_doc)` deep-copying XML and linked styles (M).

### upstream#507 — can we do GetCrossReferenceItems?
**Verdict:** resolved-in-loadfix
**Ask summary:** Enumerate cross-reference targets (numbered items) and REF fields.
**Evidence:** `fields.py` REF/PAGEREF resolution (Phase C, #115); bookmarks (#52).
**TODO (if applicable):** n/a.

### upstream#506 — How to apply text borders to text runs?
**Verdict:** resolved-in-loadfix
**Ask summary:** Need to set `w:rPr/w:bdr` on runs.
**Evidence:** `text/font.py:73-` border_color / border_style / border_size (D.20 ecosystem; HISTORY "Add Font.border_* properties (#120)").
**TODO (if applicable):** n/a.

### upstream#505 — error for add_picture with .svg extension
**Verdict:** resolved-in-loadfix
**Ask summary:** Support adding SVG images.
**Evidence:** `src/docx/image/svg.py`; HISTORY D.22 "SVG image support (#76)".
**TODO (if applicable):** n/a.

### upstream#504 — KeyError "no style with name 'Table Grid'"
**Verdict:** out-of-scope
**Ask summary:** Default template lacks 'Table Grid' style on certain Word installs.
**Evidence:** template concern; depends on `default.docx` template ship.
**TODO (if applicable):** n/a (or: ship a default template containing 'Table Grid', S).

### upstream#503 — Add/Modify Attached template in /word/_rels
**Verdict:** needs-investigation
**Ask summary:** Read/write `w:attachedTemplate` and its `settings.xml.rels` relationship.
**Evidence:** `src/docx/oxml/settings.py:371` lists `w:attachedTemplate` as a child element; no proxy accessor.
**TODO (if applicable):** Add `Settings.attached_template` property + relationship management (M).

### upstream#501 — Support for TinyMCE (HTML to docx)
**Verdict:** out-of-scope
**Ask summary:** HTML-to-docx conversion.
**Evidence:** not in scope for python-docx or loadfix.
**TODO (if applicable):** n/a.

### upstream#500 — Launch Microsoft Word on Mac
**Verdict:** out-of-scope
**Ask summary:** Python API to drive Word.app on macOS.
**Evidence:** external automation, not OOXML.
**TODO (if applicable):** n/a.

## Batch summary
- resolved-in-loadfix: 7 (#538, #536, #519, #513, #507, #506, #505)
- new-feature-needed: 6 (#543, #532, #524, #518, #517, #508)
- new-bug-needed: 2 (#515, #514)
- needs-investigation: 5 (#545, #542, #540, #510, #503)
- out-of-scope: 5 (#526, #520, #504, #501, #500)
- total: 25

### upstream#498 — how to add page number on footer?
**Verdict:** resolved-in-loadfix
**Ask summary:** User-support question about inserting a PAGE field in the footer paragraph.
**Evidence:** `src/docx/fields.py:38` example uses `paragraph.add_simple_field("PAGE", "1")`; `src/docx/text/paragraph.py:251` defines `add_simple_field`.

### upstream#497 — Extracted String Missing a Word
**Verdict:** resolved-in-loadfix
**Ask summary:** `.text` dropped a word stored inside a `<w:fldSimple>` result (a state code).
**Evidence:** `src/docx/oxml/text/paragraph.py:273` includes `w:fldSimple` in paragraph text; `src/docx/oxml/fields.py:64` implements `CT_FldSimple.text`.

### upstream#496 — access the font info for default style ('normal')
**Verdict:** needs-investigation
**Ask summary:** Reading font info for the default "Normal" style returns None when attributes come from `w:docDefaults`/`rPrDefault`.
**Evidence:** `src/docx/oxml/styles.py:339` references `w:docDefaults` but no proxy fallback for rPrDefault when style rPr is empty.
**TODO:** Expose `Styles.default_rpr` / fall back to `docDefaults/rPrDefault` when style font attrs are None — S.

### upstream#494 — Cannot access Heading 1 style
**Verdict:** needs-investigation
**Ask summary:** `document.styles['Heading 1']` raises KeyError in documents where style name was not lowercased by `BabelFish.ui2internal`.
**Evidence:** `src/docx/styles/__init__.py:16` and `src/docx/styles/styles.py:37` still force `ui2internal` lookup; no fallback to raw name.
**TODO:** In `Styles.__getitem__`, fall back to raw-name `get_by_name(key)` if internal lookup misses — S.

### upstream#493 — How to get values of a graph from a docx file
**Verdict:** resolved-in-loadfix
**Ask summary:** Read category/series values from an embedded chart.
**Evidence:** `src/docx/chart.py:94` `ChartSeries.values`, `:102` `.categories`, `:174` `Chart.categories`.

### upstream#492 — How to make document content into two columns
**Verdict:** resolved-in-loadfix
**Ask summary:** Create a two-column layout in a section (multi-column).
**Evidence:** `src/docx/section.py:918` `SectionColumns` with `count`/`equal_width`/`space`; HISTORY D.19 (#60).

### upstream#486 — Append default styles from blank doc to opened template
**Verdict:** new-feature-needed
**Ask summary:** Import/copy built-in paragraph styles (e.g. "List Bullet") from a fresh Document into a template that is missing them.
**Evidence:** no match — no `import_styles` / `copy_style` helper in `src/docx/styles/`.
**TODO:** Add `Styles.import_builtin(name)` that materializes a latent style from defaults — M.

### upstream#485 — Need a way to get paragraph outline level
**Verdict:** new-feature-needed
**Ask summary:** Expose `w:outlineLvl` on Paragraph/ParagraphFormat for inferring document structure.
**Evidence:** `src/docx/oxml/text/parfmt.py:245` has XML-level `outlineLvl` but no proxy property in `src/docx/text/parfmt.py`.
**TODO:** Add `ParagraphFormat.outline_level` read/write property — S.

### upstream#484 — how to add a new section in a special paragraph
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert a section break at a specific paragraph (not just end-of-document).
**Evidence:** `src/docx/text/paragraph.py:712` `Paragraph.insert_section_break(...)`.

### upstream#481 — how to split a table?
**Verdict:** new-feature-needed
**Ask summary:** Break one table into two with a paragraph between, mirroring Word's `SplitTable`.
**Evidence:** no match — no `Table.split` in `src/docx/table.py`.
**TODO:** Add `Table.split(before_row)` returning new Table and inserted paragraph — M.

### upstream#478 — Moving insert_paragraph_before to super class
**Verdict:** resolved-in-loadfix
**Ask summary:** Make `insert_paragraph_before` work on tables too (generalise to block-level siblings).
**Evidence:** `src/docx/table.py:96` `Table.insert_paragraph_before` implemented.

### upstream#476 — Custom table style adds borders
**Verdict:** out-of-scope
**Ask summary:** Word re-applies borders to a user-defined table style on save.
**Evidence:** no match — behaviour is Word's style-reset; library writes what user sets via `Table.borders`.

### upstream#474 — the doc has a error (table.cells(0,0))
**Verdict:** out-of-scope
**Ask summary:** Typo in docs: `table.cells(0,0)` vs `table.cell(0,0)`; API itself is fine.
**Evidence:** `src/docx/table.py:1440` confirms `.cells` is a property, not callable. Upstream docs typo, not a loadfix bug.

### upstream#473 — Content Summary in a docx file (TOC)
**Verdict:** resolved-in-loadfix
**Ask summary:** Dynamically generate a Table of Contents.
**Evidence:** `src/docx/toc.py`; `Document.add_table_of_contents` (HISTORY #116).

### upstream#471 — Paragraph.get_listnum() — rendered list number
**Verdict:** new-feature-needed
**Ask summary:** Compute the displayed number ("9.1", "(a)") for a paragraph from its numbering definition.
**Evidence:** `src/docx/numbering.py` exposes definitions/levels but no rendered-counter evaluator.
**TODO:** Add `Paragraph.list_number` computed from numbering state walk — L.

### upstream#466 — merged DOCX corrupts equation when file order swapped
**Verdict:** needs-investigation
**Ask summary:** External combine_word_documents recipe corrupts embedded equations depending on part ordering.
**Evidence:** no match — no document-merge API in loadfix; unable to reproduce without the helper code.
**TODO:** If an official `Document.append_document` is added, ensure equation relationships/embeddings are re-linked — L.

### upstream#465 — Remove lnNumType tag from sectPr
**Verdict:** resolved-in-loadfix
**Ask summary:** Programmatically remove line-numbering from a section.
**Evidence:** `src/docx/section.py:396` `Section.remove_line_numbering()`.

### upstream#463 — Add a picture at a fixed place (wp:anchor)
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert an anchored (floating) image with position offsets.
**Evidence:** `src/docx/text/paragraph.py:442` `add_floating_image`; `src/docx/shape.py:136` `FloatingImage` (HISTORY D.17 #30).

### upstream#461 — Lack of attributes in drawing_list (anchor blipFill)
**Verdict:** resolved-in-loadfix
**Ask summary:** Access `blipFill`/`r:embed` on anchored (floating) drawings in addition to inline_shapes.
**Evidence:** `src/docx/drawing/__init__.py:46,76,90,200,207` traverse both `wp:inline` and `wp:anchor` graphic data.

### upstream#460 — Add paragraph by reference across documents
**Verdict:** new-feature-needed
**Ask summary:** Copy a paragraph (with formatting and inline images) from one Document into another.
**Evidence:** no match — no cross-document `add_paragraph_by_reference` in `src/docx/document.py`.
**TODO:** Add deep-copy paragraph helper that re-embeds referenced parts (images/hyperlinks) — L.

### upstream#455 — add_picture corrupts file due to docPr id clash with header
**Verdict:** new-bug-needed
**Ask summary:** `next_id` only scans the current story's XML, so a new inline image can reuse a `wp:docPr/@id` already in use in a header/footer, corrupting the file.
**Evidence:** `src/docx/parts/story.py:131` `next_id` xpaths `//@id` on `self._element` only; ignores headers/footers.
**TODO:** Compute `next_id` across document + all header/footer parts (package-wide max) — S.

### upstream#454 — w:noBreakHyphen not in Paragraph.text
**Verdict:** resolved-in-loadfix
**Ask summary:** Include non-breaking hyphen character in `Paragraph.text` output.
**Evidence:** `src/docx/oxml/text/run.py:107,184` include `w:noBreakHyphen` in run text xpath; `:321` defines CT_NoBreakHyphen.

### upstream#451 — Inline shapes not recognized (mc:AlternateContent)
**Verdict:** new-bug-needed
**Ask summary:** Drawings wrapped in `mc:AlternateContent`/`mc:Choice` are missed by the `//w:drawing/wp:inline` lookup.
**Evidence:** no grep hit for `AlternateContent` in `src/docx/`; inline_shapes/drawing xpaths do not descend through `mc:Choice`.
**TODO:** Extend inline/anchor xpaths to traverse `mc:AlternateContent/mc:Choice` and fall back to `mc:Fallback` — M.

### upstream#449 — Parsing text boxes in Word document
**Verdict:** resolved-in-loadfix
**Ask summary:** Iterate block items inside floating text boxes / DrawingML shapes.
**Evidence:** `src/docx/drawing/__init__.py` + HISTORY D.27 "DrawingML shapes and text-box content access" (#75).

### upstream#448 — build: add Python 3.5 environment to tox
**Verdict:** out-of-scope
**Ask summary:** Add tox test matrix for Python 3.5.
**Evidence:** `pyproject.toml:34` `requires-python = ">=3.9"` — Python 3.5 unsupported and obsolete.

## Batch summary
- resolved-in-loadfix: 12 (498, 497, 493, 492, 484, 478, 473, 465, 463, 461, 454, 449)
- new-feature-needed: 5 (486, 485, 481, 471, 460)
- new-bug-needed: 2 (455, 451)
- needs-investigation: 3 (496, 494, 466)
- out-of-scope: 3 (476, 474, 448)
- total: 25

### upstream#441 — feature: Column.delete()
**Verdict:** new-feature-needed
**Ask summary:** Table has `add_column` but no way to remove a column; request for `Column.delete()`.
**Evidence:** no match — `grep -rn "delete_column\|remove_column"` returns nothing in src/docx/table.py. `Table.delete` / `Paragraph.delete` / `Run.delete` exist (HISTORY #50) but not column-level.
**TODO (if applicable):** Add `_Column.delete()` that removes `w:gridCol` and matching `w:tc` per row. S

### upstream#437 — Injecting one docx document into another docx document
**Verdict:** new-feature-needed
**Ask summary:** Merge/compose a second docx's body (paragraphs, tables, images) into a placeholder location in another docx.
**Evidence:** no match — no `Composer` / `insert_document` / `append_document` in src/docx/. Phase D.13 covers single-paragraph/table insertion but not whole-doc merge.
**TODO (if applicable):** Add `Document.append_document(other)` / `Paragraph.insert_document_before(other)` with relationship/media remapping. L

### upstream#434 — Shading Python Docx Cells Only Work on a Single Cell
**Verdict:** needs-investigation
**Ask summary:** User's XML-parsing shading approach only colors last two cells; underlying bug is reuse of a single `w:shd` element across multiple cells.
**Evidence:** src/docx/oxml/table.py:59 `CT_Shd` and `_Cell.shading` supported via D.6 (HISTORY "Cell shading and background color (#63)"). User bug is about XML-element reuse, not a library bug, but loadfix API obviates it.
**TODO (if applicable):** Confirm docs demonstrate `_Cell.shading.fill = ...` instead of shared raw XML. S

### upstream#433 — Changing Table Cell Borders and Inserting XML Elements
**Verdict:** resolved-in-loadfix
**Ask summary:** Need API to adjust `w:tcBorders` on cells/rows (remove borders around first/last rows).
**Evidence:** src/docx/oxml/table.py:127 `CT_TcBorders`; src/docx/table.py:305 `Table.borders`, `_Cell.borders` (HISTORY "Add Table.borders / _Cell.borders (#102)").

### upstream#430 — RTL attribute disables font name!
**Verdict:** needs-investigation
**Ask summary:** Setting `style.font.rtl = True` appears to drop the configured font name for the paragraph (mixed Persian/English).
**Evidence:** src/docx/text/font.py:660 rtl property and font.name both exist. RTL sets `w:rPr/w:rtl`; complex-script font mapping likely requires `w:rFonts/@w:cs`. No cs-font convenience in font.py.
**TODO (if applicable):** Investigate whether `Font.name` writes `@w:cs` for bidi runs; add `Font.complex_script_name` if missing. M

### upstream#428 — chang one row.cells.text all row.cells changed
**Verdict:** resolved-in-loadfix
**Ask summary:** Setting `.text` on one cell of a merged-cell row mutates other cells (because `row.cells` returned same object for merged span).
**Evidence:** src/docx/table.py:1560 `_Row.cells` iterates tc elements with merge-aware logic; HISTORY "Add Cell.is_merge_origin / merge_origin (#145)".

### upstream#426 — python_docx for reading and changing the font of characters
**Verdict:** out-of-scope
**Ask summary:** User wants per-character font size / line spacing but unaware runs are the granularity; this is usage help not a feature.
**Evidence:** src/docx/text/run.py and src/docx/text/font.py already expose per-run font; `Run.split` (HISTORY #94) lets callers isolate a substring run.

### upstream#425 — Feature: Add Bookmark
**Verdict:** resolved-in-loadfix
**Ask summary:** Add a bookmark create/read/delete API (incl. cross-refs).
**Evidence:** Phase C — "Add bookmarks create / read / delete (#52)" + "Add REF / PAGEREF cross-reference resolution (#115)". src/docx/bookmarks.py, src/docx/text/paragraph.py:56 `add_bookmark`.

### upstream#422 — Tables have incorrect rows' and columns' cells when there are merged cells
**Verdict:** resolved-in-loadfix
**Ask summary:** `row.cells` / `column.cells` give wrong values when rows have horizontal merges / short rows.
**Evidence:** src/docx/table.py:1560 merge-aware `_Row.cells`; `grid_cols_before/after` documented; `Cell.merge_origin` exposed (#145).

### upstream#421 — Lists in RTL style
**Verdict:** needs-investigation
**Ask summary:** RTL list (`ListNumber` + alignment=right + font.rtl) puts bullet on wrong side.
**Evidence:** RTL support at Run (#127) and Section (#148); no RTL-list preset / `Paragraph.bidi` helper beyond what font.rtl covers. Requires `w:pPr/w:bidi` for list numbering direction.
**TODO (if applicable):** Verify Paragraph-level `bidi` toggle writes `w:pPr/w:bidi` and flips numPr direction. S

### upstream#420 — KeyError: "no style with name 'BODY_TEXT (-67)'"
**Verdict:** needs-investigation
**Ask summary:** `styles[WD_STYLE.BODY_TEXT]` raises KeyError because the enum repr is used as a name lookup rather than style-id mapping.
**Evidence:** src/docx/styles/styles.py:32 `__getitem__` looks up by name via `BabelFish.ui2internal(key)`; deprecation branch calls `get_by_id` with the key (int value). WD_STYLE enum member str representation "BODY_TEXT (-67)" is fed to name lookup first.
**TODO (if applicable):** Handle WD_STYLE enum member in `Styles.__getitem__` — unwrap to `.value` and route to builtin-id lookup. S

### upstream#419 — Drop Shadow on an image
**Verdict:** new-feature-needed
**Ask summary:** Apply image "Drop Shadow Rectangle" style to an inserted picture.
**Evidence:** no match — `grep shadow` in src/docx/shape.py / drawing: no picture-effect API. `Font.shadow` exists for run text only.
**TODO (if applicable):** Add `InlineShape.effects` / `shadow` that emits `a:effectLst/a:outerShdw`. M

### upstream#415 — Replace text in paragraph keeping the runs object and styles
**Verdict:** resolved-in-loadfix
**Ask summary:** Search/replace text while preserving runs and formatting.
**Evidence:** src/docx/search.py (~560 lines) + HISTORY D.10 "Search and replace with formatting preservation (#91)" and "Document.search_regex / replace_regex / search_all / replace_all (#153, #154)".

### upstream#413 — Accessing the content of TextBox and Headers (for purpose of find and replace)
**Verdict:** needs-investigation
**Ask summary:** Find/replace should cover headers and text-box content.
**Evidence:** src/docx/search.py:362 `_iter_all_paragraphs` covers headers/footers, footnotes/endnotes/comments — but NOT `w:txbxContent` inside drawings/shapes. D.27 adds text-box content *access*, but search module doesn't traverse.
**TODO (if applicable):** Extend `_iter_all_paragraphs` to yield paragraphs inside `w:txbxContent`. M

### upstream#411 — feature: adding drawing canvas to docx
**Verdict:** new-feature-needed
**Ask summary:** Wrap image(s)/shapes in a `wdInlineShapeLockedCanvas` so multiple shapes stay aligned.
**Evidence:** src/docx/oxml/drawing.py exposes CT_TextBox (wps:txbx) but no `lockedCanvas` wrapper; grep for `LockedCanvas` returns no match.
**TODO (if applicable):** Add `add_canvas()` / grouped-shape container creation support. L

### upstream#409 — After searching and replacing text the text font size of the replaced text is set to 12pt
**Verdict:** resolved-in-loadfix
**Ask summary:** External regex replace loses run font sizing; asks for formatting-preserving replace.
**Evidence:** D.10 "Search and replace with formatting preservation (#91)"; `Document.replace_regex` (#154) preserves run properties by splitting runs instead of rewriting.

### upstream#408 — add_paragraph() performance drops as document length increases
**Verdict:** needs-investigation
**Ask summary:** `Document.add_paragraph` is O(n²): each add iterates all paragraphs.
**Evidence:** src/docx/blkcntnr.py:102 `_add_paragraph` calls `self._element.add_p()`, generated by `ZeroOrMore` (xmlchemy). `_insert_*` walks siblings to find successor insertion point — upstream behavior unchanged. No explicit fast-path in loadfix commits.
**TODO (if applicable):** Profile; consider caching last-insertion-point or using `lxml.etree.SubElement` bulk append. M

### upstream#407 — Insert image in all pages in docx or header
**Verdict:** resolved-in-loadfix
**Ask summary:** Way to add image to header on every page.
**Evidence:** src/docx/section.py `Section.header.add_paragraph().add_run().add_picture()` supported; HISTORY "Section odd/even page header-footer (#149)" and general header/footer support baseline from upstream.

### upstream#403 — how to create bookmarks?
**Verdict:** resolved-in-loadfix
**Ask summary:** Create bookmarks usable with `{ REF BOOKMARK_NAME \h }` fields.
**Evidence:** Phase C — Paragraph.add_bookmark (src/docx/text/paragraph.py:56); REF/PAGEREF resolution (#115).

### upstream#402 — Error while importing from docx (aka python-docx)
**Verdict:** out-of-scope
**Ask summary:** Install issue on Python 3.6 — SyntaxError on `from docx import Document`.
**Evidence:** Legacy Py3.6 environment issue; loadfix already supports modern Python (py.typed, 3.9+).

### upstream#400 — How to insert images behind the text or in front of the text
**Verdict:** resolved-in-loadfix
**Ask summary:** Add floating (non-inline) images with positioning.
**Evidence:** HISTORY D.17 "Floating images with wp:anchor positioning (#30)"; src/docx/shape.py:137 `FloatingImage`; src/docx/text/paragraph.py `add_floating_image`.

### upstream#397 — Save the document in text format
**Verdict:** out-of-scope
**Ask summary:** Export DOCX to plain .txt.
**Evidence:** `Document.paragraphs[*].text` already produces text; dedicated txt exporter is beyond OOXML coverage scope of this fork.

### upstream#389 — Problem about writing Chinese text into a docx file
**Verdict:** out-of-scope
**Ask summary:** Python 2 Unicode encoding confusion reading Chinese from CSV then writing via python-docx.
**Evidence:** Py2-era encoding bug; loadfix is Py3-only (HISTORY 1.0.0 "Remove Python 2 support"). N/A.

### upstream#387 — rtl support in headers
**Verdict:** needs-investigation
**Ask summary:** Applying RTL to a header paragraph overwrites the "Header" style; wants both RTL and Heading style applied.
**Evidence:** RTL on paragraph + run (#127) and Section text_direction (#148) exist; linking style + direct formatting is standard API (`paragraph.style = ...; run.font.rtl = True`). User-error flavor, but worth doc example.
**TODO (if applicable):** Add docs/user/rtl-headers.rst example showing style + bidi coexistence. S

### upstream#385 — How to add some Chinese text into a paragraph successfully?
**Verdict:** out-of-scope
**Ask summary:** Python 2 + `print`-less syntax error; CJK text addition confusion.
**Evidence:** Py2-only issue (`print exp in ...` syntax). Loadfix is Py3-only.

## Batch summary
- resolved-in-loadfix: 9 (#433, #428, #425, #422, #415, #409, #407, #403, #400)
- new-feature-needed: 4 (#441, #437, #419, #411)
- new-bug-needed: 0
- needs-investigation: 7 (#434, #430, #421, #420, #413, #408, #387)
- out-of-scope: 5 (#426, #402, #397, #389, #385)

### upstream#383 — Getting the default font for a document
**Verdict:** new-feature-needed
**Ask summary:** Expose `w:docDefaults/w:rPrDefault` so `document.styles['Normal'].font.name` reflects the document-defaults font when the Normal style omits `rFonts`.
**Evidence:** `w:docDefaults` is only a successor tuple entry in src/docx/oxml/styles.py:339; no CT_DocDefaults class / `Styles.default_font` proxy exists in src/docx/styles/.
**TODO (if applicable):** Add `Styles.document_default_font` proxy over `w:docDefaults/w:rPrDefault/w:rPr` — M.

### upstream#382 — How to set a cover to a word document?
**Verdict:** out-of-scope
**Ask summary:** User asks how to add a "cover page" (likely Word's Cover Page gallery).
**Evidence:** Usage question; WD_BUILDING_BLOCK_TYPE.COVER_PAGES exists but no add-cover-page helper, and this is really a recipe.
**TODO (if applicable):** n/a (documentation / recipe).

### upstream#381 — Streaming the content
**Verdict:** out-of-scope
**Ask summary:** Wants Django `StreamingHttpResponse`-style incremental delivery from `Document.save`.
**Evidence:** docx zip must finalise central directory at end of write; cannot stream partial output. `Document.save` accepts file-like objects (document.py:899) which already suffices.
**TODO (if applicable):** n/a.

### upstream#380 — Drawings are not parsed, and text inside drawings is not available
**Verdict:** resolved-in-loadfix
**Ask summary:** Wants to extract text from text boxes / DrawingML shapes inside a docx.
**Evidence:** src/docx/drawing/__init__.py:112 exposes `Drawing.text` / `Drawing.paragraphs` over text-frame content — HISTORY Phase D.27 (#75).
**TODO (if applicable):** n/a.

### upstream#379 — Table of deleted-document can still be accessed
**Verdict:** out-of-scope
**Ask summary:** Requests a `Document.close()` so post-save mutations fail fast.
**Evidence:** No `close()` / invalidation machinery on `Document`; pythonic lifetime relies on GC. Upstream maintainer framed this as a code-organisation issue.
**TODO (if applicable):** Optional `Document.close()` invalidating proxies — S.

### upstream#370 — Assign style to table by row?
**Verdict:** new-feature-needed
**Ask summary:** Wants `row.style = ...` or per-row conditional formatting (row-level shading / style).
**Evidence:** `_Row` has no `style` property in src/docx/table.py; conditional formatting is only via `Table.style_flags` (#144). No row-level convenience.
**TODO (if applicable):** Add `_Row.apply_shading` / bulk `_Row.cells` style/shade helper — S.

### upstream#366 — Not all Paragraph styles' font can be changed
**Verdict:** needs-investigation
**Ask summary:** Setting `paragraph.style.font.name` on Heading 1-9 / Title / Subtitle has no effect (theme-font inheritance override).
**Evidence:** Heading styles use `w:rFonts w:asciiTheme="majorHAnsi"` in default-docx-template/word/styles.xml; loadfix `Font.name` setter doesn't clear the theme attrs, so directly-set `w:ascii` is shadowed.
**TODO (if applicable):** Make `Font.name` setter clear sibling `asciiTheme`/`hAnsiTheme` — S.

### upstream#365 — Feature: don't add space between paragraphs of the same style
**Verdict:** new-feature-needed
**Ask summary:** Expose `w:contextualSpacing` as a ParagraphFormat / Style property.
**Evidence:** `w:contextualSpacing` is only listed as a successor in src/docx/oxml/text/parfmt.py:202; no `ParagraphFormat.contextual_spacing` getter/setter.
**TODO (if applicable):** Add `ParagraphFormat.contextual_spacing` bool property — S.

### upstream#363 — Feature: can open .dotx files and perhaps change type
**Verdict:** new-feature-needed
**Ask summary:** Add support for opening `.dotx` / `.dotm` Word template files and converting between template and document content types.
**Evidence:** `api.py:39` rejects anything but WML_DOCUMENT_MAIN and WML_DOCUMENT_MACRO — no WML_TEMPLATE handling. .docm is allowed (D.24).
**TODO (if applicable):** Accept `WML_TEMPLATE` / `WML_TEMPLATE_MACRO`, optionally expose `Document.content_type` setter — M.

### upstream#360 — Column widths in MS Word
**Verdict:** resolved-in-loadfix
**Ask summary:** Table column widths set via `table.columns[i].width` are ignored by MS Word (autofit override).
**Evidence:** HISTORY lists Phase D.26 "Table autofit and column-width control (#39)"; `Table.autofit` + per-cell width handling exists in src/docx/table.py.
**TODO (if applicable):** n/a.

### upstream#359 — A follow up on the issue of Figure caption
**Verdict:** resolved-in-loadfix
**Ask summary:** Provide a first-class caption API around `SEQ Figure`/`SEQ Table` fields instead of hand-rolled OxmlElement code.
**Evidence:** src/docx/captions.py + `Document.add_caption`, `Paragraph.add_caption_before/after` — HISTORY "caption helpers (#141)".
**TODO (if applicable):** n/a.

### upstream#358 — Clear all drawing elements
**Verdict:** out-of-scope
**Ask summary:** Usage question about placeholder-replace that preserves inline images when reassigning `run.text`.
**Evidence:** This is a mail-merge recipe; `Run.text` setter intentionally clears inner content. loadfix provides `search_regex`/`replace_regex` (#153/#154) which preserve formatting but this is user code.
**TODO (if applicable):** n/a.

### upstream#357 — How to insert &amp into word document
**Verdict:** out-of-scope
**Ask summary:** User confused about inserting `&` — actually works natively; lxml handles escaping.
**Evidence:** `xml:space="preserve"` handled at src/docx/oxml/text/run.py:64; no escaping bug reported.
**TODO (if applicable):** n/a (documentation).

### upstream#354 — Append new row to table
**Verdict:** out-of-scope
**Ask summary:** Usage question about incrementing a leading number column when calling `table.add_row()`.
**Evidence:** `table.add_row()` exists (table.py:86); counter management is user responsibility.
**TODO (if applicable):** n/a.

### upstream#351 — Add SVG
**Verdict:** resolved-in-loadfix
**Ask summary:** Support embedding SVG images in a document.
**Evidence:** `CT_Picture.new_svg` / `new_svg_pic_inline` with `asvg:svgBlip` extension at src/docx/oxml/shape.py:408+; HISTORY Phase D.22 (#76).
**TODO (if applicable):** n/a.

### upstream#350 — Failed to read image file while add_picture
**Verdict:** needs-investigation
**Ask summary:** `UnicodeDecodeError` in `docx/image/helpers.py` when reading certain image byte streams as UTF-8.
**Evidence:** src/docx/image/helpers.py still uses `chars.decode('UTF-8')` without error handling; likely triggered by corrupt or non-UTF8 stream in an image sniffer path.
**TODO (if applicable):** Add defensive decode with `errors="replace"` and a clearer image-format error — S.

### upstream#348 — Grid color (Table borders missing when adding to existing doc)
**Verdict:** needs-investigation
**Ask summary:** `document.add_table` on an existing docx produces a borderless table because the default `Table Normal` style has no borders there.
**Evidence:** No explicit loadfix fix for default table-style resolution when opening arbitrary templates; `Table.borders` (#102) lets caller set borders but default behavior still differs.
**TODO (if applicable):** Document the root cause; optionally default `style="Table Grid"` in `add_table` — S.

### upstream#347 — Shift pages
**Verdict:** resolved-in-loadfix
**Ask summary:** Needs ability to insert new content before an existing paragraph / page (append doesn't work at end of template).
**Evidence:** HISTORY Phase D.13 (#26) "Insert paragraph/table at arbitrary position" + `Paragraph.insert_paragraph_before`.
**TODO (if applicable):** n/a.

### upstream#346 — Cannot set Chinese character font typeface
**Verdict:** resolved-in-loadfix
**Ask summary:** Setting `font.name` doesn't change East-Asian (CJK) glyphs because `w:rFonts/@w:eastAsia` is a separate attribute.
**Evidence:** `Font.name_east_asia` (src/docx/text/font.py:581) plus `rFonts_eastAsia` accessor at oxml/text/font.py:270 — HISTORY entries #127/#128/#160.
**TODO (if applicable):** n/a.

### upstream#345 — domo.docx non readable
**Verdict:** needs-investigation
**Ask summary:** Quickstart sample produces a doc that won't open; uses `style='IntenseQuote'` / `'ListBullet'` (no-space style names).
**Evidence:** Upstream style-name normalization treats these as aliases; may fail or produce invalid `w:pStyle` under certain paths. No loadfix regression test found for these exact style aliases.
**TODO (if applicable):** Verify style-alias resolution for IntenseQuote/ListBullet/ListNumber — S.

### upstream#343 — Adding section creating an error in the xml data structure
**Verdict:** needs-investigation
**Ask summary:** `document.add_section()` on a template produces a file Word reports as corrupt (header/footer reference issue).
**Evidence:** `add_section` at src/docx/document.py:257 clones sectPr and may duplicate header/footer rId refs without relating the new sectPr. No explicit loadfix fix noted.
**TODO (if applicable):** Audit new-section header/footer ref cloning; add regression test — M.

### upstream#342 — XPath Performance - precompiling
**Verdict:** new-feature-needed
**Ask summary:** Pre-compile recurring XPath expressions via `lxml.etree.XPath` to speed up large-document generation.
**Evidence:** No precompiled XPath cache anywhere in src/docx/; each `.xpath("...")` call re-parses the expression.
**TODO (if applicable):** Introduce `_XP("expr")` cache for hot XPaths in oxml modules — M.

### upstream#340 — Revisions / track changes
**Verdict:** resolved-in-loadfix
**Ask summary:** Read original & amended text with tracked-change support.
**Evidence:** src/docx/tracked_changes.py covers `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` read, accept/reject, formatting changes, revision_marks_text — HISTORY Phase B.
**TODO (if applicable):** n/a.

### upstream#339 — Missing character (smartTag wrapped run)
**Verdict:** needs-investigation
**Ask summary:** `paragraph.runs` misses text inside `<w:smartTag>` wrapper elements.
**Evidence:** `w:smartTag` appears only in src/docx/oxml/settings.py (unrelated); `Paragraph.text` delegates to `CT_P.text` which likely doesn't descend into smartTag.
**TODO (if applicable):** Make `CT_P.text` / `paragraph.runs` recurse through `w:smartTag` — S.

### upstream#338 — Rotation text in document
**Verdict:** resolved-in-loadfix
**Ask summary:** Rotate text 90° inside table cells.
**Evidence:** `_Cell.text_direction` (src/docx/oxml/table.py:1288) + `Cell.text_direction` proxy surface `w:textDirection` (HISTORY #142), which is Word's text-rotation knob for cells.
**TODO (if applicable):** n/a.

## Batch summary
- resolved-in-loadfix: 8 (#380, #360, #359, #351, #347, #346, #340, #338)
- new-feature-needed: 5 (#383, #370, #365, #363, #342)
- needs-investigation: 6 (#366, #350, #348, #345, #343, #339)
- out-of-scope: 6 (#382, #381, #379, #358, #357, #354)
- new-bug-needed: 0

Totals: 8 + 5 + 6 + 6 + 0 = 25.

### upstream#335 — Skips over some text in a table
**Verdict:** needs-investigation
**Ask summary:** Bug: `paragraph.text` of cells skips some text (likely `w:smartTag` content not unwrapped).
**Evidence:** `src/docx/oxml/settings.py:439` only references smartTagType; no smartTag descent in `text` accessor.
**TODO (if applicable):** Unwrap `w:smartTag` runs in paragraph/cell text iteration — S.

### upstream#332 — add_picture() supports floating image and alignment
**Verdict:** resolved-in-loadfix
**Ask summary:** Allow `add_picture` to create floating/anchored images and align them.
**Evidence:** `src/docx/text/paragraph.py:442` `add_floating_image` (Phase D.17, commit f51e7a9).

### upstream#331 — How do I apply Total Row to last row in a table
**Verdict:** new-feature-needed
**Ask summary:** Enable the "Total Row" table style flag on the last row (uses `w:tblLook`/conditional formatting).
**Evidence:** `src/docx/table.py:1165` `style_flags` covers `w:tblLook` flags, but no explicit "last-row" convenience.
**TODO (if applicable):** Document or wrap `tblLook.lastRow` toggle via `Table.style_flags.last_row=True` — S.

### upstream#328 — How to get the text in the tag 'smartTag'
**Verdict:** needs-investigation
**Ask summary:** `paragraph.text` skips content inside `w:smartTag`; user wants it included.
**Evidence:** No smartTag traversal in paragraph.text; same root cause as #335.
**TODO (if applicable):** Make paragraph/run text traversal descend into `w:smartTag` — S.

### upstream#324 — How to insert image with ignoring margin?
**Verdict:** resolved-in-loadfix
**Ask summary:** Position image past page margins using anchor positioning.
**Evidence:** `src/docx/text/paragraph.py:442` `add_floating_image` with `position` relative to page (D.17).

### upstream#322 — How to set a table row to repeated as header?
**Verdict:** resolved-in-loadfix
**Ask summary:** Toggle `w:tblHeader` on a row to repeat on each page.
**Evidence:** `src/docx/table.py:1674` `_Row.is_header` getter/setter (HISTORY `_Row.is_header (#93)`).

### upstream#321 — How to set the table column different alignment?
**Verdict:** out-of-scope
**Ask summary:** Per-column alignment of cell text; Word has no native column-alignment property (set at cell/paragraph level).
**Evidence:** existing `Paragraph.alignment` / cell access covers the achievable case; no Word property to proxy.

### upstream#320 — Insert MathML as plain text for Word 2010 to render as equation
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert MathML/OMML markup so Word renders it as a native equation.
**Evidence:** `src/docx/equations.py` and `Document.equations` (commit mentions Phase Equations); OMML insert API available.

### upstream#319 — Issues with applying table styles
**Verdict:** needs-investigation
**Ask summary:** Custom table style imported from template isn't applied to new table; `add_table` style argument reverts to 'Normal Table'.
**Evidence:** `src/docx/document.py:299` `add_table` doesn't accept `_TableStyle` directly in arg check; styles module unchanged on this path.
**TODO (if applicable):** Reproduce and fix style lookup in `Document.add_table` / `_Body.add_table` when a `_TableStyle` is passed — M.

### upstream#315 — Table always fills page width
**Verdict:** needs-investigation
**Ask summary:** Newly created tables always span full width regardless of `alignment`; user expects shrink-to-contents.
**Evidence:** `src/docx/blkcntnr.py:64` `add_table(width=_block_width)` still defaults to block width; no "auto" width option.
**TODO (if applicable):** Expose `preferred_width=None` / `WD_TABLE.AUTOFIT` shortcut in `add_table` — S.

### upstream#309 — feature: Font.shading
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose run/character-style background `w:shd` as a Font property.
**Evidence:** `src/docx/text/font.py:700` `Font.shading_color` (Phase D.20 #33).

### upstream#306 — Adding Style to Table after Adding Rows in Pre-existing Document
**Verdict:** needs-investigation
**Ask summary:** Row borders/style from existing style not applied to rows added via `add_row()`.
**Evidence:** `add_row` clones `trPr` but conditional formatting from `tblStyle` not reapplied at render time; no fix in loadfix.
**TODO (if applicable):** Investigate applying `w:tblLook`/cnfStyle tracking on added rows — M.

### upstream#305 — Paragraph format in table cell
**Verdict:** needs-investigation
**Ask summary:** Setting `space_after`/`space_before = Inches(0)` on cell paragraph has no visible effect (likely style-inheritance issue with `TableNormal`).
**Evidence:** `src/docx/text/parfmt.py:352` sets pPr values correctly; upstream confusion is cell-style chain, not a bug — no loadfix change.
**TODO (if applicable):** Document style-override semantics for cell paragraph_format — S.

### upstream#300 — Color an individual cell
**Verdict:** resolved-in-loadfix
**Ask summary:** Property to set cell fill color.
**Evidence:** `src/docx/table.py:895` `_Cell.shading.fill_color` (Phase D.6 #63).

### upstream#298 — section.orientation = WD_ORIENT.LANDSCAPE don't work
**Verdict:** out-of-scope
**Ask summary:** User expected assigning orientation to auto-swap width/height; behavior is documented (setter only writes `w:orient`, width/height remain).
**Evidence:** `src/docx/section.py:254` orientation is standalone sz attribute; docs explicitly show manual swap.

### upstream#294 — Embed .xlsx file into .docx
**Verdict:** new-feature-needed
**Ask summary:** Add an API to embed .xlsx (and other OLE objects) into a document.
**Evidence:** `src/docx/embedded_objects.py` is read-only (commit: "read-only embedded OLE objects #140"); no add/create API.
**TODO (if applicable):** Implement `Run.add_embedded_object(path, content_type)` writing oleObject part + EMB rel — L.

### upstream#293 — Manually iterate to next_run or next_paragraph
**Verdict:** resolved-in-loadfix
**Ask summary:** Convenience to walk to next run/paragraph for templating; and delete paragraphs between markers.
**Evidence:** `src/docx/document.py:701` `iter_inner_content`; `Paragraph.delete` / `Run.delete` (Phase `#50`); search/replace helpers in `src/docx/search.py`.

### upstream#292 — Support for divid (paragraph unique id)
**Verdict:** resolved-in-loadfix
**Ask summary:** Read/write a stable unique id per paragraph.
**Evidence:** `src/docx/ids.py:56` `compute_stable_id`; `stable_id` on Paragraph/Run/Table/Cell (#155).

### upstream#284 — extension: support .docm (macro-enabled) Word files
**Verdict:** resolved-in-loadfix
**Ask summary:** Read and save .docm files preserving VBA macros.
**Evidence:** HISTORY: "D.24 .docm macro-enabled file support (#65)".

### upstream#279 — feature: del table.rows[i]
**Verdict:** needs-investigation
**Ask summary:** API to delete a row from a table.
**Evidence:** `src/docx/table.py:62` `Table.delete()` exists but no `_Row.delete()` / `__delitem__` found on `_Rows`.
**TODO (if applicable):** Add `_Row.delete()` (remove tr) and `_Rows.__delitem__` — S.

### upstream#275 — feature: Run.insert_run_before()
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert a new run positioned before an existing run; also run-splitting for partial formatting.
**Evidence:** `src/docx/text/run.py:268` `Run.split(offset)` (HISTORY `Run.split (#94)`); split covers the stated splitting need.

### upstream#271 — feature: BlockItemContainer.content (+ Table.cells, all_paragraphs, Paragraph.replace regex)
**Verdict:** resolved-in-loadfix
**Ask summary:** Iterate all block content; all cells; all paragraphs; regex replace preserving runs.
**Evidence:** `Document.iter_inner_content`; `Document.replace_regex` / `replace_regex_all` (Phase D.10 #91); search helpers in `src/docx/search.py`.

### upstream#270 — feature: copy table across open documents
**Verdict:** new-feature-needed
**Ask summary:** Copy a whole `Table` (with styles/relationships) from one document into another.
**Evidence:** No `Table.copy` / cross-document cloning helper in `src/docx/table.py` or `document.py`.
**TODO (if applicable):** Add `Document.add_table_from(other_table)` that deep-copies CT_Tbl and migrates rels — L.

### upstream#265 — feature: accommodate invalid BMP images at default resolution
**Verdict:** needs-investigation
**Ask summary:** `add_picture` with a BMP whose DPI headers are 0 raises `ZeroDivisionError`.
**Evidence:** `src/docx/image/bmp.py` parses BMP headers; no guard for zero DPI → same bug likely present.
**TODO (if applicable):** Default to 72 DPI when BMP horz/vert_dpi is 0 in `src/docx/image/bmp.py` — S.

### upstream#262 — feature: mailmerge fields
**Verdict:** resolved-in-loadfix
**Ask summary:** Enumerate and update MERGEFIELD / mail-merge data source in a document.
**Evidence:** `src/docx/settings.py:220` `Settings.mail_merge` + `enable_mail_merge` / `disable_mail_merge` (HISTORY `Settings.mail_merge (#130)`); complex fields via `src/docx/fields.py` (Phase C).

## Batch summary
- resolved-in-loadfix: 12 (#332, #324, #322, #320, #309, #300, #293, #292, #284, #275, #271, #262)
- new-feature-needed: 3 (#331, #294, #270)
- needs-investigation: 8 (#335, #328, #319, #315, #306, #305, #279, #265)
- out-of-scope: 2 (#321, #298)
- new-bug-needed: 0
- Total: 25

### upstream#252 — feature: Document.text
**Verdict:** new-feature-needed
**Ask summary:** Quick helper to read the entire document's plain text (e.g., `document.text`) without iterating paragraphs manually.
**Evidence:** `src/docx/document.py` has no `.text` property (grep shows only `iter_inner_content`/`revision_marks_text`). `Paragraph.text` exists but no Document-wide aggregate.
**TODO (if applicable):** Add `Document.text` returning concatenated paragraph text (optionally traversing tables/stories) — S.

### upstream#249 — feature: InlineShape.image ?
**Verdict:** new-feature-needed
**Ask summary:** Expose underlying image bytes/Image object from an `InlineShape` (blipFill→rId→ImagePart) so consumers can export embedded pictures.
**Evidence:** `src/docx/shape.py:51` `InlineShape` has width/height/type/alt_text/title but no `.image` or `.image_part`. Analogous accessor exists on `Drawing` (`src/docx/drawing/__init__.py:52`).
**TODO (if applicable):** Add `InlineShape.image` → `Image` via rId lookup, mirroring `Drawing.image` — S.

### upstream#248 — feature: Font.cs_size
**Verdict:** new-feature-needed
**Ask summary:** Expose `w:szCs` (complex-script font size) as a separate `Font` property; Hebrew/Arabic runs set only `szCs`, so `Font.size` returns None.
**Evidence:** `src/docx/oxml/text/font.py:149` declares `w:szCs` child, but `src/docx/text/font.py` exposes `size` (szCs is not surfaced); `complex_script` property exists at line 232 but no `cs_size`.
**TODO (if applicable):** Add `Font.cs_size` / `cs_size` setter backed by `w:rPr/w:szCs` — S.

### upstream#246 — fix: escape style names with embedded quotes / XPath metacharacters
**Verdict:** resolved-in-loadfix
**Ask summary:** Indexing `styles[name]` crashes when the style name contains `"` because the name is interpolated into XPath. Escape via variable binding.
**Evidence:** `src/docx/oxml/styles.py:378` uses parameterized `xpath("w:style[w:name/@w:val=$name]", name=name)`, avoiding quote-injection.
**TODO (if applicable):** n/a.

### upstream#245 — feature: Row.allow_break_across_pages
**Verdict:** resolved-in-loadfix
**Ask summary:** Wrap `w:trPr/w:cantSplit` so callers can prevent a table row from splitting across pages.
**Evidence:** `src/docx/table.py:1545` `_Row.allow_break_across_pages` getter/setter (D.16 in HISTORY.rst, referencing #51).
**TODO (if applicable):** n/a.

### upstream#243 — feature: Font.spacing (character spacing)
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose `w:rPr/w:spacing` so callers can adjust character spacing (expanded/condensed).
**Evidence:** `src/docx/text/font.py:41` `Font.character_spacing` getter + setter at line 55.
**TODO (if applicable):** n/a.

### upstream#241 — perf: XPath vs find in xmlchemy
**Verdict:** needs-investigation
**Ask summary:** Replace ElementTree `find`/`findall` in `_OxmlElementBase.first_child_found` with lxml XPath for a speed boost on large documents.
**Evidence:** `src/docx/oxml/xmlchemy.py` still uses `find` for first-child lookup; no perf overhaul committed. HISTORY.rst has no perf entry.
**TODO (if applicable):** Benchmark xmlchemy first-child lookup; switch to compiled XPath if measurable — M.

### upstream#235 — feature: add multiple pictures in the same paragraph
**Verdict:** resolved-in-loadfix
**Ask summary:** Support positioning multiple pictures in one paragraph (floating / anchored images), like pptx `add_picture(left,top,width,height)`.
**Evidence:** HISTORY "D.17 Floating images with wp:anchor positioning (#30)"; `src/docx/shape.py:136` `FloatingImage`; `Paragraph.add_floating_image` referenced at shape.py:140.
**TODO (if applicable):** n/a.

### upstream#232 — How to detect merged cells when reading tables
**Verdict:** resolved-in-loadfix
**Ask summary:** Let callers detect that two cells share a merge (vMerge / gridSpan), not just that `.text` matches.
**Evidence:** `src/docx/table.py:659` `_Cell.is_merge_origin`, line 687 `_Cell.merge_origin` (HISTORY D.* "Cell.is_merge_origin / merge_origin (#145)").
**TODO (if applicable):** n/a.

### upstream#225 — paragraph.runs misses runs inside <w:smartTag>
**Verdict:** new-bug-needed
**Ask summary:** `Paragraph.runs` only returns direct `w:r` children, omitting runs nested under `w:smartTag`/`w:customXml`.
**Evidence:** `src/docx/oxml/text/paragraph.py` `r_lst` uses `ZeroOrMore("w:r")` (direct children); no smartTag descent. `grep smartTag` hits only settings.py (smartTagType).
**TODO (if applicable):** Extend `Paragraph.runs` to descend through `w:smartTag`/`w:customXml` wrappers — M.

### upstream#224 — Feature: Read checkboxes in Word forms
**Verdict:** resolved-in-loadfix
**Ask summary:** Read the checked state of legacy Word form checkboxes (FORMCHECKBOX / ffData/checkBox).
**Evidence:** `src/docx/form_fields.py:141` `CheckboxFormField` (legacy FFDATA) + `src/docx/content_controls.py:277` SDT-checkbox read (D.14, legacy form fields #123 in HISTORY).
**TODO (if applicable):** n/a.

### upstream#223 — Document Printing
**Verdict:** out-of-scope
**Ask summary:** User asks how to print a `Document` from Python; not an OOXML feature.
**Evidence:** OS-level printing is outside python-docx's remit; no related module.
**TODO (if applicable):** n/a.

### upstream#221 — Bold + font / borders inside a table cell
**Verdict:** resolved-in-loadfix
**Ask summary:** User wants to mix fonts/bold within one cell and set cell borders — a usage question answered by existing API (runs with distinct fonts, cell borders).
**Evidence:** `src/docx/text/run.py` run.bold/font.name standard; `src/docx/table.py:552` `_Cell.borders` (CellBorders D.* #102 / #143).
**TODO (if applicable):** n/a.

### upstream#220 — Unicode title can't be set (coreprops utf-8 encode)
**Verdict:** needs-investigation
**Ask summary:** Old bug where `coreprops.py` forcibly utf-8 encoded unicode strings, breaking unicode titles on py2.
**Evidence:** `src/docx/oxml/coreprops.py` now returns str (py3-only); but worth validating round-trip of non-ASCII CoreProperties in modern loadfix test.
**TODO (if applicable):** Add regression test writing/reading a non-ASCII `core_properties.title` — S.

### upstream#217 — Lists showing up as normal paragraphs
**Verdict:** resolved-in-loadfix
**Ask summary:** Applying a "List Number"/"List Bullet" style from a newly-created doc does not produce numbering because the target document has no matching numbering definition.
**Evidence:** HISTORY "D.9 Numbering style control (#22)"; `src/docx/numbering.py:156` `Numbering.add_numbering_definition` plus `List Number` helpers in templates.
**TODO (if applicable):** n/a.

### upstream#215 — Example in docs does not run when pasted
**Verdict:** needs-investigation
**Ask summary:** Quick-start example fails verbatim due to style-name mismatch (`IntenseQuote` vs `Intense Quote`) / formatting in rendered HTML.
**Evidence:** duplicate of #198 essentially; docs at `docs/user/quickstart.rst` need review.
**TODO (if applicable):** Verify quick-start sample runs end-to-end; fix style names — S.

### upstream#214 — setting orientation fails with 0.8.5
**Verdict:** needs-investigation
**Ask summary:** Setting `section.orientation = WD_ORIENT.LANDSCAPE` on an added section does not flip page size; requires also swapping page_width/height.
**Evidence:** `src/docx/section.py:254` only toggles `@w:orient` on `w:pgSz` without swapping width/height. Python-pptx-style auto-swap not implemented.
**TODO (if applicable):** When orientation is changed, auto-swap `page_width`/`page_height` if mismatched — S.

### upstream#213 — Insert OMML equation XML into a paragraph
**Verdict:** resolved-in-loadfix
**Ask summary:** Ability to add user-authored OMML (`<m:oMath>`) as a paragraph/run.
**Evidence:** `src/docx/equations.py:39` `Equation` class; `Equation.from_omml_xml` referenced in equations.py:10 (HISTORY "Equation read + minimal create API (#113)").
**TODO (if applicable):** n/a.

### upstream#212 — Support for '.docm' OOXML format
**Verdict:** resolved-in-loadfix
**Ask summary:** Open and save macro-enabled `.docm` files (content type `...macroEnabled.main+xml`).
**Evidence:** HISTORY "D.24 .docm macro-enabled file support (#65)".
**TODO (if applicable):** n/a.

### upstream#211 — xpath cannot find <w:lastRenderedPageBreak/>
**Verdict:** resolved-in-loadfix
**Ask summary:** User cannot locate `w:lastRenderedPageBreak` via `d.element.xpath` — needed as rendered-page-break info.
**Evidence:** `src/docx/text/pagebreak.py` + `src/docx/oxml/text/pagebreak.py` expose `RenderedPageBreak`; Paragraph iter also surfaces them. (xpath miss was user namespace issue; the feature exists.)
**TODO (if applicable):** n/a.

### upstream#209 — autofit is broken (writes fixed tcW)
**Verdict:** resolved-in-loadfix
**Ask summary:** `table.autofit = True` should emit `w:tcW w:type="auto"` instead of `dxa`, but it was not honoring the setting.
**Evidence:** HISTORY "D.26 Table autofit and column-width control (#39)"; `src/docx/oxml/table.py:736` `autofit` getter/setter + column width logic.
**TODO (if applicable):** n/a.

### upstream#205 — Add row to existing table with formatting
**Verdict:** new-feature-needed
**Ask summary:** `table.add_row()` should be able to clone formatting (tc/trPr, run styling) from an existing template row.
**Evidence:** No `clone_row`/`copy_row`/`add_row_like` API in `src/docx/table.py`. Users currently must XML-deepcopy manually.
**TODO (if applicable):** Add `Table.add_row(source=row)` or `_Row.clone()` helper that deepcopies `w:tr` structure — M.

### upstream#203 — Find and replace text in footer
**Verdict:** resolved-in-loadfix
**Ask summary:** Locate and replace text occurrences inside header/footer stories.
**Evidence:** HISTORY "Document.search_regex / replace_regex / search_all / replace_all (#153, #154)"; `src/docx/document.py:930` `search_all`, `replace_all`, etc. walk headers/footers.
**TODO (if applicable):** n/a.

### upstream#201 — Ability to manipulate table borders
**Verdict:** resolved-in-loadfix
**Ask summary:** Read and set `w:tblBorders` / `w:tcBorders` via the Document API.
**Evidence:** `src/docx/table.py:305` `Table.borders`, line 1067 `TableBorders`, line 1258 `CellBorders`, line 339 `Table.set_borders` (HISTORY "Table.borders / _Cell.borders (#102)").
**TODO (if applicable):** n/a.

### upstream#198 — Quick-start example doesn't run
**Verdict:** needs-investigation
**Ask summary:** Readthedocs quick-start snippet references style names and casing that don't match built-ins (IntenseQuote vs "Intense Quote"), so copy-paste fails.
**Evidence:** `docs/user/quickstart.rst` — may need review; loadfix inherited upstream docs. Same root cause as #215.
**TODO (if applicable):** Audit and correct quick-start code snippets — S.

## Batch summary
- resolved-in-loadfix: 14 (#246, #245, #243, #235, #232, #224, #221, #217, #213, #212, #211, #209, #203, #201)
- new-feature-needed: 4 (#252, #249, #248, #205)
- new-bug-needed: 1 (#225)
- needs-investigation: 5 (#241, #220, #215, #214, #198)
- out-of-scope: 1 (#223)
Total: 25

### upstream#197 — How to use styles of another doc file?
**Verdict:** needs-investigation
**Ask summary:** User wants to apply a style from one document's styles collection to a table in another document; current `style_from_temp` reference doesn't work.
**Evidence:** no match; Style import/copy across documents not implemented. (loadfix#162 "Style mapping" is in-document linking, not cross-document.)
**TODO (if applicable):** Add `Document.import_styles(other)` to copy style definitions across docs — M.

### upstream#193 — picture format: EMF
**Verdict:** new-feature-needed
**Ask summary:** Support inserting `.emf` (Enhanced Metafile) images; currently raises `UnrecognizedImageError`.
**Evidence:** no match; `src/docx/image/` has bmp/gif/jpeg/png/tiff/svg but no emf.py.
**TODO (if applicable):** Add EMF/WMF header parser module + MIME/content-type registration — M.

### upstream#192 — feature: InlineShape.replace_image()
**Verdict:** new-feature-needed
**Ask summary:** Allow replacing the image blob behind an existing `InlineShape`/picture while preserving position and sizing.
**Evidence:** no match for `replace_image`/`replace_picture` anywhere in src/docx/.
**TODO (if applicable):** Add `InlineShape.replace_image(path_or_stream)` (and floating equiv) that swaps the image part relationship — M.

### upstream#190 — feature: Table.insert_row()
**Verdict:** new-feature-needed
**Ask summary:** Insert a row at an arbitrary index in an existing table (not just append).
**Evidence:** no match for `insert_row`; only `Table.add_row()` appends.
**TODO (if applicable):** Add `Table.insert_row(index)` and/or `_Row.insert_row_before/after()` — S.

### upstream#187 — add_picture
**Verdict:** needs-investigation
**Ask summary:** `document.add_picture('IMG_*.jpg')` intermittently fails on an iOS-produced JPEG (body truncated but ties to image stream parsing).
**Evidence:** likely same parser root-cause as #184; no specific fix landed.
**TODO (if applicable):** Harden `image/jpeg.py` against non-standard iOS JPEG APP1/EXIF segments — M.

### upstream#184 — EXIF Data Parsing Error
**Verdict:** new-bug-needed
**Ask summary:** iOS 8.3 JPEGs cause `Unexpected EOF` in `image/helpers.py`; parser doesn't respect EXIF inline-value rule (value stored in value_offset when <=4 bytes).
**Evidence:** `src/docx/image/jpeg.py` and `helpers.py` unchanged vs upstream on this code path.
**TODO (if applicable):** Fix EXIF tag reader to handle inline value vs offset semantics per TIFF spec — S.

### upstream#182 — Copy paragraphs elements from one document to another
**Verdict:** needs-investigation
**Ask summary:** User wants to reorder/copy paragraph objects into a new document; relationships (styles, images, numbering) need translation.
**Evidence:** no cross-document copy helper in src/docx/.
**TODO (if applicable):** Add `Document.append_paragraph(para)` / `copy_from(other)` that rewrites rIds and re-homes parts — L.

### upstream#181 — Feature: Section.paragraphs
**Verdict:** new-feature-needed
**Ask summary:** Expose the paragraphs contained within a given `Section` (bounded by its `sectPr`).
**Evidence:** `src/docx/section.py` has no `.paragraphs` iterator; only header/footer paragraphs.
**TODO (if applicable):** Add `Section.paragraphs` iterating body paragraphs whose section boundary is this section — S.

### upstream#180 — Getting list numbering
**Verdict:** resolved-in-loadfix
**Ask summary:** Ability to resolve `numId` -> `abstractNum` -> level info (start, format) for a paragraph.
**Evidence:** `src/docx/numbering.py` plus `Paragraph.list_format` (paragraph.py:969) expose `NumberingDefinition`/`Level`.
**TODO (if applicable):** n/a.

### upstream#179 — Feature: Add Charts to Word documents
**Verdict:** resolved-in-loadfix
**Ask summary:** Allow adding charts (DrawingML) to Word documents.
**Evidence:** `Document.add_chart()` at `src/docx/document.py:186`; HISTORY "Charts read + add_chart() (#111)".
**TODO (if applicable):** n/a.

### upstream#176 — Can't find default.docx when frozen (library.zip)
**Verdict:** needs-investigation
**Ask summary:** When packaged via cx_freeze, `_default_docx_path` fails to resolve templates inside `library.zip`.
**Evidence:** no match for cx_freeze/library.zip handling; still reads from filesystem path.
**TODO (if applicable):** Use `importlib.resources` / package loader to read template bytes regardless of packaging — S.

### upstream#174 — Large Table Creation Slow
**Verdict:** needs-investigation
**Ask summary:** Adding many rows to a table has O(n^2) behaviour — 1000 rows takes minutes.
**Evidence:** no perf-focused changes to `Table.add_row` in loadfix; still iterates `tblGrid.gridCol_lst` per row.
**TODO (if applicable):** Profile `Table.add_row` / reduce repeated XPath; consider bulk `Table.add_rows(n)` — M.

### upstream#170 — docs: update styles to indicate exception raised on style not defined
**Verdict:** new-feature-needed
**Ask summary:** Documentation should note that using an undefined style name raises `KeyError`.
**Evidence:** no doc note added in docs/user/styles*.rst.
**TODO (if applicable):** Add docs note to styles-using.rst / add_paragraph style param — S.

### upstream#169 — nested table merge cell error
**Verdict:** new-bug-needed
**Ask summary:** Merging cells in a table nested inside another table throws `ValueError` in `_grid_col` (`tc_lst.index`).
**Evidence:** no fix for nested-table grid-offset resolution; `src/docx/oxml/table.py` still uses parent-tr assumption.
**TODO (if applicable):** Fix grid-col resolution to walk only the enclosing `w:tr`, not ancestor tables — M.

### upstream#167 — feature: columnar layout in Section
**Verdict:** resolved-in-loadfix
**Ask summary:** Support newsletter-style multi-column body layout per section.
**Evidence:** `Section.columns` / `SectionColumns` / `Column` at `src/docx/section.py:60,880,918`; HISTORY "D.19 Multi-column section layout (#60)".
**TODO (if applicable):** n/a.

### upstream#166 — Documentation error: table.cells should be table.cell
**Verdict:** new-bug-needed
**Ask summary:** Docs page dev/analysis/features/cell-merge.html uses `table.cells(0,0)` instead of `table.cell(0,0)`.
**Evidence:** doc typo; content in docs/dev/analysis still needs verification.
**TODO (if applicable):** Fix code sample in cell-merge analysis doc — S.

### upstream#165 — _tr on cells
**Verdict:** needs-investigation
**Ask summary:** Legacy `_tr` attribute formerly exposed on cells is no longer present; users want it back for custom XML manipulation.
**Evidence:** `_Cell` in src/docx/table.py exposes `_tc` but not `_tr`; no restoration in loadfix.
**TODO (if applicable):** Either re-add `_Cell._tr` convenience or document the supported replacement — S.

### upstream#164 — docs: Update front-page example for table style
**Verdict:** new-bug-needed
**Ask summary:** Front-page README/quickstart table example no longer matches current default-table rendering.
**Evidence:** README/quickstart docs unchanged for this example.
**TODO (if applicable):** Update quickstart.rst table example screenshot/style name — S.

### upstream#163 — feature: cell alignment
**Verdict:** resolved-in-loadfix
**Ask summary:** Control cell vertical alignment (Center / Top / Bottom).
**Evidence:** `_Cell.vertical_alignment` (src/docx/table.py:856) with `WD_CELL_VERTICAL_ALIGNMENT`.
**TODO (if applicable):** n/a.

### upstream#161 — feature: Column.width sets cell widths
**Verdict:** resolved-in-loadfix
**Ask summary:** Assigning `Column.width` should propagate the width to every cell in that column (skipping merged cells).
**Evidence:** `_Column.width` setter at `src/docx/table.py:1431-1488` propagates to per-row `tcW`, skips `gridSpan>1`; HISTORY "D.26 Table autofit and column-width control (#39)".
**TODO (if applicable):** n/a.

### upstream#159 — feature: floating image
**Verdict:** resolved-in-loadfix
**Ask summary:** Support floating images with text wrapping (`wp:anchor`).
**Evidence:** HISTORY "D.17 Floating images with wp:anchor positioning (#30)"; `src/docx/drawing/__init__.py` handles `wp:anchor`.
**TODO (if applicable):** n/a.

### upstream#156 — How to insert table in the middle of document
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert a table at an arbitrary position (not just the document end).
**Evidence:** `Paragraph.insert_table_before/after` (paragraph.py:865,892) and `Table.insert_table_before/after` (table.py:140,164); HISTORY D.13.
**TODO (if applicable):** n/a.

### upstream#155 — Read text inside <w:sdt> tag
**Verdict:** resolved-in-loadfix
**Ask summary:** Read and modify text contained inside structured document tags (content controls).
**Evidence:** `src/docx/content_controls.py` + `src/docx/oxml/content_controls.py`; HISTORY "D.14 Content controls (SDTs) (#27)".
**TODO (if applicable):** n/a.

### upstream#154 — feature: Font.name_far_east
**Verdict:** resolved-in-loadfix
**Ask summary:** Set an East-Asian typeface (w:rFonts/@w:eastAsia) distinct from Latin font name.
**Evidence:** `Font.name_far_east` getter/setter at `src/docx/text/font.py:596-610` (plus alias).
**TODO (if applicable):** n/a.

### upstream#150 — feature: Run.shading
**Verdict:** resolved-in-loadfix
**Ask summary:** Read/set run-level background shading colour (w:shd on rPr).
**Evidence:** `Font.shading_color` at `src/docx/text/font.py:700-725`; HISTORY "D.20 Font.shading (#33)".
**TODO (if applicable):** n/a.

## Batch summary
- resolved-in-loadfix: 9 (150, 154, 155, 156, 159, 161, 163, 167, 179, 180) — actually 10
- new-feature-needed: 5 (170, 181, 190, 192, 193)
- new-bug-needed: 4 (164, 166, 169, 184)
- needs-investigation: 6 (165, 174, 176, 182, 187, 197)
- out-of-scope: 0

Corrected counts (25 total):
- resolved-in-loadfix: 10
- new-feature-needed: 5
- new-bug-needed: 4
- needs-investigation: 6
- out-of-scope: 0

### upstream#146 — feature: Cell.shading
**Verdict:** resolved-in-loadfix
**Ask summary:** Add cell-shading (background color/pattern) to table cells.
**Evidence:** `src/docx/table.py:770` `_Cell.shading` returns `CellShading`; Phase D.6 (#63) in HISTORY.rst.
**TODO (if applicable):** n/a

### upstream#145 — feature: cell borders
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose per-cell borders (incl. style variants like double line).
**Evidence:** `src/docx/table.py:1261` `_Cell.borders`; HISTORY notes `Table.borders / _Cell.borders (#102)`.
**TODO (if applicable):** n/a

### upstream#138 — feature: Support Captions
**Verdict:** resolved-in-loadfix
**Ask summary:** Add captions for figures/tables using Caption style + SEQ field for later ToC use.
**Evidence:** `src/docx/captions.py` builds `Caption` paragraphs with SEQ `w:fldSimple`; HISTORY `caption helpers (#141)`.
**TODO (if applicable):** n/a

### upstream#128 — example does not work: 'recordset' undefined, image not available
**Verdict:** out-of-scope
**Ask summary:** Upstream readthedocs landing-page example references undefined `recordset` and missing image.
**Evidence:** Problem is on upstream docs site; loadfix quickstart at `docs/user/quickstart.rst` uses runnable examples.
**TODO (if applicable):** n/a

### upstream#127 — feature: Document.variables
**Verdict:** new-feature-needed
**Ask summary:** Read/update `w:docVars` (docVar name/value pairs in `settings.xml`).
**Evidence:** `src/docx/oxml/settings.py:428` only lists `w:docVars` as a successor; no proxy or accessor.
**TODO (if applicable):** Add `Settings.doc_vars` (Mapping proxy) with read/write via CT_DocVars. S

### upstream#122 — how to create nested list?
**Verdict:** resolved-in-loadfix
**Ask summary:** Create multi-level (nested) numbered/bulleted lists.
**Evidence:** `src/docx/numbering.py:262` `AbstractNum.level(ilvl)`; `Paragraph.list_format` at `text/paragraph.py:969`; Phase D.9 (#22).
**TODO (if applicable):** n/a

### upstream#115 — feature: pgNum field
**Verdict:** resolved-in-loadfix
**Ask summary:** Support PAGE (pgNum) field insertion.
**Evidence:** `src/docx/fields.py:48` defines `WD_FIELD_TYPE.PAGE`; `Paragraph.add_simple_field` at `text/paragraph.py:254`.
**TODO (if applicable):** n/a

### upstream#112 — feature: _Row.index and _Column.index
**Verdict:** new-feature-needed
**Ask summary:** Expose public `index` property on `_Row` and `_Column`.
**Evidence:** `src/docx/table.py:1491,1709` still only private `_index`; no public `.index` property.
**TODO (if applicable):** Rename/alias `_index` to public `index` on `_Row` and `_Column`. S

### upstream#109 — feature: process bookmarks
**Verdict:** resolved-in-loadfix
**Ask summary:** Recognize, search, delete, and replace text within Word bookmarks.
**Evidence:** `src/docx/bookmarks.py` (`Bookmarks`, `Bookmark.delete`, name/get); Phase C (#52) in HISTORY.
**TODO (if applicable):** n/a

### upstream#105 — feature: horz rule (paragraph border)
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert a horizontal rule via paragraph bottom border.
**Evidence:** `ParagraphBorders` at `src/docx/text/parfmt.py:467`; Phase D.7 (#109) in HISTORY.
**TODO (if applicable):** n/a

### upstream#100 — feature: add formula
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert mathematical formulas (OMML), optionally from TeX.
**Evidence:** `src/docx/equations.py` provides `Equation.from_omml_xml` + builder helpers; HISTORY `Equation read + minimal create API (#113)`. LaTeX import explicitly out of scope.
**TODO (if applicable):** n/a

### upstream#99 — feature: replace scalar placeholders in template
**Verdict:** resolved-in-loadfix
**Ask summary:** Provide `replace(search, replace)` on documents that preserves formatting (incl. tables).
**Evidence:** `src/docx/search.py` + `Document.search_regex/replace_regex/search_all/replace_all` (HISTORY #153, #154); Phase D.10 (#91).
**TODO (if applicable):** n/a

### upstream#97 — feature: add cross-reference
**Verdict:** resolved-in-loadfix
**Ask summary:** Insert/resolve cross-references to paragraphs, titles, and page numbers.
**Evidence:** `src/docx/fields.py` REF/PAGEREF resolution (Phase C #115); `Paragraph.add_simple_field` + `add_bookmark` at `text/paragraph.py:56,254`.
**TODO (if applicable):** n/a

### upstream#96 — feature: Row.height
**Verdict:** resolved-in-loadfix
**Ask summary:** Get/set table row height (`trHeight`).
**Evidence:** `_Row.height` getter+setter at `src/docx/table.py:1668,1688`; Phase D.15 (#28).
**TODO (if applicable):** n/a

### upstream#92 — feature: Paragraph.list_format
**Verdict:** resolved-in-loadfix
**Ask summary:** Detect that a paragraph is a bullet/numbered item and retrieve its numbering info.
**Evidence:** `Paragraph.list_format` at `src/docx/text/paragraph.py:969` returns numbering definition + level; Phase D.9 (#22).
**TODO (if applicable):** n/a

### upstream#91 — feature: add custom properties support
**Verdict:** resolved-in-loadfix
**Ask summary:** Read/write/add/delete custom document properties.
**Evidence:** `src/docx/custom_properties.py` supports str/int/float/bool/datetime; Phase D.4 (#14).
**TODO (if applicable):** n/a

### upstream#86 — How to change the default font size in document
**Verdict:** needs-investigation
**Ask summary:** Support changing the document-default font size (rPrDefault in docDefaults).
**Evidence:** No match for `docDefaults`/`default_font` in `src/docx/styles/`; Normal style font editing works but not docDefaults directly.
**TODO (if applicable):** Investigate adding `Styles.default_font`/`Styles.default_paragraph_format` proxy for `w:docDefaults`. S

### upstream#83 — feature: Row.delete()
**Verdict:** resolved-in-loadfix
**Ask summary:** Remove a row from a table.
**Evidence:** `src/docx/table.py:62` `delete()` covers Row; HISTORY `Paragraph.delete / Run.delete / Table.delete (#50)`.
**TODO (if applicable):** n/a

### upstream#74 — feature: Paragraph.add_hyperlink()
**Verdict:** resolved-in-loadfix
**Ask summary:** Add external/anchor hyperlinks to a paragraph.
**Evidence:** `Paragraph.add_hyperlink` at `src/docx/text/paragraph.py:162`; Phase D.1 (#97).
**TODO (if applicable):** n/a

### upstream#72 — feature: Document.text
**Verdict:** needs-investigation
**Ask summary:** Add `Document.text` as a convenience to extract all paragraph text joined by newlines.
**Evidence:** No `Document.text` property found in `src/docx/document.py`; users must still iterate paragraphs.
**TODO (if applicable):** Add `Document.text` property returning body text (body stories only). S

### upstream#50 — docs: Working with Tables user guide page
**Verdict:** resolved-in-loadfix
**Ask summary:** Documentation page explaining reading/iterating existing tables.
**Evidence:** `docs/user/tables.rst` and `docs/user/tables-advanced.rst` exist in loadfix.
**TODO (if applicable):** n/a

### upstream#49 — Markdown Conversion
**Verdict:** out-of-scope
**Ask summary:** Support converting markdown directly into docx content.
**Evidence:** No markdown handling in `src/docx/`; upstream position (and loadfix) is that this belongs in a separate lib (pandoc etc.).
**TODO (if applicable):** n/a

### upstream#44 — features: merge/concatenate two documents
**Verdict:** new-feature-needed
**Ask summary:** Merge/concatenate two `.docx` files preserving images, styles, paragraphs.
**Evidence:** No merge/append helper present; grep for `merge_document|append_document` empty in `src/docx/`.
**TODO (if applicable):** Add `Document.append_document(other)` that imports body + relationships. L

### upstream#33 — feature: Paragraph.delete()
**Verdict:** resolved-in-loadfix
**Ask summary:** Delete a paragraph (handling last-reference cleanup for hyperlink/picture rels).
**Evidence:** `Paragraph.delete` at `src/docx/text/paragraph.py:577`; HISTORY `Paragraph.delete / Run.delete / Table.delete (#50)`.
**TODO (if applicable):** n/a

### upstream#31 — feature: add field
**Verdict:** resolved-in-loadfix
**Ask summary:** Generic support for inserting/adding WordprocessingML field codes.
**Evidence:** `src/docx/fields.py` + `Paragraph.add_simple_field` at `text/paragraph.py:254`; Phase C (#10).
**TODO (if applicable):** n/a

## Batch summary
- resolved-in-loadfix: 17
- new-feature-needed: 3
- new-bug-needed: 0
- needs-investigation: 2
- out-of-scope: 3

### upstream#30 — Can search and replace functions be added to python-docx?
**Verdict:** resolved-in-loadfix
**Ask summary:** Request for search / replace utilities over document body text.
**Evidence:** `src/docx/search.py` (search_paragraphs, replace_in_paragraphs, regex variants); HISTORY "D.10 Search and replace with formatting preservation (#91)" and `Document.search_regex / replace_regex / search_all / replace_all (#153, #154)`.

### upstream#27 — Adding elements nonsequentially
**Verdict:** resolved-in-loadfix
**Ask summary:** Request for insert-at-index and delete APIs for paragraphs / runs beyond simple append.
**Evidence:** HISTORY "D.13 Insert paragraph / table at arbitrary position (#26)"; `Paragraph.insert_paragraph_before/after`, `insert_table_before/after`, `Paragraph.delete()` at `src/docx/text/paragraph.py:577,743,758,865,892`.

### upstream#25 — Restart numbering of an ordered list
**Verdict:** new-feature-needed
**Ask summary:** User wants to restart numbering on an existing numbered list mid-document (new list instance reusing same format, or w:lvlOverride with startOverride).
**Evidence:** `add_numbering_definition` supports `start` per-level and creates a fresh w:num; `docs/user/numbering.rst:164-167` explicitly lists `w:lvlOverride` startOverride as "not exposed" and directs users to the `.element` escape hatch. No `Paragraph.restart_numbering()` / `NumberingDefinition.new_instance()` convenience.
**TODO (if applicable):** Add `NumberingDefinition.new_instance()` (fresh w:num) and/or `Paragraph.restart_numbering(level=0, start=1)` emitting w:lvlOverride/startOverride. S.

### upstream#24 — can't add emf, or eps images
**Verdict:** new-feature-needed
**Ask summary:** Support inserting EMF (and EPS) images.
**Evidence:** `src/docx/image/` has bmp/gif/jpeg/png/svg/tiff only; `src/docx/image/constants.py` has no EMF/EPS content type. SVG was added (D.22) but EMF/EPS were not.
**TODO (if applicable):** Add EMF header parser + content type `image/x-emf` (and optionally EPS `application/postscript`) to image subsystem. M.

### upstream#10 — feature: _Cell.add_picture()
**Verdict:** new-feature-needed
**Ask summary:** Shortcut to add a picture directly to a table cell (analogous to Document.add_picture).
**Evidence:** `src/docx/table.py:530` `_Cell(BlockItemContainer)` has no `add_picture`. Workaround is `cell.paragraphs[0].add_run().add_picture()`. `Document.add_picture` exists at `src/docx/document.py:238`, `Run.add_picture` at `src/docx/text/run.py:62`.
**TODO (if applicable):** Add `_Cell.add_picture(image_path_or_stream, width, height)` delegating to inner run. S.

### upstream#8 — feature: Paragraph.add_text()
**Verdict:** new-feature-needed
**Ask summary:** Convenience method to append text to the last run of a paragraph (preserving formatting), instead of `p.runs[-1].add_text()`.
**Evidence:** No `Paragraph.add_text` in `src/docx/text/paragraph.py` (only `add_text_form_field` and `add_run`); `Run.add_text` exists at `src/docx/text/run.py:91`.
**TODO (if applicable):** Add `Paragraph.add_text(text)` that appends to last run or creates a new one if none. S.

### upstream#7 — feature: document lxml element of proxy classes for advanced users
**Verdict:** resolved-in-loadfix
**Ask summary:** Documentation for accessing underlying lxml elements (_element / _p / _r / _tc) on proxy classes.
**Evidence:** `docs/user/api-concepts.rst:72-113` explicitly describes proxy `_element` (and `_p`, `_r`, `_tc`, etc.) as the escape hatch for advanced callers; also referenced from `docs/user/equations.rst:82` and `docs/user/numbering.rst:167`.

### upstream#6 — Document.paragraphs includes paragraphs nested in <w:ins>
**Verdict:** new-feature-needed
**Ask summary:** `Document.paragraphs` should reflect the "Final" view — include paragraphs inside w:ins, exclude w:del, reflect moves.
**Evidence:** `src/docx/document.py:719-726` explicitly documents that `Document.paragraphs` does NOT include paragraphs inside `w:ins` / `w:del`. Phase B adds read / accept / reject of tracked changes, but no "final view" virtual iterator over body paragraphs.
**TODO (if applicable):** Add `Document.final_paragraphs` (or option on `paragraphs`) that flattens w:ins children and hides w:del children without mutating the tree. M.

### upstream#1 — Support for footnote
**Verdict:** resolved-in-loadfix
**Ask summary:** Add footnote support (document.xml reference + footnotes.xml content).
**Evidence:** Phase A in HISTORY: "Add Document.footnotes and Footnotes / Footnote / FootnoteProperties (#1, #3, #17, #46, #48, #56, #82)"; `src/docx/footnotes.py`, `src/docx/parts/footnotes.py`, `docs/user/footnotes.rst`.

## Batch summary
- resolved-in-loadfix: 4 (#30, #27, #7, #1)
- new-feature-needed: 5 (#25, #24, #10, #8, #6)
- new-bug-needed: 0
- needs-investigation: 0
- out-of-scope: 0
- Total: 9

---

## Part 2 — Open Pull Requests (125)

### upstream-PR#1545 — Add CLI-Anything harness for python-docx

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Adds a standalone `agent-harness/` directory containing a Click-based CLI wrapper and tests for python-docx (1608 LOC).

**Evidence:**
- Adds a top-level `agent-harness/` tree — an external tool, not library changes; nothing touches `src/docx/`.
- Fork already has its own AI-agent CI pipeline (HISTORY "AI-agent CI pipeline").

**TODO (if applicable):** n/a.

### upstream-PR#1538 — feat: native support for tracked changes (w:ins, w:del) reading

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** Makes `Paragraph.text` include runs inside `w:ins`/`w:moveTo`, and registers `w:delText` as `CT_Text` so provenance tooling can read deletions.

**Evidence:**
- Loadfix `src/docx/oxml/text/paragraph.py:267` text still only walks `w:r | w:hyperlink | w:fldSimple | w:sdt` — ins/moveTo not included.
- Fork supplies `revision_marks_text()` (`src/docx/text/paragraph.py:1185`) and `accept_all_changes` as the intended API — behaviour divergence from upstream PR is a design choice worth confirming before changing `.text`.

**TODO (if applicable):** Decide whether `Paragraph.text` should reflect w:ins/w:moveTo by default vs. requiring `accept_all_changes()`. S

### upstream-PR#1537 — Feature: Add support for .dotx Word templates (#1532)

**Verdict:** new-feature-needed

**Ask summary (1-2 sentences):** Treat `.dotx` template files as valid inputs to `Document()` by adding `CT.WML_TEMPLATE_MAIN` and accepting that content type.

**Evidence:**
- `grep WML_TEMPLATE_MAIN src/docx/opc/constants.py` returns 0; fork accepts `WML_DOCUMENT_MAIN` and `WML_DOCUMENT_MACRO` (`.docm`) only (`src/docx/api.py:38`).
- No corresponding HISTORY entry; D.24 covers .docm but not .dotx.

**TODO (if applicable):** Add `CT.WML_TEMPLATE_MAIN`, register DocumentPart factory, accept it in `Document()`. S

### upstream-PR#1536 — Fix: Replace deprecated delimitedList with DelimitedList

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** Replaces deprecated `pyparsing.delimitedList` with `DelimitedList` in `tests/unitutil/cxml.py` to silence warnings on newer pyparsing.

**Evidence:**
- Test util `tests/unitutil/cxml.py` still uses `delimitedList` (confirmed by PR diff; no equivalent loadfix commit in HISTORY).
- Trivial hygiene fix, affects nothing but warning output.

**TODO (if applicable):** Swap `delimitedList` → `DelimitedList` in `tests/unitutil/cxml.py`. S

### upstream-PR#1534 — Add revision management capability + tests + docs

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** Adds write-side tracked-change API (`add_run_tracked`, `delete_tracked`, `replace_tracked`, `find_and_replace_tracked`, `settings.track_revisions`, `TrackedInsertion`/`TrackedDeletion` proxies with `accept/reject`).

**Evidence:**
- Loadfix Phase B covers this: `src/docx/tracked_changes.py` (`TrackedChange`, `_resolve_all_changes`), `Document.accept_all_changes`/`reject_all_changes` (`document.py:519,534`), `Settings.track_revisions` (`settings.py:336`).
- HISTORY Phase B: "accept / reject tracked changes (#7)", "read of tracked insertions and deletions (#53)", "move revisions (#134)".

**TODO (if applicable):** n/a.

### upstream-PR#1530 — Add full alt-text (title and description) support for pictures

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** Expose `title`/`descr` attributes on `wp:docPr` via `alt_text`/`alt_title` on InlineShape and through `add_picture(..., title=, descr=)` pass-throughs on Document/Run/StoryPart.

**Evidence:**
- InlineShape has `alt_text` and `title` already (`src/docx/shape.py:106,121`); FloatingImage also has both (`shape.py:252,272`). HISTORY: "alt_text / title on InlineShape and FloatingImage (#158)".
- BUT `Run.add_picture` / `Document.add_picture` signatures (`run.py:62`, `document.py:238`) don't accept `title`/`descr` at creation time. Workaround: set after the fact via the property.

**TODO (if applicable):** Add `title`/`descr` kwargs to `add_picture()` / `new_pic_inline()` to set at creation. S

### upstream-PR#1529 — feat(table): add Table.width property

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** Adds read/write `Table.width` property mapping to `w:tblPr/w:tblW` preferred width in dxa.

**Evidence:**
- Loadfix exposes this as `Table.preferred_width` (`src/docx/table.py:287`, setter at 301) backed by `CT_TblPr.preferred_width` (`oxml/table.py:750`). Different name, same semantic. Phase D.26 "Table autofit and column-width control (#39)".

**TODO (if applicable):** n/a (alternative implementation; consider alias if upstream naming matters).

### upstream-PR#1526 — Fix ValueError while accessing `cell._tc.bottom`

**Verdict:** needs-investigation

**Ask summary (1-2 sentences):** Rewrites `CT_Tr.tc_at_grid_offset` using a per-grid-cell dict so that grid-offsets landing mid-span still resolve to the correct `w:tc`, fixing #1458.

**Evidence:**
- Loadfix still has the original "remaining_offset" walk (`src/docx/oxml/table.py:360`) that raises `ValueError` when `grid_offset` falls inside a spanned cell. No HISTORY entry for #1458 fix.
- Needs a repro-docx check — the PR's behavioural change (return spanned tc for interior offsets) differs semantically from current (exact-start only).

**TODO (if applicable):** Reproduce #1458 with the attached test.docx; decide whether to accept spanned-hit semantics. M

### upstream-PR#1524 — feat: Add banded columns/rows attributes to table

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** Register `CT_TblLook` and expose `banded_columns`/`banded_rows` on `Table`.

**Evidence:**
- Loadfix already has `TableStyleFlags` (`src/docx/table.py:1158`) wiring `w:tblLook` with `first_row`, `last_row`, `first_column`, `last_column`, `no_horizontal_banding`, `no_vertical_banding` flags; `CT_TblLook` registered (`oxml/__init__.py:414`). HISTORY: "Table.style_flags (#144)".

**TODO (if applicable):** n/a (richer implementation already present).

### upstream-PR#1518 — Refactor parse_xml to handle external target references (#1349)

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** In `CT_Relationship.target_mode`, treat targets starting with `#` (internal bookmark hrefs) as `RTM.EXTERNAL` to avoid downstream misclassification.

**Evidence:**
- Loadfix `src/docx/opc/oxml.py:175` still does `return self.get("TargetMode", RTM.INTERNAL)` with no `#`-prefix special-case. Long-standing issue #1349/PR #1350 never landed.
- Partially overlaps with PR #1498 (skip internal/null on load) but PR #1518 addresses downstream classification.

**TODO (if applicable):** Evaluate skip-on-read (PR #1498) vs classify-as-external (PR #1518); pick one or both. S

### upstream-PR#1514 — Update documents.rst (StringIO → BytesIO)

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** Docs fix — example code uses `StringIO` where `BytesIO` is correct for docx byte streams.

**Evidence:**
- `docs/user/documents.rst` example still references `StringIO` (per PR diff context); pure documentation fix, risk zero.

**TODO (if applicable):** Replace `StringIO` with `BytesIO` in `docs/user/documents.rst`. S

### upstream-PR#1505 — fix: fallback to 72 dpi when image header is 0 (#1494)

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** Prevent `ZeroDivisionError` in `Image.width/height` when the decoded image header reports `dpi = 0` by falling back to 72.

**Evidence:**
- Loadfix `src/docx/image/image.py:89-102` still forwards `self._image_header.horz_dpi` unconditionally; division at lines 108/114 would divide by 0. No HISTORY entry for #1494/#1497.

**TODO (if applicable):** Guard `horz_dpi`/`vert_dpi` to return 72 when the header reports None or 0. S

### upstream-PR#1501 — Update Python classifiers

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Bumps `pyproject.toml` trove classifiers (3 lines).

**Evidence:**
- Build/packaging churn; fork manages its own classifiers via its release cadence (1.3.0.dev0 skeleton).

**TODO (if applicable):** n/a.

### upstream-PR#1498 — Skip null and internal bookmarks links when loading

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** In `_SerializedRelationships.load_from_xml`, skip `target_ref` values of `NULL`/`../NULL` and internal bookmark hrefs (`#...`) so broken-but-openable docs load.

**Evidence:**
- Loadfix `src/docx/opc/pkgreader.py` has no such skip; `grep null|starts.*#` in that file returns 0.
- Fork has `recover=True` but that's XML recovery, not rel-filtering.

**TODO (if applicable):** Add NULL/internal-anchor rel filtering in `pkgreader.load_from_xml`. S

### upstream-PR#1480 — Fix #1454 (NumberingPart.new NotImplementedError)

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** Replace `raise NotImplementedError` in `NumberingPart.new()` with an actual minimal-XML creator.

**Evidence:**
- Loadfix `src/docx/parts/numbering.py:27-34` implements `NumberingPart.new()` and `default()` returning a real part with empty `<w:numbering>` template; Phase D.9 "Numbering style control (#22)".

**TODO (if applicable):** n/a.

### upstream-PR#1478 — resolve issue #1475 Non-integer font sizes

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** Change `ST_HpsMeasure.convert_from_xml` to `float(str_value)` so runs with fractional half-point sizes (e.g. `"36.5625..."`) don't raise `ValueError`.

**Evidence:**
- Loadfix `src/docx/oxml/simpletypes.py:253-255` still does `Pt(int(str_value))` — would raise on non-integer half-points. No HISTORY entry for #1475.

**TODO (if applicable):** Tolerate float in `ST_HpsMeasure.convert_from_xml` (use `float(str_value)` with rounding note). S

### upstream-PR#1471 — Credit the original author

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Adds attribution lines to upstream `README.md` — not a code change.

**Evidence:**
- Purely an upstream repo hygiene/contributor-credit matter; not applicable to the loadfix fork's README.

**TODO (if applicable):** n/a.

### upstream-PR#1451 — Fix Ascii IFD reading as TIFF6.0 spec

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** In `_AsciiIfdEntry._parse_value`, when the ASCII value fits in ≤4 bytes the value is inlined into the 4-byte Value/Offset slot (left-justified), not pointed to. Current code always dereferences.

**Evidence:**
- Loadfix `src/docx/image/tiff.py:233` (see `_AsciiIfdEntry._parse_value`) unconditionally does `stream_rdr.read_str(value_count - 1, value_offset)` — missing the ≤4 byte inline branch. Issue #187.

**TODO (if applicable):** Add `if value_count <= 4: read inline at offset+8` branch in `_AsciiIfdEntry._parse_value`. S

### upstream-PR#1446 — add element as a public property not _element which is private

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Adds an `element` alias that returns `self._element` on `Paragraph`, so IDEs don't flag private-access warnings.

**Evidence:**
- Fork uses `_element` as the canonical protected handle throughout; exposing a public alias on one proxy type (Paragraph only, per diff) is cosmetic and inconsistent. Low value vs. API-surface churn.

**TODO (if applicable):** n/a.

### upstream-PR#1437 — Font name processing for w:eastAsia and w:cs

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** Try to make `font.name` setter also write `w:eastAsia` and `w:cs` rFonts attributes via extra setter kwargs (signature is broken in the PR).

**Evidence:**
- Loadfix already exposes `Font.name_cs`, `Font.name_east_asia`, `Font.name_far_east` (`src/docx/text/font.py:562,581,596`) as first-class setters mapping to `rFonts_cs`/`rFonts_eastAsia`. Cleaner than upstream PR. HISTORY: "East Asian typography (#128)".

**TODO (if applicable):** n/a.

### upstream-PR#1436 — fix duplicate docProps/core.xml UserWarning

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** Handle the alternate `.../officedocument/2006/relationships/metadata/core-properties` rel-type by discovering the existing core-properties part before creating a new one (prevents duplicate `docProps/core.xml`).

**Evidence:**
- Loadfix `grep CORE_PROPERTIES_OFFICEDOCUMENT src/docx/opc/` returns 0; `Package._core_properties_part` only looks up `RT.CORE_PROPERTIES`. Alternate-namespace rel unhandled.

**TODO (if applicable):** Add `CORE_PROPERTIES_OFFICEDOCUMENT` reltype constant and fall-back lookup in `_core_properties_part`. S

### upstream-PR#1423 — Webp support

**Verdict:** new-feature-needed

**Ask summary (1-2 sentences):** Add a `Webp` image-header class (RIFF/WEBP parser) + content-type registration to let `add_picture` accept WebP files (closes #717).

**Evidence:**
- `src/docx/image/` contains bmp/gif/jpeg/png/svg/tiff only — no `webp.py`. HISTORY D.22 only covers SVG (#76).

**TODO (if applicable):** Add `image/webp.py` header class, register signature+content-type, add fixtures. M

### upstream-PR#1407 — Add support for word documents with macros (.docm)

**Verdict:** resolved-in-loadfix

**Ask summary (1-2 sentences):** Accept `WML_DOCUMENT_MACRO` content type in `Document()` so `.docm` files open without `ValueError`.

**Evidence:**
- Loadfix `src/docx/api.py:38` accepts `CT.WML_DOCUMENT_MACRO`, and `src/docx/__init__.py:62` registers `DocumentPart` for that content type. HISTORY D.24: ".docm macro-enabled file support (#65)".

**TODO (if applicable):** n/a.

### upstream-PR#1400 — Create Questions bank

**Verdict:** out-of-scope

**Ask summary (1-2 sentences):** Adds an unrelated binary-ish file named "Questions bank" at repo root (no body). Stale junk.

**Evidence:**
- Not a python-docx change; presumably a mis-targeted PR.

**TODO (if applicable):** n/a.

### upstream-PR#1395 — Update install.rst

**Verdict:** new-bug-needed

**Ask summary (1-2 sentences):** Grammar fix in `docs/user/install.rst` ("those" → "that" where referring to the single lxml package).

**Evidence:**
- Docs typo fix; near-zero risk. Fork tracks its own docs tree so can adopt.

**TODO (if applicable):** Apply the "those" → "that" grammar fix in `docs/user/install.rst`. S

## Batch summary

- resolved-in-loadfix: 6 (1534, 1529, 1524, 1480, 1437, 1407)
- new-feature-needed: 2 (1537, 1423)
- new-bug-needed: 9 (1536, 1514, 1505, 1498, 1478, 1451, 1436, 1395, 1518)
- needs-investigation: 3 (1538, 1530, 1526)
- out-of-scope: 5 (1545, 1501, 1471, 1446, 1400)
- Total: 25

### upstream-PR#1393 — add outline level set function
**Verdict:** new-feature-needed
**Ask summary:** Adds `WD_PARAGRAPH_OUTLINELVL` enum and `outline_level` setter on ParagraphFormat. Duplicate of PR#1295 from same author.
**Evidence:** no match; loadfix has `w:outlineLvl` element registered but no public API on `ParagraphFormat`.
**TODO:** Expose `ParagraphFormat.outline_level` get/set with `WD_OUTLINELVL` enum (S).

### upstream-PR#1392 — Feature footnote support
**Verdict:** resolved-in-loadfix
**Ask summary:** Read/write footnotes at document and paragraph level, with FootnotesPart, Footnote, FootnoteReference.
**Evidence:** `src/docx/footnotes.py` (Footnotes, Footnote, FootnoteProperties), `src/docx/parts/footnotes.py`; Phase C footnotes/endnotes implemented.

### upstream-PR#1378 — Add support for macro enabled template (.dotm)
**Verdict:** resolved-in-loadfix
**Ask summary:** Accept `WML_DOCUMENT_MACRO` content type in `Document()`.
**Evidence:** `src/docx/__init__.py:62` registers `CT.WML_DOCUMENT_MACRO` for DocumentPart; `src/docx/api.py:39` allows it.

### upstream-PR#1371 — Update Table Style Example in Quickstart
**Verdict:** out-of-scope
**Ask summary:** Docs-only tweak to `docs/user/quickstart.rst` table style example.
**Evidence:** docs-only upstream copy fix; not fork-relevant.

### upstream-PR#1361 — fix(docs): update sphinx
**Verdict:** out-of-scope
**Ask summary:** Sphinx tooling bump (docs/conf.py, requirements-docs.txt).
**Evidence:** build-churn; out of scope.

### upstream-PR#1355 — Fix issue #545 (pkgwriter order)
**Verdict:** needs-investigation
**Ask summary:** Reorders `PackageWriter.write()` to emit parts and rels before `[Content_Types].xml` (Word-friendly zip ordering).
**Evidence:** `src/docx/opc/pkgwriter.py:35-38` still has original order (content-types first); not applied.
**TODO:** Evaluate whether Word requires `[Content_Types].xml` last in zip for issue #545; add regression test (S).

### upstream-PR#1353 — Add module (root) reference
**Verdict:** out-of-scope
**Ask summary:** Docs index.rst one-line tweak.
**Evidence:** docs-only.

### upstream-PR#1350 — parse_xml external target refs (issue #1349)
**Verdict:** new-bug-needed
**Ask summary:** Detect `target_ref` starting with `#` as `RTM.EXTERNAL` to avoid choking on in-document hyperlink rels.
**Evidence:** `src/docx/opc/oxml.py:170-176` target_mode does not handle `#` prefix.
**TODO:** Patch `CT_Relationship.target_mode` to return EXTERNAL for `#`-prefixed refs with test (S).

### upstream-PR#1322 — bump jinja2 2.11.3 → 3.1.3
**Verdict:** out-of-scope
**Ask summary:** Dependabot on `requirements-docs.txt`.
**Evidence:** dep-bump; out of scope.

### upstream-PR#1310 — pyinstaller hdrftr path fix
**Verdict:** needs-investigation
**Ask summary:** Replace `os.path.join(split(__file__)[0], "..", "templates", ...)` with `dirname(split(__file__)[0])` for PyInstaller compatibility.
**Evidence:** `src/docx/parts/hdrftr.py:30,47` now uses `Path(__file__).parent.parent / "templates"`; may work under PyInstaller via importlib.resources pattern but not verified.
**TODO:** Confirm PyInstaller frozen-app works; consider `importlib.resources.files` for templates (S).

### upstream-PR#1309 — Enable copying of a run's font
**Verdict:** new-feature-needed
**Ask summary:** Adds `Font.color` setter and `Run.formatting` deepcopy helper for copying run formatting.
**Evidence:** `src/docx/text/font.py:19` Font class has no color setter; no formatting copy helper on Run.
**TODO:** Add `Font.color` setter and a clean `Run.copy_formatting_from()` (M).

### upstream-PR#1295 — add outline level set function
**Verdict:** new-feature-needed
**Ask summary:** Duplicate of #1393 — enum + outline_level property on ParagraphFormat.
**Evidence:** same as #1393.
**TODO:** (See #1393).

### upstream-PR#1273 — Add custom properties support
**Verdict:** resolved-in-loadfix
**Ask summary:** docProps/custom.xml read/write of custom document properties.
**Evidence:** `src/docx/custom_properties.py`, `src/docx/parts/custom_properties.py`, `Document.custom_properties` at document.py:591.

### upstream-PR#1272 — Handle large docx files (huge_tree=True)
**Verdict:** out-of-scope
**Ask summary:** Flip `huge_tree` to True in `oxml_parser` to accept very large docs.
**Evidence:** `src/docx/oxml/parser.py:26` intentionally keeps `huge_tree=False` plus a separate `_recovery_parser`; comment says "prevents XML bombs". Fork has a deliberate security-driven policy contrary to this PR.

### upstream-PR#1270 — settings: get/set trackRevisions
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose `Settings.track_revisions` backed by `w:trackRevisions`.
**Evidence:** `src/docx/settings.py:336-347` implements `track_revisions` getter/setter.

### upstream-PR#1220 — Fix descriptor `__dict__` error on dir(CT_*)
**Verdict:** needs-investigation
**Ask summary:** Change `MetaOxmlElement` base class so `dir()` on CT_ instances does not raise in Py3.
**Evidence:** `src/docx/oxml/xmlchemy.py` in loadfix was refactored (typed rewrite); PR's exact patch no longer applies. Need to re-test `dir(CT_P())` on current loadfix.
**TODO:** Verify `dir(CT_*)` works; add regression test if still broken (S).

### upstream-PR#1219 — Ignore rels to missing files like Word does
**Verdict:** new-bug-needed
**Ask summary:** Skip header/footer rels whose target parts are absent from the zip (1C: Enterprise docs) instead of raising KeyError.
**Evidence:** `src/docx/opc/pkgreader.py:79` calls `blob_for(partname)` without try/except; missing-part handling absent.
**TODO:** Wrap `blob_for` in try/except KeyError in `_walk_phys_parts` and drop the srel (S).

### upstream-PR#1206 — Implements extended-properties (app.xml)
**Verdict:** new-feature-needed
**Ask summary:** Read/write docProps/app.xml extended properties (Company, Manager, TotalTime, etc.).
**Evidence:** no match; `grep` finds only template app.xml and constants — no `extended_properties` proxy/part in loadfix.
**TODO:** Add `ExtendedPropertiesPart` + `Document.extended_properties` proxy; see PR#1273 pattern (M).

### upstream-PR#1205 — Fix media path backslash on Windows
**Verdict:** new-bug-needed
**Ask summary:** Normalise `\` to `/` in `_SerializedRelationship._target_ref` to handle docs with Windows-style rel targets.
**Evidence:** `src/docx/opc/pkgreader.py` `_SerializedRelationship` does not normalise; no backslash handling.
**TODO:** Normalise target_ref backslashes in serialized rels with test fixture (S).

### upstream-PR#1196 — Fix table cell indexing (issue #1193)
**Verdict:** needs-investigation
**Ask summary:** Change `Table.cell()`/`column_cells()`/`row_cells()` to use 2-D cell array instead of flat index that misbehaves with merged cells.
**Evidence:** `src/docx/table.py:372-404` still uses flat `cell_idx = col_idx + row_idx*col_count` logic. Loadfix has added other merged-cell helpers (`_iter_row_cells`) but core `cell()` unchanged.
**TODO:** Reproduce issue #1193 with merged-cell table; refactor `_cells` to 2-D if confirmed (M).

### upstream-PR#1174 — fix: AttValue length too long
**Verdict:** out-of-scope
**Ask summary:** Version bump (0.8.11→0.8.11.1) plus XML parser tweak in legacy `docx/` tree.
**Evidence:** targets deleted flat-layout path; stale WIP.

### upstream-PR#1168 — Add check for Path and convert to string
**Verdict:** new-bug-needed
**Ask summary:** Accept `pathlib.Path` in `Image.from_file` (currently only `str`/file-like).
**Evidence:** `src/docx/image/image.py:36-39` only checks `isinstance(image_descriptor, str)`; Path falls through to the stream branch and crashes on `stream.seek`.
**TODO:** Accept `str | os.PathLike | IO[bytes]` in `Image.from_file` and `add_picture` chain (S).

### upstream-PR#1165 — Add mailing list to README
**Verdict:** out-of-scope
**Ask summary:** README link addition.
**Evidence:** docs-only; upstream-only concern.

### upstream-PR#1129 — Feature table enhancements (borders, cantSplit)
**Verdict:** resolved-in-loadfix
**Ask summary:** Add `Table.borders`, `_Cell.borders`, `Row.dont_split` (cantSplit).
**Evidence:** `src/docx/table.py:305,552` expose borders; `src/docx/oxml/table.py:1380-1418` implements cantSplit with inverted `allow_row_split` logic.

### upstream-PR#1126 — Fix typo for WD_PARAGRAPH_ALIGNMENT
**Verdict:** out-of-scope
**Ask summary:** Docs typo in `docs/user/text.rst`.
**Evidence:** docs-only; trivial upstream copy fix.

## Batch summary
- resolved-in-loadfix: 5 (#1392, #1378, #1273, #1270, #1129)
- new-feature-needed: 3 (#1393, #1309, #1295, #1206) -> actually 4
- new-bug-needed: 4 (#1350, #1219, #1205, #1168)
- needs-investigation: 4 (#1355, #1310, #1220, #1196)
- out-of-scope: 8 (#1371, #1361, #1353, #1322, #1272, #1174, #1165, #1126)

Totals (25): resolved=5, new-feature=4, new-bug=4, needs-investigation=4, out-of-scope=8.

### upstream-PR#1118 — Improving Paragraph.text performance by 500%~
**Verdict:** needs-investigation
**Ask summary:** Optimize Paragraph.text by avoiding repeated XPath calls. Likely still applies; benchmark before applying.
**Evidence:** src/docx/text/paragraph.py:1212; src/docx/oxml/text/paragraph.py:266 (still uses xpath over w:r|w:hyperlink|w:fldSimple|w:sdt per call).
**TODO:** Benchmark and apply Paragraph.text micro-optimization for large docs — S.

### upstream-PR#1101 — Fixing tests on Windows
**Verdict:** needs-investigation
**Ask summary:** Normalize CRLF line endings in hash-comparison tests so Windows-git-autocrlf users can run tests.
**Evidence:** no match for splitlines/rstrip in tests/opc/test_phys_pkg.py, tests/parts/test_hdrftr.py, tests/unitutil/file.py.
**TODO:** Port Windows CRLF normalization into test fixtures — S.

### upstream-PR#1097 — Add font scaling property support
**Verdict:** new-feature-needed
**Ask summary:** Expose w:w (character scaling percent) on Font as new property.
**Evidence:** no match for character_scaling/font_scaling/scale in src/docx/text/font.py; only w:w listed in _tag_seq without ZeroOrOne element.
**TODO:** Add Font.character_scale for w:rPr/w:w — S.

### upstream-PR#1064 — outlineLvl paragraph-format API
**Verdict:** new-feature-needed
**Ask summary:** Public getter/setter on ParagraphFormat for w:outlineLvl element.
**Evidence:** oxml binding exists (src/docx/oxml/text/parfmt.py:245) but no outline_level on src/docx/text/parfmt.py.
**TODO:** Add ParagraphFormat.outline_level property — S.

### upstream-PR#1051 — cache for table cells
**Verdict:** new-feature-needed
**Ask summary:** Cache Table._cells to avoid O(rows*cols) rebuild on each indexed access.
**Evidence:** src/docx/table.py:494 _cells is @property not @lazyproperty.
**TODO:** Convert Table._cells to lazyproperty with invalidation on mutation — M.

### upstream-PR#1043 — get_by_name match ui2internal-equivalent names
**Verdict:** resolved-in-loadfix
**Ask summary:** Make style lookup tolerate "Heading 1" vs internal "heading 1" name variants.
**Evidence:** src/docx/styles/styles.py:28,37,63 wraps lookups with BabelFish.ui2internal; addresses upstream #494.
**TODO:** n/a.

### upstream-PR#1036 — read merged cells info
**Verdict:** new-feature-needed
**Ask summary:** Expose helper returning tuples of merged-cell coordinates [(r1,r2,c1,c2),…].
**Evidence:** src/docx/table.py:659,687 has is_merge_origin/merge_origin per cell but no table-level list API.
**TODO:** Add Table.merged_cell_ranges convenience returning list of (r1,r2,c1,c2) — S.

### upstream-PR#1024 — Fix next_id for multiple pictures
**Verdict:** resolved-in-loadfix
**Ask summary:** Use max(existing_ids)+1 so inserting a second picture doesn't reuse an id.
**Evidence:** src/docx/parts/story.py:131-142 already uses max(used_ids)+1.
**TODO:** n/a.

### upstream-PR#1017 — Feature/fields
**Verdict:** resolved-in-loadfix
**Ask summary:** Big feature branch adding bookmarks, fields, footnotes/endnotes; superset of Phase A/B/C loadfix work.
**Evidence:** HISTORY.rst Phase A/B/C (bookmarks, simple+complex fields, footnotes, endnotes) all implemented; src/docx/{bookmarks,fields,footnotes,endnotes}.py present.
**TODO:** n/a.

### upstream-PR#958 — KeyError "Heading 1" with LibreOffice docx
**Verdict:** resolved-in-loadfix
**Ask summary:** Fix case-sensitivity in style.get_by_name so LibreOffice-edited docs don't raise KeyError.
**Evidence:** src/docx/styles/styles.py:28,37 uses BabelFish.ui2internal for name normalization.
**TODO:** n/a.

### upstream-PR#941 — w:firstLineChars support
**Verdict:** new-feature-needed
**Ask summary:** Expose w:firstLineChars ParagraphFormat attribute (East Asian first-line indent in chars).
**Evidence:** no match for firstLineChars in src/docx/.
**TODO:** Add ParagraphFormat.first_line_chars on w:ind/@w:firstLineChars — S.

### upstream-PR#908 — Add support for RGBColor font highlight
**Verdict:** resolved-in-loadfix
**Ask summary:** Allow arbitrary RGB "highlight" via w:shd instead of enum-only w:highlight.
**Evidence:** Font.shading_color (src/docx/text/font.py:700) implements w:rPr/w:shd with RGB; Phase D.20 in HISTORY.rst.
**TODO:** n/a.

### upstream-PR#863 — Exclude sectPr in sectPrChange from section list
**Verdict:** resolved-in-loadfix
**Ask summary:** Avoid treating revision-history sectPr (inside w:sectPrChange) as a real section.
**Evidence:** src/docx/oxml/document.py:61 xpath `./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr` is shallow and skips descendants.
**TODO:** n/a.

### upstream-PR#860 — API for footnote/endnote access
**Verdict:** resolved-in-loadfix
**Ask summary:** Provide footnote/endnote read and reference API at the run level.
**Evidence:** src/docx/footnotes.py, src/docx/endnotes.py; HISTORY.rst Phase A.
**TODO:** n/a.

### upstream-PR#841 — Python 2.6 set-comprehension fix
**Verdict:** out-of-scope
**Ask summary:** Rewrite set comprehension to support Python 2.6.
**Evidence:** pyproject.toml Python 3.7+; HISTORY.rst 1.0.0 removed Python 2 support.
**TODO:** n/a.

### upstream-PR#810 — Consistent reproducible binary output
**Verdict:** new-feature-needed
**Ask summary:** Make docx zip output deterministic (fixed timestamps, sorted rels + parts) for reproducible builds.
**Evidence:** src/docx/opc/phys_pkg.py uses default ZipFile (current-time datetime); no sort of rels/parts; no SOURCE_DATE_EPOCH handling.
**TODO:** Add opt-in deterministic serialization mode (fixed DOS timestamp, sorted rels/parts) — M.

### upstream-PR#798 — SVG image support (hack)
**Verdict:** resolved-in-loadfix
**Ask summary:** Allow add_picture to accept SVG images.
**Evidence:** src/docx/image/svg.py; HISTORY.rst Phase D.22.
**TODO:** n/a.

### upstream-PR#789 — Typo fix thoroughout→throughout
**Verdict:** resolved-in-loadfix
**Ask summary:** Docs typo fix.
**Evidence:** docs/api/style.rst:9 already reads "throughout".
**TODO:** n/a.

### upstream-PR#784 — Hyperlink feature
**Verdict:** resolved-in-loadfix
**Ask summary:** Provide add_hyperlink API on Paragraph.
**Evidence:** src/docx/text/paragraph.py:162 add_hyperlink; HISTORY.rst Phase D.1.
**TODO:** n/a.

### upstream-PR#781 — Unit test for eastAsia fonts
**Verdict:** needs-investigation
**Ask summary:** Adds missing unit tests for previously-added Font.name_eastAsia support.
**Evidence:** oxml binding present (src/docx/oxml/text/font.py:52); need to verify loadfix tests cover rFonts_eastAsia.
**TODO:** Verify tests/text/test_font.py covers eastAsia setter; backfill if missing — S.

### upstream-PR#755 — Section.paragraphs
**Verdict:** new-feature-needed
**Ask summary:** Access paragraphs belonging to a specific Section (between sectPr boundaries).
**Evidence:** no paragraphs attribute on src/docx/section.py (only header/footer paragraphs collection).
**TODO:** Add Section.paragraphs iterating body paragraphs owned by this section — M.

### upstream-PR#734 — Paragraph property for text after accepting changes
**Verdict:** needs-investigation
**Ask summary:** Non-mutating property returning Paragraph text with tracked insertions accepted and deletions removed.
**Evidence:** Document.accept_all_changes mutates (src/docx/document.py:519); revision_marks_text renders markers (src/docx/text/paragraph.py:1185) but no "accepted_text" accessor.
**TODO:** Add Paragraph.accepted_text / .rejected_text read-only properties — S.

### upstream-PR#731 — WMF image support
**Verdict:** new-feature-needed
**Ask summary:** Parser support for WMF vector image headers so add_picture accepts .wmf.
**Evidence:** no wmf.py in src/docx/image/.
**TODO:** Add WMF header parser alongside SVG/EMF pipeline — M.

### upstream-PR#716 — .docm (macro-enabled) support
**Verdict:** resolved-in-loadfix
**Ask summary:** Accept macroEnabled content-type and open .docm files.
**Evidence:** src/docx/api.py:39; src/docx/__init__.py:62 register CT.WML_DOCUMENT_MACRO; HISTORY.rst Phase D.24.
**TODO:** n/a.

### upstream-PR#710 — quickstart.rst wrong example (tuple vs object)
**Verdict:** new-bug-needed
**Ask summary:** docs/user/quickstart.rst table example uses item.qty/item.sku/item.desc on tuple data; should use item[0] etc.
**Evidence:** docs/user/quickstart.rst:161-181 still shows `items = ((7,'1024',...),...)` followed by `item.qty` attribute access.
**TODO:** Fix quickstart table example to use tuple indexing — S.

## Batch summary
- resolved-in-loadfix: 11 (#1043, #1024, #1017, #958, #908, #863, #860, #798, #789, #784, #716)
- new-feature-needed: 8 (#1097, #1064, #1051, #1036, #941, #810, #755, #731)
- new-bug-needed: 1 (#710)
- needs-investigation: 4 (#1118, #1101, #781, #734)
- out-of-scope: 1 (#841)
- Total: 25

### upstream-PR#392 — [feature]add add_chart for docx using chart from python-pptx.
**Verdict:** resolved-in-loadfix
**Ask summary:** Add `Document.add_chart()` for embedding python-pptx chart objects.
**Evidence:** `src/docx/document.py:186 def add_chart`, `src/docx/chart.py`, HISTORY "Charts read + add_chart() (#111)".

### upstream-PR#395 — document.story property, contains Paragraph and Table in document order
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose block iteration mixing Paragraphs and Tables in document order.
**Evidence:** `src/docx/document.py:701 iter_inner_content`, `src/docx/blkcntnr.py:77 iter_inner_content`.

### upstream-PR#417 — Added font property to paragraph
**Verdict:** new-feature-needed
**Ask summary:** Expose `Paragraph.font` that wraps `w:pPr/w:rPr` (paragraph-mark formatting).
**Evidence:** `w:rPr` listed in `_tag_seq` of CT_PPr at `src/docx/oxml/text/parfmt.py:212` but no accessor or `Paragraph.font` proxy.
**TODO (if applicable):** Add `Paragraph.font` / `CT_PPr.rPr` accessor exposing paragraph-mark Font — S.

### upstream-PR#423 — Added Canon EOS Jpeg format
**Verdict:** needs-investigation
**Ask summary:** Recognise Canon EOS JPEGs (non-JFIF/Exif signature) so add_picture doesn't reject them.
**Evidence:** `src/docx/image/__init__.py` SIGNATURES still lists only JFIF/Exif/GIF/Tiff/BMP/PNG.
**TODO (if applicable):** Evaluate adding Canon EOS / other JPEG variants to SIGNATURES — S.

### upstream-PR#429 — Add method to add picture to table cell. See #1.
**Verdict:** new-feature-needed
**Ask summary:** Provide `_Cell.add_picture()` convenience on table cells.
**Evidence:** `_Cell` at `src/docx/table.py:530` exposes `add_paragraph`/`add_table` but no `add_picture`.
**TODO (if applicable):** Add `_Cell.add_picture(image, width, height)` delegating to a new paragraph.add_run.add_picture — S.

### upstream-PR#445 — Feature/bookmarks
**Verdict:** resolved-in-loadfix
**Ask summary:** Full bookmark create/read/delete API (original Ben Timby PR).
**Evidence:** `src/docx/bookmarks.py`, `src/docx/oxml/bookmarks.py`, HISTORY Phase C "bookmarks create/read/delete (#52)".

### upstream-PR#462 — (WIP) Table style options
**Verdict:** resolved-in-loadfix
**Ask summary:** Toggle tblLook conditional flags (header/total rows, banding, first/last col).
**Evidence:** `src/docx/oxml/table.py:579 CT_TblLook`, `Table.style_flags` at `src/docx/table.py:316`, HISTORY "Table.style_flags (#144)".

### upstream-PR#468 — Merged renejsum's support for custom_properties, Fixes #91
**Verdict:** resolved-in-loadfix
**Ask summary:** Read/write of `customProperties.xml` (customProps part).
**Evidence:** `src/docx/custom_properties.py`, `src/docx/parts/custom_properties.py`, `Document.custom_properties` at `src/docx/document.py:591`, HISTORY D.4.

### upstream-PR#522 — Extend flexibility of api.py
**Verdict:** new-feature-needed
**Ask summary:** Accept `.dotx` (and likely `.dotm`) templates in `Document()` / content-type check.
**Evidence:** `src/docx/api.py:39` only accepts `WML_DOCUMENT_MAIN` and `WML_DOCUMENT_MACRO`; no `.dotx` content type constants defined.
**TODO (if applicable):** Add `WML_TEMPLATE_MAIN` / macro-enabled template content types and register with PartFactory — S.

### upstream-PR#523 — Update to __init__.py
**Verdict:** new-feature-needed
**Ask summary:** Companion `__init__.py` registration for `.dotx` template content type.
**Evidence:** no `dotx`/`template.main+xml` registration in `src/docx/__init__.py`.
**TODO (if applicable):** Covered by same `.dotx` task as PR#522 — S.

### upstream-PR#537 — Make rPr accessible from pPr
**Verdict:** new-feature-needed
**Ask summary:** Expose `w:pPr/w:rPr` (paragraph-mark run properties) as a ZeroOrOne accessor.
**Evidence:** `_tag_seq` in `CT_PPr` at `src/docx/oxml/text/parfmt.py:212` lists `w:rPr` but no descriptor generated.
**TODO (if applicable):** Add `rPr` ZeroOrOne descriptor on CT_PPr and surface via Paragraph.font — S (overlaps PR#417).

### upstream-PR#539 — Feature/bookmarks
**Verdict:** resolved-in-loadfix
**Ask summary:** Alternate bookmark implementation, same scope as PR#445.
**Evidence:** same as PR#445 — bookmarks shipped in Phase C.

### upstream-PR#564 — detect beginning of new row. if last row is incomplete, add empty cells.
**Verdict:** needs-investigation
**Ask summary:** Harden `_cells` against rows whose grid_span totals do not match tblGrid.
**Evidence:** `src/docx/table.py:495 _cells` has an orphan vMerge fallback but no explicit row-incomplete padding.
**TODO (if applicable):** Investigate malformed-table regressions; add row-padding when `sum(grid_span) < col_count` — M.

### upstream-PR#565 — use caching for faster table cell access
**Verdict:** needs-investigation
**Ask summary:** Cache `Table._cells` result for large tables (perf).
**Evidence:** `src/docx/table.py:495 _cells` recomputes every call; no `@cached_property`/`lazyproperty`.
**TODO (if applicable):** Consider cached cell grid with invalidation on structural mutation — M.

### upstream-PR#576 — read eastAsia font name
**Verdict:** resolved-in-loadfix
**Ask summary:** Read/write `w:rFonts/@w:eastAsia` on Font.
**Evidence:** `src/docx/text/font.py:581 name_east_asia` getter/setter; `rFonts.eastAsia` at `src/docx/oxml/text/font.py:52`.

### upstream-PR#579 — Table: feature : allow table looking modification
**Verdict:** resolved-in-loadfix
**Ask summary:** Modify tblLook header/footer row/col and banding at create time or later.
**Evidence:** `Table.style_flags` + `CT_TblLook` (see PR#462); HISTORY "Table.style_flags (#144)".

### upstream-PR#582 — Adding low-level support for numbering styles.
**Verdict:** resolved-in-loadfix
**Ask summary:** CT_ classes for numbering (abstractNum, Lvl, numFmt, lvlText) for workarounds.
**Evidence:** `src/docx/oxml/numbering.py` defines CT_AbstractNum / CT_Lvl / CT_LvlText / CT_Num; HISTORY Phase D.9.

### upstream-PR#608 — Add StringIO/BytesIO distinction for Python 2/3
**Verdict:** out-of-scope
**Ask summary:** Docs tweak distinguishing StringIO vs BytesIO for Python 2/3.
**Evidence:** loadfix dropped Py2 in 1.0.0; `docs/user/documents.rst:71-85` still mentions StringIO but the Py2/3 split is moot.

### upstream-PR#622 — Namespaces parameter support in xpath
**Verdict:** new-feature-needed
**Ask summary:** Allow users to pass extra `namespaces=` to `BaseOxmlElement.xpath` for custom prefixes.
**Evidence:** `src/docx/oxml/xmlchemy.py:688 xpath` hard-codes `namespaces=nsmap, **kwargs` which raises if caller also passes `namespaces=`.
**TODO (if applicable):** Merge caller-supplied namespaces dict over `nsmap` in `xpath()` — S.

### upstream-PR#643 — w:abstractNum in numbering.xml
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose CT_AbstractNum inside numbering.xml.
**Evidence:** `src/docx/oxml/numbering.py:119 class CT_AbstractNum` with `add_lvl` / `get_lvl`.

### upstream-PR#649 — Adding in support for Form Fields and Nested Chunks (alt_chunks)
**Verdict:** new-feature-needed
**Ask summary:** Two asks — (a) legacy form fields via ffData, (b) altChunk (nested docx inclusion).
**Evidence:** form fields present at `src/docx/form_fields.py` + `src/docx/oxml/form_fields.py` (HISTORY "legacy form fields (#123)"); but no `altChunk`/`AltChunkPart` anywhere in src.
**TODO (if applicable):** Add altChunk part + `Document.add_alt_chunk`/iteration of altChunks — M.

### upstream-PR#653 — Fixups to enums and their associated documentation
**Verdict:** out-of-scope
**Ask summary:** Cosmetic enum docs normalisation with intersphinx references.
**Evidence:** `docs/api/enum/*.rst` already extensive in loadfix; enum structure already rebuilt as `BaseXmlEnum` in `src/docx/enum/style.py:426` etc.

### upstream-PR#673 — Add '.docm' compatibility
**Verdict:** resolved-in-loadfix
**Ask summary:** Accept macro-enabled `.docm` documents.
**Evidence:** `src/docx/api.py:39` allows `CT.WML_DOCUMENT_MACRO`; registration at `src/docx/__init__.py:62`; HISTORY D.24.

### upstream-PR#699 — Fix tests failing on Windows due to git EOL conversion
**Verdict:** new-bug-needed
**Ask summary:** Add a `.gitattributes` to force LF for text test fixtures so Windows checkouts don't corrupt binary comparisons.
**Evidence:** no `.gitattributes` in repo root.
**TODO (if applicable):** Add `.gitattributes` entry disabling EOL conversion for test fixture dirs — S.

### upstream-PR#706 — Typo: Added a dot to fix sentences
**Verdict:** out-of-scope
**Ask summary:** Add a missing full stop to an upstream sentence in `docs/user/api-concepts.rst`.
**Evidence:** loadfix's `docs/user/api-concepts.rst` has been substantially rewritten; the targeted upstream sentence is no longer present.

## Batch summary
- resolved-in-loadfix: 11 (392, 395, 445, 462, 468, 539, 576, 579, 582, 643, 673)
- new-feature-needed: 7 (417, 429, 522, 523, 537, 622, 649)
- new-bug-needed: 1 (699)
- needs-investigation: 3 (423, 564, 565)
- out-of-scope: 3 (608, 653, 706)
- total: 25

### upstream-PR#367 — merge tables and paragraphs together
**Verdict:** needs-investigation
**Ask summary:** Adds a helper to iterate/merge tables and paragraphs in document order. Empty body, minimal context.
**Evidence:** Document/body iteration exists (`iter_inner_content`) but no explicit merge helper.
**TODO (if applicable):** Investigate scope of "merge" helper; likely superseded by existing APIs — S.

### upstream-PR#355 — feature analysis: section columns
**Verdict:** resolved-in-loadfix
**Ask summary:** Analysis doc + specimens toward multi-column sections (fixing upstream #167).
**Evidence:** `src/docx/section.py:60` SectionColumns; HISTORY D.19 "Multi-column section layout (#60)".
**TODO (if applicable):** n/a.

### upstream-PR#341 — add insert text table and picture by bookmark
**Verdict:** needs-investigation
**Ask summary:** Insert paragraph/table/picture at bookmark position, with style.
**Evidence:** `src/docx/bookmarks.py` exists; Phase C bookmarks (#52); Phase D.13 insert-at-position (#26); need to confirm bookmark-targeted insertion is exposed.
**TODO (if applicable):** Confirm or add `Bookmark.insert_paragraph/table/picture` helpers — M.

### upstream-PR#337 — Fixed failure caused by "python -OO"
**Verdict:** needs-investigation
**Ask summary:** Avoid crashes when docstrings are stripped (`-OO`).
**Evidence:** `src/docx/shared.py:262 docstring = f.__doc__` still assumes docstrings — not guarded. HISTORY 0.8.7 claims #375 fix; `-OO` regression possible.
**TODO (if applicable):** Audit `__doc__` accesses for `None` under `-OO` — S.

### upstream-PR#329 — east asia, chinese font family
**Verdict:** resolved-in-loadfix
**Ask summary:** Ensure East-Asian font name is applied (w:rFonts/@w:eastAsia).
**Evidence:** `src/docx/text/font.py:390` east_asian_language; HISTORY "East Asian typography (#128)" and "Font.language / east_asian_language / bidi_language (#160)".
**TODO (if applicable):** n/a.

### upstream-PR#326 — Enable python-docx to add border (rows)
**Verdict:** resolved-in-loadfix
**Ask summary:** Add row/table border API.
**Evidence:** HISTORY "Table.borders / _Cell.borders (#102)"; Phase D.7 paragraph borders (#109).
**TODO (if applicable):** Verify Row-level border setter exists (may need thin wrapper) — S.

### upstream-PR#317 — feature: Shape.alternative_text
**Verdict:** resolved-in-loadfix
**Ask summary:** Getter/setter for inline-shape alt-text/title.
**Evidence:** `src/docx/shape.py:106 alt_text` (getter/setter); HISTORY "alt_text / title on InlineShape and FloatingImage (#158)".
**TODO (if applicable):** n/a.

### upstream-PR#312 — feature: Cell.vAlign property
**Verdict:** resolved-in-loadfix
**Ask summary:** Cell vertical alignment getter/setter.
**Evidence:** `src/docx/table.py:856 vertical_alignment` getter+setter. (Present since 0.8.7.)
**TODO (if applicable):** n/a.

### upstream-PR#310 — Make xpath() match lxml's xpath() function signature
**Verdict:** resolved-in-loadfix
**Ask summary:** Accept kwargs (namespaces/variables) so third-party tools work on oxml elements.
**Evidence:** `src/docx/oxml/xmlchemy.py:688 def xpath(self, xpath_str, **kwargs)` already proxies kwargs.
**TODO (if applicable):** n/a.

### upstream-PR#307 — Add bidi property to ParagraphFormat
**Verdict:** resolved-in-loadfix
**Ask summary:** ParagraphFormat.bidi (w:pPr/w:bidi).
**Evidence:** `src/docx/text/parfmt.py:332` bidi property; HISTORY "RTL / bidi on Paragraph and Run (#127)".
**TODO (if applicable):** n/a.

### upstream-PR#303 — Disallowing XML entity expansion
**Verdict:** resolved-in-loadfix
**Ask summary:** XXE hardening on XMLParser.
**Evidence:** `src/docx/oxml/parser.py:19-35 resolve_entities=False, no_network=True`.
**TODO (if applicable):** n/a.

### upstream-PR#277 — Some arguably useful document manipulation features
**Verdict:** needs-investigation
**Ask summary:** Grab-bag of blkcntnr/document/section/table/paragraph manipulation helpers.
**Evidence:** Phase D.13/D.26, Paragraph.delete/insert (#50, #26) likely cover most; unclear on remainder.
**TODO (if applicable):** Diff patch vs current API to find gaps — M.

### upstream-PR#274 — non-destructive method to replace text inside a run
**Verdict:** resolved-in-loadfix
**Ask summary:** Safe text replacement preserving formatting on w:t/w:br/w:cr/w:tab.
**Evidence:** HISTORY D.10 "Search and replace with formatting preservation (#91)" plus Document.search_regex/replace_regex (#153, #154).
**TODO (if applicable):** n/a.

### upstream-PR#250 — Support some missing tags (w:shd, CT_Shd, themeColor, etc.)
**Verdict:** resolved-in-loadfix
**Ask summary:** Add w:shd on p/r/tc/tbl, ST_ThemeColor, etc.
**Evidence:** HISTORY D.6 "Cell shading and background color (#63)" and D.20 "Font.shading run-level background color (#33)".
**TODO (if applicable):** Confirm table-level w:shd coverage — S.

### upstream-PR#239 — Find key in styles dict without BabelFish translate
**Verdict:** needs-investigation
**Ask summary:** Fallback lookup when user passes internal (non-translated) style id.
**Evidence:** `src/docx/styles/styles.py:37` still uses BabelFish.ui2internal only.
**TODO (if applicable):** Add fallback to raw name lookup when ui2internal miss — S.

### upstream-PR#227 — image title and alt text
**Verdict:** resolved-in-loadfix
**Ask summary:** Set picture title/descr for accessibility.
**Evidence:** HISTORY "alt_text / title on InlineShape and FloatingImage (#158)".
**TODO (if applicable):** n/a.

### upstream-PR#226 — TrPr: row as header
**Verdict:** resolved-in-loadfix
**Ask summary:** Mark table row as repeating header.
**Evidence:** `src/docx/table.py:1674 is_header` getter+setter; HISTORY "_Row.is_header (#93)".
**TODO (if applicable):** n/a.

### upstream-PR#210 — Add ability to restart numbering
**Verdict:** resolved-in-loadfix
**Ask summary:** Expose w:lvlRestart / numbering restart control.
**Evidence:** `src/docx/oxml/numbering.py:45 w:lvlRestart`; HISTORY D.9 Numbering style control (#22).
**TODO (if applicable):** n/a.

### upstream-PR#208 — Patch 1 (index.rst modification example)
**Verdict:** out-of-scope
**Ask summary:** Duplicate of #207 — docs-only example update.
**Evidence:** Only `docs/index.rst` + images.
**TODO (if applicable):** n/a.

### upstream-PR#207 — Update index.rst (modify-existing example)
**Verdict:** out-of-scope
**Ask summary:** Add "modify existing document" example to docs index.
**Evidence:** docs-only; editorial decision upstream.
**TODO (if applicable):** n/a.

### upstream-PR#196 — EMF image support
**Verdict:** new-feature-needed
**Ask summary:** Recognize EMF images (size from EMF header, 300 DPI default).
**Evidence:** No match under `src/docx/image/` for EMF. SVG (D.22) done, EMF absent.
**TODO (if applicable):** Add EMF (and WMF) image handlers with header parsing — M.

### upstream-PR#177 — Use resource_stream to cover frozen apps
**Verdict:** needs-investigation
**Ask summary:** Use pkg_resources.resource_stream so `default.docx` loads from frozen/PyInstaller bundles.
**Evidence:** No `resource_stream`/`pkg_resources` usage in `src/docx/api.py` or `__init__.py`.
**TODO (if applicable):** Add importlib.resources fallback for `_default_docx_path` — S.

### upstream-PR#71 — add_text() on Paragraph
**Verdict:** needs-investigation
**Ask summary:** Convenience to append text to last run or a new run with inherited style.
**Evidence:** `Paragraph.add_run` exists, no `add_text`. Not listed in HISTORY.
**TODO (if applicable):** Add Paragraph.add_text as thin helper — S.

### upstream-PR#53 — Initial read-only support for notes, style objects
**Verdict:** resolved-in-loadfix
**Ask summary:** Early read-only footnotes/endnotes + style id/type/name.
**Evidence:** Phase A "Footnotes and endnotes (#1, #3, #17 … #53)" — PR#53 explicitly cited in HISTORY.
**TODO (if applicable):** n/a.

### upstream-PR#35 — Copied code example fix
**Verdict:** out-of-scope
**Ask summary:** Minor fix to a doc tutorial snippet.
**Evidence:** `docs/index.rst` only.
**TODO (if applicable):** n/a.

## Batch summary
- resolved-in-loadfix: 13 (PR#355, #329, #326, #317, #312, #310, #307, #303, #274, #250, #227, #226, #210, #53) — 14 actually
- new-feature-needed: 1 (PR#196)
- needs-investigation: 7 (PR#367, #341, #337, #277, #239, #177, #71)
- out-of-scope: 3 (PR#208, #207, #35)
- new-bug-needed: 0

(Recount: resolved=14, new-feature=1, needs-investigation=7, out-of-scope=3, total=25.)

