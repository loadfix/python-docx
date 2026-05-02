# TODO — upstream issues + PRs consolidation

Derived from ISSUES_AUDIT.md (audit of 584 open issues + 125 open PRs at
python-openxml/python-docx as of 2026-05-02). `resolved-in-loadfix` items
(264) are omitted; this file lists only actionable + investigation items,
deduped across the original audit batches.

Effort labels: **S** ≤ 1 day, **M** 1-3 days, **L** > 3 days.

---

## Part A — Bugs (29 after dedupe)

Pre-dedupe: 41 new-bug-needed entries. Ordered roughly by module.

### Image / drawing
- **Zero/invalid DPI causes ZeroDivisionError (JPEG/BMP)** — fall back to 72 DPI in `_ImageHeaderBase`/`bmp.py` when `horz_dpi`/`vert_dpi` is 0 or None. Effort **S** | `src/docx/image/image.py`, `src/docx/image/bmp.py`. Closes upstream#1494, #515, #265, upstream-PR#1505.
- **EXIF inline-value handling (TIFF6.0 ≤4 byte rule)** — fix ASCII/resolution IFD entries so values fitting in the 4-byte Value/Offset slot are read inline; guard non-rational resolution tags. Effort **S** | `src/docx/image/helpers.py`. Closes upstream#1124, #184, upstream-PR#1451.
- **docPr/@id collision across stories** — allocate `wp:docPr/@id` from a document-scoped counter spanning body + headers + footers instead of per-story scan. Effort **M** | `src/docx/parts/story.py`. Closes upstream#1121, #455.
- **InlineShape detection inside mc:AlternateContent** — extend inline/anchor xpath to traverse `mc:AlternateContent/mc:Choice` with `mc:Fallback` handling. Effort **M** | `src/docx/shape.py`. Closes upstream#451.

### Tables
- **`cell._tc.bottom` raises on gridBefore/omitted cells** — make `tc_at_grid_offset` match by range rather than equality so vMerge/gridSpan scenarios resolve. Effort **S** | `src/docx/oxml/table.py`. Closes upstream#1458.
- **`_grow_to` recursion on large merges** — convert recursive row-walk in `CT_Tc._grow_to` to an iterative loop. Effort **S** | `src/docx/oxml/table.py`. Closes upstream#1208.
- **`add_column` corrupts file** — insert a matching `w:tc` in every row alongside the new `w:gridCol`. Effort **S** | `src/docx/table.py`. Closes upstream#1102.
- **Nested-table merge cell ValueError** — fix `_grid_col` to walk only the enclosing `w:tr`, not ancestor tables. Effort **M** | `src/docx/oxml/table.py`. Closes upstream#169.
- **`table.columns[1:]` returns a single `_Column`** — add slice support to `_Columns.__getitem__` returning `list[_Column]`. Effort **S** | `src/docx/table.py`. Closes upstream#770.
- **`add_table` leaves stray `w:tbl` on bad style** — validate style before inserting, or remove on failure. Effort **S** | `src/docx/document.py`, `src/docx/blkcntnr.py`. Closes upstream#563.
- **`add_table` IndexError when body has no `w:sectPr`** — fall back to page-size default / `EMU(0)` when `sectPr_lst` is empty. Effort **S** | `src/docx/blkcntnr.py`. Closes upstream#514.

### Text / runs / font
- **Fractional twips/half-points crash conversions** — tolerate float in `ST_TwipsMeasure` and `ST_HpsMeasure` (round to int for twips, keep float for half-points). Effort **S** | `src/docx/oxml/simpletypes.py`. Closes upstream#1539, #1475, upstream-PR#1478.
- **`WD_PARAGRAPH_ALIGNMENT` missing `start`/`end`** — add START/END members mapping to `start`/`end` XML values (aliasing LEFT/RIGHT). Effort **S** | `src/docx/enum/text.py`. Closes upstream#1473.
- **Runs inside `w:smartTag` are dropped** — treat `w:smartTag` (and nested `w:customXml`/`w:sdt`) as transparent containers in run/text iteration. Effort **M** | `src/docx/oxml/text/paragraph.py`, `src/docx/text/paragraph.py`. Closes upstream#932, #225.
- **`styles[WD_STYLE.BODY_TEXT]` KeyError** — accept `WD_BUILTIN_STYLE` members in `Styles.__getitem__` (translate enum → UI name / builtin-id). Effort **S** | `src/docx/styles/styles.py`. Closes upstream#1439.

### Packaging / I/O
- **Internal bookmark hyperlink crashes package reader** — skip/ignore relationships whose target is a pure `#fragment` (or NULL) rather than looking up a package part. Effort **S** | `src/docx/opc/pkgreader.py`, `src/docx/opc/oxml.py`. Closes upstream#902, upstream#1349, upstream#678, upstream-PR#1498, upstream-PR#1518, upstream-PR#1350.
- **Ignore rels to missing package parts** — wrap `blob_for` in a KeyError-tolerant branch (and skip the srel) so Word-style loose docs still open. Effort **S** | `src/docx/opc/pkgreader.py`. Closes upstream-PR#1219.
- **Windows backslash in rel targets** — normalise `\` → `/` in `_SerializedRelationship._target_ref`. Effort **S** | `src/docx/opc/pkgreader.py`. Closes upstream-PR#1205.
- **`Document.save()` silently empty on colon-in-filename (Windows)** — raise OSError/ValueError for Windows-invalid filename chars. Effort **S** | `src/docx/document.py`, `src/docx/opc/package.py`. Closes upstream#1111.
- **Large-document parse fails on AttValue >10MB** — opt-in `Document(..., huge_tree=True)` parser config. Effort **S** | `src/docx/oxml/__init__.py`, `src/docx/api.py`. Closes upstream#1086.
- **Duplicate `docProps/core.xml` when alt core-props reltype used** — discover existing core-properties part via the alternate reltype before creating a new one. Effort **S** | `src/docx/parts/document.py`, `src/docx/opc/constants.py`. Closes upstream-PR#1436.
- **`pathlib.Path` rejected by `Image.from_file`** — accept `os.PathLike` in the image/add_picture chain (applies `os.fspath()` at entry). Effort **S** | `src/docx/image/image.py`. Closes upstream-PR#1168. (See also feature item under Image/drawing for the broader PathLike sweep.)

### Numbering / perf
- **`_next_numId` O(n²) on big docs** — use a max(+1) fast path or gap-cache in `CT_Numbering._next_numId` / `_next_abstractNumId` while preserving gap-fill. Effort **S** | `src/docx/oxml/numbering.py`. Closes upstream#940.

### Test / dev hygiene
- **`pyparsing.delimitedList` deprecated** — swap to `DelimitedList` in `tests/unitutil/cxml.py`. Effort **S**. Closes upstream-PR#1536.

### Docs
- **Docs: `table.cells(0,0)` typo (should be `table.cell`)** — fix snippets in `docs/dev/analysis/features/cell-merge.*` and quickstart table example. Effort **S** | `docs/`. Closes upstream#166, #164.
- **Docs: `StringIO` → `BytesIO` in `documents.rst`** — fix example. Effort **S** | `docs/user/documents.rst`. Closes upstream-PR#1514.
- **Docs: "those" → "that" in `install.rst`** — grammar fix. Effort **S**. Closes upstream-PR#1395.
- **Docs: quickstart tuple vs object example** — fix `item.qty` → `item[0]` etc. Effort **S** | `docs/user/quickstart.rst`. Closes upstream-PR#710.
- **Tests: `.gitattributes` to disable EOL conversion for binary fixtures** — unblock Windows checkouts. Effort **S**. Closes upstream-PR#699.

---

## Part B — Investigations (107 after dedupe)

Pre-dedupe: 131 needs-investigation entries. Each entry: what to probe + likely reclassification. Some items migrated to Part A or Part C where the action is clear; this section still holds the true "decide before acting" set.

### Tables
- **vMerge/split-cell grid decoding bugs** — reproduce split-after-merge and `column_cells(0)` cross-row leakage; if confirmed → bug in `_cells` grid reconstruction (similar to #1458 fix). Possibly bug **M**. Closes upstream#939, #1367, upstream-PR#1526, upstream-PR#564.
- **`_Table._column_count` mismatches `len(row.cells)`** — audit tblGrid vs row gridSpan totals; if off → bug **M**. Closes upstream#1334.
- **`Table.cell()` flat vs 2-D indexing with merged cells** — repro upstream #1193; refactor to 2-D grid if confirmed. Bug or feature **M**. Closes upstream-PR#1196.
- **`cell.text` drops CJK substrings** — reproduce with attached docx; likely AlternateContent/field traversal gap → bug **M**. Closes upstream#1390.
- **Missing `w:tblGrid` child raises InvalidXmlError** — harden `CT_Tbl` to synthesize an empty `tblGrid` under recover mode; bug **S**. Closes upstream#548.
- **`Table.autofit` default ambiguity** — doc fix; likely docs-only. Closes upstream#1159.
- **`Table.style = 'Table Grid'` KeyError (latent style)** — promote latent style on assignment; likely feature **M**. Closes upstream#694.
- **`Table Normal` shows no borders vs Word** — verify default style chain; either docs or default-template fix. Closes upstream#1324.
- **Custom `TableStyle` lost via `add_table(style=...)`** — reproduce style lookup path; bug **M**. Closes upstream#319.
- **Custom table style forces borders on save** — see upstream#476; likely Word-limitation doc. Closes upstream#476.
- **Row borders/style not applied to `add_row()` cells** — investigate cnfStyle/tblLook tracking. Bug **M**. Closes upstream#306.
- **`WD_TABLE_DIRECTION.RTL` no effect** — verify round-trip; add `direction` alias if needed. Closes upstream#1227.
- **`paragraph_format.space_*` ignored in cells** — investigate TableNormal inheritance override; likely docs. Closes upstream#305.
- **`run.font.size` ignored in merged cells** — repro; likely docs (target merge origin). Closes upstream#1091.
- **Hidden-column extraction helper** — consider `Table.visible_cells` / skip `w:vanish` helper. Closes upstream#1120.
- **`table.add_row()` O(n²)** — profile and add bulk `add_rows(n)` / batch API. Closes upstream#174, #703.

### Text / paragraph / run
- **`Paragraph.text` misses `w:sdt`/`w:smartTag`/`mc:AlternateContent`/fields** — audit traversal; bug **M**. Closes upstream#1327, #1389, #339, #335, #328.
- **`run.add_text("<")` possibly drops the char** — regression test; confirm `&lt;` serialization. Bug **S**. Closes upstream#1330.
- **`run.font.underline` true for visibly un-underlined runs** — verify `w:u val="none"` inheritance; possibly add `effective_underline`. Bug/feature **M**. Closes upstream#1338.
- **Track-changes view of `Paragraph.text` (w:ins/w:del/w:moveTo)** — decide whether default `.text` reflects ins/moveTo or remain accept-first; also consider `accepted_text`/`rejected_text`/`final_paragraphs` read-only. Bug or feature **S-M**. Closes upstream#6, upstream-PR#1538, upstream-PR#734.
- **`Paragraph.text` performance** — benchmark upstream #1118 micro-opts against current code. Perf **S**. Closes upstream-PR#1118.
- **`Run.add_picture` silently no-op on existing paragraph** — regression test; bug **S**. Closes upstream#981.
- **Highlight/underline skips hyperlink runs in paragraph** — verify iterator covers w:hyperlink children. Bug **S**. Closes upstream#1021.
- **Comments disappear when setting `run.text`** — regression test with comment-spanning run; fix if `w:commentRangeStart/End`/`w:commentReference` lost. Bug **M**. Closes upstream#1519.
- **Find/replace should cover headers + text-box content** — extend `_iter_all_paragraphs` to yield `w:txbxContent` / header paragraphs. Feature **M**. Closes upstream#413.
- **Styles perf hang in `_iter_styles`** — reproduce with pathological styles.xml. Bug **M**. Closes upstream#999.
- **`paragraph_format.left_indent` returns None for style-inherited indent** — add `resolved_left_indent` walking style chain. Feature **M**. Closes upstream#569.
- **ColorFormat.brightness advertised but not implemented** — add getter/setter (mirrors feature list). Bug **S** or feature. Closes upstream#665.
- **`add_paragraph` performance O(n²)** — profile and optimise insertion point caching. Perf **M**. Closes upstream#408.
- **RTL/bidi formatting interactions** — investigate `w:rtl` successors ordering, `Font.name` writing `@w:cs`, and `Paragraph.bidi` bullet flip. Bug/feature **S-M**. Closes upstream#510, #430, #421, #387.
- **`Font.name` ineffective on Heading/Title/Subtitle** — clear sibling `asciiTheme`/`hAnsiTheme`. Bug **S**. Closes upstream#366.
- **`font.rtl` disturbs `font.size` (szCs)** — ensure `Font.size` also sets szCs. Bug **S**. Closes upstream#973.

### Fields / bookmarks / TOC
- **Placeholder text in `TOC \h \z \c "Table"`** — document that placeholder is field-result text and usable via `set_result_text`. Docs **S**. Closes upstream#1418.
- **Unable to get text of MERGEFIELD placeholders** — verify `replace_all` traverses `w:instrText` / field ranges. Bug **S**. Closes upstream#1370.
- **`FollowedHyperlink` latent/missing builtin** — add helper to materialize latent built-ins. Feature **M**. Closes upstream#1376.
- **Content-control write-through to customXml** — SDT w:t edit gets overwritten by dataBinding target on open. Feature/docs **M**. Closes upstream#965.
- **`TOC_HEADING` / `INDEX_HEADING` style lookup** — verify BabelFish mapping. Bug **S**. Closes upstream#542.

### Styles / numbering
- **`styles[WD_STYLE.HEADING_1]` KeyError (casing/variants)** — fall back to raw-name lookup when internal lookup misses. Bug **S**. Closes upstream#494, #420, upstream-PR#239.
- **Default-style font access (docDefaults)** — expose `Styles.default_rpr` / fall back to `docDefaults/rPrDefault` when attrs are None. Feature **S**. Closes upstream#496, #86.
- **`paragraph.style = 'List Bullet'` KeyError in cell (style latent)** — auto-load built-in on first reference. Feature **M**. Closes upstream#570.
- **Numbered list rendering off across paragraphs** — needs repro. Bug **M**. Closes upstream#639.
- **Default-styles.xml list indents vs Word's** — review and update default template. Bug **M**. Closes upstream#1443.
- **Heading numbering scheme helper** — `Numbering.set_heading_scheme(...)` applying a multi-level scheme to Heading1..Heading9. Feature **M**. Closes upstream#1535.
- **`Settings.attached_template`** — new Settings property + rel mgmt. Feature **M**. Closes upstream#503.
- **Custom-table-style font vs Normal-style override** — investigate `w:tblStylePr/w:rPr` precedence. Bug **M**. Closes upstream#1161.
- **Quickstart style-name regressions** — IntenseQuote/ListBullet/ListNumber alias audit. Bug **S**. Closes upstream#345, #215, #198.

### Sections / page layout
- **`add_section()` corrupts template header/footer refs** — audit header/footer clone path. Bug **M**. Closes upstream#343.
- **Orientation setter not swapping page_width/height** — auto-swap when orientation changes. Bug **S**. Closes upstream#214.
- **RTL header style overwritten** — docs example covering style + bidi coexistence. Docs **S**. Closes upstream#387.

### Images
- **EXIF orientation ignored by `add_picture`** — honour orientation tag when computing dimensions/rendering. Bug **M**. Closes upstream#540.
- **Image size wrong until moved in Word** — repro `<wp:extent>`/`<a:ext>` consistency. Bug **M**. Closes upstream#1164.
- **JPEG sniffer rejects valid JPEGs lacking JFIF/Exif markers** — relax to accept any valid SOI+SOF sequence; covers Canon EOS variants too. Bug **S**. Closes upstream#1430, #187, #350, upstream-PR#423.
- **Watermark image shifts on reopen** — verify anchoring round-trip against current watermark module. Bug **S**. Closes upstream#1474.
- **Hidden inline image scenario** — document common causes (anchor vs inline). Docs **S**. Closes upstream#667.
- **`add_picture` full `title`/`descr` pass-through** — add kwargs to `add_picture()` / `new_pic_inline()`. Feature **S**. Closes upstream-PR#1530.

### Packaging / I/O / build
- **libmagic mime detection returns octet-stream** — investigate zip-ordering / uncompressed mimetype stream. Feature **S**. Closes upstream#545, upstream-PR#1355.
- **Frozen-app default template lookup** — `importlib.resources` fallback for `_default_docx_path`. Bug **S**. Closes upstream#176, upstream-PR#1310, upstream-PR#177.
- **Reproducible binary output** — investigate zeroing zip mtimes, sorting rels, dropping `w:rsid*`. Feature **M**. Closes upstream#811.
- **Deterministic zipping (see Part C for feature)** — see upstream#1042 / upstream-PR#810 in features.
- **Google Docs fails to convert generated .docx** — investigate compat setting(s). Bug **M**. Closes upstream#607.
- **`doc.part.related_parts[rId]._blob` returns None after header/footer change** — reload/regression test. Bug **M**. Closes upstream#606.
- **Memory retention on large doc loads** — investigate dropping `Package._rels` refs / explicit `_element.clear()`. Perf **M**. Closes upstream#1428.
- **Warn-and-skip missing part references (WSL / broken docs)** — extend recover mode so missing part references log instead of KeyError. Bug **S**. Closes upstream#604.
- **`int(None)` on docx→html attribute** — awaits repro; add None-guard when source identified. Bug **S**. Closes upstream#1369.
- **`PyCharm` debugger error on `tc.bottom`** — repro without IDE. Bug **S**. Closes upstream#1433.
- **`dir(CT_*)` descriptor-dict error** — verify still broken on current MetaOxmlElement. Bug **S**. Closes upstream-PR#1220.
- **`-OO` mode crashes on stripped docstrings** — audit `__doc__` accesses. Bug **S**. Closes upstream-PR#337.
- **Windows CRLF test normalisation** — port PR#1101 fixture changes. Tests **S**. Closes upstream-PR#1101.
- **`xmlchemy.first_child_found` XPath vs find** — benchmark and switch to compiled XPath if measurable. Perf **M**. Closes upstream#241.

### Metadata / core props / settings
- **Naive-datetime core_properties** — document "treated as UTC"; optional DeprecationWarning. Docs/warning **S**. Closes upstream#1542.
- **Unicode title round-trip** — regression test for `coreprops.title` with non-ASCII. Tests **S**. Closes upstream#220.
- **Duplicate `cp:lastModifiedBy` on set** — regression test that existing docs with duplicate entry get single-element result. Bug **S**. Closes upstream#1037.
- **`Document.save` overwrite behaviour undocumented** — note in docs + docstring. Docs **S**. Closes upstream#1252.
- **Preserve original `settings.xml` on merge** — merging docs writes minimal settings that breaks mc:AlternateContent offsets. Bug **M**. Closes upstream#573.
- **Protection classification (sensitivity label)** — custom-xml based. Feature **M**. Closes upstream#602.
- **`ElementProxy` missing from Sphinx API docs** — add `.. autoclass::` directive. Docs **S**. Closes upstream#1178.
- **Docs: tables formatting recipe** — "Formatting cell content" section with header-row bold/centered example. Docs **S**. Closes upstream#1160.
- **Docs: hyphen/dash in Table.style docstring** — clarify wording. Docs **S**. Closes upstream#1094.
- **Docs: horizontal-rule paragraph + Title border override** — usage note. Docs **S**. Closes upstream#1055.
- **Docs: Font.name vs theme fonts on Headings** — add inheritance note. Docs **S**. Closes upstream#1089.
- **Docs: Style-override semantics for cell paragraph_format** — docs note. Docs **S**. Closes upstream#305.
- **Docs: Borderless-table root cause for `add_table` in existing doc** — note; optionally default style='Table Grid'. Docs **S**. Closes upstream#348.
- **Docs: shading per cell** — `_Cell.shading.fill = ...` example. Docs **S**. Closes upstream#434.
- **Docs: nested-table iteration recipe** — iter_block_items snippet. Docs **S**. Closes upstream#670.
- **Table direction/left_indent / fields on form cells / heading number** — misc investigations below.
- **Cell text missing on form-field cells** — add regression test that Cell.text includes form-field result text. Bug **S**. Closes upstream#689.
- **Math equation default font (`w:mathFont`)** — expose on settings/equation. Feature **M**. Closes upstream#1341.
- **`WD_BUILTIN_STYLE.TOC_HEADING` usage** — document with `add_table_of_contents`. Docs **S**. Closes upstream#542.
- **Semi-structured content extraction guidance** — clarify intent; likely docs. Closes upstream#1448.
- **Quickstart example runs verbatim** — end-to-end; fix style names. Docs **S**. Closes upstream#215, #198.
- **`character format application issue`** — awaits minimal repro. Closes upstream#1484.
- **JPEG repro for #1493 crash** — awaits attachment; likely dup of zero-DPI bug. Closes upstream#1493.
- **`_tr` attribute on `_Cell`** — re-add convenience or document replacement. Docs/S. Closes upstream#165.
- **`Paragraph.insert_paragraph_before/after` after text-box** — sugar. Feature **S**. Closes upstream#1119.
- **`del table.rows[i]`** — `_Row.delete()` + `_Rows.__delitem__`. Feature **S**. Closes upstream#279.
- **Demo regression** — audit `src/docx/templates/default.docx` against bundled demo. Bug **M**. Closes upstream#1059.
- **Table spans page width by default** — `preferred_width=None` shortcut in `add_table`. Feature **S**. Closes upstream#315.
- **`Section.paragraphs` iteration** — audit (also covered in features). Closes upstream-PR#755 (cross-ref).
- **Bookmark-targeted insertion helper** — re-evaluate PR#341 scope vs Phase D.13. Feature **M**. Closes upstream-PR#341.
- **Grab-bag manipulation helpers (PR#277)** — diff against current API. Closes upstream-PR#277.
- **Empty-body merge/iteration helper (PR#367)** — likely superseded. Closes upstream-PR#367.
- **`eastAsia` font unit tests** — backfill if missing. Closes upstream-PR#781.
- **Duplicate pages regression** — repro and trace. Closes upstream#989.
- **`Paragraph.add_text` helper** — decide and add. Feature **S**. Closes upstream-PR#71.
- **`Styles.import_from` scope** — see also feature list. Closes upstream#197.
- **Cross-doc paragraph copy** — revisit `Document.append_paragraph(para)` / `copy_from(other)` rewiring rIds. Feature **L**. Closes upstream#182.
- **Equation rels corrupted during external merge recipe** — ensure any future `Document.append_document` re-links equation rels. Bug **L**. Closes upstream#466.

---

## Part C — Features (91 after dedupe)

Pre-dedupe: 140 new-feature-needed entries. Organized by module.

### Image / drawing
- **PathLike acceptance across image entry points** — widen `add_picture`, `Image.from_file`, `Document()` ctor to accept `os.PathLike` (apply `os.fspath()` at entry). Effort **S**. Closes upstream#1544, upstream-PR#1168.
- **WebP image support** — add WebP header parser + content-type registration. Effort **M**. Closes upstream#717, upstream-PR#1423.
- **EMF / WMF / EPS image support** — add EMF + WMF header parser + content-type registration (optional EPS fallback). Effort **M**. Closes upstream#1391, #193, #24, upstream-PR#731, upstream-PR#196.
- **Picture outline/border API** — `InlineShape.outline` / `FloatingImage.outline` (width/rgb/transparency) writing `pic:spPr/a:ln`. Effort **M**. Closes upstream#1510.
- **Picture crop API (`a:srcRect`)** — set at insert time and read/apply on export. Effort **M**. Closes upstream#1463, #1331.
- **Image opacity / alphaModFix** — setter on inline + floating. Effort **S**. Closes upstream#1316.
- **Release aspect-ratio lock** — `lock_aspect_ratio` setter clearing `noChangeAspect`. Effort **S**. Closes upstream#1314.
- **Image drop-shadow / effects** — `InlineShape.effects` / `shadow` writing `a:effectLst/a:outerShdw`. Effort **M**. Closes upstream#419.
- **Linked (external) pictures** — `add_picture(..., link=True, save_with_document=False)` with `r:link` blip + external relationship; also URL/linked-web variant. Effort **M**. Closes upstream#916, #1002.
- **InlineShape.delete() / FloatingImage.delete()** — drop `w:drawing`, prune unused rId + optional part. Effort **M**. Closes upstream#1425, #518.
- **InlineShape.replace_image()** — swap image part while preserving position/size. Effort **M**. Closes upstream#192.
- **`InlineShape.image`** — expose underlying `Image` via rId. Effort **S**. Closes upstream#249.
- **Floating shape by coordinates** — `Paragraph.add_floating_shape(...)` with h/v anchor + offset. Effort **M**. Closes upstream#1414.
- **Preset DrawingML shape writer** — `Document.add_shape(WD_SHAPE, ...)` for roundRect etc. Effort **M**. Closes upstream#1112, #517.
- **Drawing canvas container** — `add_canvas()` for grouped/anchored shapes. Effort **L**. Closes upstream#411.
- **Text box authoring** — `Document.add_text_box()` / `Run.add_text_box()` creating `wps:wsp/wps:txbx` with fill/border/position. Effort **L**. Closes upstream#524.
- **`_Cell.add_picture()`** — convenience on table cells. Effort **S**. Closes upstream#10, upstream-PR#429.

### Tables
- **Column deletion API** — `Table.delete_column(index)` and `_Column.delete()` removing `w:gridCol` + each row's `w:tc` + grid-span updates. Effort **M**. Closes upstream#1500, #441.
- **Row insert at index** — `Table.insert_row(index)` / `_Row.insert_row_before/after()`. Effort **S**. Closes upstream#190.
- **Row add-with-template** — `Table.add_row(source_row=...)` / `_Row.clone()` cloning tc/trPr/run formatting. Effort **M**. Closes upstream#1189, #205.
- **Cell add-to-row** — `_Row.add_cell()` / `_Row.insert_cell(index)`. Effort **S**. Closes upstream#532.
- **Cell split (unmerge)** — `_Cell.split()` clearing gridSpan/vMerge. Effort **M**. Closes upstream#733.
- **`_Cell.add_table(style=...)`** — add `style` kwarg for nested tables. Effort **S**. Closes upstream#1285.
- **Table caption / description (alt-text)** — `Table.alt_text` / `alt_description` via `w:tblCaption` / `w:tblDescription`. Effort **S**. Closes upstream#1048, #921.
- **Table left-indent** — `Table.indent` / `Table.left_indent` mapping `w:tblInd`. Effort **S**. Closes upstream#1144, #586.
- **Table-level default cell margins** — `Table.cell_margins` mirroring `_Cell.margins`. Effort **S**. Closes upstream#1401.
- **Split table** — `Table.split(before_row)` returning new Table + paragraph. Effort **M**. Closes upstream#481.
- **Merged-cell info helper** — `Table.merged_cell_ranges` returning `[(r1,r2,c1,c2), …]`. Effort **S**. Closes upstream-PR#1036.
- **Cached cell grid** — public cached `Table.cells` / lazyproperty `_cells` with invalidation. Effort **M**. Closes upstream#1209, upstream-PR#1051, upstream-PR#565.
- **Fast read-only table iteration** — `Table.iter_rows_fast()` / `TableSnapshot` for huge tables. Effort **M**. Closes upstream#1516.
- **Row-level shading/style** — `_Row.apply_shading` / bulk row styling. Effort **S**. Closes upstream#370.
- **Total-row flag** — expose `tblLook.lastRow` toggle. Effort **S**. Closes upstream#331.
- **Cross-document table copy** — `Document.add_table_copy(table)` / `Document.add_table_from(other_table)` copying CT_Tbl + rewiring rIds. Effort **L**. Closes upstream#612, #270.

### Text / paragraph / run
- **Rendered list-label renderer** — `Paragraph.list_label` / `Document.list_labels()` traversing body in order and formatting per-level counters via `lvlText`. Effort **L**. Closes upstream#1454, #1372, #1365, #614, #554, #590, #471.
- **Paragraph restart-numbering** — `Paragraph.restart_numbering(level, start)` emitting `w:lvlOverride/startOverride`; `NumberingDefinition.new_instance()` too. Effort **S**. Closes upstream#25.
- **Multi-run hyperlinks** — `Hyperlink.add_run(text, style)` writer. Effort **S**. Closes upstream#1515.
- **Hyperlink wrap existing text** — `Run.make_hyperlink` / `Paragraph.insert_hyperlink_at(run, url)` splitting runs. Effort **M**. Closes upstream#610.
- **Hyperlink address/anchor setters** — `Hyperlink.address` / `Hyperlink.fragment` setters updating `w:anchor` and rel target. Effort **S**. Closes upstream#1176.
- **`w:sym` character extraction** — `CT_Sym.__str__` returns derived char; optional `Run.text_with_symbols`. Effort **S**. Closes upstream#1528.
- **DOCPROPERTY field resolution in `paragraph.text`** — substitute core/custom property values. Effort **M**. Closes upstream#1482.
- **`RGBColor.from_string` 3-hex form** — accept 3-digit shorthand. Effort **S**. Closes upstream#1466.
- **`Font.copy_to` / `Run.copy_formatting_from`** — clone `w:rPr` (+ `Font.color` setter). Effort **M**. Closes upstream#1308, upstream-PR#1309.
- **`Paragraph.add_text(text)`** — append-to-last-run convenience. Effort **S**. Closes upstream#8, upstream-PR#71.
- **`Paragraph.font` (paragraph-mark rPr)** — expose `w:pPr/w:rPr` as a Font proxy; also expose `pPr.rPr` directly. Effort **S**. Closes upstream-PR#417, upstream-PR#537.
- **`ParagraphFormat.outline_level`** — get/set `w:outlineLvl` with `WD_OUTLINELVL` enum. Effort **S**. Closes upstream#485, upstream-PR#1393, upstream-PR#1295, upstream-PR#1064.
- **`ParagraphFormat.contextual_spacing`** — `w:contextualSpacing` bool. Effort **S**. Closes upstream#365.
- **`ParagraphFormat.first_line_chars`** — `w:ind/@w:firstLineChars`. Effort **S**. Closes upstream-PR#941.
- **`ParagraphFormat.auto_space_de/dn`** — East-Asian auto-space flags. Effort **S**. Closes upstream#1071.
- **`ParagraphFormat.shading_color`** — mirror `Font.shading_color` on paragraphs. Effort **S**. Closes upstream#1238.
- **`Font.cs_size`** — complex-script font size (`w:szCs`). Effort **S**. Closes upstream#248.
- **`Font.character_scale`** — `w:rPr/w:w` percent. Effort **S**. Closes upstream-PR#1097.
- **`Font.ligatures`** — map `w:rPr/w14:ligatures/@w14:val`. Effort **S**. Closes upstream#1150.
- **`.next_block` / `.previous_block` navigation** — on Paragraph/Table. Effort **S**. Closes upstream#583.
- **`Document.text` helper** — concatenated body text. Effort **S**. Closes upstream#252, #72.
- **SDT-aware body iteration** — optional `include_sdt` / flatten mode surfacing SDT-wrapped paragraphs (also covers TOC). Effort **M**. Closes upstream#1280.
- **`Paragraph.element` / `Table.element` / `_Cell.element` public aliases** — drop the private-access warnings. Effort **S**. Closes upstream#1445.

### Fields / bookmarks / TOC
- **TOC refresh / update-fields-on-open** — `TOC.mark_dirty()` / `Settings.update_fields_on_open`. Effort **S**. Closes upstream#1403.
- **List of Figures / Tables** — `Document.add_list_of_figures` / `add_list_of_tables` emitting `TOC \c "Figure"` / `"Table"`. Effort **M**. Closes upstream#723.
- **altChunk write + read** — `Document.add_alt_chunk`, iteration + form-field altChunk support. Effort **M**. Closes upstream#1317, #1103, upstream-PR#649.

### Sections / page layout
- **`Section.paragraphs`** — paragraphs bounded by a given sectPr. Effort **M**. Closes upstream#181, upstream-PR#755.
- **`Section.delete()` / `sections.pop()`** — merge sectPr into neighbour. Effort **M**. Closes upstream#1348.
- **`Section.copy_header_from(other_section)` + footer twin** — copy header/footer between sections/documents. Effort **M**. Closes upstream#668.
- **Page-break detection** — `Paragraph.page_breaks_inside` / `Table.spans_page_break` using `w:lastRenderedPageBreak`. Effort **M**. Closes upstream#744.

### Charts / equations / OLE / attachments
- **OLE object embedding (write)** — `Run.add_ole_object(path, prog_id, icon)` + embedding part; covers xlsx/pdf/zip attachments. Effort **L**. Closes upstream#1130, #1127, #1023, #743, #713, #294.
- **altChunk extraction** — read-side `Document.attachments` + altChunk iteration. Effort **M**. Closes upstream#1103.
- **Equation edit helpers** — traversal/edit on OMML. Effort **M**. Closes upstream#1235.
- **Chart.replace_data** — programmatic chart data edits. Effort **L**. Closes upstream#1141.

### Numbering / styles
- **Cross-document style import** — `Styles.import_from(source_doc, names=None)` / `import_style(style)` deep-copy + linked styles. Effort **M**. Closes upstream#1375, #1083, #508, #701, #197.
- **Import builtin latent styles** — `Styles.import_builtin(name)` materialising e.g. "List Bullet" from defaults (also FollowedHyperlink latent materialise). Effort **M**. Closes upstream#486.
- **Document default font proxy** — `Styles.document_default_font` over `w:docDefaults/w:rPrDefault`. Effort **M**. Closes upstream#383.
- **Embed fonts (fontTable parts)** — `FontTable.add_embedded_font(path)` wiring parts + rels; preserve embedded fonts on save. Effort **L**. Closes upstream#1231, #1307.
- **`next_paragraph_style` auto-apply** — document or auto-apply on `add_paragraph`. Effort **S**. Closes upstream#888.

### Packaging / I/O / documents
- **Strict OOXML support** — detect Strict namespaces and rewrite to Transitional on package open. Effort **L**. Closes upstream#1520, #693.
- **Flat-OPC support** — `<pkg:package>` reader/writer. Effort **L**. Closes upstream#892.
- **.dotx / .dotm templates** — accept `WML_TEMPLATE_MAIN` + macro variant in `Document()`; `Document.from_template(...)` helper; optional content-type switch on save. Effort **M**. Closes upstream#1532, #363, upstream-PR#1537, upstream-PR#522, upstream-PR#523.
- **Deterministic / reproducible save** — `Document.save(reproducible=True)` with fixed zip timestamps, sorted rels/parts. Effort **M**. Closes upstream#1042, upstream-PR#810.
- **File-not-found vs not-a-ZIP diagnostics** — distinct exceptions in `Document()`. Effort **S**. Closes upstream#1410.
- **Cross-document body/paragraph/image copy** — `Document.append_document(other)` / `Document.append_body(other)` / `Document.append_paragraph(para)` importing body, styles, numbering, images, watermarks, with rId rewiring. Effort **L**. Closes upstream#1457, #558, #543, #437, #460, #44, #709.
- **Custom xpath namespaces** — let callers pass `namespaces=` to `BaseOxmlElement.xpath`. Effort **S**. Closes upstream-PR#622.
- **XPath pre-compilation cache** — `_XP("expr")` cache for hot xpaths. Effort **M**. Closes upstream#342.

### Settings / metadata
- **Extended properties (app.xml)** — `Document.extended_properties` (Company, Manager, Application, TotalTime, …) with `ExtendedPropertiesPart`. Effort **M**. Closes upstream#911, #572, upstream-PR#1206.
- **`DocumentStatistics.pages`** — page count from cached app.xml property. Effort **M**. Closes upstream#1084.
- **Document-level language** — `Settings.theme_font_language` (+ optional `Document.set_language` convenience). Effort **M**. Closes upstream#727.
- **Spell/grammar-check toggles** — `Settings.hide_spelling_errors` / `hide_grammatical_errors`. Effort **S**. Closes upstream#1177.
- **Auto-hyphenation toggle** — `Settings.auto_hyphenation` + related flags. Effort **S**. Closes upstream#680.
- **`Settings.doc_vars`** — `w:docVars` Mapping proxy. Effort **S**. Closes upstream#127.
- **Timezone-aware comment timestamps** — optional `date` param on `Comments.add_comment` / `Comment.add_reply`. Effort **S**. Closes upstream#1533.
- **Strip metadata on new docs** — `Document(include_metadata=False)` / `core_properties.clear_all()`. Effort **S**. Closes upstream#1464.

### Tracked changes / revisions
- **Author/date on add_paragraph/add_run** — wrap in `w:ins` when track-changes is on. Effort **M**. Closes upstream#1025.

### Other / cross-cutting
- **Public `_Row.index` / `_Column.index`** — rename/alias `_index`. Effort **S**. Closes upstream#112.
- **Docs: raise on missing style** — note in `styles-using.rst`. Effort **S**. Closes upstream#170.

---

## Part D — Out-of-scope (134, grouped)

One line per item.

### User-support / usage questions (38)
- upstream#1512 — answer is `Document.replace_regex`/search module.
- upstream#1469 — matplotlib + image layout question.
- upstream#1449 — needs layout engine (physical pages).
- upstream#1444 — "dump document as reconstructing code" generator.
- upstream#1441 — Kivy/KivyMD compat question.
- upstream#1483 — RTL in tkinter/customtkinter (not python-docx).
- upstream#1438 — "keep table on one page" layout/pagination.
- upstream#1431 — pympler sizing hits lxml internals, not library.
- upstream#1422 — Word compat setting, user self-resolved.
- upstream#1409 — read-API discoverability / docs.
- upstream#1406 — bolding substrings via existing Paragraph API.
- upstream#1359 — field-insertion snippet help.
- upstream#1357 — docx↔XML round-trip how-to.
- upstream#1354 — user self-resolved.
- upstream#1344 — "split on Heading 1" recipe.
- upstream#1332 — runtime page number (layout).
- upstream#1233 — filter runs by style + size (no gap).
- upstream#1108 — user's `.isnumeric` bug.
- upstream#1099 — wrong PyPI package name.
- upstream#1079 — tutorial-level try/except question.
- upstream#1075 — cell linear-index tutorial.
- upstream#1070 — local-ref misunderstanding.
- upstream#947 — copy runs between documents usage.
- upstream#922 — `font.size` None when inherited.
- upstream#918 — generic copy/transform how-to.
- upstream#738 — `.text` on List Paragraph usage.
- upstream#736 — stacking-tables layout question.
- upstream#730 — how-to on custom style.
- upstream#721 — `.doc` vs `.docx` in uwsgi.
- upstream#674 — centering picture usage.
- upstream#682 — colored/underlined `paragraph.text` usage.
- upstream#698 — general editing question.
- upstream#707 — XML-copy across docs.
- upstream#616 — `document.save(path)` confusion.
- upstream#615 — empty body / unclear.
- upstream#585 — empty body, style-setting usage.
- upstream#382 — "cover page" gallery usage.
- upstream#358 — placeholder-replace usage.
- upstream#354 — row-number increment usage.
- upstream#426 — per-char font granularity tutorial.
- upstream#321 — per-column alignment Word limitation.
- upstream#298 — orientation setter documented behaviour.

### Meta / project-status / contribution
- upstream#1492 — upstream status.
- upstream#1477 — user clarification post, no ask.
- upstream#1461 — stray cross-post.
- upstream#1481 — repo-policy security improvements.
- upstream#1404 — OSS-Fuzz harness offer.
- upstream#1426 — YouTube tutorial feedback.
- upstream#1154 — upstream status question.
- upstream#1068 — wheels on PyPI request (upstream only).
- upstream#1019 — replace Lena test image (Debian hygiene).
- upstream#644 — contributor intro.
- upstream#926 — bulk black reformat request.
- upstream#526 — license question (MIT).
- upstream#448 — add Python 3.5 to tox.
- upstream#609 — Python 2.6 restoration.

### Word limitations / rendering engine
- upstream#1468 — TOC clickability after unoconv.
- upstream#1250 — chart icon thumbnail.
- upstream#1228 — automerging adjacent tables.
- upstream#1047 — detect page-break inside table (layout).
- upstream#966 — remove "last page" (no static pages).
- upstream#979 — auto "(continue)" when cell splits.
- upstream#995 — Word DPI interpretation.
- upstream#588 — x/y character position.
- upstream#989 — visual duplicate pages (kept also in investigations).
- upstream#662 — overlay text on image.

### Dependency / install / packaging issues
- upstream#1077 — lxml on M1.
- upstream#1034 — PyInstaller packaging.
- upstream#983 — PyInstaller template bundling (there is an investigation variant).
- upstream#402 — SyntaxError on py3.6 install.
- upstream#520 — pip-in-wrong-interpreter.
- upstream#657 — install without pip.

### Unrelated conversions / integrations
- upstream#1479 — HTML iframe embed inside docx.
- upstream#1465 — S3 incremental streaming.
- upstream#1429 — docx→PDF converter help.
- upstream#1339 — pandas `DataFrame.style` → styled docx.
- upstream#1320 — ipynb→docx conversion.
- upstream#1303 — cross-doc Table copy recipe (feature dup).
- upstream#1268 — ZIP attachment via OLE package.
- upstream#1435 — Visio (.vsd/.vsdx) embedding via OLE (covered by OLE feature).
- upstream#1082 — tables from python-pdf2docx.
- upstream#1080 — cv2/docx import-order.
- upstream#1057 — python-docx-template escaping.
- upstream#1054 — Document object → PDF without disk.
- upstream#1050 — docx → JPG/GIF.
- upstream#1049 — run VBA macro from Python.
- upstream#1033 — change List Bullet size via Word style.
- upstream#1027 — rPr at pPr level (out-of-scope-tagged; feature var).
- upstream#913 — OLE stream extraction via olefile.
- upstream#549 — 3D matplotlib from Word table.
- upstream#550 — duplicate of #549.
- upstream#591 — open/save encrypted docx.
- upstream#596 — transliterate Unicode.
- upstream#501 — HTML→docx via TinyMCE.
- upstream#500 — drive Word.app on macOS.
- upstream#397 — export docx → txt.
- upstream#381 — streaming HTTP delivery of save.
- upstream#49 — markdown→docx conversion.
- upstream#223 — printing a Document.

### Stale / duplicate / superseded / editorial
- upstream#1517 — incomplete issue body.
- upstream#1447 — wheel `RECORD` already present.
- upstream#1352 — Sphinx object.inv intersphinx tweak.
- upstream#1033, #666 etc (above).
- upstream#993 — quickstart style-id UserWarning (already in Bugs docs area).
- upstream#990 — quickstart snippet relies on undefined namedtuple.
- upstream#959 — caller's source encoding.
- upstream#929 — empty body.
- upstream#504 — "Table Grid" style missing in default Word installs.
- upstream#476 — Word re-borders custom table style on save.
- upstream#474 — docs typo (covered in Bugs Part A).
- upstream#389 / #385 — py2 unicode confusion.
- upstream#379 — `Document.close()` optional (small feature) — accept if trivial.
- upstream#357 — amp escaping already handled.
- upstream#128 — upstream-docs landing-page example.

### PRs out of scope
- upstream-PR#1545 — standalone `agent-harness/` tree (fork has its own CI agents).
- upstream-PR#1501 — classifier bumps.
- upstream-PR#1471 — README attribution.
- upstream-PR#1446 — `element` alias PR (see feature #1445).
- upstream-PR#1400 — stray "Questions bank" file.
- upstream-PR#1371 — quickstart table-style tweak.
- upstream-PR#1361 — Sphinx tooling bump.
- upstream-PR#1353 — index.rst one-line tweak.
- upstream-PR#1322 — Dependabot jinja2 bump.
- upstream-PR#1272 — `huge_tree=True` flip (see Bugs opt-in variant).
- upstream-PR#1174 — legacy-tree version bump.
- upstream-PR#1165 — README mailing-list link.
- upstream-PR#1126 — text.rst typo.
- upstream-PR#841 — Python 2.6 fix.
- upstream-PR#608 — StringIO/BytesIO py2/3 docs tweak.
- upstream-PR#653 — enum docs normalisation.
- upstream-PR#706 — missing full-stop in api-concepts.rst.
- upstream-PR#208 — docs example dup of PR#207.
- upstream-PR#207 — "modify existing document" docs example.
- upstream-PR#35 — minor docs snippet fix.
