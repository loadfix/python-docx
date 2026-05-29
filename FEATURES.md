# Features

`loadfix/python-docx` is a fork of
[python-docx](https://github.com/python-openxml/python-docx) that extends the
library with footnotes and endnotes, tracked changes, bookmarks, fields,
content controls, charts, equations, SmartArt, watermarks, digital signatures,
accessibility tooling, cross-document operations, and many more OOXML
capabilities that were previously out of reach.

This document is the full, single-page catalogue of what the library can do
today. Each section covers one feature area, opens with a short overview,
shows a copy-pasteable snippet against a fresh `Document()`, and then lists
the public methods, properties, and classes that make up that surface.
Items marked `[Added in 2026.05.0]` are additions from this fork — every
other item is inherited from the upstream base.

**Table of contents**

- [Opening and saving documents](#opening-and-saving-documents)
- [Paragraphs](#paragraphs)
- [Runs and text](#runs-and-text)
- [Fonts and character formatting](#fonts-and-character-formatting)
- [Paragraph formatting](#paragraph-formatting)
- [Hyperlinks](#hyperlinks)
- [Tables](#tables)
- [Lists and numbering](#lists-and-numbering)
- [Styles](#styles)
- [Inline images](#inline-images)
- [Floating images and shapes](#floating-images-and-shapes)
- [Charts](#charts)
- [SmartArt](#smartart)
- [Equations](#equations)
- [Sections and page layout](#sections-and-page-layout)
- [Headers and footers](#headers-and-footers)
- [Comments](#comments)
- [Footnotes and endnotes](#footnotes-and-endnotes)
- [Bookmarks](#bookmarks)
- [Fields and cross-references](#fields-and-cross-references)
- [Table of contents](#table-of-contents)
- [Tracked changes](#tracked-changes)
- [Content controls (SDT)](#content-controls-sdt)
- [Form fields](#form-fields)
- [Watermarks](#watermarks)
- [Captions](#captions)
- [Mail merge](#mail-merge)
- [Document properties](#document-properties)
- [Settings](#settings)
- [Themes](#themes)
- [Permissions and protection](#permissions-and-protection)
- [Ink annotations](#ink-annotations)
- [Embedded objects and attachments](#embedded-objects-and-attachments)
- [Font table](#font-table)
- [Web settings](#web-settings)
- [Glossary (building blocks)](#glossary-building-blocks)
- [Digital signatures](#digital-signatures)
- [Accessibility](#accessibility)
- [Document outline](#document-outline)
- [Document statistics](#document-statistics)
- [Readability metrics](#readability-metrics)
- [Search and replace](#search-and-replace)
- [CSS-selector queries](#css-selector-queries)
- [Cross-document operations](#cross-document-operations)
- [Packaging and I/O options](#packaging-and-io-options)
- [API concepts](#api-concepts)

---

## Opening and saving documents

The top-level `docx.Document()` factory opens a `.docx`, `.docm`, `.dotx`, or
`.dotm` package, or — when called with no argument — creates a fresh document
from the bundled default template. Strict-OOXML packages are transparently
translated to Transitional on open and Flat-OPC (`<pkg:package>`) XML input is
auto-detected. `Document.save()` serialises back to a path or stream, with
optional Flat-OPC or reproducible (byte-identical) output. Documents support
the context-manager protocol and expose a `huge_tree` escape hatch for very
large files plus a `recover=True` mode that tolerates malformed XML.

```python
from docx import Document

# open with default template
with Document() as document:
    document.add_heading("Hello", level=1)
    document.add_paragraph("A paragraph.")
    document.save("out.docx")

# open an existing file (path may be str, pathlib.Path, or file-like)
document = Document("report.docx", huge_tree=False, include_metadata=True)

# derive a new document from a template
document = Document.from_template("corporate.dotx")

# reproducible save (byte-identical for the same content)
document.save("out.docx", reproducible=True)

# Flat-OPC single-XML output
document.save("out.xml", flat_opc=True)
```

- `docx.Document(docx=None, recover=False, huge_tree=False, include_metadata=True, password=None, strict=False)` — Factory returning a `docx.document.Document`. `recover=True`, `huge_tree=True`, `include_metadata=False`, `os.PathLike` paths, `.dotx`/`.dotm` templates, Strict-OOXML, and Flat-OPC inputs are all `[Added in 2026.05.0]`. `password=` decrypts an ECMA-376 Agile-Encryption (password-protected) `.docx` via the optional `python-ooxml-crypto` dependency. `[Added in 2026.05.10]`. `strict=True` opts into ECMA-376 Strict conformance tracking so `Document.is_strict` returns `True` and a subsequent `save(strict=None)` preserves the class; Strict packages are always auto-detected and translated to Transitional on open regardless of the flag. `[Added in 2026.05.11]`
- `docx.Document.from_template(template)` — Open a `.dotx`/`.dotm` and return a document whose main-part content-type is switched to the matching non-template variant. `[Added in 2026.05.0]`
- `Document.save(path_or_stream, flat_opc=False, reproducible=False, password=None, strict=None, compatibility=None)` — Write the document. `flat_opc` and `reproducible` are `[Added in 2026.05.0]`. `password=` encrypts the output using ECMA-376 Agile Encryption via the optional `python-ooxml-crypto` dependency. `[Added in 2026.05.10]`. `strict=None` (default) preserves the package's current `is_strict` flag; `True` / `False` override per call. Byte-level emission is always Transitional today; the flag is recorded on the package for round-trip preservation. `[Added in 2026.05.11]`. `compatibility=` opts into older-Word compatibility mode by stamping `settings.xml/w:compat/w:compatSetting[@w:name="compatibilityMode"]` with the matching integer (`"Word 2003"` → 11, `"Word 2007"` → 12, `"Word 2010"` → 14, `"Word 2013"` → 15, `"Word 2016"` → 16; raw ints accepted). When the target predates Word 2010, the modern threaded-comments parts (`commentsIds.xml` / `commentsExtensible.xml` / `commentsExtended.xml`) are stripped so the older client opens the file without parser errors. **A compat-mode save is best-effort:** it tells Word to *open* the file as if authored under the older release and prevents the modern UI affordances from showing, but features the older client cannot render (SmartArt, OMML equations, content controls, …) are left in place and may render as placeholders. `[Added in 2026.05.dev0]` — closes #94.
- `Document.is_strict` — `True` when the package was loaded as ECMA-376 Strict (auto-detected by the Strict-namespace sniff, or via explicit `strict=True` on the factory). Writable; assigning to it flips the class recorded on the underlying `OpcPackage`. `[Added in 2026.05.11]`
- `Document.close()` — Drop transient state (tracked-changes contexts). Safe to call more than once. `[Added in 2026.05.0]`
- `Document.__enter__` / `Document.__exit__` — Context-manager support. `[Added in 2026.05.0]`
- `Document.recovery_warnings` — List of parser warnings collected when `recover=True` was used. `[Added in 2026.05.0]`
- `Document.repair(path_or_stream, strategy='best-effort')` — Best-effort recovery loader for damaged packages. Returns a `(Document, RepairReport)` tuple. `strategy='best-effort'` (default) drops unparseable parts, fixes common XML defects (orphan `w:bookmarkEnd`, bad encoding declarations, illegal control bytes), reconstructs truncated zips, and prunes dangling rel targets; `strategy='strict'` matches the existing `Document(...)` factory and raises on the first defect; `strategy='truncate'` keeps everything that parses and discards the rest. `RepairReport.repaired` / `unrecoverable` / `parts_dropped` enumerate what happened. `[Added in 2026.05.13]` — closes #92.
- `Document.from_html(source, clean=True)` / `Document.from_html_string(html, clean=True)` — Build a `Document` from an HTML file (path, `os.PathLike`, binary or text file-like) or in-memory string. Stdlib-only parser (`html.parser`) — no `BeautifulSoup` dependency. Element mapping: `<h1>`-`<h6>` → `Heading 1`-`Heading 6`; `<strong>`/`<b>` → bold; `<em>`/`<i>` → italic; `<u>` → underline; `<a href>` → hyperlink (`http`/`https`/`mailto` only — other schemes drop to plain text); `<ul>`/`<ol>` → `List Bullet`/`List Number` items; `<table>`/`<tr>`/`<td>` → Word table (cell content is plain text in this minimal importer); `<img src="data:…">` → embedded picture (remote URLs degrade to alt-text — they are not fetched); `<blockquote>` → `Quote` style; `<code>`/`<pre>` → `Courier New` runs / paragraphs (`<pre>` preserves whitespace). `clean=True` (default) strips `<script>`/`<style>`/comments and drops `class`/`id` attributes; `style` attributes are honoured only for `color: #aabbcc` (best-effort). LaTeX import is **not** supported by this method — use `docx.math.OMath` for programmatic equation authoring. `[Added in 2026.05.14]` — closes #95.
- `Document.stream(source)` — Bounded-memory **read-only** loader for very large `.docx` packages. Returns a `docx.streaming.StreamingDocument` usable as a context manager. `source` may be a path, `os.PathLike`, binary file-like, or a `bytes` payload. The body of `word/document.xml` is iterated via `lxml.iterparse` so peak memory stays bounded regardless of body size — paragraphs and tables are released as soon as the consumer has yielded. `StreamingDocument.paragraphs` and `.tables` are forward-only generators (NOT sequences). `.headers`, `.footers`, `.sections`, and `.styles` remain eager (small parts). `StreamingDocument.save()` raises `StreamingNotMutableError` — re-open via `docx.Document(...)` to mutate. Decision tree: prefer `Document(...)` for typical bodies and any mutation; prefer `Document.stream()` for hundreds-of-MB read-only passes. `[Added in 2026.05.13]` — closes #93.
- `docx.streaming.StreamingDocument` / `docx.streaming.StreamingNotMutableError` — Re-exports from `docx.streaming` for explicit imports. `[Added in 2026.05.13]`
- `docx.exceptions.EncryptedDocumentError` — Raised when opening a password-protected `.docx` without a correct password, or when `python-ooxml-crypto` is required but not installed. `[Added in 2026.05.0]`
- `docx.exceptions.RmsProtectedDocumentError` — Subclass of `EncryptedDocumentError`, raised when opening a file wrapped in Azure RMS / AIP / IRM protection (not decryptable with a password). `[Added in 2026.05.10]`
- `docx.exceptions.PythonDocxError` / `InvalidSpanError` / `InvalidXmlError` — Library-specific exceptions.
- `docx.exceptions.DocxError` — Structured, LLM-friendly base error carrying machine-readable `code`, human `message`, fuzzy-match `suggestion`, conceptual `location`, and `operation` name. Subclasses (`StyleNotFoundError`, `StyleDuplicateError`, `StyleTypeMismatchError`, `LatentStyleNotFoundError`, `BuiltinStyleNotFoundError`, `BookmarkNotFoundError`, `FontNotFoundError`, `FontFamilyInvalidError`, `FontEmbedEmptyError`, `ThemeTokenInvalidError`, `InvalidColorError`, `InvalidBrightnessError`, `OutOfRangeError`, `ValueOutOfRangeError`, `NotAWordFileError`, `NotAWordTemplateError`) multi-inherit from the legacy built-in (`KeyError` / `ValueError` / `IndexError`) so existing `except` blocks continue to catch them. `__all_codes__` enumerates every stable error code; `to_dict()` returns a JSON-serialisable view for LLM repair loops. `[Added in 2026.05.13]` — closes #12.

---

## Paragraphs

Paragraphs are the most common block in the document body, each cell, and
every header/footer/footnote/endnote/comment story. The fork extends the
classic `Document.add_paragraph()` / `insert_paragraph_before()` API with a
symmetric `insert_paragraph_after()`, `insert_table_before()` /
`insert_table_after()`, `delete()`, page-break helpers, caption insertion,
TOC insertion, field insertion, content-control insertion, bookmarks,
permission ranges, `w:next`-style auto-application, and a stable-ID
fingerprint for tools that track paragraphs across save/load.

```python
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

document = Document()
p = document.add_paragraph("First paragraph.", style="Normal")
p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

# insert before and after
before = p.insert_paragraph_before("inserted before")
after = p.insert_paragraph_after("inserted after")
after.style = "Intense Quote"

# delete & replace
before.delete()

# stable id survives save/reload so long as position + text don't change
print(p.stable_id)

# bookmark the paragraph
p.add_bookmark("chapter-1")
document.save("out.docx")
```

- `Document.add_paragraph(text="", style=None, track_author=None, bind_to=None)` — Append a new paragraph. `track_author` wraps the inserted run in `w:ins`. `bind_to` records a smart-placeholder record so `{customer.name}`, `{date:short}`, `{property:Title}`, `{i}`, and `{customer.address.line1}` tokens in `text` resolve at every save against the bound record / document properties / iteration index, while preserving the source string for round-trip re-binding (#68). `[Added in 2026.05.0]` for `track_author` and `w:next` auto-style handling. `[Added in 2026.05.13]` for `bind_to`.
- `Document.bind(record=None, iteration=None)` — Bind / re-bind a record to the document for smart-placeholder resolution. Returns `self` for chaining; subsequent `save()` re-resolves every previously-stamped token against the new record. `[Added in 2026.05.13]`
- `Document.add_heading(text, level=1)` — Shortcut for `add_paragraph` with `"Heading N"` / `"Title"` style.
- `Document.add_page_break()` — Append a paragraph containing only a page break.
- `Document.add_caption(text, label="Figure", style="Caption")` — Append a numbered `SEQ`-field caption paragraph. `[Added in 2026.05.0]`
- `Paragraph.insert_paragraph_before(text=None, style=None)` — Insert immediately before this paragraph.
- `Paragraph.insert_paragraph_after(text=None, style=None)` — Insert immediately after. `[Added in 2026.05.0]`
- `Paragraph.insert_table_before(rows, cols, style=None, width=None)` / `insert_table_after(...)` — Insert a sibling table. `[Added in 2026.05.0]`
- `Paragraph.add_caption_before(text, label="Figure", style="Caption")` / `add_caption_after(...)` — Insert a caption adjacent to this paragraph. `[Added in 2026.05.0]`
- `Paragraph.insert_table_of_contents_before(levels=(1,3))` / `insert_table_of_contents_after(...)` — Insert a TOC paragraph adjacent to this one. `[Added in 2026.05.0]`
- `Paragraph.insert_section_break(start_type=WD_SECTION.NEW_PAGE)` / `remove_section_break()` — Add/remove a `w:sectPr` inside this paragraph. `[Added in 2026.05.0]`
- `Paragraph.delete()` — Remove this paragraph from its parent. `[Added in 2026.05.0]`
- `Paragraph.clear()` — Remove all content while keeping paragraph-level formatting.
- `Paragraph.alignment` — `WD_PARAGRAPH_ALIGNMENT` (see [enum](docs/api/enum/WdAlignParagraph.rst)).
- `Paragraph.style` — Read/write paragraph style as a `ParagraphStyle` or name.
- `Paragraph.text` — Read/write plain text (replaces all content on set).
- `Paragraph.paragraph_format` — `ParagraphFormat` proxy — indent, spacing, borders, frame, etc.
- `Paragraph.font` — Paragraph-mark `rPr` proxy. `[Added in 2026.05.0]`
- `Paragraph.runs` / `Paragraph.all_runs` — Direct-child runs vs every visible run (descends into hyperlinks, fields, SDTs, ins/del). `all_runs` is `[Added in 2026.05.0]`.
- `Paragraph.hyperlinks` — List of `Hyperlink` proxies.
- `Paragraph.drawings` — List of `Drawing` children. `[Added in 2026.05.0]`
- `Paragraph.floating_images` — List of `FloatingImage` for each `wp:anchor`. `[Added in 2026.05.0]`
- `Paragraph.fields` / `Paragraph.form_fields` — Fields and legacy form fields in this paragraph. `[Added in 2026.05.0]`
- `Paragraph.content_controls` — Inline SDTs. `[Added in 2026.05.0]`
- `Paragraph.equations` — OMML expressions. `[Added in 2026.05.0]`
- `Paragraph.ink_annotations` / `Paragraph.embedded_objects` — Read-only proxies. `[Added in 2026.05.0]`
- `Paragraph.rendered_page_breaks` / `Paragraph.page_breaks_inside` / `Paragraph.contains_page_break` / `Paragraph.has_page_break` / `Paragraph.clear_page_breaks()` — Page-break introspection and mutation.
- `Paragraph.next_block` / `Paragraph.previous_block` — Walk sibling blocks. `[Added in 2026.05.0]`
- `Paragraph.iter_inner_content()` — Yield runs and hyperlinks in document order.
- `Paragraph.rsid` — Word's editing-session revision-save ID. `[Added in 2026.05.0]`
- `Paragraph.stable_id` — 16-char hex fingerprint stable across save/reload. `[Added in 2026.05.0]`
- `Paragraph.has_section_break` — True if paragraph carries a `w:sectPr`. `[Added in 2026.05.0]`
- `Paragraph.element` — Public alias for the underlying `w:p` element. `[Added in 2026.05.0]`

---

## Runs and text

A run (`<w:r>`) is the smallest styled unit of text. The fork adds
`Run.split()` for mid-run edits, `Run.delete()`, `Run.make_hyperlink()`,
`Run.add_symbol()` and `.symbols`, ruby-annotation access, stable IDs, and a
`copy_formatting_from()` helper.

```python
from docx import Document

document = Document()
p = document.add_paragraph()
r = p.add_run("Hello world")
r.bold = True
r.italic = True
r.font.size = 140000  # EMU (or use Pt(14))

# split at char offset 5 → two runs: "Hello" and " world"
left, right = r.split(5)
right.italic = False

# insert a symbol (Unicode char code) in a specific font
left.add_symbol(0x2603, font="Segoe UI Symbol")  # snowman

# delete a run entirely
right.delete()

document.save("out.docx")
```

- `Paragraph.add_run(text=None, style=None, track_author=None)` — Append a run. `track_author` is `[Added in 2026.05.0]`.
- `Paragraph.add_text(text)` — Append `text` onto the last run instead of creating a new one. `[Added in 2026.05.0]`
- `Run.text` — Read/write. `\t` / `\n` / `\r` map to `w:tab` / `w:br`.
- `Run.bold` / `Run.italic` / `Run.underline` — Tri-state (True/False/None for "inherit").
- `Run.style` — Character style.
- `Run.clear()` — Remove all child text/runs.
- `Run.add_tab()` — Insert a tab.
- `Run.add_break(break_type=WD_BREAK.LINE)` — Line / page / column / wrap break.
- `Run.add_picture(path_or_stream, width=None, height=None, link=False, save_with_document=True, url=None)` — Inline picture in this run. `link`, `save_with_document`, `url` and `os.PathLike` support are `[Added in 2026.05.0]`.
- `Run.add_text_box(width=None, height=None, text=None)` — Append a DrawingML text box. `[Added in 2026.05.0]`
- `Run.add_ole_object(ole_path, prog_id, icon_path=None)` — Embed an OLE payload. `[Added in 2026.05.0]`
- `Run.add_symbol(char_code, font)` — Insert a `w:sym`. `[Added in 2026.05.0]`
- `Run.symbols` — Iterator of `Symbol` proxies. `[Added in 2026.05.0]`
- `Run.text_with_symbols` — Text including symbol glyphs. `[Added in 2026.05.0]`
- `Run.equations` / `Run.ruby_annotations` — Inline OMML / ruby. `[Added in 2026.05.0]`
- `Run.split(offset)` — Split into two runs at `offset`, preserving formatting. `[Added in 2026.05.0]`
- `Run.delete()` — Remove this run. `[Added in 2026.05.0]`
- `Run.make_hyperlink(url=None, anchor=None)` — Wrap this run in a hyperlink. `[Added in 2026.05.0]`
- `Run.mark_comment_range(last_run, comment_id)` — Place `commentRangeStart`/`commentRangeEnd` markers.
- `Run.copy_formatting_from(source)` — Copy `rPr` from another run. `[Added in 2026.05.0]`
- `Run.contains_page_break` / `Run.iter_inner_content()` — Inline content iteration.
- `Run.formatting_change` — `FormattingChange` for `w:rPrChange`. `[Added in 2026.05.0]`
- `Run.rsid` / `Run.stable_id` — Editing-session ID and stable fingerprint. `[Added in 2026.05.0]`
- `docx.ruby.RubyAnnotation` — Base text, ruby text, alignment, language. `[Added in 2026.05.0]`
- `docx.text.symbol.Symbol` — Font and character-code reader. `[Added in 2026.05.0]`

---

## Fonts and character formatting

`Font` (via `Run.font` or `Paragraph.font`) exposes every `w:rPr` child
covered by OOXML, including the fork additions for run borders, run-level
background shading, East-Asian layout, explicit language tags, character
scale, ligatures, and kerning.

```python
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX, WD_UNDERLINE, WD_BORDER_STYLE

document = Document()
run = document.add_paragraph().add_run("Styled text")

font = run.font
font.name = "Calibri"
font.size = Pt(12)
font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
font.underline = WD_UNDERLINE.SINGLE
font.highlight_color = WD_COLOR_INDEX.YELLOW
font.small_caps = True

# fork-era extras
font.shading_color = RGBColor(0xFF, 0xFF, 0xCC)
font.border_style = WD_BORDER_STYLE.SINGLE
font.border_color = RGBColor(0x00, 0x00, 0x00)
font.border_width = Pt(0.5)
font.language = "en-US"
font.character_scale = 90   # 90 %
font.kerning = Pt(10)

document.save("out.docx")
```

- `Font.name` / `Font.size` / `Font.color` / `Font.highlight_color` — Core identity and size.
- `Font.bold` / `italic` / `underline` / `strike` / `double_strike` / `superscript` / `subscript` / `all_caps` / `small_caps` / `shadow` / `outline` / `emboss` / `imprint` / `hidden` / `web_hidden` / `math` / `snap_to_grid` / `no_proof` / `spec_vanish` — Tri-state boolean toggles.
- `Font.character_spacing` / `Font.kerning` / `Font.character_scale` — Letter-spacing controls (`character_scale` and `ligatures` are `[Added in 2026.05.0]`).
- `Font.ligatures` — `"all"`, `"standardContextual"`, etc. `[Added in 2026.05.0]`
- `Font.cs_size` / `Font.complex_script` / `Font.cs_bold` / `Font.cs_italic` — Complex-script properties.
- `Font.shading_color` — Run-level background color. `[Added in 2026.05.0]`
- `Font.border_style` / `Font.border_color` / `Font.border_width` / `Font.border_space` / `Font.remove_border()` — Run borders. `[Added in 2026.05.0]`
- `Font.name_cs` / `Font.name_east_asia` / `Font.name_far_east` — Script-specific typeface overrides.
- `Font.language` / `Font.east_asian_language` / `Font.bidi_language` / `Font.remove_language()` — Per-run language tags. `[Added in 2026.05.0]`
- `Font.rtl` / `Font.right_to_left` — Right-to-left flag. `[Added in 2026.05.0]`
- `Font.east_asian_layout` / `Font.set_east_asian_layout(...)` / `Font.remove_east_asian_layout()` — Two-lines-in-one, kinsoku, etc. `[Added in 2026.05.0]`
- `Font.copy_to(target)` — Copy every `rPr` property onto another `Font`. `[Added in 2026.05.0]`
- `docx.text.font.EastAsianLayout` — Proxy for `w:eastAsianLayout`. `[Added in 2026.05.0]`
- `docx.dml.color.ColorFormat` — RGB / theme-color / tint / shade.
- Enums: `WD_COLOR_INDEX`, `WD_UNDERLINE`, `WD_BORDER_STYLE`.

---

## Paragraph formatting

`Paragraph.paragraph_format` exposes `ParagraphFormat`, the Word
"Paragraph…" dialog mapped to OOXML. The fork adds paragraph borders,
text-frame controls, contextual spacing, outline level, RTL, kinsoku /
word-wrap, first-line-chars, and auto-space-DE/DN.

```python
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_BORDER_STYLE, WD_LINE_SPACING

document = Document()
p = document.add_paragraph("A well-formatted paragraph.")
pf = p.paragraph_format

pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
pf.line_spacing = 1.15
pf.space_after = Pt(8)
pf.first_line_indent = Inches(0.5)
pf.keep_with_next = True

# fork-era additions
pf.borders.top.style = WD_BORDER_STYLE.SINGLE
pf.borders.top.width = Pt(0.5)
pf.contextual_spacing = True
pf.right_to_left = False

document.save("out.docx")
```

- `ParagraphFormat.alignment` — `WD_PARAGRAPH_ALIGNMENT`.
- `ParagraphFormat.first_line_indent` / `left_indent` / `right_indent` — Lengths.
- `ParagraphFormat.line_spacing` / `line_spacing_rule` — Spacing controls.
- `ParagraphFormat.space_before` / `space_after` — Paragraph spacing.
- `ParagraphFormat.keep_together` / `keep_with_next` / `widow_control` / `page_break_before` — Pagination toggles.
- `ParagraphFormat.contextual_spacing` — `[Added in 2026.05.0]`
- `ParagraphFormat.outline_level` — `WD_OUTLINELVL` or int 0–9. `[Added in 2026.05.0]`
- `ParagraphFormat.right_to_left` / `kinsoku` / `word_wrap` / `auto_space_de` / `auto_space_dn` / `first_line_chars` — Bidi and East-Asian typography. `[Added in 2026.05.0]`
- `ParagraphFormat.tab_stops` — `TabStops` collection.
- `ParagraphFormat.borders` — `ParagraphBorders` (top/bottom/left/right/between/bar). `[Added in 2026.05.0]`
- `ParagraphFormat.frame` / `ParagraphFormat.set_frame(...)` / `ParagraphFormat.remove_frame()` — Text frames. `[Added in 2026.05.0]`
- `Paragraph.drop_cap` / `Paragraph.add_drop_cap(letter, mode=DROP, lines=3)` — Drop caps (`w:framePr/@w:dropCap`). `.drop_cap` returns a `DropCap` proxy with `mode`/`lines`/`x`/`y`/`width`/`height`/`wrap`/`horizontal_anchor`/`vertical_anchor`, or `None` when the paragraph is not a drop-cap frame. `.add_drop_cap(letter)` splits the paragraph: inserts a drop-cap frame paragraph before `self` containing `letter`, then strips that leading character from `self`'s first `w:t`. `[Added in 2026.05.0]`

---

## Hyperlinks

Hyperlinks can be created from scratch, read off existing paragraphs, or
wrapped around an existing run slice. Both external URLs and internal
anchors (bookmark names) are supported, and URL fragments (`#section`) are
exposed as a first-class attribute.

```python
from docx import Document

document = Document()
p = document.add_paragraph("Visit ")
link = p.add_hyperlink(url="https://example.com/#intro", text="our site",
                       tooltip="example homepage")
p.add_run(".")

# external URL with the ergonomic wrapper (mailto/tel/http auto-prepended)
document.add_paragraph().add_url("alice@example.com")          # -> mailto:
document.add_paragraph().add_url("+1 555 0100")                # -> tel:
document.add_paragraph().add_url("www.example.com")            # -> http://

# autolink: split free text on URL/email matches
document.add_paragraph().add_text_with_links(
    "See https://example.com or email alice@example.com for more."
)

# internal anchor (string, Bookmark, or heading Paragraph)
document.add_paragraph().add_hyperlink(anchor="chapter-1", text="Chapter 1")
heading = document.add_heading("Q1 Review")
document.add_paragraph().add_link_to(heading, text="back to Q1")

# wrap part of an existing run as a hyperlink
r = document.add_paragraph().add_run("click here to read more")
p2 = r._parent
p2.insert_hyperlink_at(r, url="https://docs.example", start=0, end=10)

document.save("out.docx")
```

- `Paragraph.add_hyperlink(url=None, text=None, style="Hyperlink", anchor=None, tooltip=None)` — Append a new hyperlink. `[Added in 2026.05.0]`. `tooltip` arg `[Added in 2026.05.12]`.
- `Paragraph.add_link_to(target, text=None, style="Hyperlink", tooltip=None)` — Internal-link wrapper accepting a `Bookmark`, a heading `Paragraph` (auto-bookmarks the heading), or a bookmark-name string. `[Added in 2026.05.12]`.
- `Paragraph.add_url(url, text=None, style="Hyperlink", tooltip=None)` — External-link wrapper that auto-prepends `mailto:` / `tel:` / `http://` for email-shape, phone-shape, and `www.` arguments. `[Added in 2026.05.12]`.
- `Paragraph.add_text_with_links(text, style="Hyperlink")` — Append `text` and auto-detect URLs / emails as hyperlinks. Returns the new runs and hyperlinks in document order. `[Added in 2026.05.12]`.
- `Paragraph.insert_hyperlink_at(run, url=None, anchor=None, start=None, end=None)` — Wrap (part of) an existing run in a hyperlink, splitting as needed. `[Added in 2026.05.0]`
- `Run.make_hyperlink(url=None, anchor=None)` — Wrap a run as a hyperlink. `[Added in 2026.05.0]`
- `Paragraph.hyperlinks` — List of `Hyperlink` in document order.
- `Hyperlink.url` / `Hyperlink.address` / `Hyperlink.fragment` — URL parts; `address`/`fragment` are editable.
- `Hyperlink.tooltip` — Read/write `w:tooltip` attribute (the popup hover text). `[Added in 2026.05.12]`.
- `Hyperlink.runs` / `Hyperlink.text` / `Hyperlink.contains_page_break` / `Hyperlink.add_run(...)` — Content access and extension.

---

## Cross-format linked content (INCLUDETEXT)

`Paragraph.link_to(target_url)` writes a Word `INCLUDETEXT` complex
field that points at an external Excel cell, Excel table column, or
PowerPoint slide. Word re-resolves these on open;
`Document.update_links()` re-resolves them in-process via the sibling
`xlsx` (and, when present, `pptx`) packages so the cached field
result reflects the live workbook value at save time. `[Added in
2026.05.13]`.

```python
from docx import Document

doc = Document()
para = doc.add_paragraph("Q1 revenue: ")
para.link_to("revenue.xlsx#RevenueQ1!B5")          # specific cell
para.link_to("revenue.xlsx#RevenueQ1[Total]")      # table column total
para.link_to("summary.pptx#slide-3")               # specific slide

doc.update_links(base_dir="./reports")             # refresh on save

for link in doc.linked_targets:
    print(link.kind, link.url, link.cached_text)

doc.save("report.docx")
```

- `Paragraph.link_to(target_url, cached_result=None, mark_dirty=True)` — Append an `INCLUDETEXT` complex field that points at an external resource and return a `LinkedTarget` proxy. `[Added in 2026.05.13]`.
- `Document.linked_targets` — Iterate every `LinkedTarget` in the document body in document order. `[Added in 2026.05.13]`.
- `Document.update_links(base_dir=None)` — Best-effort refresh: re-resolves every link's target via the sibling `xlsx` / `pptx` packages and rewrites the cached field result. Returns the count of fields updated. Failed resolutions leave the cached text alone. `[Added in 2026.05.13]`.
- `LinkedTarget.url` / `LinkedTarget.kind` / `LinkedTarget.parsed` / `LinkedTarget.cached_text` / `LinkedTarget.field` — Read-only views of the underlying field's URL, parsed shape, cached display text, and raw `Field` proxy.
- `LinkedTarget.resolve(base_dir=None)` — Best-effort fetch the live value at the URL without writing it back. Returns `None` when the target can't be resolved.
- `LinkedTarget.refresh(base_dir=None)` — `resolve()` then write the result into the field's cached text.

Recognised URL shapes:

| URL                                  | `kind`              |
| ------------------------------------ | ------------------- |
| `workbook.xlsx#Sheet1!B5`             | `xlsx-cell`         |
| `workbook.xlsx#'Sheet With Spaces'!A1` | `xlsx-cell`        |
| `workbook.xlsx#Table1[Column]`         | `xlsx-table-column` |
| `deck.pptx#slide-3`                    | `pptx-slide`        |
| anything else                          | `unknown`           |

`unknown` URLs round-trip cleanly (saved + reloaded with no loss) but
`update_links()` skips them — there's no resolver. PowerPoint slides
return `"[Slide N]"` (or `"[Slide N: <title>]"` when the sibling
`pptx` package is installed); real slide rendering needs PowerPoint
itself, which is out of scope for the in-process path.

---

## Tables

Tables are first-class blocks. The fork extends them with per-cell and
whole-table borders, cell margins, text direction, merged-cell reads, row
height setters, header-row repeat, table-style flags (banded rows/columns),
autofit behavior, alt text, copy (including cross-document), split at a row,
CRUD operations, and stable IDs.

```python
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_BORDER_STYLE

document = Document()
tbl = document.add_table(rows=2, cols=3, style="Table Grid")
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.alt_text = "Quarterly results"

tbl.cell(0, 0).text = "Region"
tbl.cell(0, 1).text = "Q1"
tbl.cell(0, 2).text = "Q2"

# merge two cells
tbl.cell(1, 0).merge(tbl.cell(1, 1))

# borders + cell shading
tbl.borders.top.style = WD_BORDER_STYLE.SINGLE
tbl.borders.top.color = RGBColor(0x44, 0x44, 0x44)
tbl.cell(0, 0).shading.fill_color = RGBColor(0xEE, 0xEE, 0xEE)

# row height
tbl.rows[0].height = Inches(0.4)
tbl.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
tbl.rows[0].is_header = True

document.save("out.docx")
```

- `Document.add_table(rows, cols, style=None)` — Append a new table.
- `Document.add_table_copy(other_table)` / `Document.add_table_from(...)` — Deep-copy a table (possibly from another document), rewiring images and importing styles. `[Added in 2026.05.0]`
- `Table.alignment` / `Table.direction` / `Table.table_direction` — Placement controls.
- `Table.autofit` / `Table.autofit_behavior` / `Table.allow_autofit` / `Table.preferred_width` — Layout. `autofit_behavior`, `allow_autofit`, `preferred_width` are `[Added in 2026.05.0]`.
- `Table.autofit_to_content()` / `Table.autofit_to_window()` / `Table.fixed_width_columns(widths)` / `Table.autofit_char_width_twips` — Real column-width computation. `autofit_to_content()` walks every cell, measures the longest line, and assigns each column a width proportional to its widest content (default estimate: `len(line) * 100` twips, roughly Calibri-11pt average advance; override via `autofit_char_width_twips`). Writes `w:gridCol/@w:w` and `w:tc/w:tcPr/w:tcW` in lock-step and switches the table to `w:tblLayout w:type="fixed"`. Returns the list of computed `Length` widths. `autofit_to_window()` emits `w:tblLayout w:type="autofit"` plus `w:tblW w:type="pct" w:w="5000"` and writes each single-span `w:tcW w:type="pct"` with its equal share of the 5000 (`5000 // n`). `fixed_width_columns(widths)` writes `w:tblLayout w:type="fixed"` plus explicit per-column widths (raises `ValueError` on length mismatch). `[Added in 2026.05.11]` (R15-3)
- `Table.indent` / `Table.left_indent` — Table indentation. `indent` is `[Added in 2026.05.0]`.
- `Table.style` / `Table.style_flags` — Style application and banding flags (`first_row`, `last_row`, `first_column`, `last_column`, `no_horizontal_banding`, `no_vertical_banding`). `style_flags` is `[Added in 2026.05.0]`.
- `Table.borders` / `Table.set_borders(...)` — `TableBorders` proxy. `[Added in 2026.05.0]`
- `Table.cell_margins` — Per-table cell-margin defaults. `[Added in 2026.05.0]`
- `Table.alt_text` / `Table.alt_description` — Accessibility fields. `[Added in 2026.05.0]`
- `Table.cell(row, col)` / `Table.row_cells(i)` / `Table.column_cells(i)` / `Table.rows` / `Table.columns` / `Table.cells` — Access.
- `Table.add_row(source_row=None)` / `Table.insert_row(index)` / `Table.add_column(width)` / `Table.delete_column(index)` — CRUD. `insert_row`, `delete_column` are `[Added in 2026.05.0]`.
- `Table.split(before_row)` — Split into two tables at a boundary. `[Added in 2026.05.0]`
- `Table.delete()` — Remove from document. `[Added in 2026.05.0]`
- `Table.merged_cell_ranges` — Tuples of `(top_row, top_col, bottom_row, bottom_col)`. `[Added in 2026.05.0]`
- `Table.merge_range(row0, col0, row1, col1)` — Merge a rectangular block in one call; returns the origin `_Cell`. Corners may be supplied in any diagonal order. `[Added in 2026.05.11]`
- `Table.spans_page_break` — `True` if the rendered table crosses a page break.
- `Table.stable_id` / `Table.formatting_change` — Stable fingerprint and tracked-formatting proxy. `[Added in 2026.05.0]`
- `Table.next_block` / `Table.previous_block` — Block-level navigation. `[Added in 2026.05.0]`
- `_Cell.add_paragraph(...)` / `_Cell.add_table(...)` / `_Cell.add_picture(...)` — Nested content.
- `_Cell.merge(other)` / `_Cell.split()` / `_Cell.is_merge_origin` / `_Cell.merge_origin` / `_Cell.grid_span` — Merge handling. Merge-origin APIs are `[Added in 2026.05.0]`.
- `_Cell.merge_down(count=1)` / `_Cell.unmerge_vertical()` — Vertical-merge authoring. `merge_down` marks this cell `w:vMerge="restart"` and the `count` cells below as `w:vMerge="continue"`; `unmerge_vertical` strips `w:vMerge` across the full span (walking up to the origin if called on a continuation). `[Added in 2026.05.11]`
- `_Cell.is_merged_origin` / `_Cell.is_merged_continuation` — Plain-boolean companions to the tri-state `is_merge_origin`. Useful when iterating and the "not merged" case should be |False| rather than |None|. `[Added in 2026.05.11]`
- `_Cell.borders` / `_Cell.margins` / `_Cell.set_margins(...)` / `_Cell.remove_margins()` — Cell-level borders and margins. `[Added in 2026.05.0]`
- `_Cell.shading.fill_color` / `_Cell.shading.pattern` — Background. `[Added in 2026.05.0]`
- `_Cell.text_direction` / `_Cell.vertical_alignment` / `_Cell.width` / `_Cell.text` — Cell properties (`text_direction` is `[Added in 2026.05.0]`).
- `_Cell.is_tracked_insertion` / `_Cell.is_tracked_deletion` / `_Cell.formatting_change` / `_Cell.stable_id` — Track-changes and stable-id hooks. `[Added in 2026.05.0]`
- `_Row.height` / `_Row.height_rule` / `_Row.is_header` / `_Row.allow_break_across_pages` — Row-level properties. `height` setter, `is_header`, `allow_break_across_pages` are `[Added in 2026.05.0]`.
- `TableBorders` — `top` / `bottom` / `left` / `right` / `inside_h` / `inside_v` → `BorderElement.style` / `.width` / `.color` / `.space`. `[Added in 2026.05.0]`
- `CellBorders`, `CellShading`, `CellMargins`, `TableCellMargins`, `TableStyleFlags` — Helper proxies. `[Added in 2026.05.0]`
- `Document.add_dataframe(df, style="executive", alternating_rows=None, header_color=None, header_text_color=None, autofit=True, align=None, number_format=None, show_total_row=False, table_style=None)` — Append a `pandas.DataFrame` as a styled Word table. Built-in styles: `executive` (bold header bar in theme primary, alternating row tint, total row at bottom), `minimal` (header underline only, no fills, monospace numbers), `boxed` (full grid borders, light header tint), `striped` (zebra rows, no borders). Number-format DSL accepts the standard Python format-spec mini-language (`$,.1f`, `0.0%`, `,d` …) for numeric columns plus a small DSL for date columns (`MMM YYYY`, `YYYY-MM-DD HH:mm:ss` …). Total-row aggregator accepts `True` / `"sum"` / `"mean"` / `"count"` / `"none"` or a per-column `{col: op}` mapping. `pandas` is **not** a hard dependency — DataFrame input is sniffed at runtime and the helper raises `ImportError` when pandas is missing. `[Added in 2026.05.13]`

```python
import pandas as pd
from docx import Document

df = pd.DataFrame(
    {
        "Region": ["AMER", "APAC", "EMEA"],
        "Revenue": [1234.5, 987.6, 654.3],
        "Growth": [0.087, 0.121, -0.034],
    }
)
doc = Document()
doc.add_dataframe(
    df,
    style="executive",
    alternating_rows=True,
    align={"Revenue": "right", "Region": "left"},
    number_format={"Revenue": "$,.1f", "Growth": ".1%"},
    show_total_row=True,
)
```

---

## Lists and numbering

`Document.numbering` exposes a read/write wrapper around `numbering.xml` so
that you can create new abstract numbering definitions, allocate instances,
apply them to paragraphs, restart numbering on demand, and ask Word-style
"what label would this paragraph show?" for each paragraph.

```python
from docx import Document
from docx.enum.text import WD_NUMBER_FORMAT

document = Document()
numbering = document.numbering
definition = numbering.add_numbering_definition(
    levels=[
        {"format": WD_NUMBER_FORMAT.DECIMAL, "text": "%1.", "start": 1},
        {"format": WD_NUMBER_FORMAT.LOWER_LETTER, "text": "%2)", "start": 1},
    ]
)

p1 = document.add_paragraph("First")
p2 = document.add_paragraph("Second")
p3 = document.add_paragraph("A sub-item")
definition.apply_to(p1, level=0)
definition.apply_to(p2, level=0)
definition.apply_to(p3, level=1)
p2.restart_numbering(start=5)

# rendered label text (e.g. "1.", "5.", "a)")
print(p2.list_label)
print(document.list_labels())

document.save("out.docx")
```

- `Document.numbering` — `Numbering` proxy. Creates `numbering.xml` on demand. `[Added in 2026.05.0]`
- `Numbering.add_numbering_definition(levels=...)` — Add an abstract definition. `[Added in 2026.05.0]`
- `Numbering.add_abstract_definition(format, start=1, lvl_text='%1.', alignment=WD_ALIGN_PARAGRAPH.LEFT)` — Shortcut for the common single-level list (no `w:num` is created — pair with `add_definition()`). `[Added in 2026.05.10]`
- `Numbering.add_definition(abstract_num_id, style_name=None)` — Allocate a fresh `w:num` instance; optionally bind a paragraph-style name onto level-0's `w:pStyle`. Returns a `NumInstance`. `[Added in 2026.05.10]`
- `Numbering.abstract_definitions` — Alias of `.definitions`. `[Added in 2026.05.10]`
- `Numbering.abstract_definition(abstract_num_id)` — Lookup by id. `[Added in 2026.05.10]`
- `Numbering.num_instances` / `Numbering.num_instance(num_id)` — `NumInstance` proxies over `w:num`. `[Added in 2026.05.10]`
- `Numbering.next_num_id()` / `Numbering.next_abstract_num_id()` — Public id allocators (peek without mutating). `[Added in 2026.05.10]`
- `Numbering.definitions` / iteration — Existing definitions.
- `NumberingDefinition.apply_to(paragraph, level=0)` — Apply a numbering to a paragraph. `[Added in 2026.05.0]`
- `NumberingDefinition.new_instance()` / `NumberingDefinition.levels` / `NumberingDefinition.level(ilvl)` — Instance and level access.
- `AbstractNumberingDefinition` — Alias of `NumberingDefinition` for clarity when you mean *the* abstract definition. `[Added in 2026.05.10]`
- `NumInstance` — Per-`w:num` proxy (`num_id`, `abstract_num_id`, `definition`, `level_overrides`, `set_level_override(ilvl, start)`). `[Added in 2026.05.10]`
- `LevelOverride` — Proxy for `w:lvlOverride` (`ilvl`, `start_override`). `[Added in 2026.05.10]`
- `Level.number_format` / `Level.text` / `Level.start` / `Level.indent` / `Level.ilvl` — Per-level properties.
- Assigning a list-backed paragraph style via `paragraph.style = "MyList"` allocates a fresh `w:num` so the new list restarts from `1` rather than continuing a sibling list — only when the paragraph has no explicit `w:numPr`. `[Added in 2026.05.10]`
- `Paragraph.list_level` / `Paragraph.list_format` / `Paragraph.numbering_format` / `Paragraph.list_label` — Read paragraph's current list settings. `[Added in 2026.05.0]`
- `Paragraph.restart_numbering(level=None, start=1)` — Restart the counter. `[Added in 2026.05.0]`
- `Document.list_labels()` — `{id(p): label}` for every numbered paragraph in the body (one pass). `[Added in 2026.05.0]`
- `Document.add_list_of_figures(caption_label="Figure")` / `Document.add_list_of_tables(caption_label="Table")` — Append `TOC \c` fields. `[Added in 2026.05.0]`
- `docx.numbering.ListLabelRenderer` — Low-level label renderer used by the properties above. `[Added in 2026.05.0]`
- Enum: `WD_NUMBER_FORMAT` (decimal, roman, letter, bullet, etc.).

---

## Styles

`Document.styles` is a `Styles` collection covering paragraph, character,
table, and numbering styles. The fork adds style import across documents,
`link_style` / `next_style` / `is_redefined`, a document-default font
accessor, a `Style.delete()`, and direct access to the XML style element.

```python
from docx import Document
from docx.enum.style import WD_STYLE_TYPE

document = Document()
styles = document.styles

# create a new paragraph style
my_style = styles.add_style("Summary", WD_STYLE_TYPE.PARAGRAPH)
my_style.base_style = styles["Normal"]
my_style.font.bold = True

# apply it
document.add_paragraph("Summary text", style="Summary")

# query
for style in styles:
    print(style.name, style.type)

# document-wide default font
styles.document_default_font.name = "Calibri"

document.save("out.docx")
```

- `Document.styles` — `Styles` collection.
- `Styles.add_style(name, style_type, builtin=False)` — Add a new style.
- `Styles.default(style_type)` — The document's default for a given type.
- `Styles.get_by_id(style_id, style_type)` / `Styles.get_style_id(style_or_name, style_type)` — Id lookups.
- `Styles.document_default_font` — `Font` proxy for `docDefaults/rPrDefault`. `[Added in 2026.05.0]`
- `Styles.import_from(other_doc, names)` / `Styles.import_style(style)` / `Styles.import_builtin(name)` — Cross-document import. `[Added in 2026.05.0]`
- `Styles.latent_styles` — `LatentStyles` collection.
- `BaseStyle.name` / `.style_id` / `.type` / `.builtin` / `.priority` / `.hidden` / `.locked` / `.quick_style` / `.unhide_when_used` — Common metadata.
- `BaseStyle.delete()` — Remove the style.
- `BaseStyle.link_style` / `BaseStyle.next_style` / `BaseStyle.is_redefined` — Style-mapping properties. `[Added in 2026.05.0]`
- `CharacterStyle.base_style` / `CharacterStyle.font` — Character-style specifics.
- `ParagraphStyle.paragraph_format` / `ParagraphStyle.next_paragraph_style` — Paragraph-style specifics.
- `LatentStyles` / `_LatentStyle` — Latent style collection.
- `docx.styles.BabelFish` — UI ↔ internal style-name translation.
- Enum: `WD_STYLE_TYPE`, `WD_BUILTIN_STYLE`.

---

## Inline images

`Document.add_picture()` / `Run.add_picture()` append an inline image. All
common formats are supported, including the fork additions of **SVG**,
**WebP**, **EMF**, **WMF**, and **EPS**. Linked (external) pictures, image
replacement, outline/border, crop, opacity, shadow, aspect-ratio lock, alt
text, and delete are all first-class.

```python
from docx import Document
from docx.shared import Inches, Pt, RGBColor

document = Document()
shape = document.add_picture("logo.png", width=Inches(2))
shape.alt_text = "Company logo"
shape.title = "Logo"

shape.outline.color = RGBColor(0, 0, 0)
shape.outline.width = Pt(1)
shape.outline.transparency = 0.25

shape.crop.set(left=0.05, right=0.05)    # fractions (0..1)
shape.opacity = 0.9
shape.effects.shadow.apply(blur_radius=Pt(4), distance=Pt(2))
shape.lock_aspect_ratio = True

# replace the image bytes, keep the drawing
shape.replace_image("new-logo.png")

document.save("out.docx")
```

- `Document.add_picture(path_or_stream, width=None, height=None, link=False, save_with_document=True, url=None)` — Append a new paragraph with an inline picture. `[Added in 2026.05.0]` for `link`, `save_with_document`, `url`, and `os.PathLike`.
- `Run.add_picture(...)` — Inline picture in an existing run.
- `Document.inline_shapes` — `InlineShapes` collection for iteration/indexing.
- `InlineShape.width` / `InlineShape.height` / `InlineShape.type` / `InlineShape.image` — Core picture data.
- `InlineShape.alt_text` / `InlineShape.title` — Accessibility metadata. `[Added in 2026.05.0]`
- `InlineShape.opacity` / `InlineShape.lock_aspect_ratio` — Visual controls. `[Added in 2026.05.0]`
- `InlineShape.locks` / `FloatingImage.locks` — `ShapeLocks` proxy exposing `no_select`, `no_move`, `no_resize`, `no_rotate`, `no_change_aspect`, `no_edit_points`, `no_adjust_handles`, `no_change_arrowheads`, `no_change_shape_type`, `no_group`, `no_ungroup`, `no_text_edit`, aggregate `locked`, and `lock_all()` / `unlock_all()`. Writes to `pic:cNvPicPr/a:picLocks`; setting a lock to `False` removes the attribute so it does not survive a round-trip. `[Added in 2026.05.10]`
- `InlineShape.outline` — `PictureOutline` (style, color, width, transparency). `[Added in 2026.05.0]`
- `InlineShape.crop` — `PictureCrop` (left/top/right/bottom, `set(...)`). `[Added in 2026.05.0]`
- `InlineShape.effects.shadow` — `ShadowFormat` (blur, distance, angle, color, `apply(...)`, `clear()`). `[Added in 2026.05.0]`
- `InlineShape.delete(part=None)` — Remove the drawing and prune orphan image parts. `[Added in 2026.05.0]`
- `InlineShape.replace_image(path_or_stream)` — Swap the blob, keeping the drawing. `[Added in 2026.05.0]`
- `docx.drawing.Picture` — Generic picture proxy for canvas/group contexts.
- Supported formats: PNG, JPEG, GIF, BMP, TIFF, **SVG** (via `docx.image.svg`), **WebP**, **EMF**, **WMF**, and **EPS**. SVG/WebP/EMF/WMF/EPS are `[Added in 2026.05.0]`.

---

## Floating images and shapes

Floating images are anchored (`<wp:anchor>`) rather than inline, with
horizontal/vertical anchor frames, offsets, and wrap style. The fork also
adds DrawingML preset shapes, text boxes, canvases, and read-only access to
group shapes.

```python
from docx import Document
from docx.shared import Inches
from docx.enum.shape import WD_SHAPE, WD_ANCHOR_H, WD_ANCHOR_V, WD_WRAP_TYPE

document = Document()
p = document.add_paragraph()

# floating image at a specific page offset
img = p.add_floating_shape(
    "banner.png",
    x=Inches(1), y=Inches(2),
    width=Inches(4), height=Inches(2),
    h_anchor=WD_ANCHOR_H.PAGE, v_anchor=WD_ANCHOR_V.PAGE,
    wrap=WD_WRAP_TYPE.SQUARE,
)

# inline preset shape
shape = document.add_shape(
    WD_SHAPE.ROUNDED_RECTANGLE,
    width=Inches(3), height=Inches(1),
    text="A rounded rectangle",
)

# text box
tb = document.add_text_box(width=Inches(3), height=Inches(1), text="Note")

# canvas with two sub-shapes
canvas = document.add_canvas(width=Inches(5), height=Inches(3))

document.save("out.docx")
```

- `Paragraph.add_floating_image(path, width=None, height=None, position=None)` — Add `wp:anchor` image. `[Added in 2026.05.0]`
- `Paragraph.add_floating_shape(path, x=0, y=0, width=None, height=None, h_anchor=..., v_anchor=..., wrap=...)` — Coordinate-first helper. `[Added in 2026.05.0]`
- `Paragraph.add_shape(shape_type, width=None, height=None, text=None)` / `Document.add_shape(...)` — Append a DrawingML preset shape. `[Added in 2026.05.0]`
- `Document.add_text_box(...)` / `Run.add_text_box(...)` — Append a text-box shape. `[Added in 2026.05.0]`
- `Document.add_canvas(width=None, height=None)` — Append a canvas (`wpc:wpc`). `[Added in 2026.05.0]`
- `FloatingImage.width` / `.height` / `.horizontal_anchor` / `.vertical_anchor` / `.horizontal_offset` / `.vertical_offset` / `.offset` / `.position` / `.wrap_type` / `.type` / `.opacity` / `.lock_aspect_ratio` / `.alt_text` / `.title` / `.outline` / `.crop` / `.effects` / `.delete(part=None)` — Floating picture surface. `[Added in 2026.05.0]`
- `docx.drawing.Drawing` — Base proxy for `<w:drawing>`; exposes `.picture`, `.shape`, `.smart_art`, `.chart`, etc.
- `docx.drawing.WordprocessingShape` — DrawingML shape with `add_paragraph()`, `text`, `paragraphs`. `[Added in 2026.05.0]`
- `docx.drawing.GroupShape` — Read-only group (`wpg:grpSp`) iteration. `[Added in 2026.05.0]`
- `docx.drawing.Canvas` — Canvas proxy with `add_shape(...)`. `[Added in 2026.05.0]`
- Enums: `WD_SHAPE`, `WD_ANCHOR_H`, `WD_ANCHOR_V`, `WD_WRAP_TYPE`, `WD_DRAWING_TYPE`.

---

## Charts

`Document.add_chart()` creates a `.chartPart` with numeric data for the
supported chart types (bar, stacked bar, column, stacked column, line, pie);
`Document.charts` reads existing charts of any type; `Chart.replace_data()`
rewrites the category and series values in place. `[Added in 2026.05.0]`.

```python
from docx import Document
from docx.shared import Inches
from docx.chart import WD_CHART_TYPE

document = Document()
chart = document.add_chart(
    WD_CHART_TYPE.COLUMN,
    categories=["Q1", "Q2", "Q3", "Q4"],
    series_data={
        "North": [10, 12, 14, 11],
        "South": [8, 13, 9, 15],
    },
    width=Inches(5), height=Inches(3),
)
print(chart.chart_type)
for s in chart.series:
    print(s.name, s.values)

# replace values later
chart.replace_data(categories=["Q1", "Q2", "Q3", "Q4"],
                   series_data={"Total": [18, 25, 23, 26]})

document.save("out.docx")
```

- `Document.add_chart(chart_type, categories, series_data, width=None, height=None)` — Append a new chart. `[Added in 2026.05.0]`
- `Document.charts` — List of `Chart` proxies in document order. `[Added in 2026.05.0]`
- `Chart.chart_type` / `Chart.title` / `Chart.has_legend` / `Chart.series` / `Chart.categories` — Reads. `[Added in 2026.05.0]`
- `Chart.replace_data(categories, series_data)` — Rewrite all data in place. `[Added in 2026.05.0]`
- `ChartSeries.name` / `.values` / `.categories` — Per-series reads. `[Added in 2026.05.0]`
- `ChartSeries.format` — `ooxml_chart.ChartFormat` proxy over the
  series' `c:spPr`. Authors DrawingML fill / gradient directly on the
  chart series via the shared `python-ooxml-chart` 0.5 API:
  `series.format.fill.apply_gradient(stops=[(0.0, "FF0000"), (1.0,
  "0000FF")], angle=45.0)` writes a multi-stop `<a:gradFill>` that
  survives `Document.save` / reopen. `ooxml-chart` 0.5 gradient-fill
  support adopted: `FormatFill.gradient`, `FormatFill.apply_gradient`,
  `GradientFill.stops` / `.angle` / `.type`, `FILL_TYPE`, and
  `XL_GRADIENT_FILL_TYPE` are exported from `ooxml_chart`. The
  shared `a:gradFill` / `a:gs` / `a:gsLst` / `a:lin` `CT_*` classes
  are registered in docx's element-class lookup so read-back from a
  saved chart part reconstructs typed proxies. `[Added in 2026.05.11]`
- Enum: `docx.chart.WD_CHART_TYPE` (`BAR`, `BAR_STACKED`, `COLUMN`, `COLUMN_STACKED`, `LINE`, `PIE`).
- `Document.add_chart_inline(kind, data, x=None, y=None, title=None, subtitle=None, size=None, show_values=False, show_legend="auto", secondary_axis=None)` — Ergonomic chart authoring with three input shapes (dict, list-of-dicts, `pandas.DataFrame`) and 13 chart kinds: `bar`, `column`, `line`, `area`, `pie`, `donut`, `scatter`, `bubble`, `combo`, `stacked-bar`, `stacked-column`, `stacked-area`, `sparkline` (plus `grouped-bar` / `grouped-column` aliases). `pandas` is **not** a hard dependency — DataFrame input is sniffed at runtime. `secondary_axis=[<series-name>, ...]` plots the named series against a right-hand value-axis (typically used with `kind="combo"`). `[Added in 2026.05.13]`

```python
from docx import Document

document = Document()
document.add_chart_inline(
    kind="bar",
    data={"AMER": 14.2, "APAC": 8.1, "EMEA": 9.0},
    title="Q1 Revenue by Region",
    subtitle="($B)",
    size=(6.0, 4.0),
)

# Multi-series with secondary axis (pandas optional)
document.add_chart_inline(
    kind="combo",
    data=[
        {"r": "AMER", "rev": 100.0, "mar": 18.0},
        {"r": "APAC", "rev": 80.0, "mar": 22.0},
    ],
    x="r",
    y=["rev", "mar"],
    secondary_axis=["mar"],
)

document.save("out.docx")
```

Chart-kind decision tree (also in the `docx.chart_inline` module docstring):

| Goal | Kind |
| --- | --- |
| Compare values across categories | `bar` / `column` |
| Same, totals share a band | `stacked-bar` / `stacked-column` |
| Trend over a continuous x | `line` / `area` |
| Trend with totals stacked | `stacked-area` |
| Whole-of-100% breakdown | `pie` / `donut` |
| Two numeric variables, no time order | `scatter` |
| Three numeric variables (x, y, size) | `bubble` |
| Different y-scales on the same chart | `combo` (with `secondary_axis`) |
| Tiny in-line trend, no axes / labels | `sparkline` |

---

## SmartArt

SmartArt is read *and* authorable. `Document.smart_art` walks every
`<w:drawing>` that references a `dgm:relIds` diagram and returns a
`SmartArt` proxy carrying `.nodes`, `.text`, and the underlying
diagram-data partname. `Document.add_smart_art(layout_name)` appends a
new SmartArt at the end of the document body, backed by four freshly
minted companion parts (`data`, `layout`, `colors`, `quickStyle`) under
`word/diagrams/`. The returned `SmartArt` is populated one content node
at a time via `SmartArt.add_node(text)`. Supported layout families are
`"list"`, `"cycle"` and `"process"` — each selects a Word built-in
layout algorithm keyed by its canonical `loTypeId` URN; Word's own
layout engine handles rendering, so the embedded `layout1.xml` copy
exists to satisfy package requirements rather than to drive geometry.
`[Added in 2026.05.0]` (read-side); `add_smart_art` / `add_node`
`[Added in 2026.05.8]`.

```python
from docx import Document

# -- authoring --
document = Document()
diagram = document.add_smart_art("process")
diagram.add_node("Plan")
diagram.add_node("Build")
diagram.add_node("Ship")
document.save("with-smartart.docx")

# -- reading --
document = Document("with-smartart.docx")
for diagram in document.smart_art:
    print(diagram.data_partname)
    print(diagram.text)
    for node in diagram.nodes:
        print(" " * node.level, node.text)
```

- `Document.add_smart_art(layout_name, width=None, height=None)` — Return a new
  empty `SmartArt`. `layout_name` is one of `"list"`, `"cycle"`, `"process"`
  (case-insensitive). `[Added in 2026.05.8]`
- `SmartArt.add_node(text)` — Append a top-level content node and return
  its `SmartArtNode`. `[Added in 2026.05.8]`
- `Document.smart_art` / `Document.smart_arts` — List of `SmartArt` (plural
  alias added in 2026.05.10). `[Added in 2026.05.0]`
- `Document.iter_smart_arts()` — Generator yielding every `SmartArt` in
  document order. `[Added in 2026.05.10]`
- `InlineShape.is_smart_art` — |True| when the shape wraps a DrawingML
  diagram (``a:graphicData/@uri`` matches the SmartArt URI). `[Added in
  2026.05.10]`
- `InlineShape.smart_art` — `SmartArt` proxy when
  `is_smart_art` is |True|; |None| otherwise. `[Added in 2026.05.10]`
- `SmartArt.color_transform` / `SmartArt.style_transform` — Typed read-
  side proxies over the companion `diagrams/colorsN.xml` /
  `diagrams/quickStyleN.xml` parts, routed through `python-ooxml-smartart`
  0.3. |None| when the corresponding companion part is not resolvable.
  `[Added in 2026.05.10]`
- `SmartArt.graphic_frame_xml` — Raw bytes of the wrapping `w:drawing`
  element for consumers that want to migrate a SmartArt graphic to
  python-pptx without round-tripping through the docx writer. |None|
  when the SmartArt was constructed without a host drawing. `[Added in
  2026.05.10]`
- `SmartArt.data_partname` / `SmartArt.dm_rId` / `SmartArt.nodes` / `SmartArt.text`. `[Added in 2026.05.0]`
- `SmartArtNode.text` / `.level` / `.model_id` / `.children`. `[Added in 2026.05.0]`

---

## Equations

OMML (Office Math) equations are both read and writable via a minimal
builder API. You can drop a literal OMML string onto a paragraph, or
assemble common structures (fractions, superscripts, radicals) from small
factory functions. `[Added in 2026.05.0]`.

```python
from docx import Document
from docx.equations import build_fraction, build_radical

document = Document()
p = document.add_paragraph("Pythagoras: ")

omml = (
    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    + build_radical(
        build_fraction(numerator_text="a^2 + b^2", denominator_text="1"),
    )
    + "</m:oMath>"
)
p.add_equation(omml)

for eq in document.equations:
    print(eq.text, eq.is_display_mode)

document.save("out.docx")
```

- `Paragraph.add_equation(omml_xml, display_mode=False)` — Append an OMML expression. Accepts an OMML XML string **or** a `docx.math.MathExpr` proxy (raw operators are auto-wrapped in `<m:oMath>`). `[Added in 2026.05.0]` (MathExpr accepted `[Added in 2026.05.12]`)
- `Document.equations` / `Paragraph.equations` / `Run.equations` — Read iterators. `[Added in 2026.05.0]`
- `Equation.text` / `.raw_xml` / `.xml_element` / `.is_display_mode` / `.set_text(...)` / `.replace_identifier(old, new)` / `.swap_children(a, b)` / `Equation.from_omml_xml(...)`. `[Added in 2026.05.0]`
- Builders: `build_identifier`, `build_fraction`, `build_superscript`, `build_subscript`, `build_radical`. `[Added in 2026.05.0]`

### Pythonic equation construction (`docx.math`)

`docx.math` re-exports the `ooxml_math` 0.3.0 proxy layer so callers can
build equations with typed Python objects instead of hand-writing OMML
XML. `[Added in 2026.05.12]`

```python
from docx import Document
from docx.math import Fraction, Lit, Sum, Var, oMath

doc = Document()
p = doc.add_paragraph("sum: ")

expr = oMath(
    Sum(
        body=Fraction(Var("x"), Lit(2)),
        lower=Var("i"),
        upper=Lit("n"),
    )
)
p.add_equation(expr)
```

- `docx.math.MathExpr` — Abstract base for every proxy.
- Leaves: `Var`, `Lit`, `Text`, `Raw`.
- Operator tree: `Fraction`, `Radical`, `Sub`, `Sup`, `SubSup`, `Pre`, `Sum`, `Product`, `Integral`, `Nary`, `Limit`, `FuncApply`, `Delimiter`, `Matrix`, `Accent`, `Bar`, `Box`, `BorderBox`, `Phantom`, `GroupChar`, `EqArray`. `[Bar / Box / BorderBox / Phantom / GroupChar / EqArray added in 2026.05.10 on ooxml-math 0.4.0]`
- Root container: `oMath`.
- Parse dispatch: `from_element(element)` — returns the matching proxy for any OMML element.
- `Paragraph.math_expressions` — Generator yielding a `MathExpr` for each `<m:oMath>` / `<m:oMathPara>` in the paragraph. `[Added in 2026.05.10]`
- `Document.iter_math_expressions()` — Document-wide walk yielding a `MathExpr` for each body equation. `[Added in 2026.05.10]`
- `Paragraph.add_math(expr)` — Insert a math block before the first run (or append when no runs exist). Returns a `MathExpr` proxy around the inserted element. `[Added in 2026.05.10]`

### LaTeX-to-OMML translation (`docx.latex_math`)

`docx.latex_math` ships a minimal LaTeX-to-OMML translator for the common case
of authoring equations in LaTeX. The supported subset covers variables, digit
literals, `+ - * /`, `=`, superscripts and subscripts, `\frac{a}{b}`,
`\sqrt{x}`, parentheses, common Greek letters (`\alpha` … `\omega`,
`\Gamma` … `\Omega`), and `\begin{align}…\end{align}` equation arrays.
Everything else (matrices, integrals with limits, custom commands, full
LaTeX-to-MathML) raises `NotImplementedError` pointing back at the supported
subset. `[Added in 2026.05.11]`

```python
from docx import Document
from docx.latex_math import latex_to_omml

doc = Document()
p = doc.add_paragraph("Euler: ")
p.add_math_from_latex(r"e^{i \pi} + 1 = 0")

# standalone form — returns a CT_OMath element
omath = latex_to_omml(r"\frac{a+b}{2}")
doc.save("out.docx")
```

- `docx.latex_math.latex_to_omml(latex)` — Translate a LaTeX math body to a
  `CT_OMath` element. Raises `LatexMathError` on malformed input,
  `NotImplementedError` on unsupported constructs. `[Added in 2026.05.11]`
- `docx.latex_math.LatexMathError` — `ValueError` subclass raised for
  malformed input (unbalanced braces, stray separators, …). `[Added in 2026.05.11]`
- `Paragraph.add_math_from_latex(latex)` — Append an OMML expression built
  from *latex*. Returns the inserted `MathExpr` proxy. `[Added in 2026.05.11]`

---

## Sections and page layout

Every document has at least one section. `Document.sections` is a sequence
with indexing, iteration, and `pop()`; each `Section` carries page size and
orientation, margins, gutter, header/footer distances, columns, page
borders, line numbering, paper source, document grid, text direction,
even/odd headers, first-page headers/footers, and (fork-era) footnote /
endnote overrides, watermarks, and section-copy helpers.

```python
from docx import Document
from docx.enum.section import WD_ORIENTATION, WD_SECTION_START
from docx.shared import Inches, Pt

document = Document()
section = document.sections[0]
section.page_height = Inches(11)
section.page_width = Inches(8.5)
section.orientation = WD_ORIENTATION.PORTRAIT
section.left_margin = Inches(1)
section.right_margin = Inches(1)

# two-column layout with a divider line
section.set_columns(
    count=2,
    widths=[Inches(3), Inches(3)],
    space=Inches(0.5),
    separator=True,
)

# group every pgMar attribute on a single proxy
section.page_margins.gutter = Inches(0.25)
section.page_margins.header = Inches(0.5)

# page size with the paper-size code for round-tripping
section.page_size.width = Inches(8.5)
section.page_size.height = Inches(11)
section.page_size.code = 1  # Windows DEVMODE.dmPaperSize code for "Letter"

# page border
section.set_page_border("top", style="single", width=Pt(1))

# line numbering from 1, every 5 lines
section.set_line_numbering(count_by=5, start=1, distance=Inches(0.2))

# right-to-left binding (gutter on the right)
section.rtl_gutter = True

# append a new section that breaks to a new page
document.add_section(WD_SECTION_START.NEW_PAGE)

document.save("out.docx")
```

- `Document.sections` — `Sections` sequence. `pop(index=-1)` is `[Added in 2026.05.0]`.
- `Document.add_section(start_type=WD_SECTION.NEW_PAGE)` — Append a new section.
- `Section.start_type` / `Section.orientation` / `Section.page_height` / `Section.page_width` / `Section.left_margin` / `Section.right_margin` / `Section.top_margin` / `Section.bottom_margin` / `Section.header_distance` / `Section.footer_distance` / `Section.gutter` — Page metrics.
- `Section.page_margins` — Grouped proxy exposing every `w:pgMar` attribute (`top`/`right`/`bottom`/`left`/`header`/`footer`/`gutter`) as a |Length|. `[Added in 2026.05.10]`
- `Section.page_size` — Grouped proxy for `w:pgSz` exposing `width`/`height`/`orientation`/`code` (`w:code` = Windows printer paper-size code). `[Added in 2026.05.10]`
- `Section.vertical_alignment` — Vertical alignment of text on the page (`WD_VERTICAL_ALIGNMENT.TOP` / `.CENTER` / `.BOTH` / `.BOTTOM`); maps to `w:sectPr/w:vAlign` (ECMA-376 17.6.22). `[Added in 2026.05.6]`
- `Section.columns` / `Section.set_columns(count, equal_width=None, space=None, widths=None, separator=None)` — Multi-column layout. `widths=` (per-column `Length` sequence) and `separator=` (draw a vertical divider line, `w:cols/@w:sep`) are `[Added in 2026.05.10]`.
- `Section.page_borders` / `Section.set_page_border(side, ...)` / `Section.remove_page_borders()` — Page-level borders. `[Added in 2026.05.0]`
- `Section.line_numbering` / `Section.set_line_numbering(...)` / `Section.remove_line_numbering()` — `[Added in 2026.05.0]`
- `Section.page_numbering` / `Section.set_page_numbering(...)` / `Section.remove_page_numbering()` — `w:pgNumType` with `fmt`/`start`/`chapter_style`/`chapter_separator`. `[Added in 2026.05.3]`. Alias `Section.page_number_format` is `[Added in 2026.05.10]`.
- `Section.first_page_paper_source` / `Section.other_pages_paper_source` — Paper-source bin ids. `[Added in 2026.05.0]`
- `Section.document_grid` / `Section.doc_grid` / `Section.set_document_grid(...)` / `Section.remove_document_grid()` — East-Asian grid controls. `[Added in 2026.05.0]`; `doc_grid` alias `[Added in 2026.05.10]`.
- `Section.text_direction` / `Section.right_to_left` / `Section.rtl_gutter` — `[Added in 2026.05.0]`; `rtl_gutter` (maps to `w:sectPr/w:rtlGutter`, places the gutter on the right for RTL binding) `[Added in 2026.05.10]`.
- `Section.different_first_page_header_footer` / `Section.different_odd_and_even_pages_header_footer` — Toggle variant headers/footers (`w:titlePg`, settings `w:evenAndOddHeaders`).
- `Section.first_page_header` / `Section.first_page_footer` / `Section.even_page_header` / `Section.even_page_footer` / `Section.header` / `Section.footer` — Header/footer access.
- `Section.footnote_properties` / `Section.add_footnote_properties()` / `Section.remove_footnote_properties()` / `Section.endnote_properties` / `Section.add_endnote_properties()` / `Section.remove_endnote_properties()` — Section-level overrides. `[Added in 2026.05.0]`
- `Section.add_text_watermark(text, ...)` / `Section.add_image_watermark(image, ...)` / `Section.remove_watermark()` / `Section.watermark` — Watermark per section. `[Added in 2026.05.0]`
- `Section.copy_header_from(other)` / `Section.copy_footer_from(other)` — Cross-section header/footer copy. `[Added in 2026.05.0]`
- `Section.delete()` — Remove this section break. `[Added in 2026.05.0]`
- `Section.iter_inner_content()` / `Section.paragraphs` / `Section.tables` — Content iteration.
- `Section.formatting_change` — `FormattingChange` for `w:sectPrChange`. `[Added in 2026.05.0]`
- `SectionColumns` / `Column` — Column collection; `count`, `equal_width`, `space`, `separator`, `set_widths(widths)`, per-column `width` / `space`. `separator` and `set_widths` are `[Added in 2026.05.10]`; remainder `[Added in 2026.05.0]`.
- `PageMargins`, `PageSize`, `PageBorders`, `PageBorder`, `LineNumbering`, `PageNumbering`, `DocumentGrid` — Helper proxies. `PageMargins` / `PageSize` are `[Added in 2026.05.10]`; `PageNumbering` is `[Added in 2026.05.3]`; remainder `[Added in 2026.05.0]`.
- Enums: `WD_SECTION`, `WD_SECTION_START`, `WD_ORIENTATION`, `WD_VERTICAL_ALIGNMENT`, `WD_BORDER_DISPLAY`, `WD_BORDER_OFFSET_FROM`, `WD_LINE_NUMBERING_RESTART`, `WD_CHAPTER_SEPARATOR`, `WD_DOC_GRID_TYPE`, `WD_HEADER_FOOTER_INDEX`.

---

## Headers and footers

Headers and footers live on sections and inherit from the previous section
by default (`is_linked_to_previous`). The primary flavour is always
available; even-page and first-page variants require toggling the
corresponding section property first.

```python
from docx import Document

document = Document()
section = document.sections[0]
section.different_first_page_header_footer = True

header = section.header
header.paragraphs[0].text = "Regular header"

first = section.first_page_header
first.paragraphs[0].text = "First page only"

footer = section.footer
footer.paragraphs[0].text = "Page footer"

document.save("out.docx")
```

- `Section.header` / `Section.footer` — Primary header/footer.
- `Section.first_page_header` / `Section.first_page_footer` — First-page variant (requires `section.different_first_page_header_footer=True`).
- `Section.even_page_header` / `Section.even_page_footer` — Even-page variant (requires `different_odd_and_even_pages_header_footer=True`). `[Added in 2026.05.0]`
- `_Header.is_linked_to_previous` / `_Footer.is_linked_to_previous` — Read/write inheritance flag. Assigning `False` drops any existing reference *and* creates a fresh, empty `/word/headerN.xml` part; assigning `True` removes the reference so the section inherits its ancestor's definition. `[Round-tripped through save/reopen — 2026.05.10]`
- Schema — each variant is persisted as an independent `w:headerReference` / `w:footerReference` on the section's `w:sectPr`, with `@w:type` one of `default`, `first`, or `even`, pointing at a distinct `/word/headerN.xml` or `/word/footerN.xml` content part. A section with all three variants therefore writes three separate parts. `[Round-trip-tested in 2026.05.10]`
- `Section.different_first_page_header_footer` ↔ `w:sectPr/w:titlePg` — per-section toggle for the first-page variant.
- `Section.different_odd_and_even_pages_header_footer` ↔ `w:settings/w:evenAndOddHeaders` — document-level toggle (alias for `Settings.even_and_odd_headers`; exposed on `Section` for discoverability).
- `_Header.paragraphs` / `_Header.tables` / `_Header.add_paragraph(...)` / `_Header.add_table(...)` — BlockItemContainer API.

---

## Comments

Full comments support: create a comment anchored to one or more runs, add
replies, edit `author` / `initials`, and read the (timezone-aware)
timestamp.

```python
from docx import Document

document = Document()
p = document.add_paragraph("Hello ")
r1 = p.add_run("world")
p.add_run("!")

comment = document.add_comment(
    r1, text="Consider 'globe' instead.",
    author="Ben", initials="BH",
)
comment.add_reply(text="Agreed.", author="Alex", initials="AX")

for c in document.comments:
    print(c.author, c.timestamp, "—", c.text)
    for reply in c.replies:
        print(" ↳", reply.author, reply.text)

document.save("out.docx")
```

- `Document.add_comment(runs, text="", author="", initials="", date=None)` — Add a comment with a reference range. `date` kwarg is `[Added in 2026.05.5]`.
- `Document.comments` — `Comments` collection.
- `Comments.add_comment(...)` / `Comments.get(comment_id)` / iteration / `len()`.
- `Comment.text` / `Comment.author` / `Comment.initials` / `Comment.comment_id` / `Comment.timestamp` — Core properties. `author` and `initials` are writable. `timestamp` is timezone-aware.
- `Comment.add_reply(text=None, author="", initials="")` / `Comment.reply(...)` / `Comment.replies` — Threaded replies. `.reply` alias `[Added in 2026.05.10]`.
- `Comment.is_resolved` / `Comment.resolve()` / `Comment.reopen()` — Word 2013+ resolved/reopened state via `word/commentsExtended.xml` (`w15:commentEx/@done`). The extended-comments part is created on first call and the per-comment entry is materialised/updated automatically. `[Added in 2026.05.10]`
- `Comment.parent_comment` — The parent `Comment` in a thread, resolved via `w16cid:paraIdParent` (falls back to `w15:commentEx/@paraIdParent`). Returns `None` for root comments. `[Added in 2026.05.10]`
- `Comment.add_paragraph(...)` — Multi-paragraph comment bodies.
- `Run.mark_comment_range(last_run, comment_id)` — Low-level anchor helper.
- `CommentsPart.comments_extended_part` / `CommentsPart.comments_extended_part_or_add()` — Low-level accessors for the `word/commentsExtended.xml` part (`w15:commentsEx` root with `<w15:commentEx>` and `<w15:presenceInfo>` children). `[Added in 2026.05.10]`
- `Document.comments_ids` — `CommentIds` proxy over `word/commentsIds.xml` (`w16cid:commentsIds` / `w16cid:commentId`) — paragraph-id registry that Office 365 uses to re-attach classic comments across edit sessions. The part is created on first access; :meth:`iter_ids` yields `(comment_id, paragraph_id)` tuples in document order. `[Added in 2026.05.10]` (R13-2)
- `Document.comments_extensible` — `CommentsExtensible` proxy over `word/commentsExtensible.xml` (`w16cex:commentsExtensible`) — durable GUID registry companion to `commentsIds`. Also lazily materialised. `[Added in 2026.05.10]` (R13-2)
- `Comment.paragraph_id` — Read/write. The `w16cid:paraId` token recorded in `commentsIds.xml` for this comment. Returns `""` when no registry entry exists (no `commentsIds` part related, or the comment predates the registry); assigning a value materialises the part and the entry on first write. `[Added in 2026.05.10]` (R13-2)
- `Comment.durable_id` — Read/write. The `w16cex:durableId` recorded in `commentsExtensible.xml` for this comment (positional mapping against `commentsIds`). Returns `""` when no registry entry exists; assigning a value materialises both parts and the entry on first write. `[Added in 2026.05.10]` (R13-2)
- `Comments.add_comment(...)` / `Comment.add_reply(...)` / `Comment.reply(...)` — **auto-mint** a fresh 32-bit hex `paragraph_id` and a canonical-shape GUID `durable_id` for every newly-created comment, so Office 365 clients don't renumber them on the next edit. The auto-minted values match the formats Word writes (`[0-9A-F]{8}` and `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`). `[Added in 2026.05.10]` (R13-2)
- `CommentsPart.comments_ids_part` / `CommentsPart.comments_ids_part_or_add()` / `.comments_extensible_part` / `.comments_extensible_part_or_add()` — Low-level accessors for the two modern comment-id parts. Follow the same create-on-demand pattern as `comments_extended_part_or_add`. `[Added in 2026.05.10]` (R13-2)

---

## Footnotes and endnotes

Footnote and endnote parts are lazily created; the collections can be
iterated and mutated. Each note is a `BlockItemContainer` you can fill with
paragraphs, runs, and pictures just like the body. Per-document and
per-section numbering properties are exposed through `FootnoteProperties`
and `EndnoteProperties`. `[Added in 2026.05.0]`.

```python
from docx import Document
from docx.enum.text import (
    WD_NUMBER_FORMAT, WD_FOOTNOTE_RESTART, WD_FOOTNOTE_POSITION,
)

document = Document()
p = document.add_paragraph("See the note")
r = p.add_run(".")

fn = document.footnotes.add(r, text="This is the footnote text.")
print(fn.footnote_id, fn.text)

# ergonomic one-call form (Added in 2026.05.7) — appends a reference run
# to `p` and seeds the new footnote with the given text in a single step.
p.add_footnote("Source: AWS Annual Review 2026, p.42")
p.add_endnote("Collected at end of document instead of bottom of page.")

# friendly shorthand for numbering / restart on the document collection
document.footnotes.numbering = "i, ii, iii"   # WD_NUMBER_FORMAT.LOWER_ROMAN
document.footnotes.restart = "section"        # WD_FOOTNOTE_RESTART.EACH_SECTION
document.endnotes.numbering = "*, dagger, double-dagger"  # WD_NUMBER_FORMAT.CHICAGO

# document-wide restart at each section, Roman numerals
props = document.add_footnote_properties()
props.number_format = WD_NUMBER_FORMAT.LOWER_ROMAN
props.numbering_restart = WD_FOOTNOTE_RESTART.EACH_SECTION  # alias `.restart_rule`
props.position = WD_FOOTNOTE_POSITION.BOTTOM_OF_PAGE
props.start_number = 1

# section-level override — section-level wins over document-level in Word
first_section = document.sections[0]
sec_props = first_section.add_footnote_properties()
sec_props.position = WD_FOOTNOTE_POSITION.END_OF_SECTION  # sectEnd

# document-level separator / continuation-separator / continuation-notice refs.
# Each w:id points to a w:footnote in the footnotes part whose w:type gives its role.
props.separator_id = 0
props.continuation_separator_id = 1
props.continuation_notice_id = 2

# endnotes mirror the same API
document.endnotes.add(r, text="An endnote.")

document.save("out.docx")
```

- `Document.footnotes` / `Document.endnotes` — `Footnotes` / `Endnotes` collections. `[Added in 2026.05.0]`
- `Footnotes.add(run, text="")` / `Endnotes.add(run, text="")` / iteration / `len()`. `[Added in 2026.05.0]`
- `Paragraph.add_footnote(text="")` / `Paragraph.add_endnote(text="")` — ergonomic one-call authoring. Appends a reference run to the paragraph and seeds a fresh `w:footnote` / `w:endnote` body with `text`; returns the new `Footnote` / `Endnote` for further population. Refuses to nest a note inside another note. `[Added in 2026.05.7]`
- `Footnotes.numbering` / `Footnotes.restart` (and the matching `Endnotes` setters) — shorthand pass-throughs to `FootnoteProperties.number_format` / `.restart_rule` accepting friendly strings (`"1, 2, 3"`, `"i, ii, iii"`, `"*, dagger, double-dagger"`, `"arabic"`, `"chicago"`, `"section"`, `"page"`, …), `WD_NUMBER_FORMAT` / `WD_FOOTNOTE_RESTART` enum members, or raw OOXML tokens. Auto-create the `w:footnotePr` / `w:endnotePr` element on first set. `[Added in 2026.05.7]`
- `Footnote.text` / `.footnote_id` / `.add_paragraph(...)` / `.clear()` / `.delete()` — and analogous `Endnote` members. `[Added in 2026.05.0]`
- `Document.footnote_properties` / `Document.add_footnote_properties()` / `Document.endnote_properties` / `Document.add_endnote_properties()` — Document-level (`w:settings/w:footnotePr` etc.). `[Added in 2026.05.0]`
- `Section.footnote_properties` / `Section.endnote_properties` / `Section.add_*` / `Section.remove_*` — Section-level overrides (`w:sectPr/w:footnotePr` etc.). `[Added in 2026.05.0]`
- `FootnoteProperties.number_format` / `.start_number` / `.restart_rule` / `.numbering_restart` (alias) / `.position` — Writable properties. `[Added in 2026.05.0]`
- `FootnoteProperties.separator_id` / `.continuation_separator_id` / `.continuation_notice_id` — Document-level separator-note refs (`w:footnote/@w:id` with `w:type` = `separator` / `continuationSeparator` / `continuationNotice`). Analogous `EndnoteProperties` members for `w:endnote` refs. `[Added in 2026.05.0]`
- `EndnoteProperties` — Same shape as `FootnoteProperties` with `WD_ENDNOTE_POSITION` for `.position`. `[Added in 2026.05.0]`
- Enums: `WD_NUMBER_FORMAT`, `WD_FOOTNOTE_RESTART` (`CONTINUOUS` / `EACH_SECTION` / `EACH_PAGE`), `WD_FOOTNOTE_POSITION` (`BOTTOM_OF_PAGE` / `BENEATH_TEXT` / `END_OF_SECTION` / `END_OF_DOCUMENT`), `WD_ENDNOTE_POSITION` (`END_OF_SECTION` / `END_OF_DOCUMENT`).

---

## Bookmarks

Bookmarks span one or more runs (or an entire paragraph) and carry a unique
name used by cross-references (`REF`, `PAGEREF`). Adding, reading, renaming,
and deleting are all supported, and the collection models the whole
ECMA-376 range-marker family (`w:bookmarkStart`/`End`, `w:moveFromRangeStart`/
`End`, `w:moveToRangeStart`/`End`, `w:commentRangeStart`/`End`) for unique-ID
allocation. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document()
p = document.add_paragraph()
r1 = p.add_run("Chapter 1")
r2 = p.add_run(" begins here.")

# whole-paragraph bookmark via Paragraph.add_bookmark
p.add_bookmark("chapter-1")

# multi-run (or cross-paragraph) via Document.add_bookmark (first + last run)
p2 = document.add_paragraph()
first = p2.add_run("Start")
p2.add_run(" middle ")
last = p2.add_run("end")
document.add_bookmark([first, last], name="span")

# dict-like add() accepts runs or paragraphs for start / end
document.bookmarks.add("whole", p, p2)

# inspect
bm = document.bookmarks["chapter-1"]
print(bm.text)                   # plain text between start and end markers
print(bm.start_paragraph.text)   # first paragraph overlapped
print(bm.end_paragraph.text)     # last paragraph overlapped
print([p.text for p in bm.paragraphs])  # every overlapped paragraph

bm.name = "ch-1"                 # rename
document.bookmarks.remove("span")  # delete by name
print(document.bookmarks.next_id())  # next free @w:id across all range markers

document.save("out.docx")
```

- `Paragraph.add_bookmark(name, start_run=None, end_run=None)` — Add a bookmark. `[Added in 2026.05.0]`
- `Document.add_bookmark(runs, name)` — Multi-run / cross-paragraph bookmark. `[Added in 2026.05.0]`
- `Document.bookmarks` — `Bookmarks` collection. `[Added in 2026.05.0]`
- `Bookmarks.get(name)` / `Bookmarks[name]` / `name in Bookmarks` / `iter(Bookmarks)` / `len(Bookmarks)`. `[Added in 2026.05.0]`
- `Bookmarks.add(name, start, end=None)` — `start` / `end` may be `Run` or `Paragraph`; `end` defaults to `start`. `[Added in 2026.05.0]`
- `Bookmarks.remove(name)` — Raises `KeyError` if the name is unknown. `[Added in 2026.05.0]`
- `Bookmarks.next_id()` — Next unused `@w:id` across every range-marker element (`w:bookmarkStart`/`End`, `w:moveFromRangeStart`/`End`, `w:moveToRangeStart`/`End`, `w:commentRangeStart`/`End`). `[Added in 2026.05.0]`
- `Bookmark.name` (read/write) / `Bookmark.bookmark_id` / `Bookmark.delete()`. `[Added in 2026.05.0]`
- `Bookmark.start_paragraph` / `Bookmark.end_paragraph` / `Bookmark.paragraphs` — `Paragraph` objects overlapped by the bookmark (or `None` / `[]` for an orphan). `[Added in 2026.05.0]`
- `Bookmark.text` — Concatenated text of every `w:t` between the start and end markers. `[Added in 2026.05.0]`

---

## Fields and cross-references

Simple (`w:fldSimple`) and complex (`w:fldChar`) fields can be added,
enumerated, and resolved. `REF` and `PAGEREF` are resolved against real
bookmarks (`PAGEREF` returns `?` because python-docx has no layout engine);
`DOCPROPERTY`, `AUTHOR`, `TITLE`, `SUBJECT`, `KEYWORDS`, `COMMENTS`,
`LASTSAVEDBY` are resolved from core properties. A typed
`Paragraph.add_cross_reference(ref_type, target_name, ...)` helper covers
`REF`, `PAGEREF`, `NOTEREF`, `SEQREF`, and `STYLEREF` with `\h` / `\r` /
`\p` switches and returns a `CrossReference` proxy that resolves the
target back to a `Bookmark`. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document()
p1 = document.add_paragraph()
p1.add_run("Jump to ")
p1.add_simple_field(r'REF heading1 \h', text="the heading")
p1.add_run(".")

document.add_heading("Introduction", level=1)
document.paragraphs[-1].add_bookmark("heading1")

# resolve cross-refs in place (REF text ← bookmark text)
n = document.resolve_cross_references()
print(f"resolved {n} cross-references")

# update-fields-on-open hint
document.settings.update_fields_on_open = True

document.save("out.docx")
```

- `Paragraph.add_simple_field(instr, text=None)` — Append a `w:fldSimple`. `[Added in 2026.05.0]`
- `Paragraph.add_complex_field(instr, result_text=None)` — Append `begin`/`separate`/`end`. `[Added in 2026.05.0]`
- `Paragraph.add_field(instruction, cached_result=None)` — Ergonomic R3-9 shim over `add_complex_field`; emits the five-run `fldChar` sequence with `instrText` and an optional cached result. `[Added in 2026.05.10]`
- `Paragraph.add_cross_reference(ref_type, target_name, insert_as_hyperlink=False, insert_paragraph_number=False, insert_relative_position=False, cached_result=None)` — Typed builder for `REF`, `PAGEREF`, `NOTEREF`, `SEQREF`, `STYLEREF` complex fields. Builds the instruction string, quoting names with spaces, and emits the `\h` / `\r` / `\p` switches when requested. Returns a `CrossReference`. `[Added in 2026.05.10]`
- `CrossReference` (subclass of `Field`) — proxy for REF-family fields exposing `ref_type`, `target_name`, `insert_as_hyperlink`, `insert_paragraph_number`, `insert_relative_position`, and `target_bookmark(document) -> Bookmark | None`. Obtainable from any `Field` via `Field.as_cross_reference`. `[Added in 2026.05.10]`
- `build_cross_reference_instruction(ref_type, target_name, insert_as_hyperlink=False, insert_paragraph_number=False, insert_relative_position=False, extra_switches=None)` — Standalone helper returning the field-code string (e.g. `'REF heading1 \\h \\r'`). Quotes names that contain characters outside `[A-Za-z0-9_]`. `[Added in 2026.05.10]`
- `Paragraph.add_toc(heading_range=(1, 3), hyperlinks=True, hide_in_web=True, use_outline_levels=True, omit_page_numbers_range=None, separator=None, custom_styles=None, bookmark_name=None, cached_result=None, mark_dirty=True)` — Typed builder for a `TOC` complex field. Emits the conventional `\o "min-max" \h \z \u` switches plus the optional `\n`, `\p`, `\t`, `\b` switches. Returns a `TocField`. `[Added in 2026.05.10]`
- `Paragraph.add_table_of_figures(caption_label="Figure", hyperlinks=True, cached_result=None, mark_dirty=True)` — Emits a `TOC \c "<label>" \h` field (the shape Word uses for *List of Figures* / *List of Tables*). Returns a `TableOfFiguresField`. `[Added in 2026.05.10]`
- `Paragraph.add_table_of_authorities(category=None, hyperlinks=False, cached_result=None, mark_dirty=True)` — Emits a `TOA` field for legal briefs. Returns a `TableOfAuthoritiesField` with a `.category` accessor. `[Added in 2026.05.10]`
- `TocField` (subclass of `Field`) — proxy for TOC-family fields exposing `heading_range: tuple[int,int] | None`, `hyperlinks_enabled: bool`, `hide_in_web: bool`, `use_outline_levels: bool`, `omit_page_numbers_range: tuple[int,int] | None`, `separator: str | None`, `custom_styles: list[tuple[str,int]]`, `caption_label: str | None`, `bookmark_name: str | None`. Subclasses: `TableOfFiguresField` (for `TOC \c "<label>"` and bare `TOF`) and `TableOfAuthoritiesField` (for `TOA`, adding a `.category: int | None`). Obtainable from any `Field` via `Field.as_toc`. `[Added in 2026.05.10]`
- `build_toc_field_instruction(field_type="TOC", heading_range=(1, 3), hyperlinks=True, hide_in_web=True, use_outline_levels=True, omit_page_numbers_range=None, separator=None, custom_styles=None, caption_label=None, bookmark_name=None, extra_switches=None)` — Standalone helper returning the TOC field-code string (e.g. `'TOC \\o "1-3" \\h \\z \\u'`). `[Added in 2026.05.10]`
- `parse_toc_instruction(text)` — TOC-specific parser. Unlike `parse_field_instruction`, treats `\o`, `\n`, `\p`, `\t`, `\c`, `\b`, `\s`, `\l`, `\d`, `\e`, `\g`, `\a` as argument-taking (per ECMA-376 § 17.16.5.68) so `\o "1-3"` round-trips as `switches["O"] == "1-3"` rather than spilling into positional args. `[Added in 2026.05.10]`
- `Paragraph.fields` — Mixed list of simple and complex fields. `[Added in 2026.05.0]`
- `Document.fields` — All fields in the body (simple + complex) in document order. `[Added in 2026.05.10]`
- `Run.parent_field` — The enclosing complex |Field| when this run sits between a `begin` and `end` marker, else `None`. `[Added in 2026.05.10]`
- `Field.instruction` / `Field.type` / `Field.result_text` / `Field.is_complex` / `Field.is_dirty` / `Field.mark_dirty()` / `Field.update_result_text(new_text)` / `Field.resolve(document)`. `[Added in 2026.05.0]`
- `Field.field_type` (alias for `type`) / `Field.result` (alias for `result_text`). `[Added in 2026.05.10]`
- `parse_field_instruction(text)` → `ParsedFieldInstruction(name, args, switches)` — Tokeniser for field-code strings. Supports quoted args, `{nested}` field groups, and the ECMA-376 argument-taking switches `\*`, `\@`, `\#`, `\f`. `[Added in 2026.05.10]`
- `Field.evaluate(context)` — Evaluate `IF` (with nested `{MERGEFIELD}`), `MERGEFIELD`, `HYPERLINK`, `= <expr>` arithmetic formula, and `PAGE` / `NUMPAGES` / `DATE` / `TIME` placeholders against a caller-supplied mapping. `[Added in 2026.05.8]`
- `Document.resolve_cross_references()` — Walk the body, resolve `REF`/`PAGEREF`/`DOCPROPERTY`/core-property fields, return count updated. `[Added in 2026.05.0]`
- `Document.evaluate_fields(context)` — Batch-apply `Field.evaluate` across every field in the body; writes the evaluated text back in place and returns the number of fields updated. `[Added in 2026.05.8]`
- Field type detection: `docx.fields.WD_FIELD_TYPE` constants. Covers `PAGE`, `NUMPAGES`, `DATE`, `TIME`, `AUTHOR`, `TITLE`, `FILENAME`, `REF`, `TOC`, `SEQ`, `HYPERLINK`, `PAGEREF`, `NOTEREF`, `SEQREF`, `MERGEFIELD`, `STYLEREF`, `NUMBEREDHEADERS`. `[Added in 2026.05.0, extended in 2026.05.10]`

```python
# data-driven field evaluation (mail-merge-style)
document = Document()
p = document.add_paragraph()
p.add_simple_field('IF {MERGEFIELD status} = "active" "Active" "Archived"', "?")

n = document.evaluate_fields({"status": "active"})
print(f"{n} field(s) updated")  # → "Active"
```

```python
# programmatic REF / PAGEREF / NOTEREF cross-references
document = Document()
heading = document.add_heading("Introduction", level=1)
heading.add_bookmark("intro")

p = document.add_paragraph("See ")
xref = p.add_cross_reference(
    "REF", "intro", insert_as_hyperlink=True, cached_result="Introduction"
)
p.add_run(" on page ")
p.add_cross_reference(
    "PAGEREF", "intro", insert_as_hyperlink=True, cached_result="1"
)

# inspect the cross-reference
assert xref.ref_type == "REF"
assert xref.target_name == "intro"
assert xref.insert_as_hyperlink is True
assert xref.target_bookmark(document).name == "intro"
```

---

## Table of contents

`Document.add_table_of_contents()` emits a `TOC` complex field whose
*cached result text* previews the body's headings; Word rebuilds the real
TOC on open. Sibling helpers insert the TOC before or after a specific
paragraph, and List-of-Figures / List-of-Tables emit the matching
`TOC \c "Label"` fields. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document()
document.add_heading("Contents", level=1)
document.add_table_of_contents(levels=(1, 3))

document.add_heading("Chapter One", level=1)
document.add_heading("A sub-heading", level=2)
document.add_paragraph("Body text...")

document.add_list_of_figures(caption_label="Figure")
document.add_list_of_tables(caption_label="Table")

document.save("out.docx")
```

- `Document.add_table_of_contents(levels=(1, 3))` — Append a TOC. `[Added in 2026.05.0]`
- `Paragraph.insert_table_of_contents_before(levels=(1,3))` / `insert_table_of_contents_after(...)` — Place TOC adjacent to a paragraph. `[Added in 2026.05.0]`
- `Document.add_list_of_figures(caption_label="Figure")` / `Document.add_list_of_tables(caption_label="Table")` — `[Added in 2026.05.0]`
- `Document.include_sdt_flat` iteration flag on `iter_inner_content()` surfaces TOC-wrapper content. `[Added in 2026.05.0]`
- `Paragraph.add_toc(...)` / `Paragraph.add_table_of_figures(...)` / `Paragraph.add_table_of_authorities(...)` — Typed TOC / TOF / TOA builders with full switch coverage. Each returns a `TocField` subclass so callers can inspect `heading_range`, `custom_styles`, `caption_label`, `category`, etc. without reparsing the instruction text. See the *Fields and cross-references* section for the full signature. `[Added in 2026.05.10]`

```python
# typed TOC with custom-style mapping and restricted heading range
document = Document()
toc = document.paragraphs[-1].add_toc(
    heading_range=(1, 4),
    custom_styles=[("Quote", 2), ("Intense Quote", 3)],
    separator="-",
)
assert toc.heading_range == (1, 4)
assert toc.hyperlinks_enabled is True
assert toc.custom_styles == [("Quote", 2), ("Intense Quote", 3)]
assert toc.is_dirty is True  # → Word re-renders on open
```

---

## Tracked changes

Read and resolve tracked insertions, deletions, move revisions, and
formatting changes. Accept / reject can be applied per-change or
document-wide; a context-manager wraps new content as tracked insertions by
the named author; `revision_marks_text()` renders a text preview with
bracket markers. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document()

# write new content as tracked insertions by "Reviewer"
with document.tracked_changes(author="Reviewer"):
    document.add_paragraph("Added under review.")

# per-paragraph inspection
for p in document.paragraphs:
    for tc in p.tracked_changes:
        print(tc.type, tc.author, tc.date, repr(tc.text))

# CLI preview
print(document.revision_marks_text())

# accept everything in one shot
n = document.accept_all_revisions()
print(f"resolved {n} revisions")

# or accept just one author's edits
n = document.accept_revisions_by_author("Reviewer")

document.save("out.docx")
```

- `Document.tracked_changes(author, date=None)` — Context manager that wraps new runs in `w:ins`. `[Added in 2026.05.0]`
- `Document.accept_all_changes()` / `Document.reject_all_changes()` — Resolve every change in the body. `[Added in 2026.05.0]`
- `Document.accept_revisions()` / `Document.reject_revisions()` — ECMA-376-spelled aliases of the above. `[Added in 2026.05.11]`
- `Document.accept_all_revisions()` / `Document.reject_all_revisions()` — Bulk-resolve spellings aligned with `Document.revisions`. Equivalent to `accept_all_changes` / `reject_all_changes`. `[Added in 2026.05.13]`
- `Document.accept_revisions_by_author(author)` / `Document.reject_revisions_by_author(author)` — Selectively resolve the revisions whose `w:author` equals the given string; revisions by other authors survive. Covers run-level, cell-level, and formatting-level revisions. `[Added in 2026.05.13]`
- `Document.revisions` / `Paragraph.revisions` / `Run.revisions` — Typed-proxy read of every run-level revision (`Insertion`, `Deletion`, `Move`) plus, on `Run`, the local `FormattingChange` from `w:rPrChange`. `[Added in 2026.05.11]`
- `Document.revision_marks_text(open_ins="[+", close_ins="+]", open_del="[-", close_del="-]")` — Body-text preview with markers. `[Added in 2026.05.0]`
- `Paragraph.tracked_changes` / `Paragraph.revision_marks_text(...)` / `Paragraph.formatting_change` — Per-paragraph reads. `[Added in 2026.05.0]`
- `Run.formatting_change` / `_Cell.is_tracked_insertion` / `_Cell.is_tracked_deletion` / `_Cell.formatting_change` / `Table.formatting_change` / `Section.formatting_change` — Change detection on other types. `[Added in 2026.05.0]`
- `Revision` (alias of `TrackedChange`) / `Insertion` / `Deletion` / `Move` — Typed subclasses wrapping `w:ins`, `w:del`, and `w:moveFrom`/`w:moveTo`. `[Added in 2026.05.11]`
- `TrackedChange.author` / `.date` / `.text` / `.type` / `.accept()` / `.reject()`. `[Added in 2026.05.0]`
- `MoveRevision.name` / `MoveRevision.peer` — Move source ↔ destination pairing. `[Added in 2026.05.0]` (now an alias of `Move`. `[Added in 2026.05.11]`)
- `FormattingChange.author` / `.date` / `.old_properties` — `w:rPrChange` / `w:pPrChange` / `w:sectPrChange` reader. `[Added in 2026.05.0]`
- `Settings.track_revisions` / `Settings.rsid_root` / `Settings.rsids` — Revision-ID plumbing. `[Added in 2026.05.0]` ([`RsidList`](#revision-save-ids-rsids) proxy is `[Added in 2026.05.12]`)

### Revision-save IDs (rsids)

Word stamps every paragraph, run, and section that changed in an editing
session with a session-wide 8-hex-digit "revision-save ID" (`w:rsidR`,
`w:rsidP`, `w:rsidRPr`, `w:rsidSect`, `w:rsidRDefault`). The complete set
is recorded in `w:settings/w:rsids`, with the first-ever id promoted to
`w:rsidRoot`. Rsids are separate from tracked changes (`w:ins`/`w:del`)
— Word always writes them and doesn't expose a toggle.

`Settings.rsids` returns a live `RsidList` (a `list[str]` subclass for
backward-compat) that also exposes `.root`, `.ids`, and a
`.new_session()` minter; `Document.tag_revisions()` stamps every
paragraph, run, `w:pPr`, nested `w:rPr`, and `w:sectPr` in the body with
a given (or freshly minted) rsid. This is the authoring counterpart of
the existing read accessors. `[Added in 2026.05.12]`

```python
from docx import Document

document = Document("draft.docx")

# inspect the document's recorded rsids
print("root rsid:", document.settings.rsids.root)
print("all rsids:", document.settings.rsids.ids)
print("legacy list view:", list(document.settings.rsids))

# mint an rsid for this editing session and tag every edit site with it
rsid = document.tag_revisions()          # or: tag_revisions(rsid="00ABCDEF")
assert rsid in document.settings.rsids.ids

document.save("draft.docx")
```

- `Settings.rsids` — `RsidList` proxy. `[Added in 2026.05.0]`, richer API `[Added in 2026.05.12]`
- `RsidList.root` / `RsidList.ids` / `RsidList.new_session()` / `RsidList.add(rsid)` — Read-and-mint rsid helpers over `w:rsids`. `[Added in 2026.05.12]`
- `Document.tag_revisions(rsid=None)` — Stamp every paragraph/run/pPr/rPr/sectPr in the body with an editing-session rsid. `[Added in 2026.05.12]`

---

## Content controls (SDT)

Structured Document Tags (SDTs) — rich text, plain text, date, checkbox,
combo, dropdown, picture. Block-level and inline controls are both
supported, and custom-XML data binding can be attached or removed.
`[Added in 2026.05.0]`. Additional type markers — `w15:repeatingSection`
(repeating-section template) and `w:docPartObj`/`w:docPartList`
(building-block gallery) — plus per-type proxy subclasses with typed
accessors (`DropDownListControl.items`, `DateControl.full_date`,
`RepeatingSectionControl.rows`, `BuildingBlockControl.gallery`, etc.) and
the SDT write-protection `w:lock` property are all surfaced.
`[Added in 2026.05.10]`.

```python
from docx import Document
from docx.content_controls import ContentControlType

document = Document()

# block-level rich-text placeholder
cc = document.add_content_control(
    ContentControlType.RICH_TEXT, tag="description", title="Description",
)

# inline checkbox
p = document.add_paragraph("I agree: ")
chk = p.add_content_control(ContentControlType.CHECKBOX, tag="agree")
chk.checked = True

# data binding onto a customXml part
cc.set_data_binding(
    xpath="/root/desc",
    prefix_mappings="xmlns:ns0='http://example.com/schema'",
    store_item_id="{ITEM-ID}",
)

for control in document.content_controls:
    print(control.type, control.tag, control.title)

# typed proxies dispatch automatically based on the SDT's `w:sdtPr` marker
dropdown = document.add_content_control(ContentControlType.DROPDOWN, tag="Color")
dropdown.items = ["Red", "Green", "Blue"]        # DropDownListControl

date_cc = document.add_content_control(ContentControlType.DATE, tag="Signed")
date_cc.full_date = "2026-05-09"                # DateControl
date_cc.date_format = "yyyy-MM-dd"

# repeating-section template, e.g. invoice line items
rs = document.add_content_control(
    ContentControlType.REPEATING_SECTION, tag="LineItems",
)
rs.section_title = "Item"
rs.add_row()
rs.add_row()                                     # RepeatingSectionControl

# building-block gallery picker (e.g. cover pages)
bb = document.add_content_control(
    ContentControlType.BUILDING_BLOCK, tag="Cover",
)
bb.gallery = "Cover Pages"
bb.category = "Built-In"
bb.unique = True                                 # BuildingBlockControl

# write-protect the rich-text control so end-users can edit but not delete
cc.lock = "sdtLocked"

document.save("out.docx")
```

- `Document.add_content_control(type, tag=None, title=None)` — Block-level SDT. `[Added in 2026.05.0]`
- `Paragraph.add_content_control(type, tag=None, title=None)` — Inline SDT. `[Added in 2026.05.0]`
- `Document.add_text_control(kind="rich-text", name=None, placeholder=None, value=None, locked=None, bind_to=None, items=None, title=None)` — Ergonomic block-level SDT authoring; accepts `"text"` / `"rich-text"` / `"dropdown"` / `"combo"` / `"date"` / `"checkbox"` / `"picture"` / `"repeating-section"` strings or `ContentControlType` members. `[Added in 2026.05.13]`
- `Paragraph.add_text_control(kind="text", ...)` — Inline ergonomic counterpart. `[Added in 2026.05.13]`
- `Document.add_repeating_section(name=None, section_title=None, schema=None, locked=None)` — Schema-aware repeating-section authoring; the returned `RepeatingSectionControl` exposes `.set_schema(...)` and `.add(item)` for templated row stamping. `[Added in 2026.05.13]`
- `Document.content_controls` / `Paragraph.content_controls` — Collections. `[Added in 2026.05.0]`
- `ContentControl.type` / `.tag` / `.title` / `.sdt_id` / `.text` / `.checked` / `.element`. `[Added in 2026.05.0]`
- `ContentControl.data_binding` / `.set_data_binding(xpath, prefix_mappings="", store_item_id=None)` / `.remove_data_binding()`. `[Added in 2026.05.0]`
- `DataBinding.prefix_mappings` / `.xpath` / `.store_item_id`. `[Added in 2026.05.0]`
- Enum: `ContentControlType` (`RICH_TEXT`, `PLAIN_TEXT`, `DATE`, `CHECKBOX`, `COMBO_BOX`, `DROPDOWN`, `PICTURE`, `REPEATING_SECTION`, `BUILDING_BLOCK`). `REPEATING_SECTION` and `BUILDING_BLOCK` are `[Added in 2026.05.10]`; others `[Added in 2026.05.0]`
- `ContentControl.proxy_for(sdt)` — Factory dispatching to the matching proxy subclass. `[Added in 2026.05.10]`
- `ContentControl.lock` — SDT write-protection (`unlocked` / `sdtContentLocked` / `sdtLocked` / `contentLocked`). `[Added in 2026.05.10]`
- Typed proxy subclasses returned by `proxy_for()` and the `content_controls` collections:
  - `RichTextControl` / `PlainTextControl` (`.multi_line`) / `PictureControl` / `CheckboxControl`. `[Added in 2026.05.10]`
  - `DateControl` (`.full_date`, `.date_format`). `[Added in 2026.05.10]`
  - `DropDownListControl` / `ComboBoxControl` (`.items: list[str]`, `.add_item(display_text, value=None)`, `ComboBoxControl.last_value`). `[Added in 2026.05.10]`
  - `BuildingBlockControl` (`.gallery`, `.category`, `.unique`). `[Added in 2026.05.10]`
  - `RepeatingSectionControl` (`.section_title`, `.rows`, `.add_row()`). `[Added in 2026.05.10]`
- `Document.custom_xml_parts` — Read-only list of bound `CustomXmlPart` data sources. `[Added in 2026.05.0]`
- `Document.bind_data_source(path, name, schema=None)` — Attach (or replace) a custom-XML data source under a logical id. Re-binding with the same `name` swaps the underlying payload while preserving the store-item id, so SDTs already wired to the source resolve against the new payload on the next save. Optional `schema` validates the payload via `python-ooxml-customxml`; failures raise `DataSourceValidationError`. `[Added in 2026.05.13]`
- `Document.data_sources` — List of |DataSource| proxies for every bound source. `[Added in 2026.05.13]`
- `Paragraph.add_text_control(..., bind_source=name)` / `Document.add_text_control(..., bind_source=name)` — When `bind_source` is supplied alongside `bind_to`, the emitted `<w:dataBinding>` is anchored to that source's store-item id and the resolved value is inlined into the SDT (closes #80). `[Added in 2026.05.13]`

---

## Bibliography and citations

A bibliography of citation sources is stored in a `/customXml/item{N}.xml`
part with a `<b:Sources>` root element. python-docx exposes the read path
via `Document.bibliography` and the write path via `Document.add_citation`
plus `Paragraph.add_citation_reference`. The bibliography part (and its
sibling `itemProps{N}.xml` datastore part) is materialized lazily on first
use. `[Added in 2026.05.8]`.

```python
from docx import Document

document = Document()

# Add a source. `tag` is the citation key used by references; `source_type`
# defaults to "Book". Extra kwargs become text-only <b:Capitalized> children.
document.add_citation(
    "smith2020",
    title="Distributed Systems",
    author="Smith, John",
    year=2020,
    city="London",
    publisher="Acme",
)
document.add_citation(
    "einstein1905",
    source_type="JournalArticle",
    title="Zur Elektrodynamik bewegter Koerper",
    author="Einstein, Albert",
    year=1905,
)

# Insert a citation SDT that points at the source by tag.
p = document.add_paragraph("As argued in ")
p.add_citation_reference("smith2020")
p.add_run(", ...")

# Or insert a bare CITATION complex field with page / prefix / suffix switches.
p2 = document.add_paragraph("Also ")
p2.add_citation("einstein1905", pages="891-921", prefix="cf. ")

# Read back.
for source in document.bibliography:
    print(source.tag, source.author, source.year, source.title)

hit = document.bibliography.get_by_tag("smith2020")
assert hit is not None and hit.year == "2020"

# Walk every CITATION field in the body.
for citation in document.citations:
    print(citation.source_tag, citation.pages, citation.prefix, citation.suffix)

document.save("out.docx")
```

- `Document.bibliography` — Returns a |Bibliography| proxy; lazily creates the customXml part. `[Added in 2026.05.8]`
- `Document.add_citation(tag, title=None, author=None, year=None, source_type="Book", **extra)` — Append a |Source| and return it. `[Added in 2026.05.8]`
- `Document.citations` — Walk every `CITATION` field in the body (bare or SDT-wrapped) and return a list of |Citation| proxies. `[Added in 2026.05.10]`
- `Paragraph.add_citation_reference(tag, result_text=None, locale_id=1033)` — Insert a `<w:sdt>` with a `CITATION` field referencing `tag`. `[Added in 2026.05.8]`
- `Paragraph.add_citation(source_tag, pages=None, prefix=None, suffix=None, result_text=None)` — Insert a bare complex `CITATION` field with optional `\p` / `\f` / `\s` switches (for page-range, prefix, suffix overrides). Returns a |Field|. `[Added in 2026.05.10]`
- `Bibliography.sources` — List of every |Source|. `[Added in 2026.05.8]`
- `Bibliography.add_source(tag, source_type="Book", title=None, author=None, year=None, **extra)` — Append a |Source|, validating `source_type` against the ECMA-376 catalogue (`Book`, `JournalArticle`, `ConferenceProceedings`, `Report`, `Misc`, `InternetSite`, `Film`, `SoundRecording`, `Performance`, `Art`, `DocumentFromInternetSite`, `ElectronicSource`, `Case`, `Patent`, `Interview`, ...). `[Added in 2026.05.8, catalogue validation in 2026.05.10]`
- `Bibliography.get_by_tag(tag)` — Lookup; returns |Source| or |None|. `[Added in 2026.05.8]`
- `Bibliography.selected_style` / `.style_name` — APA / MLA / etc. style selector. `[Added in 2026.05.8]`
- `Source.tag` / `.title` / `.author` / `.year` / `.source_type` / `.publisher` / `.city` / `.field(name)` / `.element`. `[Added in 2026.05.8; publisher/city/field() added in 2026.05.10]`
- `Citation.source_tag` / `.pages` / `.prefix` / `.suffix` / `.field` — Read-only access to the switch values parsed from the CITATION instruction plus the underlying |Field|. `[Added in 2026.05.10]`

---

## Form fields

Legacy `w:ffData` form fields (`FORMTEXT`, `FORMCHECKBOX`, `FORMDROPDOWN`)
can be authored directly and read back with a unified `FormField`
interface. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document()
p = document.add_paragraph("Name: ")
text_ff = p.add_text_form_field("name", default="Unknown", maxlength=32)

p = document.add_paragraph("Subscribe: ")
chk_ff = p.add_checkbox_form_field("subscribe", checked=False)

p = document.add_paragraph("Size: ")
dd_ff = p.add_dropdown_form_field(
    "size", options=["Small", "Medium", "Large"], default_index=1,
)

for ff in document.form_fields:
    print(ff.type, ff.name, "=", ff.value)

document.save("out.docx")
```

- `Paragraph.add_form_field(kind, name, **kwargs)` — Unified dispatcher; `kind` is a `WD_FORM_FIELD_TYPE` or one of `"text"` / `"checkbox"` / `"dropdown"`. Returns the appropriate typed subclass. `[Added in 2026.05.10]`
- `Paragraph.add_text_form_field(name, default="", maxlength=None)` — Add a `FORMTEXT`. `[Added in 2026.05.0]`
- `Paragraph.add_checkbox_form_field(name, checked=False)` — Add a `FORMCHECKBOX`. `[Added in 2026.05.0]`
- `Paragraph.add_dropdown_form_field(name, options, default_index=0)` — Add a `FORMDROPDOWN`. `[Added in 2026.05.0]`
- `Document.form_fields` / `Paragraph.form_fields` — Collections. `[Added in 2026.05.0]`
- `FormField.type` / `.name` / `.help_text` / `.status_text` / `.enabled` / `.calc_on_exit` / `.value` / `.current_value` — Unified read. `.current_value` alias added in `2026.05.10`. `[Added in 2026.05.0]`
- `FormField.text_input` / `FormField.checkbox` / `FormField.dropdown` — Typed views. `[Added in 2026.05.0]`
- `FormField.proxy_for(begin_run)` — classmethod returning a typed `TextInputField` / `CheckBoxField` / `DropDownListField`. `[Added in 2026.05.10]`
- `FormField.to_sdt()` — Replace the legacy form field in place with an equivalent `w:sdt` (`w:text` / `w14:checkbox` / `w:dropDownList`). Maps `w:name` → `w:tag`, `w:helpText` → `w:alias`. `[Added in 2026.05.10]`
- `TextInputFormField.default` / `.max_length` / `.format` / `.type`. `.type` (one of `"regular"`, `"number"`, `"date"`, `"currentTime"`, `"currentDate"`, `"calculated"`) added in `2026.05.10`. `[Added in 2026.05.0]`
- `CheckboxFormField.default` / `.checked` / `.size_auto` / `.size`. `.size_auto` / `.size` (half-points) added in `2026.05.10`. `[Added in 2026.05.0]`
- `DropdownFormField.options` / `.default_index` / `.result_index`. `[Added in 2026.05.0]`
- Typed subclasses: `TextInputField`, `CheckBoxField`, `DropDownListField`. `[Added in 2026.05.10]`
- Enum: `WD_FORM_FIELD_TYPE`. `[Added in 2026.05.0]`

---

## Watermarks

Text and image watermarks are attached to a section's header via
`Section.add_text_watermark()` / `Section.add_image_watermark()`.
`[Added in 2026.05.0]`.

```python
from docx import Document
from docx.shared import Inches

document = Document()
section = document.sections[0]

section.add_text_watermark(
    text="CONFIDENTIAL",
    font="Calibri",
    color="C0C0C0",
    size=72,
)

# or an image watermark
# section.add_image_watermark("watermark.png", width=Inches(4))

wm = section.watermark
if wm is not None:
    print(wm.type, wm.text)

section.remove_watermark()
document.save("out.docx")
```

- `Section.add_text_watermark(text, font=None, size=None, color=None, bold=False, italic=False, semi_transparent=True)` — `[Added in 2026.05.0]`
- `Section.add_image_watermark(image_path_or_stream, width=None, height=None)` — `[Added in 2026.05.0]`
- `Section.remove_watermark()` / `Section.watermark` — `[Added in 2026.05.0]`
- `Watermark.type` / `Watermark.text` — Read-only introspection. `[Added in 2026.05.0]`
- `Watermark.remove()` — Detach the watermark's paragraph from its header. `[Added in 2026.05.0]`

### Document-level watermark helpers

`Document.add_text_watermark()` / `Document.add_picture_watermark()` wire
the same watermark into **every** section in one call. `Document.watermarks`
enumerates the current set. `[Added in 2026.05.0]`

```python
from io import BytesIO
from docx import Document

document = Document()

# -- text watermark, diagonal silver-grey by default --
document.add_text_watermark(
    "DRAFT",
    font_name="Calibri",
    font_size=36,
    color_rgb="808080",
    diagonal=True,
)

# -- picture watermark (path or BytesIO of a PNG/JPEG) --
# document.add_picture_watermark(BytesIO(png_bytes), scale=0.5)

# -- enumerate / remove --
for wm in document.watermarks:
    print(wm.type, wm.text)
document.watermarks[0].remove()

# -- document-wide page background colour --
document.page_background_color = "4472C4"
```

- `Document.add_text_watermark(text, font_name="Calibri", font_size=36, color_rgb="808080", diagonal=True)` — `[Added in 2026.05.0]`
- `Document.add_picture_watermark(image_path_or_stream, scale=1.0)` — `[Added in 2026.05.0]`
- `Document.watermarks` — list of |Watermark| proxies across all sections. `[Added in 2026.05.0]`
- `Document.page_background_color` — `"RRGGBB"` hex string view of `w:background/@w:color`. `[Added in 2026.05.0]`
- `Document.background_color` — same underlying element, exposed as |RGBColor|. `[Added in 2026.05.0]`

---

## Captions

Captions are paragraphs styled `"Caption"` that carry a `SEQ` field for
auto-numbering (`Figure 1`, `Table 7`, etc.). Helpers append or insert
captions relative to a figure or table. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document()
# picture followed by a caption below
picture_p = document.add_paragraph()
picture_p.add_run().add_picture("diagram.png")

document.add_caption("Architecture overview", label="Figure")

# table with caption above
tbl = document.add_table(rows=2, cols=2)
tbl_p = tbl._element.addprevious  # conceptually
tbl._element.getparent()  # paragraph helpers work on the surrounding paragraph
```

- `Document.add_caption(text, label="Figure", style="Caption")` — Append a numbered caption paragraph. `[Added in 2026.05.0]`
- `Paragraph.add_caption_before(text, label="Figure", style="Caption")` / `Paragraph.add_caption_after(...)` — Insert adjacent caption. `[Added in 2026.05.0]`
- `docx.captions.new_caption_paragraph(paragraph, text, label, style)` — Low-level helper. `[Added in 2026.05.0]`
- Caption sequences automatically include the `SEQ {label} \* ARABIC` field; Word renumbers on open.

---

## Mail merge

Mail-merge main-document settings are readable and writable via
`Settings.mail_merge`. `[Added in 2026.05.0]`.

```python
from docx import Document
from docx.enum.text import WD_MAIL_MERGE_TYPE  # if present
from docx.settings import WD_MAIL_MERGE_DESTINATION, WD_MAIL_MERGE_TYPE

document = Document()
document.settings.enable_mail_merge(
    main_document_type=WD_MAIL_MERGE_TYPE.FORM_LETTERS,
    destination=WD_MAIL_MERGE_DESTINATION.NEW_DOCUMENT,
)

mm = document.settings.mail_merge
print(mm.main_document_type, mm.destination)

document.settings.disable_mail_merge()
document.save("out.docx")
```

- `Settings.mail_merge` — `MailMerge` proxy or `None`. `[Added in 2026.05.0]`
- `Settings.enable_mail_merge(main_document_type=..., destination=..., data_type=...)` — Turn it on. `[Added in 2026.05.0]`
- `Settings.disable_mail_merge()` — Remove the `w:mailMerge`. `[Added in 2026.05.0]`
- `MailMerge.main_document_type` / `.destination` / `.data_type` — Per-property reads and writes. `[Added in 2026.05.0]`
- `MailMerge.connect_string` / `.query` / `.mail_subject` / `.address_field_name` /
  `.active_record` / `.check_errors` — Per-property string and integer reads
  and writes. `[Added in 2026.05.0]`
- `MailMerge.data_source` / `.header_source` — rId references to the external
  merge data-source / header-source parts (``w:mailMerge/w:dataSource/@r:id``
  and ``.../w:headerSource/@r:id``). `[Added in 2026.05.10]`
- `MailMerge.odso` — `OdsoSettings` proxy or `None`. `[Added in 2026.05.10]`
- `MailMerge.add_odso()` / `.remove_odso()` — Create or drop the ODSO manifest. `[Added in 2026.05.10]`
- Enums: `WD_MAIL_MERGE_TYPE` (aka `WD_MAIL_MERGE_DOCUMENT_TYPE`), `WD_MAIL_MERGE_DESTINATION`, `WD_MAIL_MERGE_DATA_TYPE`, `WD_ODSO_TYPE`. `[Added in 2026.05.0 / .10]`

### ODSO — Office Data Source Object

`MailMerge.odso` exposes the `w:odso` manifest describing the merge data
source: the UDL path, table/view name, column delimiter, ODSO source
category, first-row-as-header flag, optional relationship-referenced
source part, and the field-mapping dict between Word merge-field names
and external column names.

```python
from docx import Document
from docx.enum.text import (
    WD_MAIL_MERGE_DOCUMENT_TYPE,
    WD_MAIL_MERGE_DATA_TYPE,
    WD_ODSO_TYPE,
)

document = Document()
document.settings.enable_mail_merge(
    main_document_type=WD_MAIL_MERGE_DOCUMENT_TYPE.FORM_LETTERS,
    data_type=WD_MAIL_MERGE_DATA_TYPE.ODBC,
    connect_string="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=customers.accdb",
    query="SELECT * FROM Customers",
)

mm = document.settings.mail_merge
odso = mm.add_odso()
odso.udl = "customers.udl"
odso.table = "Customers"
odso.column_delimiter = 44
odso.type = WD_ODSO_TYPE.DATABASE
odso.first_row_has_column_names = True
odso.field_mapping = {
    "FirstName": "First_Name",
    "LastName": "Last_Name",
    "Email": "Email_Address",
}
```

- `OdsoSettings.udl` / `.table` — UDL file path and table/view name. `[Added in 2026.05.10]`
- `OdsoSettings.src` — rId of the source-file relationship. `[Added in 2026.05.10]`
- `OdsoSettings.column_delimiter` — ASCII code of the column delimiter (`44`
  for comma, `9` for tab). `[Added in 2026.05.10]`
- `OdsoSettings.type` — `WD_ODSO_TYPE` source-category enum. `[Added in 2026.05.10]`
- `OdsoSettings.first_row_has_column_names` — `w:fHdr` boolean. `[Added in 2026.05.10]`
- `OdsoSettings.field_mapping` — `dict[str, str]` mapping merge-field names to
  external column names; assigning replaces the entire `w:fieldMapData` list. `[Added in 2026.05.10]`

---

## Document properties

Core (Dublin-Core), custom (typed user-defined), and extended (application)
properties are all exposed. `CustomProperties` is a dict-like typed mapping;
`ExtendedProperties` covers `docProps/app.xml` (Company, Manager, Pages,
Words, TotalTime, AppVersion...).

```python
from docx import Document

document = Document()

cp = document.core_properties
cp.author = "Ben"
cp.title = "Quarterly Report"

# typed custom properties
document.custom_properties["ReviewerCount"] = 3
document.custom_properties["IsDraft"] = True
document.custom_properties["ReleaseDate"] = "2026-05-01"

# extended properties
ep = document.extended_properties
ep.set("Company", "Example Inc")
print(ep.get("Application"))

document.save("out.docx")
```

- `Document.core_properties` — `CoreProperties` (author, title, subject, keywords, category, comments, content_status, identifier, language, version, created, last_modified_by, last_printed, modified, revision).
- `Document.custom_properties` — `CustomProperties` mapping. `[Added in 2026.05.0]`
- `CustomProperties.__getitem__` / `__setitem__` / `__delitem__` / `__contains__` / `__len__` / `__iter__` / `.add(name, value)` / `.get(name, default=None)` / `.names()` / `.items()` — Full mapping interface. Supports `str`, `int`, `float`, `bool`, and date strings. `[Added in 2026.05.0]`. `datetime.date` values serialise as `vt:date` (ISO-8601 `YYYY-MM-DD`) and `datetime.datetime` values as `vt:filetime`. `[Added in 2026.05.8]`
- `Document.extended_properties` — `ExtendedProperties` (`app.xml`) proxy. `[Added in 2026.05.0]`
- `ExtendedProperties.get(name)` / `.set(name, value)` / `.clear_all()` — Generic reads/writes; typed property accessors (`company`, `manager`, `pages`, `words`, `characters`, `total_time`, `application`, `app_version`, `template`, etc.) are generated from a declarative spec. `[Added in 2026.05.0]`

---

## Settings

`Document.settings` is a rich proxy over `word/settings.xml`. The fork adds
compatibility flags, doc-vars, theme-font language, mail merge, view,
spell/grammar toggles, auto-hyphenation, and explicit footnote/endnote
properties. Document protection is exposed as a structured object with
Word-compatible password hashing.

```python
from docx import Document
from docx.settings import WD_VIEW, WD_PROTECTION

document = Document()
settings = document.settings

settings.view = WD_VIEW.WEB
settings.track_revisions = True
settings.update_fields_on_open = True
settings.hide_spelling_errors = True
settings.auto_hyphenation = True
settings.compat_flags["allowSpaceOfSameStyleInTable"] = True
settings.doc_vars["GreetingName"] = "World"
settings.theme_font_language = ("en-US", "ja-JP", None)

# protection (filling forms, tracked changes, comments, read-only)
settings.enable_protection(WD_PROTECTION.READ_ONLY, password="secret")

document.save("out.docx")
```

- `Document.settings` — `Settings` proxy.
- `Settings.compatibility_mode` / `Settings.compat_settings` / `Settings.compat_flags` — Compatibility plumbing. `[Added in 2026.05.0]` for `compat_settings` and `compat_flags`.
- `Settings.default_tab_stop` / `Settings.zoom_percent` / `Settings.view` — Layout & view. `view` is `[Added in 2026.05.0]`.
- `Settings.track_revisions` / `Settings.rsid_root` / `Settings.rsids` — Track changes. `[Added in 2026.05.0]` for rsids.
- `Settings.update_fields_on_open` — Tell Word to refresh fields on open. `[Added in 2026.05.0]`
- `Settings.odd_and_even_pages_header_footer` / `Settings.even_and_odd_headers` — Odd/even header footer flag.
- `Settings.theme_font_language` — `(latin, east_asian, bidi)` tuple. `[Added in 2026.05.0]`
- `Settings.hide_spelling_errors` / `Settings.hide_grammatical_errors` / `Settings.auto_hyphenation` / `Settings.do_not_hyphenate_caps` / `Settings.consecutive_hyphen_limit` / `Settings.hyphenation_zone` — Proofing and hyphenation. `[Added in 2026.05.0]`
- `Settings.doc_vars` — `DocVars` dict-like (w:docVars). `[Added in 2026.05.0]`
- `Settings.mail_merge` / `.enable_mail_merge(...)` / `.disable_mail_merge()` — See [Mail merge](#mail-merge). `[Added in 2026.05.0]`
- `Settings.footnote_properties` / `Settings.endnote_properties` / `Settings.add_*` / `Settings.remove_*` — Document-level note properties. `[Added in 2026.05.0]`
- `Settings.document_protection` / `Settings.enable_protection(mode, password=None)` / `Settings.disable_protection()` — See [Permissions](#permissions-and-protection). `[Added in 2026.05.0]`
- `Settings.write_protection` / `Settings.enable_write_protection(recommended=False, password=None)` / `Settings.disable_write_protection()` — Password-to-modify (`w:writeProtection`). `[Added in 2026.05.10]`
- `Settings.remove_personal_information` / `Settings.remove_date_and_time` — Privacy-toggle pair (`w:removePersonalInformation` / `w:removeDateAndTime`). Strip author names and timestamps from comments and revisions when Word saves the document. `[Added in 2026.05.10]` (R5-21)
- `Settings.characters_with_numbers_width` — `w:charactersWithNumbersWidth` toggle used by some East-Asian layouts. `[Added in 2026.05.10]` (R5-21)
- `CompatSettings.find(name, uri=None)` / `.set(name, uri, val)` / `.as_dict()` — URI-aware read/write over `w:compatSetting` entries (the bare `proxy[name]` form only matches by name). `[Added in 2026.05.10]` (R5-21)
- `CompatSettings` / `CompatFlags` / `DocVars` — Dict-like subtype helpers. `[Added in 2026.05.0]`
- Enums: `WD_VIEW`, `WD_PROTECTION`, `WD_MAIL_MERGE_TYPE` (aka `WD_MAIL_MERGE_DOCUMENT_TYPE`), `WD_MAIL_MERGE_DESTINATION`, `WD_MAIL_MERGE_DATA_TYPE`, `WD_ODSO_TYPE`.

---

## Themes

`Document.theme` exposes the `theme1.xml` part read-only. Theme colors and
theme fonts are accessible as typed structures. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document("branded.docx")
theme = document.theme
if theme is not None:
    print("Theme:", theme.name)
    print("Major (Headings) font:", theme.fonts.major_latin)
    print("Minor (Body) font:", theme.fonts.minor_latin)
    print("Accent 1 color:", theme.colors.accent_1)
    print("Hyperlink color:", theme.colors.hyperlink)
```

- `Document.theme` — `Theme` proxy or `None`. `[Added in 2026.05.0]`
- `Theme.name` / `.colors` / `.fonts`. `[Added in 2026.05.0]`
- `ThemeColors.dark_1` / `.dark_2` / `.light_1` / `.light_2` / `.accent_1` ... `.accent_6` / `.hyperlink` / `.followed_hyperlink` / `ThemeColors[name]`. `[Added in 2026.05.0]`
- `ThemeFonts.major_latin` / `.minor_latin` / `.major_east_asian` / `.minor_east_asian` / `.major_cs` / `.minor_cs` / `.name`. `[Added in 2026.05.0]`

---

## Permissions and protection

Document-wide protection (read-only, filling-forms, comments,
tracked-changes) is controlled through `Settings.document_protection` and
its `enable_protection()` / `disable_protection()` helpers, or the
`Document.protect()` / `Document.unprotect()` shortcuts. A separate
`Settings.write_protection` proxy drives the password-to-modify
(`w:writeProtection`) marker, which Word reads as "read-only recommended"
and enforces on save. Range-level permissions
(`w:permStart`/`w:permEnd`) restrict edits to a specific user or group
within the document. `[Added in 2026.05.0]` across the board;
`Document.protect/unprotect` and `WriteProtection` are
`[Added in 2026.05.10]`.

```python
from docx import Document
from docx.enum.text import WD_PROTECTION

document = Document()

# global edit-lock: only allow tracked-changes edits, passworded
document.protect(WD_PROTECTION.TRACKED_CHANGES, password="s3cret!")

# password-to-modify (independent of the edit lock)
document.settings.enable_write_protection(recommended=True, password="saveP!")

# range-level: a paragraph only editable by "alex@example.com"
p = document.add_paragraph("Restricted section.")
p.add_permission_range(user="alex@example.com")

for pr in document.permission_ranges:
    print(pr.id, pr.user, pr.edit_group)

document.unprotect()  # clears documentProtection, leaves writeProtection
document.save("out.docx")
```

- `Settings.document_protection` — `DocumentProtection` proxy. `[Added in 2026.05.0]`
- `DocumentProtection.mode` / `.enforce` / `.formatting_locked` / `.password_hash` / `.password_salt` / `.crypto_*` / `.spin_count` — Read/write. `.set_password(password)` hashes with Word's algorithm. `[Added in 2026.05.0]`
- `Settings.enable_protection(mode, password=None)` / `Settings.disable_protection()` — High-level shortcuts. `[Added in 2026.05.0]`
- `Document.protect(edit_mode=None, password=None, enforcement=True)` / `Document.unprotect()` — Document-level shortcuts wrapping the settings helpers. `[Added in 2026.05.10]`
- `Settings.write_protection` — `WriteProtection` proxy over `w:writeProtection`. `[Added in 2026.05.10]`
- `WriteProtection.recommended_read_only` / `.enforcement` / `.password_hash` / `.password_salt` / `.crypto_*` / `.spin_count` / `.present` — Read/write. `.set_password(password)` applies Word's SHA-1 password algorithm. `[Added in 2026.05.10]`
- `Settings.enable_write_protection(recommended=False, password=None)` / `Settings.disable_write_protection()` — High-level shortcuts. `[Added in 2026.05.10]`
- `Paragraph.add_permission_range(name=None, user=None, edit_group=None)` — Wrap a paragraph in a `w:permStart`. `[Added in 2026.05.0]`
- `Paragraph.permission_ranges` / `Document.permission_ranges` — Collections. `[Added in 2026.05.0]`
- `PermissionRange.id` / `.user` / `.edit_group` / `.displaced_by_custom_xml` / `.delete()`. `[Added in 2026.05.0]`
- Enum: `WD_PROTECTION` (`READ_ONLY`, `COMMENTS`, `TRACKED_CHANGES`, `FORMS`). `[Added in 2026.05.0]`

---

## Ink annotations

Ink annotations (`<w:contentPart>` pointing at an `inkml` part) are
read-only — you can iterate them, read the raw ink-ML blob, and see how
many strokes they hold. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document("with-ink.docx")
for ink in document.ink_annotations:
    print(ink.partname, ink.stroke_count, len(ink.blob), "bytes")
```

- `Document.ink_annotations` / `Paragraph.ink_annotations` — Iterators over `InkAnnotation`. `[Added in 2026.05.0]`
- `InkAnnotation.blob` / `.partname` / `.stroke_count` / `.paragraph`. `[Added in 2026.05.0]`

---

## Embedded objects and attachments

Embedded OLE objects (Excel sheets, PDFs, arbitrary files) can be added via
`Run.add_ole_object()`. `altChunk` attachments — arbitrary foreign payloads
(HTML, RTF, another docx) that Word merges on open — are added with
`Document.add_alt_chunk()`. Both are also exposed read-only as
`embedded_objects` / `attachments` / `alt_chunks`. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document()
p = document.add_paragraph("See attached: ")
r = p.add_run()
r.add_ole_object("model.xlsx", prog_id="Excel.Sheet.12")

# HTML altChunk
document.add_alt_chunk(
    "<html><body><h1>HTML content</h1></body></html>",
    content_type="text/html",
)

# RTF altChunk with matchSrc so Word preserves the RTF's source formatting
document.add_alt_chunk(
    rtf_bytes,
    content_type="application/rtf",
    match_src=True,
)

for ole in document.embedded_objects:
    print(ole.prog_id, len(ole.blob))

for chunk in document.alt_chunks:
    print(chunk.content_type, len(chunk.blob), "match_src=", chunk.match_src)

document.save("out.docx")
```

Supported altChunk payload content-types (anything else is accepted but
stored verbatim): `text/html`, `application/xhtml+xml`, `application/rtf`
/ `text/rtf`, `text/plain`, `application/msword`, `message/rfc822`
(MHTML) / `multipart/related`, and a WordprocessingML document fragment
(`application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml`).

Since `2026.05.10` (R14-5) four format-specific helpers wrap
`add_alt_chunk()` with the right content-type pinned:

```python
document.add_html_chunk("<p>hello</p>")                  # application/xhtml+xml
document.add_text_chunk("plain paragraph")               # text/plain
document.add_rtf_chunk(b"{\\rtf1 ...}")                  # application/rtf
document.add_mhtml_chunk(b"MIME-Version: 1.0\r\n...")    # message/rfc822
```

None of the helpers sanitise the payload. altChunks are rendered by
Word's native import filters on open, which historically have been
RCE vectors (CVE-2017-0199, CVE-2023-21716, and others). See
[SECURITY.md](SECURITY.md) before embedding untrusted HTML or RTF.

- `Run.add_ole_object(ole_path_or_stream, prog_id, icon_path_or_stream=None)` — Embed an OLE payload inline. `[Added in 2026.05.0]`
- `Document.embedded_objects` / `Paragraph.embedded_objects` — Collections of `EmbeddedObject`. `[Added in 2026.05.0]`
- `EmbeddedObject.blob` / `.embedded_partname` / `.prog_id` / `.r_id` / `.type` / `.paragraph`. `[Added in 2026.05.0]`
- `Document.add_alt_chunk(content, content_type="text/html", match_src=None)` — Append a `w:altChunk`; pass `match_src=True` to write a `w:altChunkPr/w:matchSrc` child. `[Added in 2026.05.0]`
- `Document.add_html_chunk(html, match_src=None)` — `w:altChunk` with `application/xhtml+xml`. `[Added in 2026.05.10]`
- `Document.add_text_chunk(text, encoding="utf-8", match_src=None)` — `w:altChunk` with `text/plain`. `[Added in 2026.05.10]`
- `Document.add_rtf_chunk(rtf, match_src=None)` — `w:altChunk` with `application/rtf`. `[Added in 2026.05.10]`
- `Document.add_mhtml_chunk(mhtml, match_src=None)` — `w:altChunk` with `message/rfc822`. `[Added in 2026.05.10]`
- `Document.alt_chunks` — List of `AltChunk` proxies. `[Added in 2026.05.0]`
- `AltChunk.rId` / `.part` / `.content_type` / `.content` / `.match_src` (get/set). `[Added in 2026.05.0]`
- `Document.attachments` — List of `Attachment` (same underlying `altChunk` elements, read-oriented). `[Added in 2026.05.0]`
- `Attachment.r_id` / `.content_type` / `.blob` / `.partname`. `[Added in 2026.05.0]`

---

## Font table

`fontTable.xml` describes the fonts referenced by the document and can
embed TTF bytes for private fonts. The fork exposes a read-only view of the
table plus `add_embedded_font()` for authoring. `[Added in 2026.05.0]`.

Since `2026.05.10` (R5-23) the font table also supports Word's native
obfuscated-font embedding format (ECMA-376 Part 1 §17.8). Pass raw
TrueType bytes to `FontTable.embed_font(name, regular=..., bold=...,
italic=..., bold_italic=...)` — python-docx generates a fresh
`fontKey` GUID per variant, XOR-obfuscates the first 32 bytes, and
stores the result as an `application/vnd.openxmlformats-officedocument.obfuscatedFont`
part. `FontMetadata.embedded_regular` / `.embedded_bold` / `.embedded_italic`
/ `.embedded_bold_italic` deobfuscate on read and return the raw TTF
bytes (or `None` when the variant is not embedded).

```python
from docx import Document

document = Document("with-fonts.docx")
ft = document.font_table
if ft is not None:
    for meta in ft:
        print(meta.name, meta.family, meta.embed_regular)
    print("Calibri" in ft)

# embed Word-style obfuscated TTF bytes when authoring
ft2 = document.font_table_or_new
with open("Acme-Regular.ttf", "rb") as fh:
    ft2.embed_font("Acme", regular=fh.read())
# ...and read back the deobfuscated bytes later:
with open("Acme-Regular.ttf", "rb") as fh:
    assert ft2["Acme"].embedded_regular == fh.read()
```

- `Document.font_table` — `FontTable` or `None`. `[Added in 2026.05.0]`
- `Document.font_table_or_new` — Same, but creates an empty part if missing. `[Added in 2026.05.0]`
- `FontTable.__iter__` / `__len__` / `__contains__` / `__getitem__` / `.get(name)` / `.add_embedded_font(path, family="regular")`. `[Added in 2026.05.0]`
- `FontTable.fonts` — `{name: FontMetadata}` snapshot. `[Added in 2026.05.10]`
- `FontTable.embed_font(name, regular=bytes, bold=bytes, italic=bytes, bold_italic=bytes)` — obfuscated-font authoring (R5-23). `[Added in 2026.05.10]`
- `FontMetadata.name` / `.family` / `.charset` / `.pitch` / `.panose` / `.alt_name` / `.embed_regular` / `.embed_bold` / `.embed_italic` / `.embed_bold_italic`. `[Added in 2026.05.0]`
- `FontMetadata.embedded_regular` / `.embedded_bold` / `.embedded_italic` / `.embedded_bold_italic` — deobfuscated TTF bytes (or `None`). `[Added in 2026.05.10]`
- `docx.font_obfuscation.obfuscate_font_bytes(data, guid)` / `.deobfuscate_font_bytes(data, guid)` / `.generate_font_key()` — low-level helpers for ECMA-376 §17.8. `[Added in 2026.05.10]`

---

## Web settings

`webSettings.xml` is exposed read-oriented via `Document.web_settings`
(reader was `2026.05.0`; writer surface extended in `2026.05.10` to
cover `rely_on_vml` plus a read-only enumeration of the frameset
frames for R5-22).

```python
from docx import Document

document = Document("some.docx")
ws = document.web_settings
if ws is not None:
    print(ws.encoding, ws.optimize_for_browser, ws.allow_png)
    print(ws.rely_on_vml)
    for frame in ws.frames:
        print("frame:", frame)
```

- `Document.web_settings` — `WebSettings` proxy or `None`. `[Added in 2026.05.0]`
- `WebSettings.encoding` / `.optimize_for_browser` / `.allow_png` / `.do_not_save_as_single_file`. `[Added in 2026.05.0]`
- `WebSettings.rely_on_vml` (read/write bool) — toggle VML fallback for legacy browsers. `[Added in 2026.05.10]`
- `WebSettings.frames` — list of `<w:frame>` children under the root `<w:frameset>` (empty when no frameset). `[Added in 2026.05.10]`

---

## Glossary (building blocks)

The glossary document (AutoText / Quick Parts / cover pages) is exposed via
`Document.glossary`. You can enumerate building blocks, filter by category
or gallery, and inspect each entry's paragraphs and tables. `[Added in
2026.05.0]`. Since `2026.05.10` the glossary is also writable — lazy-create
it with `Document.ensure_glossary()` and then add or remove building
blocks.

```python
from docx import Document
from docx.enum.text import WD_BUILDING_BLOCK_GALLERY

# read-only inspection
document = Document("template-with-glossary.dotx")
g = document.glossary
if g is not None:
    print("%d building blocks" % len(g))
    for bb in g:
        print(bb.name, "→", bb.category.category_name, "/", bb.category.gallery)
        print("  type:", bb.type, "behaviors:", bb.behaviors)
    print("categories:", g.categories)
    print("galleries:", g.galleries)

# write: lazy-create a glossary and add an entry
document = Document()
g = document.ensure_glossary()
g.add_building_block(
    "MyQuickPart",
    category="Custom",
    gallery=WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS,
    content="Canned paragraph text.",
)
g.remove_building_block("MyQuickPart")
document.save("with-glossary.docx")
```

- `Document.glossary` — `Glossary` or `None`. `[Added in 2026.05.0]`
- `Document.ensure_glossary()` — returns a `Glossary`, lazy-creating the
  `glossaryDocument` part if needed. `[Added in 2026.05.10]`
- `Glossary.__iter__` / `__len__` / `__getitem__(name)` / `.building_blocks` / `.categories` / `.galleries` / `.by_category(name=None, gallery=None)`. `[Added in 2026.05.0]`
- `Glossary.add_building_block(name, category="General", gallery=WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS, content=None)` / `.remove_building_block(name)`. `content` accepts `str`, an existing `Paragraph`, or `None`. `[Added in 2026.05.10]`
- `BuildingBlock.name` / `.category` / `.description` / `.guid` / `.paragraphs` / `.tables`. `[Added in 2026.05.0]`
- `BuildingBlock.uuid` (alias of `.guid`) / `.type` / `.types` / `.behaviors` / `.content_paragraphs` (alias of `.paragraphs`). `[Added in 2026.05.10]`
- `BuildingBlockCategory.category_name` / `.gallery` / `.gallery_value`. `[Added in 2026.05.0]`

---

## Digital signatures

Digital signatures are detected and enumerated; no cryptographic
verification is performed. python-docx can also emit **unsigned**
signature-line placeholders via
`Document.add_signature_line(...)` — useful for authoring a document
that still needs to be signed in Word or by a separate signing tool
(e.g. `ooxml_signatures.Signer` from `python-ooxml-signatures` 0.2+).

```python
from docx import Document

document = Document("signed.docx")
if document.is_signed:
    for sig in document.signatures:
        print(sig.partname, sig.signer, sig.signed_at)

# Authoring — append an unsigned placeholder that round-trips
# through save/reload. Downstream tools fill in the real
# <SignatureValue> to make the signature cryptographically valid.
document = Document()
document.add_signature_line(
    "CN=Alice Example, O=Acme",
    signer_title="Chief Example Officer",
    email="alice@acme.test",
)
document.save("unsigned-placeholder.docx")
```

- `Document.is_signed` — `True` when `_xmlsignatures/*` parts exist. `[Added in 2026.05.0]`
- `Document.signatures` — List of `SignatureInfo`. `[Added in 2026.05.0]`
- `Document.add_signature_line(signer_name, signer_title=None, email=None)` —
  Attach an unsigned signature-line placeholder part. Returns the
  `SignatureInfo` for the new part. `[Added in 2026.05.10]`
- `SignatureInfo.partname` / `.blob` / `.signer` / `.signed_at`. `[Added in 2026.05.0]`

---

## Accessibility

`Document.validate_heading_structure()` returns a list of `HeadingIssue`
objects describing outline problems — skipped levels, multiple H1s, empty
headings, starting-below-H1. `InlineShape.alt_text` / `Table.alt_text`
expose the accessibility fields. `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document("document.docx")

# image alt text
for shape in document.inline_shapes:
    shape.alt_text = shape.alt_text or "An image"

# heading outline
issues = document.validate_heading_structure()
for issue in issues:
    print(issue.code, "@", issue.paragraph_index, issue.message)
```

- `Document.validate_heading_structure()` — List of `HeadingIssue`. `[Added in 2026.05.0]`
- `HeadingIssue` — `code`, `message`, `paragraph_index`, `heading_level`, `heading_text`. `[Added in 2026.05.0]`
- `InlineShape.alt_text` / `.title` / `FloatingImage.alt_text` / `.title` — Accessibility metadata. `[Added in 2026.05.0]`
- `Table.alt_text` / `Table.alt_description` — Table alt text. `[Added in 2026.05.0]`

---

## Document outline

`Document.outline()` returns a hierarchical heading-tree snapshot of the body
— the docx parallel of pptx's `deck.summarize()` / `skeleton()`. Each
`OutlineNode` carries `level` (0 for `Title`, 1..9 for `Heading N`), `text`,
`paragraph_index` (position in `Document.paragraphs`), a stable 8-char `id`,
the section's `word_count`, and a list of nested `children`. Page numbers
are intentionally omitted because python-docx has no layout engine; the
document-wide `total_pages_estimated` reads Word's cached
`docProps/app.xml` `<Pages>` value when present.
`Document.slice(start, end)` returns a new `Document` containing one
heading-bounded section, copied via `append_paragraph` so image / hyperlink
/ style references are rewired into the slice. `[Added in 2026.05.7]`.

```python
from docx import Document

document = Document("report.docx")
outline = document.outline()

# pretty-print the structure
for node in outline.walk():
    print("  " * node.level + node.text)

# JSON-serialisable (for LLM tools)
import json
print(json.dumps(outline.to_dict(), indent=2))

# pull a single section out as its own .docx
methodology = document.slice(start="Methodology", end="Results")
methodology.save("methodology.docx")
```

- `Document.outline()` — `Outline` with `sections`, `title`, `total_paragraphs`, `total_pages_estimated`. `[Added in 2026.05.7]`
- `Outline.walk()` / `__iter__` / `__len__` / `find(heading)` / `to_dict()`. `[Added in 2026.05.7]`
- `OutlineNode.walk()` / `to_dict()` — `level`, `text`, `paragraph_index`, `id`, `word_count`, `children`. `[Added in 2026.05.7]`
- `Document.slice(start, end=None)` — Returns a new `Document` for one heading-bounded section. `[Added in 2026.05.7]`

---

## Document statistics

`Document.statistics` returns a `DocumentStatistics` namedtuple with the
counts Word displays in its "Word Count" dialog. The body story is counted;
headers, footers, footnotes, endnotes, and comments are not. Pages are
sourced from the cached value in `docProps/app.xml` (python-docx does not
lay the document out). `[Added in 2026.05.0]`.

```python
from docx import Document

document = Document("report.docx")
stats = document.statistics
print("paragraphs:", stats.paragraphs)
print("words:     ", stats.words)
print("characters:", stats.characters)
print("characters (no spaces):", stats.characters_no_spaces)
print("pages:     ", stats.pages)  # may be None
```

- `Document.statistics` — `DocumentStatistics(paragraphs, words, characters, characters_no_spaces, pages)`. `[Added in 2026.05.0]`

---

## Readability metrics

`Document.readability()` returns a `ReadabilityReport` with the six standard
readability scores (Flesch Reading Ease, Flesch-Kincaid Grade, Gunning Fog,
SMOG, Coleman-Liau, Automated Readability Index) plus the underlying word /
sentence / syllable / character / complex-word counts, both for the whole
body and broken down per `Heading 1` section. Body text before the first
`Heading 1` is grouped under a synthetic `Introduction` section.
Tokenisation uses a stdlib-only heuristic (vowel-group syllable counting,
regex sentence splitting) so no extra dependency is added. `[Added in 2026.05.12]`.

```python
from docx import Document

metrics = Document("paper.docx").readability()
print(metrics.overall.flesch_reading_ease)       # 62.3
print(metrics.overall.flesch_kincaid_grade)      # 9.4
print(metrics.overall.gunning_fog)               # 11.2
print(metrics.overall.word_count)                # 4231
for section in metrics.sections:
    print(section.title, section.flesch_kincaid_grade, section.word_count)
```

- `Document.readability()` — `ReadabilityReport(overall, sections)`. `[Added in 2026.05.12]`
- `ReadabilityReport.to_dict()` — JSON-serialisable snapshot. `[Added in 2026.05.12]`
- `ReadabilityScores` — `flesch_reading_ease`, `flesch_kincaid_grade`, `gunning_fog`, `smog`, `coleman_liau`, `automated_readability`, `word_count`, `sentence_count`, `syllable_count`, `character_count`, `complex_word_count`, `avg_words_per_sentence`, `avg_syllables_per_word`. `[Added in 2026.05.12]`
- `SectionScores` — `title`, `paragraph_index`, `scores` plus pass-through accessors for every metric and count. `[Added in 2026.05.12]`

---

## Search and replace

Plain-text and regex-based search + replace work against body paragraphs
(`search` / `replace` / `search_regex` / `replace_regex`) or across every
story in the document (`_all` variants — body plus headers, footers,
footnotes, endnotes, and comments). All variants preserve run formatting
of the first character's run. `[Added in 2026.05.0]`.

```python
import re
from docx import Document

document = Document()
document.add_paragraph("Hello world")
document.add_paragraph("Hello again")

# plain text
matches = document.search("Hello", case_sensitive=False)
for m in matches:
    print(m.paragraph_index, m.start, m.end)

# replace in body only
document.replace("Hello", "Hi")

# regex replace everywhere (headers / footers / footnotes too)
document.replace_regex_all(re.compile(r"\bHi\b"), "Hiya")

document.save("out.docx")
```

- `Document.search(text, case_sensitive=True, whole_word=False)` — Body-only matches. `[Added in 2026.05.0]`
- `Document.search_all(text, case_sensitive=True, whole_word=False)` — Every story. `[Added in 2026.05.0]`
- `Document.search_regex(pattern, flags=0)` / `Document.search_regex_all(pattern, flags=0)` — Regex search. `[Added in 2026.05.0]`
- `Document.replace(old, new, case_sensitive=True, whole_word=False)` / `Document.replace_all(...)` — Body / all-stories replacement. `[Added in 2026.05.0]`
- `Document.replace_regex(pattern, replacement, flags=0)` / `Document.replace_regex_all(...)` — Regex replacement. `[Added in 2026.05.0]`
- `SearchMatch.paragraph` / `.paragraph_index` / `.run_indices` / `.start` / `.end` / `.location` — Match metadata; `location` identifies the story. `[Added in 2026.05.0]`
- Story location strings include `"body"`, `"table:0:row:1:col:2"`, `"header:section0:primary"`, `"footnote:2"`, `"endnote:3"`, `"comment:5"`.

---

## CSS-selector queries

`Document.select(selector)` and `Document.select_one(selector)` return
proxy objects matching a CSS-style selector. The grammar is a small,
zero-dependency subset of CSS that maps onto the python-docx proxy
graph: eight element kinds (`p`, `r`, `tbl`, `tr`, `td`, `hyperlink`,
`bookmark`, `comment`), the four attribute operators
(`=` / `^=` / `$=` / `*=` plus bare `[name]`), descendant / child
(`>`) / adjacent-sibling (`+`) combinators, and the
`:first-child` / `:last-child` / `:nth-child(N)` / `:not(...)`
pseudo-classes. `[Added in 2026.05.13]` (closes #78).

```python
from docx import Document

doc = Document("report.docx")

# Every H1
for p in doc.select('p[style="Heading 1"]'):
    print(p.text)

# Bold runs anywhere under a heading
for run in doc.select('p[style^="Heading "] r[bold]'):
    print(run.text)

# Cells in the second column of every table
cells = doc.select('tbl tr td:nth-child(2)')

# First paragraph immediately after each heading
intros = doc.select('p[style^="Heading "] + p')

# All external hyperlinks
links = doc.select('hyperlink[address^="https://"]')
```

- `Document.select(selector)` — Return a list of matching proxies in document order. `[Added in 2026.05.13]`
- `Document.select_one(selector)` — Return the first match or `None`. `[Added in 2026.05.13]`
- `docx.selectors.compile_selector(text)` — Parse once, reuse the AST on hot paths. `[Added in 2026.05.13]`
- `docx.selectors.SelectorSyntaxError` — Raised for malformed selectors. `[Added in 2026.05.13]`

See the `docx.selectors` module docstring for the full cheatsheet,
including the supported attribute names per element kind (`style`,
`text`, `name`, `id`, `address`, `tooltip`, `author`, `level`, plus the
boolean run flags `bold` / `italic` / `underline` / `hidden`).

---

## Cross-document operations

Whole documents can be appended, paragraph-by-paragraph copies can be
imported with their style / numbering / image dependencies, and individual
tables, headers, and footers can be copied between sections or documents.
`[Added in 2026.05.0]`.

```python
from docx import Document

merged = Document()
chapter1 = Document("chapter1.docx")
chapter2 = Document("chapter2.docx")

# append whole bodies (images, styles, numbering, etc. all follow)
merged.append_document(chapter1)
merged.append_document(chapter2)

# copy a single paragraph
another = Document("snippets.docx")
merged.append_paragraph(another.paragraphs[0])

# copy a single table (including styles + images)
merged.add_table_copy(another.tables[0])

# copy a header between sections
merged.sections[1].copy_header_from(merged.sections[0])

# import a style from another document
merged.styles.import_from(another, names=["BodyQuote"])

merged.save("book.docx")
```

- `Document.append_document(other)` / `Document.append_body(other)` — Append another document's body. Returns the number of block elements copied. `[Added in 2026.05.0]`
- `Document.append_paragraph(paragraph)` — Copy a single paragraph with dependencies. `[Added in 2026.05.0]`
- `Document.add_table_copy(other_table)` / `Document.add_table_from(other_table)` — Deep-copy a table. `[Added in 2026.05.0]`
- `Section.copy_header_from(other_section)` / `Section.copy_footer_from(other_section)` — `[Added in 2026.05.0]`
- `Styles.import_from(other_doc, names)` / `Styles.import_style(style)` / `Styles.import_builtin(name)` — Style import with `basedOn` / `next` / `link` resolution. `[Added in 2026.05.0]`

---

## Document diff (semantic compare)

Compare two documents by semantic content rather than raw XML — the
output is a structured set of paragraph add/remove/modify findings,
table mutations, image counts, and (at ``level="formatting"``) style
changes. Useful for PR review on docx artefacts and for verifying
generator changes against a baseline. `[Added in 2026.05.13]`

```python
from docx import Document

old = Document("q1-review-v1.docx")
new = Document("q1-review-v2.docx")

diff = old.diff(new)                       # default level="content"
print(diff.summary)
# {'paragraphs_added': 3, 'paragraphs_removed': 1, 'paragraphs_modified': 7,
#  'tables_modified': 1, 'images_added': 0, 'styles_changed': 0,
#  'total_changes': 12}

for change in diff.changes:
    print(change.kind, change.target, change.before, change.after)

# pick a granularity level
old.diff(new, level="structural")          # add/remove only
old.diff(new, level="content")             # + per-paragraph text edits
old.diff(new, level="formatting")          # + style / font / colour changes

# render formats
md = diff.to_markdown()                    # PR-comment-friendly Markdown
html = diff.to_html()                      # web-UI fragment with CSS classes
review_doc = diff.to_word_track_changes()  # third Document with [INS]/[DEL] markers
review_doc.save("review.docx")
```

Example PR-comment-friendly Markdown output:

```markdown
### Document diff (`level=content`)

| Kind | Count |
| --- | ---: |
| Paragraphs added | 3 |
| Paragraphs removed | 1 |
| Paragraphs modified | 7 |
| **Total changes** | **11** |

#### paragraph_added (3)

- `paragraph[3]` -> `NEW line A`
- `paragraph[10]` -> `NEW line B`
- `paragraph[31]` -> `NEW line C`

#### paragraph_modified (7)

- `paragraph[2]`: `Section 02: original line content` -> `Section 02: original line content (revised)`
- ...
```

- `Document.diff(other, level="content")` — Returns a `SemanticDiff`. `level` is one of `"structural"`, `"content"` (default), `"formatting"`. `[Added in 2026.05.13]`
- `SemanticDiff.summary` — Counts dictionary (`paragraphs_added`, `paragraphs_removed`, `paragraphs_modified`, `paragraphs_moved`, `tables_added`, `tables_removed`, `tables_modified`, `images_added`, `images_removed`, `styles_added`, `styles_removed`, `styles_changed`, `formatting_changed`, `total_changes`).
- `SemanticDiff.changes` — List of `Change(kind, target, before, after, detail)` records.
- `SemanticDiff.filter(*kinds)` — Subset by change-kind.
- `SemanticDiff.to_markdown(max_per_kind=25)` — PR-comment-friendly Markdown summary.
- `SemanticDiff.to_html()` — HTML5 fragment with `diff-added` / `diff-removed` / `diff-modified` CSS classes.
- `SemanticDiff.to_word_track_changes()` — Best-effort: emits a fresh `Document` whose paragraphs carry `[INS]` / `[DEL]` / `[~MOD]` markers. Native `w:ins` / `w:del` revision marks are out of scope for this exporter (markers are visible plain text rather than reviewable Word revisions).

---

## Markdown export (`Document.to_markdown()`)

A minimal, preview-grade GitHub-Flavoured-Markdown exporter. Walks the
document body and emits a GFM string suitable for handoff to PR
comments, issue trackers, static-site generators, and LLM ingestion
pipelines. Not round-trippable: there is no Markdown -> docx import
path. `[Added in 2026.05.29]`

```python
from docx import Document

doc = Document("q1-review.docx")
print(doc.to_markdown())
# # Q1 Review
#
# **Revenue grew 8.7% YoY**
#
# - AMER: $14.2B
# - APAC: $8.1B
#
# [full report](https://example.com/report) with _emphasis_.
#
# > a wise quote
#
# | Region | Revenue |
# | --- | --- |
# | AMER | $14.2B |
#
# ---
#
# Important claim[^1].
#
# [^1]: Source: 2024 study.
```

Mapping:
- `Heading 1` .. `Heading 6` -> `#` .. `######`
- bold runs -> `**text**`, italic -> `_text_`, monospace runs / `Code`
  / `HTMLCode` style -> `` `text` ``
- hyperlinks -> `[text](url)`, with `(` / `)` percent-encoded so a
  Wikipedia-style URL like `Foo_(bar)` survives intact.
- bullet lists -> `- `, decimal lists -> `1. `, with nested levels
  indented by two spaces per level.
- tables -> GFM `| col | col |` (header row + separator). `|`
  characters inside cells are escaped, paragraph breaks within a cell
  collapse to a single space.
- block quotes (paragraphs styled `Quote` / `Intense Quote` /
  `BlockQuote`) -> `> `.
- inline pictures -> `![alt](path)` where `path` is the .docx
  zip-relative archive path (e.g. `word/media/image1.png`).
- hard page breaks -> a thematic-break `---` line.
- footnote / endnote references -> `[^N]` markers with a `[^N]: text`
  block at the end of the document. The same footnote referenced
  twice reuses its index.

Lossy conversions (Markdown is a strict subset of Word's
expressiveness):
- run-level fonts, sizes, and colours collapse — only bold, italic,
  inline-code survive.
- paragraph alignment, indentation, and spacing collapse.
- drawing anchors, text boxes, OMML equations, fields, and SmartArt
  are skipped.
- multi-paragraph cells flatten to space-joined text.
- image bytes are not embedded; the path is the in-zip archive
  reference.

API:
- `Document.to_markdown()` — Returns a GFM string. `[Added in 2026.05.29]`
- `docx.markdown_export.document_to_markdown(document)` — Module-level
  entry point if you want to compose the renderer outside the
  `Document` class.

---

## PDF/A archival export (`Document.save_as_pdf_a()`)

A best-effort PDF/A (ISO 19005) archival exporter. Renders the
document body to a self-describing PDF with an XMP metadata packet
declaring the requested PDF/A conformance level. Suitable for
long-term archival pipelines where the requirement is "an
archive-ready PDF that opens in any reader" rather than a strict
spec-validating output. `[Added in 2026.05.29]`

```python
from docx import Document

doc = Document("report.docx")
doc.save_as_pdf_a("report.pdf", level="3a")
# other accepted levels: "1a", "1b", "2a", "2b", "3a", "3b"
```

Rendering covers:
- Paragraphs (with `Heading 1` .. `Heading 6` styles promoted to
  larger font sizes that step from 22pt down to 12pt).
- Inline runs (`bold` / `italic` / `underline` survive; font name,
  colour, and size collapse to the renderer defaults).
- Tables — basic grid with even column widths and space-joined cell
  text.
- Inline images (`w:drawing` with a resolvable `r:embed` rId placed at
  the paragraph flow position; anchored drawings, EMF / WMF, and bare
  SVG are skipped).
- Hard page breaks (`run.add_break(WD_BREAK.PAGE)` -> fresh PDF page).
- Bullet / numbered list items rendered with leading marker and
  per-level indentation.
- XMP metadata stream declaring `pdfaid:part` and `pdfaid:conformance`
  matching the requested level, plus `dc:title` / `xmp:CreatorTool` /
  `pdf:Producer`.

Out of scope (skipped silently; future work):
- **Font embedding fidelity** — uses ReportLab's stock Helvetica
  family rather than a fully-embedded Unicode TTF. Strict PDF/A
  validators flag this; ship a bundled Liberation Sans / DejaVu Sans
  in a future revision.
- **Output intent (sRGB ICC profile)** — not declared.
- **Footnotes, endnotes, equations, fields, drawings beyond inline
  pictures, change-tracking marks, section headers / footers / page
  numbers** — silently skipped.
- **Hyperlinks** lose their target URL (text only).
- **Spec validation** — running output through a PDF/A validator
  (veraPDF, Acrobat Pro Preflight) will surface conformance errors.

Backed by the optional `[pdfa]` extra (pulls in `reportlab`):

```
pip install 'python-docx[pdfa]'
```

When `reportlab` is not importable, `save_as_pdf_a()` raises an
`ImportError` pointing at the extra.

API:
- `Document.save_as_pdf_a(path_or_stream, level="3a")` — Render to
  the supplied path or file-like object. `[Added in 2026.05.29]`
- `docx.pdf_a_export.document_to_pdf_a(document, path_or_stream, level)`
  — Module-level entry point. `[Added in 2026.05.29]`

---

## Packaging and I/O options

`Document.save()` supports:
- A regular `.docx` / `.docm` save (default).
- A **reproducible** zip layout with a fixed timestamp, sorted members, and
  no extra metadata — byte-identical output for the same content.
  `[Added in 2026.05.0]`
- A **Flat-OPC** (`<pkg:package>`) single-XML serialisation.
  `[Added in 2026.05.0]`
- Path objects (`os.PathLike`). `[Added in 2026.05.0]`
- `.docm` macro-enabled output (auto-detected from the loaded part
  content-type). `[Added in 2026.05.0]`

Opening supports:
- `.docx`, `.docm`, `.dotx`, `.dotm` packages (`Document()`).
  `.dotx` / `.dotm` template discrimination is `[Added in 2026.05.0]`.
- Strict-OOXML packages translated on the fly. `[Added in 2026.05.0]`
- Flat-OPC input auto-detected. `[Added in 2026.05.0]`
- `recover=True` tolerating malformed XML with warnings on
  `Document.recovery_warnings`. `[Added in 2026.05.0]`
- `Document.repair(path, strategy='best-effort')` — full-fat
  recovery loader for damaged `.docx` files (truncated zip, orphan
  bookmarks, illegal control bytes, dangling rel targets, bad
  encoding declarations). Returns `(Document, RepairReport)`.
  `[Added in 2026.05.13]`
- `huge_tree=True` relaxing lxml's XML-bomb safety limits.
  `[Added in 2026.05.0]`
- `include_metadata=False` stripping the default template's core /
  extended properties on load. `[Added in 2026.05.0]`
- `EncryptedDocumentError` raised for password-protected packages when
  no `password=` is supplied. `[Added in 2026.05.0]`
- `password=` kwarg decrypts an ECMA-376 Agile-Encryption
  (password-protected) `.docx` via the optional `python-ooxml-crypto`
  dependency. `[Added in 2026.05.10]`

```python
from docx import Document

# reproducible save
doc = Document()
doc.add_paragraph("This will be byte-identical on every save.")
doc.save("rep.docx", reproducible=True)

# Flat-OPC
doc.save("rep.xml", flat_opc=True)

# macro-enabled: open a .docm and save as .docm
macro = Document("macros.docm")
print(macro.has_macros)
macro.save("macros-out.docm")

# recover mode
with open("bad.docx", "rb") as f:
    broken = Document(f, recover=True)
print(broken.recovery_warnings)

# best-effort repair (returns the salvaged document + a structured report)
doc, report = Document.repair("corrupted.docx")            # default = "best-effort"
print(report.repaired)            # ['/word/document.xml: closed orphan w:bookmarkStart id=42', ...]
print(report.parts_dropped)       # ['/word/junk.xml: unparseable XML — dropped', ...]
print(report.unrecoverable)       # []
doc.save("repaired.docx")

# strict mode mirrors the default Document(...) factory and raises on first defect
doc, report = Document.repair("clean.docx", strategy="strict")

# truncate mode keeps every part that parses, drops everything from the first defect on
doc, report = Document.repair("partial.docx", strategy="truncate")

# password-protected (requires optional `python-ooxml-crypto`)
doc = Document()
doc.add_paragraph("confidential")
doc.save("protected.docx", password="hunter2")
reopened = Document("protected.docx", password="hunter2")
```

- `Document.save(path_or_stream, flat_opc=False, reproducible=False, password=None)` — Save options as above. `password=` encrypts the output using ECMA-376 Agile Encryption via `python-ooxml-crypto`. `[Added in 2026.05.10]`
- `Document.has_macros` — `True` when a VBA project is present. `[Added in 2026.05.0]`
- `docx.exceptions.EncryptedDocumentError` — `[Added in 2026.05.0]`
- `docx.exceptions.RmsProtectedDocumentError` — `[Added in 2026.05.10]`
- `docx.package.Package` — `Package.open(...)` / `.is_signed` / `.recovery_warnings` / `.signatures` — Low-level package access.
- Flat-OPC helpers: `docx.opc.flat_opc.write_flat_opc` / `is_flat_opc`.

### Password-protected documents

python-docx supports both reading and writing ECMA-376 Agile-Encryption
password-protected `.docx` files via the optional
[`python-ooxml-crypto`](https://github.com/loadfix/python-ooxml-crypto)
dependency. Install it with `pip install 'python-docx[encryption]'` (or
directly with `pip install python-ooxml-crypto`).

```python
from docx import Document
from docx.exceptions import EncryptedDocumentError

# encrypt on save
doc = Document()
doc.add_paragraph("confidential")
doc.save("protected.docx", password="hunter2")

# decrypt on load
reopened = Document("protected.docx", password="hunter2")

# wrong-password / missing-password cases raise EncryptedDocumentError
try:
    Document("protected.docx", password="wrong")
except EncryptedDocumentError as e:
    print(e)
```

Azure RMS / AIP / IRM-wrapped files (whose payload is keyed to the user's
Microsoft 365 identity, not a password) cannot be decrypted by
python-ooxml-crypto and raise `docx.exceptions.RmsProtectedDocumentError`
(a subclass of `EncryptedDocumentError`). Delegate decryption to
Microsoft Office automation or the Microsoft Information Protection SDK
before opening such files with python-docx. `[Added in 2026.05.10]`

---

## docx.kit — high-level authoring helpers

The `docx.kit` namespace bundles small, opinionated pattern helpers that
compose the primitive python-docx APIs into common authoring shapes.
Pure in-tree (no new packages, no new mandatory dependencies); ships
under the `[kit]` extras flag so future kit-only deps can be opted in
without inflating the base install.

```bash
pip install 'python-docx[kit]'
```

### Front-matter helpers

`docx.kit.front_matter` exposes seven helpers for the conventional
front-matter sections of a long-form document (title page, copyright
page, dedication, preface, table of contents, list of figures, list of
tables). Each helper appends its section at the end of the document
body and returns the list of paragraphs it created (in document order,
including a trailing page-break paragraph by default) so the caller can
post-process them. `[Added in 2026.05.dev0]`

```python
from docx import Document
from docx.kit import front_matter

doc = Document()

front_matter.add_title_page(
    doc,
    title="Annual Report 2026",
    subtitle="Underlying performance",
    author="Acme Corp",
    date="March 2026",
)
front_matter.add_copyright_page(
    doc, holder="Acme Corp", year=2026, edition="First Edition"
)
front_matter.add_dedication(doc, text="To everyone who shipped on time.")
front_matter.add_preface(
    doc, title="Preface", body="This document outlines the year's results."
)
front_matter.add_table_of_contents(doc)        # uses existing TOC machinery
front_matter.add_list_of_figures(doc)          # TOC field filtered to "Figure" SEQ
front_matter.add_list_of_tables(doc)           # TOC field filtered to "Table" SEQ

doc.save("annual-report.docx")
```

Each helper accepts `page_break=False` to suppress the trailing page
break (useful when stitching helpers into a custom layout). The
title-page / copyright-page / dedication helpers prefer Word's built-in
styles (`Title`, `Subtitle`, `Quote`) and silently fall back to
`Normal` when a custom template lacks them — the spirit of a kit is
"works out of the box, customise as you like".

### Chapter opener pages

`docx.kit.chapter.add_chapter_opener` emits a section break followed by
the canonical chapter-start layout: a small "Chapter N" line, a large
`Heading 1` title, an italic centered epigraph, an optional decorative
image, and an optional 3-line `w:framePr` drop cap on the first letter
of the chapter body. `[Added in 2026.05.29]`

```python
from docx import Document
from docx.kit import chapter

doc = Document()

chapter.add_chapter_opener(
    doc,
    chapter_number="Chapter 1",
    title="The First Light",
    epigraph='"In the beginning, there was..." -- Genesis 1:1',
    drop_cap=True,
    image="chapter1-opener.png",
    color="primary",
)

# Page break is automatic; the next add_paragraph after this starts the
# chapter body. When drop_cap=True, the first paragraph is split into a
# 1-character framePr drop-cap paragraph plus a body paragraph holding
# the remainder, matching Word's "Insert -> Drop Cap (Dropped)" output.
doc.add_paragraph("It was a dark and stormy night...")
```

`color` accepts a named preset (`"primary"`, `"secondary"`, `"accent"`,
`"muted"`, `"black"`), an `RGBColor`, or a 6-character hex string. All
non-required arguments (`chapter_number`, `epigraph`, `image`,
`drop_cap`, `color`) are optional; the helper degrades gracefully for
documents that only need a subset of the layout.

### Section dividers and chapter ornaments

`docx.kit.dividers` exposes four small composition helpers for
inserting fleurons (Unicode ornament glyphs), three-star asterisms,
dashed / dotted / wave glyph rows, short underline rules, and full
chapter breaks (vertical-whitespace + ornament + vertical-whitespace)
between long-form-document sections. `[Added in 2026.05.29]`

```python
from docx import Document
from docx.kit.dividers import (
    add_divider,
    add_fleuron,
    add_three_stars,
    add_chapter_break,
)
from docx.shared import Pt

doc = Document()
doc.add_paragraph("End of scene one.")

add_divider(doc, kind="line")           # short underlined rule
add_divider(doc, kind="dashed")         # nine em-dashes
add_divider(doc, kind="dots")           # seven em-spaced middle dots
add_divider(doc, kind="wave")           # nine tildes

add_fleuron(doc, glyph="❦")             # FLORAL HEART (default)
add_three_stars(doc)                     # ✦ ✦ ✦

add_chapter_break(doc, ornament="line", spacing=Pt(36))
```

Every helper appends a single centred paragraph (or, for
`add_chapter_break`, three paragraphs — leading gap, ornament,
trailing gap) and returns the appended paragraph(s). `add_chapter_break`
accepts any of the four `add_divider` kinds *plus* `"fleuron"` and
`"stars"`, and forwards an optional `glyph=` to the underlying helper.

### Resume / CV template family

`docx.kit.resume` ships three template factories that build a fully
styled |Document| from plain-Python keyword arguments, plus three
visual styles (`modern` / `classic` / `minimal`). `[Added in 2026.05.29]`

```python
from docx.kit.resume import (
    resume_chronological,
    resume_functional,
    resume_technical,
)

doc = resume_chronological(
    name="Ben Hooper",
    title="Senior Software Engineer",
    contact={
        "email": "ben@example.com",
        "phone": "+61 2 1234 5678",
        "linkedin": "in/benhooper",
        "github": "benh",
    },
    summary="15+ years of distributed-systems experience.",
    experience=[
        {"company": "Acme Corp", "title": "Staff Engineer",
         "start": "2020-03", "end": "present",
         "bullets": ["Led X migration", "Shipped Y to production"]},
    ],
    education=[
        {"school": "UNSW", "degree": "BE (Hons) Software", "year": 2010},
    ],
    skills=["Python", "Go", "AWS"],          # or {"Cloud": [...], ...}
    style="modern",                            # modern | classic | minimal
)

resume_functional(
    name="Ben Hooper",
    focus_areas=["Engineering Leadership", "Distributed Systems"],
    skills={"Languages": ["Python", "Go"], "Cloud": ["AWS"]},
    style="classic",
)

resume_technical(
    name="Ben Hooper",
    projects=[{"name": "monorepo-tool", "tech": ["Python", "Rust"],
               "url": "github.com/x/monorepo-tool",
               "bullets": ["Saved 10s on every CI run"]}],
    tech_stack={"Languages": ["Python"], "Cloud": ["AWS"]},
    style="minimal",
)
```

The contact block accepts six recognised keys (`email`, `phone`,
`linkedin`, `github`, `website`, `location`) — recognised link kinds
become hyperlinks (`mailto:`, normalised LinkedIn / GitHub URLs,
`https://` for bare website domains). Unrecognised keys are rendered
verbatim after the recognised ones, in caller-insertion order.

Every factory raises `ValueError` when `name` is empty or `style` is
not one of the three built-ins. The returned `Document` is a fresh
instance — callers further mutate it (add a header via
`docx.kit.letterhead.set_letterhead`, append more sections, etc.)
or save via `Document.save(...)`.

### Mail-merge engine

`docx.kit.mail_merge.merge` bulk-renders one personalised `Document`
per record from a single template, composing the
smart-placeholder machinery from the
[smart placeholders](#smart-placeholders) section with an
ergonomic one-line API. `[Added in 2026.05.29]`

```python
from docx.kit.mail_merge import merge

records = [
    {"first_name": "Alice", "role": "Engineer", "salary": "$120k"},
    {"first_name": "Bob",   "role": "Manager",  "salary": "$140k"},
]

# In-memory: returns a list[Document] in record order.
docs = merge(template="offer-letter-template.docx", records=records)
for doc, record in zip(docs, records):
    doc.save(f"offer-{record['first_name']}.docx")

# Direct-to-disk: pass output_dir + filename_template.
paths = merge(
    template="offer-letter-template.docx",
    records=records,
    output_dir="out/",
    filename_template="offer-{first_name}.docx",
)
```

`template` accepts a path, an open binary file-like object, or an
already-loaded `Document`. The `{i}` token resolves to the current
0-based row index inside both document text and the filename
template.

### Invoice / quote / statement templates (AUS GST)

`docx.kit.invoices` ships three template factories — `invoice`,
`quote`, `statement` — that build complete billing documents with
ATO-compliant tax-invoice layout and auto-computed subtotal / GST /
grand total. `[Added in 2026.05.29]`

```python
from docx.kit.invoices import invoice, quote, statement

doc = invoice(
    invoice_number="INV-2026-0042",
    issue_date="2026-03-15",
    due_date="2026-04-14",
    seller={
        "name": "Acme Corp",
        "abn": "12 345 678 901",
        "address": "123 Pitt Street\nSydney NSW 2000",
        "phone": "+61 2 1234 5678",
        "email": "billing@acme.com",
    },
    buyer={"name": "Beta Pty Ltd", "abn": "98 765 432 109"},
    items=[
        {"description": "Consulting (March)",
         "quantity": 40, "unit_price": 250, "gst_rate": 0.10},
        {"description": "Travel reimbursement",
         "quantity": 1,  "unit_price": 580, "gst_rate": 0.10},
    ],
    payment_terms="Net 30",
    bank_details={"bsb": "062-001", "account": "1234 5678",
                  "name": "Acme Corp Pty Ltd"},
)
doc.save("INV-2026-0042.docx")

quote(quote_number="QU-2026-0099",
      issue_date="2026-03-15",
      valid_until="2026-04-15",
      seller={"name": "Acme Corp", "abn": "12 345 678 901"},
      buyer={"name": "Beta Pty Ltd"},
      items=[{"description": "Discovery", "quantity": 1, "unit_price": 5000}])

statement(period_start="2026-03-01",
          period_end="2026-03-31",
          buyer={"name": "Beta Pty Ltd"},
          invoices=[
              {"invoice_number": "INV-001", "date": "2026-03-05",
               "amount": 1000, "balance": 0,    "status": "Paid"},
              {"invoice_number": "INV-002", "date": "2026-03-15",
               "amount": 2000, "balance": 2000, "status": "Outstanding"},
          ])
```

Defaults follow ATO tax-invoice rules: a 10% GST rate is applied to
any line item that omits `gst_rate`; the header reads "Tax Invoice"
when at least one line carries GST and falls back to plain "Invoice"
when every line is GST-free. International callers opt out by passing
`default_gst_rate=0` (or per-line `gst_rate=0`); pass
`default_gst_rate=0.15` for an NZ-style 15% GST.

Each factory auto-computes `subtotal`, `gst_total`, and
`grand_total`, renders money values as `"$1,234.56"` (two decimals,
comma thousands), and right-aligns every numeric column in the
line-item table so the currency columns read cleanly. `quote` labels
its grand total "Estimated Total" so the reader doesn't treat it as
a binding bill; `statement` aggregates the supplied invoices into a
"Total Balance Owing" row at the foot. Every factory raises
`ValueError` on missing required identifiers, malformed line items
(non-numeric `unit_price`/`quantity`/`amount`, out-of-range
`gst_rate`), or non-mapping rows.

### Scientific paper templates (IEEE / ACM / APA / Nature)

`docx.kit.scientific` ships four template factories that build a
fully-shaped scientific-paper draft in one call —
`ieee_paper`, `acm_paper`, `apa_paper`, `nature_paper`. Each captures
the venue's structural skeleton (title block, author list, abstract,
keywords / index terms, body sections, references) and applies the
matching layout (IEEE and Nature switch the body to a two-column
section via a continuous section break; APA stays single-column with
double line spacing; ACM stays single-column at draft time and lets
the `acmart` stylesheet do the camera-ready two-column rendering).
`[Added in 2026.05.29]`

```python
from docx.kit.scientific import (
    ieee_paper,
    acm_paper,
    apa_paper,
    nature_paper,
)

doc = ieee_paper(
    title="A Distributed Consensus Algorithm",
    authors=[
        {"name": "Alice", "affiliation": "Acme Corp",
         "email": "alice@acme.com"},
        {"name": "Bob",   "affiliation": "Beta Labs",
         "email": "bob@beta.io"},
    ],
    abstract="We present a distributed consensus algorithm.",
    keywords=["consensus", "distributed systems"],
    sections=[
        {"heading": "Introduction", "body": "..."},
        {"heading": "Algorithm",    "body": "..."},
        {"heading": "Conclusion",   "body": "..."},
    ],
    references=[
        {"authors": "Lamport, L.",
         "title":   "The Part-Time Parliament",
         "venue":   "TOCS",
         "year":    1998},
    ],
)
doc.save("paper.docx")
```

Every factory raises `ValueError` when `title` is empty or any author
or section entry is malformed. Reference and section bodies accept
either a single string (one paragraph) or a sequence of strings (one
paragraph per item). `nature_paper` omits keywords by Nature house
style; `acm_paper` exposes a `ccs_concepts` kwarg for the CCS-Concepts
block ACM camera-ready rendering requires.

### Legal industry templates (court paper / brief / declaration / TOA)

`docx.kit.legal` ships four template factories that build the
legal-industry document shapes lawyers, paralegals, and
litigation-support staff produce every day — `court_paper`, `brief`,
`declaration`, `table_of_authorities`. The shapes follow common
Australian / NSW litigation practice (Federal Court of Australia /
NSW Supreme Court front-sheet layout, AGLC4 case-citation house style
for the TOA, Oaths Act 1900 (NSW) declaration jurat conventions).
`court_paper` and `brief` honour a `line_numbering=True` flag that
wires up Word's built-in line numbering via
`Section.set_line_numbering` (`w:sectPr/w:lnNumType`) so the numbers
appear in the left margin in Word and Print Preview. Output is a
*starting point only, not legal advice* — every document carries the
same disclaimer the module docstring carries. `[Added in 2026.05.29]`

```python
from docx.kit.legal import (
    court_paper,
    table_of_authorities,
    brief,
    declaration,
)

doc = court_paper(
    court="Federal Court of Australia",
    division="New South Wales District Registry",
    case_no="NSD 1234 of 2026",
    parties=[
        {"role": "Plaintiff", "name": "Acme Corp Pty Ltd"},
        {"role": "Defendant", "name": "Beta Pty Ltd"},
    ],
    document_type="Statement of Claim",
    line_numbering=True,
    body=[
        {"heading": "Background",      "paragraphs": ["..."]},
        {"heading": "Cause of Action", "paragraphs": ["..."]},
    ],
)

toa = table_of_authorities(
    citations=[
        {"case": "Donoghue v Stevenson [1932] AC 562", "first_pin": 580},
        {"case": "Smith v Jones (2020) 270 CLR 100",   "first_pin": 105},
    ],
)
```

`table_of_authorities` renders the supplied citations as a numbered
list (read-only fallback) and appends a Word `TOA` complex field
(`Paragraph.add_table_of_authorities`) that Word replaces with a
live, page-aware Table of Authorities the first time the user
presses F9. Pass `category=1` to filter to Cases (or 2 = Statutes,
3 = Other Authorities, …); the default emits the field without `\c`
so Word includes every category. Every factory raises `ValueError`
when its required argument is missing or malformed — `parties` for
`court_paper`, `matter` and `counsel` for `brief`, `declarant` for
`declaration`, citation `case` for `table_of_authorities`.
### Medical clinical-note templates (SOAP / discharge / referral)

`docx.kit.medical` ships three template factories that build a
clinical-document draft in one call — `soap_note`,
`discharge_summary`, `referral_letter`. Each renders the conventional
section structure for its document type (SOAP uses Subjective /
Objective / Assessment / Plan; discharge summaries use Admission /
Investigations / Procedures / Discharge medications / Follow-up;
referral letters use a salutation / clinical question / history /
examination / requested action / closing). When a `vitals` mapping is
supplied, the helper renders a structured "Parameter / Value" table
with one row per recognised vital (`bp`, `hr`, `temp`, `rr`, `spo2`,
`weight`, `height`, `bmi`, `pain`); unrecognised keys round-trip
verbatim. Every output document carries an explicit "template only —
not a medical record / not legal advice" disclaimer rendered into the
first page. `[Added in 2026.05.29]`

```python
from docx.kit.medical import (
    soap_note,
    discharge_summary,
    referral_letter,
)

doc = soap_note(
    patient={"name": "John Doe", "dob": "1980-05-15", "mrn": "1234567"},
    encounter_date="2026-03-15",
    provider={"name": "Dr. Alice Smith", "role": "GP",
              "practice": "Sydney Medical Centre"},
    subjective="Patient reports persistent cough for 3 weeks.",
    objective={
        "vitals": {"bp": "120/80", "hr": 72, "temp": 36.8,
                   "rr": 18, "spo2": 98},
        "examination": "Chest clear, no rales.",
        "labs": ["FBC: WNL", "CRP: 12 mg/L"],
    },
    assessment=[
        {"text": "Acute bronchitis", "code": "J20.9"},
        "Hypertension stable",
    ],
    plan=[
        "Amoxicillin 500mg TDS for 7 days",
        "Review in 1 week",
        "Continue current antihypertensive regimen",
    ],
)
doc.save("doe-john-2026-03-15.docx")
```

Diagnoses accept either a string (rendered verbatim) or a
`{"text": ..., "code": ...}` mapping (rendered as
`"text (CODE)"` matching ICD-10-AM presentation). Medications accept
either a string or a `{"name", "dose", "frequency", "duration"}`
mapping (missing fields are elided). Every factory raises `ValueError`
when `patient` or `provider`/`referrer` is missing or has no `name`.

### Correction of Error / post-mortem template

`docx.kit.coe.coe(doc, ...)` appends a structured Correction of Error
(also known as post-mortem or incident review) section to an existing
`Document`. The shape follows the conventional Amazon / SRE-style COE:
an incident metadata block (date / severity / duration / customer
impact), a one-paragraph summary, a chronological timeline table, the
*Five Whys* cascade rendered as a question / answer table, a
contributing-factors bullet list, an action-items table (item / owner /
due), and a lessons-learned bullet list. Returns the list of newly-
appended `Paragraph` and `Table` objects in document order. `page_break=True`
(the default) emits a trailing page break. `[Added in 2026.05.29]`

```python
from docx import Document
from docx.kit import coe

doc = Document()
coe.coe(
    doc,
    title="DB-2026-05-29 — primary failover delay",
    incident_date="2026-05-29",
    severity="Sev2",
    duration="47 minutes",
    customer_impact="20% of users saw 5xx errors during the failover window.",
    summary="One-paragraph summary of what happened.",
    timeline=[
        ("14:32 UTC", "Heartbeat alert fires"),
        ("14:34 UTC", "On-call paged"),
        ("14:42 UTC", "Failover initiated"),
        ("14:55 UTC", "Failover failed; rollback initiated"),
        ("15:19 UTC", "Service restored"),
    ],
    five_whys=[
        ("Why did the service fail?", "The primary replica fell behind."),
        ("Why did the primary fail?", "Disk hit 100% util."),
        ("Why did the disk fill?", "A rogue analytics query backfilled to it."),
        ("Why did the query run on primary?",
         "Routing rule misconfigured 6 weeks ago."),
        ("Why was it not caught?",
         "We don't alert on routing rule changes."),
    ],
    contributing_factors=[
        "Routing rule misconfigured",
        "Lack of canary on rule changes",
    ],
    action_items=[
        {"item": "Add canary on routing rule changes",
         "owner": "SRE", "due": "2026-06-15"},
        {"item": "Backfill alert on primary disk util",
         "owner": "DBA", "due": "2026-06-08"},
    ],
    lessons_learned=[
        "Always canary routing changes",
        "Alert on every disk reaching > 80% utilisation",
    ],
)
doc.save("coe.docx")
```

`timeline` and `five_whys` accept `list[tuple[str, str]]`; `action_items`
accepts `list[dict]` with required `item` and optional `owner` / `due`
keys. `contributing_factors` and `lessons_learned` are `list[str]`. The
helper falls back silently to `Normal` when the loaded template lacks
`Title` / `List Bullet` / `Table Grid` styles. `ValueError` is raised
when `title` is empty, when a timeline / five-whys entry is not a
2-tuple, or when an action item lacks a non-empty `item` key.

### Brand asset manager (YAML-driven)

`docx.kit.brand.BrandAssets` loads a corporate brand-asset bundle —
colours, fonts, logo path variants, conventional spacing values — from
a YAML manifest and exposes it as a typed, attribute-accessible object
so kit helpers and ad-hoc authoring code can compose against a single
source of truth. `[Added in 2026.05.29]`

```python
from docx import Document
from docx.kit.brand import BrandAssets
from docx.kit.letterhead import set_letterhead

brand = BrandAssets.load("aws-brand.yaml")

doc = Document()
doc.add_picture(brand.logos.full_color)
para = doc.add_paragraph("AWS")
para.runs[0].font.color.rgb = brand.colors.primary  # AWS Smile Orange
para.runs[0].font.name = brand.fonts.heading

# Brand-aware kit composition:
set_letterhead(
    doc,
    logo=brand.logos.full_color,
    return_address="410 Terry Ave N\nSeattle WA 98109",
)
```

The YAML schema is permissive — every block is optional, and unknown
keys are preserved on each sub-view's `extras` mapping rather than
dropped or rejected:

```yaml
name: AWS
colors:
  primary: '#FF9900'      # AWS Smile Orange
  secondary: '#232F3E'    # Squid Ink
  accent: '#0073BB'       # Lightning Blue
  background: '#FAFAFA'
fonts:
  heading: 'Amazon Ember Display'
  body: 'Amazon Ember'
logos:
  full_color: 'logos/aws-full-color.png'
  monochrome: 'logos/aws-mono.png'
  reverse: 'logos/aws-reverse.png'         # for dark backgrounds
spacing:
  paragraph: 12pt
  section: 24pt
```

Resolution rules:

- Colours parse as `RGBColor` from any of `"#FF9900"`, `"FF9900"`,
  `"#F90"`, `"F90"`, an existing `RGBColor`, or a 3-int list/tuple.
- Logo paths resolve against the YAML file's directory at load time,
  so a brand kit that ships `aws-brand.yaml` alongside a `logos/`
  directory works regardless of the caller's CWD. Absolute paths in
  the manifest are passed through verbatim.
- Spacing parses as `Length` with the unit suffix in the value
  (`12pt`, `0.5in`, `24mm`, `3cm`, `914400emu`, `720twips`); bare
  numbers / numeric strings are interpreted as points.
- Fonts are font-family name strings — the kit doesn't embed fonts;
  the caller is responsible for ensuring the named family is
  available in the rendering environment.

`BrandAssets.from_dict(data, base_dir=...)` constructs the same
object from an already-parsed mapping for callers who load the
manifest from a non-YAML source (TOML / JSON / env-var harness) or
who want to construct a brand programmatically. PyYAML is required
only for `BrandAssets.load(yaml_path)`; opt in via the
`[brand]` extras flag (`pip install 'python-docx[brand]'`).
### Brand-guideline validator

`docx.kit.brand.validate_brand` lints a document against a brand palette
and returns a list of `BrandFinding` records (`severity` / `location` /
`rule` / `message`). Five rules are checked: `font-not-on-brand`
(warning) — paragraph or run uses a font outside the palette;
`color-not-on-brand` (warning) — text colour not in the palette
(`#000000` is surfaced as the canonical "text-default" mismatch);
`wrong-logo` (error) — inline image looks like a logo but is not
registered under `brand.logos`; `heading-style-mismatch` (warning) —
heading paragraph drifts even when body text is correct;
`inconsistent-spacing` (info) — paragraph format diverges from
`brand.spacing`. The validator is read-only and never mutates the
document. `[Added in 2026.05.29]`

```python
from docx import Document
from docx.kit.brand import validate_brand

doc = Document("draft.docx")
findings = validate_brand(doc, rules="aws-brand.yaml")
for f in findings:
    print(f.severity, f.location, f.rule, f.message)
# warning paragraph 5 font-not-on-brand: 'Times New Roman' (allowed: ...)
# error   image 'logo.jpg' wrong-logo: image looks like a logo ...
```

`rules` accepts a YAML path, a pre-loaded `dict`, or any
`BrandAssets`-shaped object that exposes `fonts` / `colors` / `logos` /
`spacing`. Colours may be supplied either as a list of hex strings or a
`{hex: human-name}` mapping; logos may be supplied as a `path`, a
`sha1` digest, or both. PyYAML is loaded lazily — when unavailable, a
small built-in subset parser handles the schemas the issue documents.

---

## API concepts

`python-docx` is organised in three layers:

- **Document API** (`src/docx/document.py`, `src/docx/text/*.py`,
  `src/docx/table.py`, `src/docx/section.py`, etc.) — proxy objects wrapping
  OOXML elements. This is where the overwhelming majority of user code
  lives.
- **Parts layer** (`src/docx/parts/*.py`) — `XmlPart` subclasses that own
  the XML trees for each of the document's constituent parts (document,
  numbering, styles, comments, footnotes, endnotes, chart, settings,
  custom-xml, font-table, ...) and manage the relationships between them.
- **oxml layer** (`src/docx/oxml/*.py`) — `CT_*` classes extending
  `lxml.etree.ElementBase` and mapping directly onto schema element names.

`lxml` handles the XML parsing, serialisation, and XPath work beneath the
library. `docx.shared` carries `Length` subclasses (`Inches`, `Cm`, `Mm`,
`Pt`, `Emu`, `Twips`), `RGBColor`, `ElementProxy`, and `StoryChild`.

```python
from docx import Document
from docx.shared import Inches, Cm, Pt, RGBColor

document = Document()
# any Length is just a typed int — freely interchangeable
width = Inches(2)
print(int(width), Cm(5.08))          # same length, different constructor
print(Pt(12), RGBColor(0x2E, 0x74, 0xB5))
document.sections[0].left_margin = Cm(2.5)
document.save("out.docx")
```

- `docx.shared.Length` / `Inches` / `Cm` / `Mm` / `Pt` / `Emu` / `Twips` — Length constructors and arithmetic.
- `docx.shared.RGBColor` — `(r, g, b)` triple with `from_string()`, `rgb`, hex output.
- `docx.shared.ElementProxy` / `Parented` / `StoryChild` — Proxy base classes.
- `docx.opc.constants.CONTENT_TYPE` / `RELATIONSHIP_TYPE` — Content-type and rel-type constants used by the parts layer.
- `docx.oxml.ns.qn(tag)` — Clark-notation tag expansion; only needed when dropping into the oxml layer.

---

*This file is generated and maintained by hand — see `HISTORY.rst` for the
full change log, `docs/user/*.rst` for narrative tutorials, and
`docs/api/*.rst` for per-class API reference pages.*
