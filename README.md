# python-docx

A Python library for reading, creating, and updating Microsoft Word
2007+ (`.docx`) files.

This repository is a fork of [python-docx](https://github.com/python-openxml/python-docx)
by Steve Canny. It builds on their original work by extending coverage
to 100+ additional OOXML features — footnotes and endnotes, tracked
changes, bookmarks, fields, content controls, charts, equations,
SmartArt, watermarks, digital signatures, accessibility tooling, and
cross-document operations. Forked at upstream `1.2.0` (2025-06-16).
Credit for the foundational library goes to the original author.

## Installation

```
pip install git+https://github.com/loadfix/python-docx.git
```

Requires Python 3.9+.

Not yet published to PyPI. Install from source only.

## Usage

```python
from docx import Document

document = Document()
document.add_paragraph("It was a dark and stormy night.")
document.save("dark-and-stormy.docx")

document = Document("dark-and-stormy.docx")
print(document.paragraphs[0].text)
# It was a dark and stormy night.
```

The package is imported as `docx`, matching upstream. Existing
upstream code runs unchanged against this fork.

## API

See [`FEATURES.md`](FEATURES.md) for the full catalogue — 43 sections
covering every public capability, with fork additions marked
`[Added in 2026.05.0]`.

Summary of areas extended beyond upstream `1.2.0`:

- Footnotes, endnotes, and their numbering properties
- Tracked changes (read, accept, reject, insertions, deletions, moves,
  formatting changes, cell/row changes, revision IDs)
- Bookmarks (create, read, delete, cross-paragraph)
- Fields (simple, complex, REF/PAGEREF cross-references, DOCPROPERTY
  resolution, table of contents, list of figures/tables)
- Content controls (SDTs: rich text, plain text, date, checkbox, combo,
  dropdown, picture; custom XML data binding)
- Bibliography and citations (`Document.bibliography`,
  `Document.add_citation`, `Paragraph.add_citation_reference` — backed by
  the `customXml/item{N}.xml` + `itemProps{N}.xml` part pair)
- Form fields (text input, checkbox, dropdown)
- Cross-format linked content (`Paragraph.link_to(target_url)` —
  Excel cells, Excel table columns, PowerPoint slides via `INCLUDETEXT`;
  `Document.update_links()` re-resolves via sibling `xlsx`/`pptx`)
- Charts (read + create for bar/line/pie; `Chart.replace_data()`)
- SmartArt (read + create for list/cycle/process layout families)
- Equations (OMML read + builders for identifier, fraction, superscript,
  subscript, radical)
- Watermarks, captions, ink annotations, embedded OLE objects, alt-chunks
- Tables (borders, shading, margins, autofit, merged-cell helpers, style
  flags, caption/description, indent, row height, header rows,
  cross-document copy, CRUD on rows/columns/cells, `Document.add_dataframe`
  styled DataFrame import with optional `pandas`)
- Sections (page borders, line numbering, document grid, paper source,
  columns, text direction, odd/even and first-page header/footer,
  copy between sections)
- Images (PNG/JPEG/GIF/BMP/TIFF/SVG/WebP/EMF/WMF/EPS; linked, floating,
  outline, crop, opacity, shadow, alt text, delete, replace)
- Shapes (preset DrawingML shapes, text boxes, canvas)
- Numbering (custom definitions, restart, rendered list labels)
- Styles (cross-document import, builtin latent materialisation,
  document-default font, next-paragraph auto-apply)
- Fonts (cs size, character scale, ligatures, shading, borders,
  language, East Asian layout, symbols, ruby)
- Accessibility (alt text, heading-structure validation)
- Search and replace (plain, regex, across tables/headers/footers/footnotes)
- CSS-selector queries (`Document.select` / `Document.select_one` —
  paragraphs, runs, tables, hyperlinks, bookmarks, comments by
  attribute / combinator / pseudo-class)
- Cross-document operations (`append_document`, `add_table_copy`,
  `copy_header_from`)
- Semantic diff (`Document.diff(other)`, three granularity levels,
  Markdown / HTML / Word output formats — review-friendly compare for
  PR workflows)
- Packaging (`.dotx` / `.dotm` templates, Strict OOXML translation,
  Flat-OPC read/write, reproducible save, `huge_tree` opt-in, recover
  mode, `Document.repair()` best-effort recovery for damaged packages,
  password-protected read/write via optional `python-ooxml-crypto`,
  `Document.stream()` bounded-memory reader for very large documents,
  `Document.from_html()` / `from_html_string()` HTML import, `os.PathLike`
  support)
- Settings and metadata (compat flags, view, mail merge,
  `Document.extended_properties`, doc vars, page stats, spell/grammar
  toggles, auto-hyphenation, timezone-aware comments)
- Themes, web settings, font table (with font embedding), glossary,
  digital-signature detection
- High-level authoring helpers under `docx.kit` — pattern-level
  compositions over the primitive APIs. Ships
  `docx.kit.front_matter` (title page, copyright page, dedication,
  preface, table of contents, list of figures, list of tables),
  `docx.kit.chapter.add_chapter_opener` (section break + Heading 1 title
  + epigraph + decorative image + drop cap),
  `docx.kit.dividers` (`add_divider` / `add_fleuron` / `add_three_stars`
  / `add_chapter_break` for section dividers and chapter ornaments —
  fleurons, three-stars, dashed/dotted/wave/line breaks),
  `docx.kit.letterhead.set_letterhead` (branded header + footer with
  three styles), `docx.kit.resume`
  (`resume_chronological` / `resume_functional` / `resume_technical`
  factories returning fully-styled CV documents in three visual
  styles — `modern` / `classic` / `minimal`),
  `docx.kit.mail_merge.merge` (bulk-render N personalised documents
  from a single template + iterable of records),
  `docx.kit.contracts` (`nda` / `msa` / `sow` /
  `contractor_agreement` boilerplate factories — *starting points
  only, not legal advice*), `docx.kit.invoices` (`invoice` /
  `quote` / `statement` factories with AUS GST defaults — 10% GST,
  override per-line via `gst_rate=0` for international callers,
  auto-computed subtotal / GST / grand total, right-aligned line-item
  table; output complies with ATO tax-invoice rules when the seller
  carries an ABN), `docx.kit.memos` (`investment_memo`
  with McKinsey-style SCQA executive summary, and `business_case`
  with options-analysis table), `docx.kit.templates`
  (`brief` / `coe` / `rfp_response` / `white_paper` document-template
  registry covering short briefs, Centre of Excellence charters,
  RFP responses with a pricing table, and white papers with abstract
  and references), `docx.kit.scientific`
  (`ieee_paper` / `acm_paper` / `apa_paper` / `nature_paper`
  scientific-paper template factories — IEEE / Nature switch the body
  to two-column layout, APA applies double line spacing, ACM stays
  single-column for the `acmart` stylesheet), `docx.kit.legal`
  (`court_paper` / `brief` / `declaration` / `table_of_authorities`
  legal industry template factories with Federal Court of Australia /
  NSW Supreme Court front-sheet layout, Word built-in line numbering
  via `w:sectPr/w:lnNumType`, and a live `TOA` complex field —
  *starting points only, not legal advice*), and `docx.kit.medical`
  (`soap_note` / `discharge_summary` / `referral_letter`
  clinical-note template factories with Subjective / Objective /
  Assessment / Plan structure and a structured vitals table —
  *template only, not a medical record*), and `docx.kit.brand`
  (`BrandAssets.load(yaml_path)` — YAML-driven manifest loader for
  brand colours, font pairs, logo path variants, and conventional
  spacing values; composes with `set_letterhead`, `add_chapter_opener`,
  and the rest of the kit so an organisation declares its brand once
  and reuses it everywhere). Lives under the optional `[kit]` extras
  flag (`pip install python-docx[kit]`); `BrandAssets.load` additionally
  needs PyYAML, which the optional `[brand]` extras pulls in
  (`pip install 'python-docx[brand]'`).
  (`validate_brand` brand-guideline linter that walks a document and
  returns `BrandFinding` records covering font / colour / logo /
  heading-style / spacing drift against a YAML, dict, or
  `BrandAssets`-shaped palette). Lives under the optional
  `[kit]` extras flag (`pip install python-docx[kit]`).

API and user-guide documentation lives under `docs/` and builds with
Sphinx. The theme is Furo.

```
pip install Sphinx furo
python -m sphinx -b html docs docs/_build/html
```

## Reproducible builds

`Document.save(path, reproducible=True)` produces a byte-identical
`.docx` for byte-identical inputs across machines and runs:

```python
from docx import Document

doc = Document()
doc.add_paragraph("Hello")
doc.save("out.docx", reproducible=True)
```

The flag stamps every zip-member with the fixed 1980-01-01 timestamp,
emits members in sorted order, normalises external file attributes,
and disables the rsid-family churn attributes that Word otherwise
mints on every save — the four sources of cross-machine and cross-
session nondeterminism. Use it for source-control-friendly diffs,
fixture regeneration, and content-addressable artefact pipelines.
The matching keyword is also accepted by the sibling `python-pptx`,
`python-xlsx`, and `python-vsdx` parents so cross-format build
pipelines share a single idiom (issue #150).

## Round-trip support

A central design goal of this fork is **round-trip fidelity** — load a
real-world `.docx`, mutate a few elements, save, and have nothing else
change. Charts, comments, custom XML parts, math, ink, signatures,
bibliography, and the rest of the loadfix-extended feature surface
must all survive.

The cross-monorepo round-trip gate lives at
[`tests/round_trip/`](../tests/round_trip/README.md) and runs as the
`round-trip-fidelity` CI job. The full per-feature support matrix
(what's "fully preserved" / "preserved with caveats" / "lossy")
across all four parent formats lives at
[`docs/round-trip-fidelity.md`](../docs/round-trip-fidelity.md).

## Status

Unstable. Not yet published to PyPI. Current version: `2026.05.10`
(first release as an independent fork). Versioning is CalVer
(`YYYY.MM.patch`). Public API tracks upstream `1.2.0` for the
inherited surface; fork additions are considered experimental until
the next calendar release.

## Contributing

Issues and pull requests are tracked at
<https://github.com/loadfix/python-docx/issues>. Please file issues
against this fork; upstream's tracker is for upstream-shared concerns
only.

When contributing:

- Run the tests: `pytest tests/ -q` and `uv run behave features/`.
- Keep `FEATURES.md` current when adding, modifying, or removing public
  API (see `CLAUDE.md` for contributor conventions).
- Consult `spec/` (XSD schemas and the ISO/IEC 29500 PDFs) for
  authoritative element ordering and cardinality when implementing new
  `CT_*` classes.

## License

MIT. See `LICENSE`. Inherited from upstream `python-openxml/python-docx`.

## Related projects

Part of a family of document-rendering libraries:

- [docxjs](https://github.com/loadfix/docxjs) — browser-side DOCX → HTML renderer (TypeScript)
- [pptxjs](https://github.com/loadfix/pptxjs) — browser-side PPTX → HTML renderer (TypeScript)
- [xlsxjs](https://github.com/loadfix/xlsxjs) — browser-side XLSX → HTML renderer (TypeScript)
- [python-pptx](https://github.com/loadfix/python-pptx) — Python PPTX parser/generator
- [python-xlsx](https://github.com/loadfix/python-xlsx) — Python XLSX parser/generator
- [ooxml-validate](https://github.com/loadfix/ooxml-validate) — Python/.NET OOXML validator (wraps Microsoft Open XML SDK + LibreOffice)
