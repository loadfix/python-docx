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
- Charts (read + create for bar/line/pie; `Chart.replace_data()`)
- SmartArt (read + create for list/cycle/process layout families)
- Equations (OMML read + builders for identifier, fraction, superscript,
  subscript, radical)
- Watermarks, captions, ink annotations, embedded OLE objects, alt-chunks
- Tables (borders, shading, margins, autofit, merged-cell helpers, style
  flags, caption/description, indent, row height, header rows,
  cross-document copy, CRUD on rows/columns/cells)
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
  `os.PathLike` support)
- Settings and metadata (compat flags, view, mail merge,
  `Document.extended_properties`, doc vars, page stats, spell/grammar
  toggles, auto-hyphenation, timezone-aware comments)
- Themes, web settings, font table (with font embedding), glossary,
  digital-signature detection

API and user-guide documentation lives under `docs/` and builds with
Sphinx. The theme is Furo.

```
pip install Sphinx furo
python -m sphinx -b html docs docs/_build/html
```

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
