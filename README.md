# python-docx

A Python library for reading, creating, and updating Microsoft Word
2007+ (`.docx`) files.

Based on [python-openxml/python-docx](https://github.com/python-openxml/python-docx)
by Steve Canny and contributors. Forked at upstream `1.2.0` (2025-06-16)
and extended with 100+ additional OOXML features — footnotes and
endnotes, tracked changes, bookmarks, fields, content controls, charts,
equations, SmartArt, watermarks, digital signatures, accessibility
tooling, cross-document operations, and more.

## Status

Unstable. Not yet published to PyPI. Install from source only.

Current version: `2026.05.0` (first release as an independent fork).
Versioning is CalVer (`YYYY.MM.patch`).

## Installation

```
pip install git+https://github.com/loadfix/python-docx.git
```

Requires Python 3.9+.

## Example

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

## Features

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
- Form fields (text input, checkbox, dropdown)
- Charts (read + create for bar/line/pie; `Chart.replace_data()`)
- SmartArt (read)
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
- Cross-document operations (`append_document`, `add_table_copy`,
  `copy_header_from`)
- Packaging (`.dotx` / `.dotm` templates, Strict OOXML translation,
  Flat-OPC read/write, reproducible save, `huge_tree` opt-in, recover
  mode, encrypted-file detection, `os.PathLike` support)
- Settings and metadata (compat flags, view, mail merge,
  `Document.extended_properties`, doc vars, page stats, spell/grammar
  toggles, auto-hyphenation, timezone-aware comments)
- Themes, web settings, font table (with font embedding), glossary,
  digital-signature detection

## Documentation

API and user-guide documentation lives under `docs/` and builds with
Sphinx. The theme is Furo.

```
pip install Sphinx furo
python -m sphinx -b html docs docs/_build/html
```

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

This project is part of a series of OOXML libraries under the loadfix
org:

- [loadfix/python-docx](https://github.com/loadfix/python-docx) — Word (this repo)
- [loadfix/python-pptx](https://github.com/loadfix/python-pptx) — PowerPoint
- [loadfix/python-xlsx](https://github.com/loadfix/python-xlsx) — Excel
