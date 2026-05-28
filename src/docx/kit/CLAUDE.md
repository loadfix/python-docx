# CLAUDE.md — `docx.kit` submodule

Opinionated authoring helpers built on top of python-docx's core
public API. Lives **inside** `python-docx` (not as a sibling
package) per the project's Wave-4 scoping memo: kit content is
Python helpers tightly coupled to a single parent's API with no
cross-parent reuse story, so an in-parent submodule under a
`[kit]` extras flag is the right home.

## Conventions

- **Compose, don't reach down.** Every kit helper must call only
  python-docx's *public* API (`Document.add_paragraph`,
  `Paragraph.add_complex_field`, etc.). No `_element` /
  `oxml` / `etree` access. If a helper needs something the public
  API doesn't expose, raise it on the core API first, then add the
  kit helper.
- **Return paragraphs in document order.** Each helper appends a
  *section* of paragraphs at the end of the body and returns the
  list of newly-appended `Paragraph` objects (in order, including
  trailing page breaks). Callers post-process by iterating the
  returned list.
- **`page_break=True` by default.** Front-matter sections each end
  with their own page break so the next section starts on a fresh
  page. Pass `page_break=False` to suppress.
- **Style fallback to `Normal`.** Helpers prefer Word's built-in
  styles (`Title`, `Subtitle`, `Quote`, etc.). When the loaded
  template lacks one, fall back to `Normal` rather than raise. The
  spirit of a kit is "works out of the box".
- **Tests live under `tests/kit/`.** Same `Describe*` /
  `it_*` BDD naming as the rest of the suite. Each helper gets a
  fixture-driven test that asserts (a) the right paragraphs were
  appended, (b) styles are the expected ones, (c) any field
  instructions are correct.
- **Public API surfaces in `docx/kit/__init__.py`.** Re-export the
  submodule (e.g. `from docx.kit import front_matter`) so
  `from docx.kit import front_matter; front_matter.add_title_page(doc, ...)`
  works.

## What goes here

- High-level "recipe" helpers — front matter, layout patterns,
  content shorthands (`add_callout_box`, `add_kpi_row`,
  `add_two_column_section`, …).
- Conventional shapes that combine 5+ public-API calls into one.

## What does NOT go here

- Anything that needs new XML element classes — those land in
  `src/docx/oxml/` first.
- Anything cross-format — kit is per-parent. Cross-format helpers
  belong in a shared `python-ooxml-*` package.
- Anything that doesn't add value over a single existing API call.

## Adding a new kit module

1. Create `src/docx/kit/<name>.py` with `from __future__ import annotations`.
2. Re-export from `src/docx/kit/__init__.py`.
3. Add tests under `tests/kit/test_<name>.py`.
4. Update the kit overview docstring in `__init__.py`.
5. If the module needs runtime deps, add them to the `kit = []`
   list in `pyproject.toml`.
