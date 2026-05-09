# CLAUDE.md

python-docx fork (loadfix/python-docx) — extending python-docx with footnotes, endnotes, track changes, fields, bookmarks, and other missing OOXML capabilities.

This project is one of a sibling series of OOXML libraries under the loadfix org:

- **loadfix/python-docx** — Word `.docx` (this repo)
- **loadfix/python-pptx** — PowerPoint `.pptx`
- **loadfix/python-xlsx** — Excel `.xlsx`
- **loadfix/ooxml-validate** — cross-format validator (Microsoft Open XML SDK + LibreOffice)

The three reading/writing libraries share an architectural lineage (three-layer proxy/part/oxml pattern over lxml) and OOXML spec conventions. When implementing a feature that exists across the trio, consult the sibling repos for naming and API-shape precedent. When verifying that library output is correct, use `ooxml-validate`.

## Architecture

Three-layer pattern:

```
Document API  (src/docx/document.py, src/docx/footnotes.py, etc.)
    |  Proxy objects wrapping oxml elements
Parts Layer   (src/docx/parts/*.py)
    |  XmlPart subclasses owning XML trees, managing relationships
oxml Layer    (src/docx/oxml/*.py)
    |  CT_* element classes extending lxml.etree.ElementBase
lxml          (XML parsing/serialization)
```

## Source Layout

```
src/docx/           Main package (src-layout, NOT flat)
src/docx/oxml/      CT_* element classes (low-level XML wrappers)
src/docx/parts/     Part classes (document, numbering, comments, styles, etc.)
src/docx/text/      Text-related proxy classes (paragraph, run, font, parfmt)
src/docx/styles/    Style proxy classes
src/docx/enum/      Enumerations (WD_ALIGN, WD_STYLE_TYPE, etc.)
src/docx/templates/ Default XML templates for new parts
tests/              pytest test suite
features/           behave acceptance tests
```

## Key Patterns

### CT_ Element Classes (oxml layer)

Define in `src/docx/oxml/`, register in `src/docx/oxml/__init__.py`.

```python
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrOne, ZeroOrMore, OptionalAttribute
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String

class CT_Footnote(BaseOxmlElement):
    """``<w:footnote>`` element."""
    pPr = ZeroOrOne("w:pPr", successors=("w:r",))
    r = ZeroOrMore("w:r", successors=())
    id = RequiredAttribute("w:id", ST_DecimalNumber)
```

- `ZeroOrOne(tag, successors=(...))` — generates getter, `_add_*()`, `get_or_add_*()`, `_remove_*()`, `_insert_*()`
- `ZeroOrMore(tag, successors=(...))` — generates `*_lst` property, `add_*()`, `_insert_*()`
- `successors` tuple must match XSD schema ordering exactly — consult `../ooxml-reference-corpus/spec/ecma-376-5/part-1/xsd/wml.xsd` (WordprocessingML), `dml-wordprocessingDrawing.xsd` (anchor/inline drawing), or `shared-math.xsd` (OMML) for authoritative ordering. The sibling `rnc/` directory has RELAX NG Compact variants that are easier to read than the XSDs.
- Register: `register_element_cls("w:footnote", CT_Footnote)` in `oxml/__init__.py`

### Part Classes

Extend `XmlPart` or `StoryPart`. Follow `CommentsPart` as a model:

```python
class FootnotesPart(StoryPart):
    @classmethod
    def default(cls, package):
        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        element = cast("CT_Footnotes", parse_xml(cls._default_xml()))
        return cls(partname, content_type, element, package)
```

Wire into `DocumentPart` with lazy creation:
```python
@property
def _footnotes_part(self):
    try:
        return self.part_related_by(RT.FOOTNOTES)
    except KeyError:
        part = FootnotesPart.default(self.package)
        self.relate_to(part, RT.FOOTNOTES)
        return part
```

Register in `src/docx/__init__.py`:
```python
PartFactory.part_type_for[CT.WML_FOOTNOTES] = FootnotesPart
```

### Proxy Objects

Wrap CT_ elements. Inherit from `ElementProxy`, `StoryChild`, or `BlockItemContainer`:

```python
class Footnote(BlockItemContainer):
    @property
    def footnote_id(self):
        return self._element.id
```

### Constants

- Content types: `src/docx/opc/constants.py` — `CT.WML_FOOTNOTES` and `CT.WML_ENDNOTES` already defined
- Relationship types: same file — `RT.FOOTNOTES` and `RT.ENDNOTES` already defined
- Namespaces: `src/docx/oxml/ns.py` — `qn("w:footnote")` for Clark notation

## OOXML feature workflow (required before implementing any new feature)

Every OOXML feature is defined by a manifest in the shared corpus
repository `loadfix/ooxml-reference-corpus` (sibling checkout at
`../ooxml-reference-corpus/`). Before implementing any new feature:

1. **Read the manifest.** Look under
   `../ooxml-reference-corpus/features/docx/` for a JSON manifest whose
   `assertions` block defines what "passing" means. Example:
   `features/docx/bold-text.json` — the XPath there is the canonical
   definition of "bold".

2. **Consult the ECMA-376 5th edition spec** (corpus-only — the spec
   archive is intentionally NOT duplicated into this repo):
   - PDF: `../ooxml-reference-corpus/spec/ecma-376-5/part-1/Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf`
   - RNC schemas (easier to read): `../ooxml-reference-corpus/spec/ecma-376-5/part-1/rnc/`
   - XSD schemas (authoritative for validators): `../ooxml-reference-corpus/spec/ecma-376-5/part-1/xsd/`

3. **If no manifest exists yet**, author one *first*. The manifest is
   the definition of done; the library code is an attempt at it. Follow
   the schema at
   `../ooxml-reference-corpus/features/manifest.schema.json`. Commit:
   the JSON manifest, a `scripts/gen_<name>.py` generator, the
   machine-generated fixture under `fixtures/docx/`, and (optional but
   valuable) an Office-authored `.office.docx` companion.

4. **Verify conformance.** Implementation passes when
   `ooxml_validate.conformance.run_feature(manifest, library="python-docx",
   fixture_path=output)` returns `status="pass"`.

## OOXML spec vs Microsoft Word reality

Microsoft Word does NOT strictly implement ISO/IEC 29500 / ECMA-376. Treat the spec as a starting point, not ground truth.

- Word writes the **Transitional** flavor, not **Strict**. The 4th/5th/6th editions of ISO 29500-1 tightened the spec toward Strict; Word still emits Transitional namespaces that trace back to the original 1st edition / ECMA-376 2006.
- Word emits Microsoft extensions in the `w14:`, `w15:`, `w16:`, `w16cid:`, `w16se:` namespaces (Word 2010/2013/2016+), gated by `mc:Ignorable`. These are documented in the `[MS-DOCX]` / `[MS-OE376]` extension series, not in the corpus-level spec.
- Word's reader tolerates out-of-order, extra, and missing elements that the spec forbids. Word's writer emits shapes the spec doesn't mandate. A spec-valid file is not automatically a file Word will open cleanly.
- **When the spec and Word disagree, match Word.** The canonical way to resolve ambiguity is: save a minimal `.docx` from Word, unzip it, and inspect the XML. The corpus-level `.xsd` files tell you what is *allowed*; the `.office.docx` companions in the corpus tell you what is *interoperable*.

## Test Conventions

- Framework: pytest with BDD-style naming
- Test classes: `Describe*` pattern
- Test methods: `it_*`, `its_*`, `they_*` prefixes
- Test XML: `cxml.element("w:footnotes/(w:footnote{w:id=1})")` — compact XML expression language
- Mocks: `class_mock(request, "dotted.path")`, `instance_mock(request, Class)`, `method_mock(request, Class, "name")`
- Test utilities in `tests/unitutil/`
- Acceptance tests live under `features/` (behave, Gherkin `.feature` files plus step modules).

Example:
```python
class DescribeCT_Footnotes:
    def it_can_add_a_footnote(self):
        footnotes = cast(CT_Footnotes, element("w:footnotes"))
        footnote = footnotes.add_footnote()
        assert footnote.id == 2
```

## Commands

```bash
# Run tests
pytest tests/ -v

# Run a specific test
pytest tests/unit/test_footnotes.py -v

# Run acceptance tests
behave features/

# Type check
pyright src/

# Install in dev mode
pip install -e ".[dev]"
```

### Running the conformance (round-trip) harness

`tests/conformance/` hosts the byte-identical round-trip harness — the
docx counterpart of the vsdx harness at
`python-vsdx/tests/conformance/`. It iterates every `*.office.docx`
fixture under `~/code/ooxml-reference-corpus/fixtures/docx/`, loads
each via `docx.Document`, re-serialises via `.save(BytesIO)`, and
asserts that every zip entry is byte-identical to the original's.

The harness is **pure instrumentation** — it surfaces drift without
fixing it. When a fixture fails, file a fidelity bug with the
drifting part name and a short preview of the divergence; do **not**
relax the harness's byte-equality contract to mask a drop. Every
relaxation risks hiding a real silent data loss.

```bash
# Run only the conformance harness
pytest -m conformance tests/conformance/ -v

# Run everything except the conformance harness (fast unit loop — the CI default)
pytest -m 'not conformance' tests/

# Override the corpus lookup path
DOCX_CORPUS_ROOT=/path/to/alt/corpus pytest -m conformance tests/conformance/

# See the list of fixtures pytest discovered without running them
pytest -m conformance --collect-only -q tests/conformance/
```

The harness **skips cleanly** when no `*.office.docx` fixtures are
present (clean checkout, CI without the corpus mounted) — landing or
removing fixtures never causes a green run to go red for
infrastructure reasons.

Per-entry diff: when a part's bytes drift, the failure names the
drifting zip entry and shows the first 200 chars of the original and
saved XML side-by-side. Whole-file dumps are deliberately avoided —
the investigator gets a scannable hint, not a megabyte of XML. See
`tests/conformance/diff.py` for the format.

Only Office-authored fixtures (filename suffix `.office.docx`)
participate; generator-authored `*.docx` fixtures are synthesised by
the same writer under test and so don't exercise the `reader →
writer` fidelity path this harness enforces (the existing
`test_reproducible_save.py` covers writer-writer determinism).

#### Spec-validity companion (`test_validate.py`)

`tests/conformance/test_validate.py` is an orthogonal lane to the
byte-round-trip harness. For each fixture, it pipes python-docx's
save output through
[`loadfix/ooxml-validate`](https://github.com/loadfix/ooxml-validate)
— a Python wrapper around the Microsoft Open XML SDK validator — and
asserts the issue list is empty. This catches a class of bug the
byte-equality contract cannot: output that is *spec-invalid* but that
Microsoft Word still opens (Word is famously lenient about its own
schema).

```bash
# Install ooxml-validate + its .NET 8+ runtime dep
pip install -e '.[conformance]'
sudo apt-get install -y dotnet-runtime-8.0  # or brew install --cask dotnet

# Run just the validator lane
pytest -m conformance tests/conformance/test_validate.py -v
```

The lane skips cleanly when `ooxml-validate` is not installed (default
`[dev]` extras) or when the `dotnet` executable is missing. Fixtures
the reader refused are skipped (nothing to validate). Fixtures whose
byte-round-trip *fails* are still validated — spec-validity is
independent of fidelity.

Non-goal: this harness does not fix validator findings. When a fixture
fails, file a ticket; do not relax the assertion to mask a real
emission bug. See `tests/conformance/test_validate.py` module docstring
for the full contract.

## What NOT to do

- Don't amend or force-push to `master`, and never force-push to an upstream remote under any circumstance.
- Don't commit secrets, API tokens, local scratch output, or generated docs.
- Don't add runtime dependencies lightly — every new dep affects a large user base. If you must, raise it first.
- Don't introduce backwards-incompatible API changes without a HISTORY/FEATURES note and a transition plan (deprecation warning where possible).
- Don't silence warnings with broad `filterwarnings` ignores — they exist to catch real problems.
- Don't delete `py.typed`; removing it silently breaks downstream type-checking.
- Don't bypass the xmlchemy descriptor layer with raw `lxml.etree` access in production code — the descriptors carry namespace, type, and default semantics.
- Don't move unit tests out of their current location or rename test methods away from the `Describe*` / `it_*` BDD convention — test discovery relies on it.

## Common workflows

### Adding a new public method on an existing class
1. Implement in the appropriate `src/docx/…` module.
2. Add unit tests in the mirrored test file under `tests/`.
3. Add a behave scenario under `features/` if the capability is user-visible.
4. Update `FEATURES.md` — refresh the entry and snippet, verify the snippet runs against a fresh `Document()`.

### Adding a new enum value
- Enums live in `src/docx/enum/`. Read a neighboring enum first to see the XML-mapping pattern.
- Update any doc reference that enumerates the valid values.

### Adding a new XML element class
- Custom element classes live in `src/docx/oxml/…`. They use the `xmlchemy` descriptor layer (`ZeroOrOne`, `OneAndOnlyOne`, `ZeroOrMore`, `RequiredAttribute`, `OptionalAttribute`, …).
- Consult `../ooxml-reference-corpus/spec/ecma-376-5/part-1/xsd/*.xsd` (or the easier-to-read sibling `rnc/*.rnc`) for authoritative element ordering before declaring `successors`.
- Register the class with `register_element_cls("w:tag", CT_Tag)` at the bottom of `src/docx/oxml/__init__.py`.
- Save a minimal `.docx` from Word that exercises the element, unzip it, and compare — **when the spec and Word disagree, match Word**.

### Keep README.md and TODO.md current

Whenever a feature is added, removed, or a public option changes, update both of these files *in the same PR* as the code change — stale docs have bitten us before.

- **`README.md`** — the API block reflects the real public surface. If you add/remove a function or option, add/remove the matching entry. If you add or remove an export, reflect it in the API section. Any prose sections (Status, Contributing, project-specific sections) should also match reality.
- **`TODO.md`** — if the change resolves a tracked issue, move that entry into a "Resolved in fork" / "Done" section with a one-line description and the PR/commit reference. Update any counts table at the top and bump the "last updated" date.

Minimum check before every PR that touches source: `grep -n "<feature name>" README.md TODO.md` to catch stale references.

## Important

- Before implementing a new feature or element class, consult `../ooxml-reference-corpus/spec/ecma-376-5/` for authoritative schema information: `part-1/xsd/*.xsd` (W3C XSD grammars), `part-1/rnc/*.rnc` (RELAX NG Compact equivalents, easier to read), and the four `Ecma Office Open XML Part N.pdf` files (markup-language reference, OPC packaging, markup compatibility, transitional migration features). These live in the corpus repo rather than here to avoid duplicating ~50 MB of PDFs across six sibling projects — they are not runtime dependencies, just the canonical sources for element ordering, attribute types, and cardinality.
- Keep `FEATURES.md` current when adding, modifying, or deleting public API. It is a single-page catalogue of every public feature (43 sections, ~1800 lines) with fork additions marked `[Added in 1.3.0.dev0]`. For each change: add the new entry (or update/remove the existing one) under the relevant section, refresh its snippet if the API surface shifted, and verify the snippet still runs against a fresh `Document()`.
- Always run tests after changes: `pytest tests/ -v`
- The successors tuple in element declarations MUST match XSD ordering
- Footnote IDs 0 and 1 are reserved (separator, continuation separator)
- Use `src/` layout — all code is under `src/docx/`, not `docx/`
- Follow existing code style: no docstring on test methods, BDD-style names
- XML templates go in `src/docx/templates/`
