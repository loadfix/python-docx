.. :changelog:

Release History
---------------

Unreleased — ``docx.kit.lint`` walks every story (#673)
+++++++++++++++++++++++++++++++++++++++++++++++++++++++

- ``docx.kit.lint.lint(document)`` no longer iterates ``document.paragraphs``
  (body-only) for paragraph-/run-/picture-scoped rules. Every rule
  whose scope is "every paragraph in the document" now composes
  against the public cross-story walker
  (:meth:`Document.iter_all_paragraphs` / :meth:`iter_all_runs` /
  :meth:`iter_all_pictures`, shipped in #708) so headers, footers,
  footnotes, endnotes, comments, and table-cell paragraphs are
  surfaced too. A footer with ``Page  X of  Y`` (double space), a
  header with ``TODO: insert client name``, or a footnote with mixed
  quotes is now flagged the same as the equivalent body defect.

  Rules migrated: ``multiple-spaces``, ``trailing-whitespace``,
  ``tab-instead-of-indent``, ``leading-spaces-instead-of-indent``,
  ``mixed-quotes``, ``empty-paragraph``,
  ``inconsistent-heading-levels``, ``mixed-fonts``,
  ``over-long-paragraph``, ``placeholder-text``, ``bare-url``,
  ``missing-alt-text``. ``excessive-font-size-variation`` was
  already cross-story (per #646). ``missing-document-title``
  (document-level), ``table-without-header-row`` (table-scoped),
  ``trailing-empty-paragraph`` (end-of-body trailing), and
  ``trailing-heading`` (body-only end-of-document) keep their
  previous scope.

  Per-story state — the ``empty-paragraph`` consecutive-run
  accumulator and the ``inconsistent-heading-levels`` previous-level
  tracker — resets at every story boundary so a body→header
  transition never pairs the trailing body empty with the header's
  first paragraph or compares the body's last heading to a header
  heading.

  Finding shape: ``Finding.location`` now carries the walker's tag
  verbatim for non-body findings (``"header:section0:primary"``,
  ``"footnote:2"``, ``"table:0:row:0:col:0"``, …) and keeps the
  legacy ``"paragraph N"`` shape for body findings.
  ``Finding.paragraph_index`` is populated only for body findings so
  existing autofix lookups (``document.paragraphs[paragraph_index]``)
  keep working. Non-body findings carry
  ``autofix_available=False`` — the autofix path is body-only in
  this PR; cross-story autofixes are a follow-up. Closes #673.

Unreleased — ``docx.kit.lint`` — structured locator + quantitative fields on ``Finding.details`` (#675)
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

- Every built-in rule that historically buried a locator or
  quantitative datum inside the human-readable ``message`` /
  ``location`` strings now exposes the same data as a structured
  field on ``Finding.details``. Callers driving autofixes or
  building structured reports can read ``finding.details[...]``
  instead of regex-parsing the message prose. Specifically:

  - ``tab-instead-of-indent`` — ``details["tab_count"]``.
  - ``trailing-whitespace`` — ``details["trailing_count"]``.
  - ``inconsistent-heading-levels`` — adds
    ``details["actual_level"]`` and ``details["expected_level"]``
    alongside the existing ``level`` / ``previous_level`` /
    ``skipped`` fields.
  - ``empty-paragraph`` — ``details["consecutive_count"]`` and
    ``details["run_start"]`` for the trailing run of empties.
  - ``missing-alt-text`` — adds ``details["shape_index"]`` (the
    first occurrence's index) and ``details["cnvpr_id"]`` (the
    ``wp:docPr/@id`` integer) alongside the existing
    ``occurrence_count`` / ``additional_locations`` fields.
  - ``bare-url`` — ``details["url"]`` (one finding per match).
  - ``excessive-font-size-variation`` — ``details["distinct_sizes"]``
    (the sorted tuple of distinct point sizes) and
    ``details["threshold"]``.
  - ``placeholder-text`` — adds ``details["placeholder_text"]`` as
    a key alias for the existing ``details["placeholder"]`` so the
    naming aligns with the other locator-fields rules.

  The ``message`` strings are unchanged so callers that still string-
  parse keep working. Closes #675.

Unreleased — Docs: rewrite ``bind_tokens`` module docstring to match MCE reality (#735)
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

- The ``docx.bind_tokens`` module docstring previously justified
  emitting a fork-internal ``<lfxbind:src>`` element directly inside
  ``<w:r>`` by appealing to a stated OOXML convention that "Word and
  every other OOXML consumer follow the 'preserve but ignore unknown
  children' convention". That premise is wrong — ECMA-376 has no such
  blanket rule. The actual contract is the Markup Compatibility and
  Extensibility (MCE) framework defined in ECMA-376 Part 3, under
  which unknown-namespace elements are only safely ignored when their
  prefix is listed in ``mc:Ignorable`` on a containing element or
  when they sit inside an ``<mc:AlternateContent>`` wrapper.
  Microsoft Word's loader rejects bare unknown-namespace elements as
  schema-violating, which is the failure mode that drove #733.

  The "Round-trip preservation" section of the module docstring is
  rewritten to drop the false convention claim, state the real MCE
  contract, and document the strategy taken after #733: an opt-in
  gate on bound record / persisted marker, root-declared prefix
  hoisted via ``etree.cleanup_namespaces`` (never inline), and a
  defensive guard in ``_write_source_marker`` that suppresses
  emission when the prefix declaration is missing. The same
  correction is applied to the ``lfxbind`` registration comment in
  ``docx.oxml.ns``. Pure documentation; no behaviour change.
  Closes #735.

Unreleased — Word-compatibility fix: gate ``apply_bind_tokens`` on opt-in (#733)
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

- ``Document.save()`` no longer runs ``docx.bind_tokens.apply_bind_tokens``
  unconditionally. The resolver previously fired on every save and stamped
  a fork-internal ``<lfxbind:src xmlns:lfxbind="...">…</lfxbind:src>``
  marker inside any ``<w:r>`` whose text contained a brace-quoted
  identifier (``{customer-code}``, ``{role}`` in template instructions,
  ``{var}`` in code samples). The marker's namespace is not declared on
  the document root and is not wrapped in ``mc:AlternateContent``, so
  Microsoft Word rejected the saved package as malformed — every saved
  ``.docx`` containing matching prose became Word-unloadable after a
  single python-docx round-trip, even when the caller never opted into
  the bind-tokens feature.

  The resolver now fires only when ``Document.bind`` /
  ``add_paragraph(..., bind_to=...)`` (or ``set_bound_record``) has bound
  a record, *or* when the document already carries a ``<lfxbind:src>``
  marker from a previous bound save (so previously-bound documents
  continue to round-trip cleanly). For callers who never opted in,
  brace-quoted prose now round-trips byte-clean. New helper
  ``docx.bind_tokens.has_persisted_marker(document)`` exposes the
  second half of the gate. A defensive guard in ``_write_source_marker``
  also suppresses any marker emission when the document root lacks the
  ``lfxbind`` prefix declaration — the inline ``xmlns:lfxbind`` form is
  what triggered Word's namespace-mismatch rejection. When binding is
  active, ``set_bound_record`` and ``apply_bind_tokens`` now hoist the
  ``lfxbind`` prefix onto the document root via
  ``etree.cleanup_namespaces`` so subsequent marker writes inherit the
  root declaration. Closes #733.

Unreleased — Hyperlink ergonomics
+++++++++++++++++++++++++++++++++

- **``docx.kit.lint`` — ``mixed-quotes`` exposes structured counts**
  (#672). The ``mixed-quotes`` finding now populates
  ``details["smart_count"]`` and ``details["straight_count"]`` with
  the integer counts of curly (``“”‘’``) and straight (``"`` / ``'``)
  quote characters in the offending paragraph. Agents driving an
  autofix can read ``finding.details`` directly to make a dominance
  decision instead of re-walking the paragraph's runs or text-parsing
  the human-readable message. The legacy message text
  (``paragraph N mixes smart (curly) and straight quotes``) is
  preserved unchanged for back-compat with rendered reports.

- **``docx.kit.lint`` — ``multiple-spaces`` exempts intentional
  formatting** (#645). Heading-styled paragraphs (`Heading 1` …
  `Heading 9`, `Title`, `Subtitle`) whose multi-space run sits
  immediately after a leading ``\d+(\.\d+)*`` numeric prefix
  (e.g. ``4.1  Three-LZA topology``) and List-styled paragraphs
  (`List Bullet`, `List Number`, `List Paragraph`, `List Continue`,
  `Body Text Indent`, `Quote`) whose match starts at the very
  beginning of the paragraph (a hanging indent before the bullet
  glyph) no longer fire. The exemption is per-match, not
  per-paragraph: a heading that combines an intentional numbering
  gap with a mid-sentence double-space defect still flags the
  defect, and the autofix repairs only the non-exempt run. Real
  Word documents using these template conventions are no longer
  silently ruined by the autofix.

- **``docx.kit.lint`` — ``over-long-paragraph`` + ``trailing-whitespace``
  per-style exemptions and threshold kwarg** (#649). The
  ``over-long-paragraph`` rule's exemption matching now tolerates
  Word's numbered style variants — ``List Bullet`` matches ``List
  Bullet 2`` / ``List Bullet 3``, ``Caption`` matches ``Caption 2``,
  etc. — and the default exemption set grows to cover ``List
  Continue``, ``Endnote Text``, and ``Intense Quote`` alongside the
  existing list / caption / footnote / quote families. The CenITex
  baseline regressed a deliberate compound bullet point styled
  ``List Bullet`` whose 1000+ character body was editorial intent;
  the prefix-with-trailing-space match silences it without losing
  the rule's value on body prose. The ``trailing-whitespace`` rule
  gains a verbatim-style exemption — paragraphs whose style is
  ``Code`` / ``Preformatted`` / ``Plain Text`` / ``HTML
  Preformatted`` / ``Macro Text`` (case-insensitive, plus numbered
  variants) keep their trailing whitespace because in those styles
  it's load-bearing rather than drift. Two new predicates surface
  the per-rule logic: ``_is_overlong_exempt_style(paragraph,
  exemptions)`` and ``_is_trailing_whitespace_exempt_style(paragraph)``.
  ``lint()`` also accepts a convenience ``over_long_threshold=N``
  kwarg so callers can tune the most-touched knob without building a
  full :class:`LintConfig`; the kwarg wins over a config-supplied
  value when both are passed.

- **Skip empty ``docProps/custom.xml`` on save** (#721). A fresh
  ``Document().save()`` no longer materialises an empty
  ``docProps/custom.xml`` part with the matching
  ``[Content_Types].xml`` override and ``_rels/.rels`` entry. The
  save path's bind-tokens helper still touches
  ``document.custom_properties`` to build a property map, but a
  library-authored part with zero ``<cst:property>`` children is now
  pruned at save time (alongside the existing thumbnail prune in
  ``OpcPackage._drop_unused_package_rels``). Round-trip fidelity is
  preserved: a docx that arrived with an empty custom-properties
  part keeps it across read+save (``_loaded_from_package = True``
  guard), so files authored by Microsoft Office or older python-docx
  releases survive a reload unchanged. Reading a document with or
  without the part continues to work — the public
  ``Document.custom_properties`` accessor still creates the part on
  demand the first time a property is added.

- **``docx.kit.lint`` — ``missing-alt-text`` decorative + dedupe
  enrichments** (#651). Each finding now carries
  ``details["occurrence_count"]`` (how many shapes share the image
  binary) and ``details["additional_locations"]`` (the sibling
  locator strings beyond the first occurrence) so callers rendering
  reports can list every duplicate without recomputing the dedupe.
  Decorative detection already honoured python-docx's
  ``a11y_role == "decorative"`` and Office 365's
  ``<a16:decorative val="1"/>`` extension marker on
  ``wp:docPr/a:extLst``; this round adds the structured ``details``
  payload and the dedicated regression tests for the five-occurrence
  "logo on every header" pattern.

- **``docx.kit.lint`` — new ``trailing-heading`` rule** (#644). Flags
  heading paragraphs at the end of the document with no body content
  beneath them — a section title that promises content the document
  never delivers. Catches both the unfinished-section pattern (the
  motivating case: the last paragraph of a draft is
  ``Heading 1 '11. Glossary'`` with nothing under it) and the common
  authoring bug where ``Heading`` style is auto-applied to a final
  pasted line. Severity ``info``; no autofix (the rule cannot guess
  what content the author intended). Walks ``Document.iter_inner_content``
  so a trailing table counts as body content even when its cells are
  empty; only paragraphs without text count toward "no body content".
  Adjacent trailing headings are each flagged independently — a
  following heading is the next section title, not body content for
  this one.

- **``docx.kit.lint`` — ``multiple-spaces`` detects cross-run double
  spaces** (#657). The detector and autofix now operate on the joined
  text of a paragraph's runs instead of scanning each run in
  isolation, so the common bold/italic mid-phrase pattern (``"emphasized "``
  + ``" text"``) — whose joined ``paragraph.text`` shows a visible
  double space but whose individual runs each carry only a single
  trailing/leading space — is flagged and fixable. The autofix
  follows a deterministic tie-break rule: the surviving single space
  lands in the *first* run that contributed a space, so the
  formatting (bold, italic, font) of the run that "owned" the gap is
  preserved. The new ``Finding.details`` payload exposes
  ``run_indices``, ``match_start``, and ``match_end`` for tooling that
  needs to reason about the affected span. The new
  ``_collapse_cross_run_spaces`` helper is the underlying primitive.

- **``docx.kit.lint`` — new ``leading-spaces-instead-of-indent``
  rule** — sibling to ``tab-instead-of-indent`` for authors who fake
  an indent with the spacebar (common from web / markdown copy-paste).
  Fires on body paragraphs whose first run starts with four-or-more
  leading ``ASCII space`` characters (configurable via the module-
  level ``LEADING_SPACES_MIN_RUN``); severity ``info``. The autofix
  strips the leading spaces and stamps a compensating
  ``paragraph_format.left_indent`` of 36 pt (≈ 0.5 inch) per
  threshold-wide block, layered on top of any existing direct
  indent. Skips the same heading / TOC / list / hanging-indent
  paragraphs the tab variant skips. Closes #676.

- **``docx.kit.lint`` adds the ``bare-url`` rule** (#647). Flags
  paragraphs whose plain text contains a raw URL (``https://...``,
  ``http://...``, ``www.…``) that is not wrapped in a
  ``<w:hyperlink>`` element. Severity is ``info`` and the finding is
  read-only — choosing the visible link text and the relationship
  target is meaning-bearing, so an autofix would routinely guess
  wrong. URLs already covered by an existing
  :class:`~docx.text.hyperlink.Hyperlink` (matched against the
  paragraph's ``hyperlinks`` collection) are skipped, and trailing
  sentence punctuation (``.,;:!?)]}>"'``) is stripped from the
  reported URL so a sentence-ending period is not treated as part
  of the link. Multiple bare URLs in the same paragraph yield one
  finding per URL.

- **Round-trip fidelity: paragraph-mark formatting mirror is opt-in
  (#734)** — ``Document.save()`` no longer copies the first run's
  ``<w:rPr>`` onto every paragraph's ``<w:pPr><w:rPr>`` by default.
  The mirror was firing unconditionally on every save (including
  no-op ``Document(p).save(p2)`` round-trips) and inflated body XML
  by ~16% on real documents — a 173 KB fixture grew from 172 KB to
  200 KB on a single round-trip with no edits. The mirror also
  silently extended ``<w:bCs/>``/``<w:iCs/>`` (auto-written by the
  ``Font.bold`` / ``Font.italic`` setters) onto paragraph marks the
  source document never marked as complex-script. Pass
  ``mirror_paragraph_marks=True`` to ``Document.save()`` to recover
  the historical "keep typing in bold" emission shape; even with the
  flag enabled the complex-script tags are now mirrored only when
  the source paragraph mark already carried them. Closes #734.

- **``docx.kit.lint`` — new ``excessive-font-size-variation`` rule**
  (#646). A document-scoped check that aggregates every explicit
  ``run.font.size`` across non-heading paragraphs and emits a single
  ``info`` finding when the count of distinct sizes exceeds the
  module-level threshold ``_EXCESSIVE_FONT_SIZE_THRESHOLD`` (default
  ``4``). The motivating case is the regenerated CenITex masters whose
  body runs carry seven distinct sizes (9, 12, 13, 15, 18, 20, 34 pt)
  — a strong signal of inconsistent direct formatting that survived
  editing. The finding's ``message`` lists the distinct sizes in
  ascending order so the author can scan for the stragglers.
  Heading paragraphs are skipped (their sizes are intentionally
  varied) and runs whose ``font.size`` is ``None`` (inheriting from
  their style) are not counted as drift. Read-only — collapsing sizes
  is a meaning-bearing decision the author must make, so no autofix.

- **``Document.iter_all_paragraphs()`` /
  ``Document.iter_all_runs()`` / ``Document.iter_all_pictures()``** —
  promote the previously-private ``docx.search._iter_all_paragraphs``
  walker to a documented public surface (#662). Each method yields
  ``(item, location)`` tuples covering every story in the document —
  body, body-level table cells, every non-inherited
  header / footer on every section, plus the footnote / endnote /
  comment parts — paired with a stable ``location`` tag drawn from
  the same vocabulary used by ``docx.search.search_all_paragraphs``
  (``"body"``, ``"table:0:row:1:col:2"``,
  ``"header:section0:primary"``, ``"footer:section0:even_page"``,
  ``"footnote:2"``, ``"endnote:3"``, ``"comment:5"``). All three
  methods share keyword-only ``include_tables=`` /
  ``include_headers_footers=`` / ``include_footnotes=`` /
  ``include_endnotes=`` / ``include_comments=`` flags; the body group
  is always yielded. Foundation for the cross-story ``lint()`` walker
  in #673. The lower-level ``docx.search.iter_all_paragraph_groups``
  building block is also public; the underscore alias
  ``_iter_all_paragraphs`` is retained as a deprecated forwarder for
  third-party callers.

- **``Document.from_html(source, clean=True)`` /
  ``Document.from_html_string(html, clean=True)``** — read-side
  companion to ``Document.to_html()`` (#95). Stdlib-only HTML parser
  (``html.parser``) — no ``BeautifulSoup`` dependency. Maps the major
  block + inline structural elements (``<h1>``-``<h6>``, ``<p>``,
  ``<strong>``/``<b>``, ``<em>``/``<i>``, ``<u>``, ``<a href>``,
  ``<ul>``/``<ol>``/``<li>``, ``<table>``, ``<img>``, ``<blockquote>``,
  ``<code>``, ``<pre>``) onto Word equivalents. ``clean=True`` (the
  default) strips ``<script>``/``<style>``/comments and drops
  ``class``/``id`` attributes; ``style`` attributes are honoured only
  for ``color: <hex>`` (best-effort). Hyperlink and image schemes are
  restricted to ``http``/``https``/``mailto`` so a round-trip
  ``from_html`` → ``to_html`` cannot promote attacker-controlled
  schemes. Remote ``<img src>`` URLs are *not* fetched — they degrade
  to alt-text — only ``data:`` URLs are decoded and embedded as
  pictures. LaTeX import is intentionally out of scope; use
  ``docx.math.OMath`` for programmatic equations. Closes #95.

- **``Paragraph.add_markdown(md)`` / ``_Cell.add_markdown(md)``** —
  inline Markdown rendering for paragraphs and table cells (#23).
  Supported subset: bold (``**...**`` / ``__...__``), italic
  (``*...*`` / ``_..._``), inline code (``` `...` ``), inline links
  (``[text](url)``), bullet lists (``-`` / ``*``) → ``List Bullet``
  paragraphs, numbered lists (``1.`` …) → ``List Number`` paragraphs,
  ATX headings (``#`` … ``######``, top-of-input only) →
  ``Heading N`` paragraphs, blank-line paragraph separator, and
  single-``\n`` soft line-breaks (``w:br``). Tables, images, fenced
  code blocks, footnotes, blockquotes, and horizontal rules are out
  of scope (escape-hatch: pass raw text via :meth:`add_run`). The
  Markdown source is *not* preserved on the document — after
  ``add_markdown`` returns, the paragraph holds the equivalent OOXML
  and a subsequent read does not recover the original markdown.
  Implementation is a small hand-rolled tokenizer (no new runtime
  dep). Closes #23.

- **``Document.add_paragraph(text, bind_to=record)`` /
  ``Document.bind(record=...)``** smart-placeholder fields (#68). Token
  strings such as ``{customer.name}``, ``{customer.address.line1}``,
  ``{date:short}``, ``{date:'MMM d, yyyy'}``, ``{property:Title}``,
  and ``{i}`` (iteration index) are resolved at every
  :meth:`Document.save` against the bound record / document
  properties. The original token-source string is preserved in a
  fork-scoped ``<lfxbind:src>`` child of each affected ``<w:r>`` so
  that ``load -> bind -> save`` cycles re-resolve cleanly against
  the new record instead of carrying the previously-resolved
  literal forward. Mirrors the page-token mechanism shipped on
  python-pptx for issue #38; same machinery, docx namespace.
  ``[Added in 2026.05.13]``
- **``Document.diff(other, level="content")``** computes a semantic
  diff between two documents (#75) — paragraph add / remove / modify,
  table mutations, image presence, and (at ``level="formatting"``)
  style / font / colour changes. The returned
  :class:`docx.semantic_diff.SemanticDiff` exposes :attr:`summary`
  (counts dictionary) and :attr:`changes` (list of
  :class:`Change` records), plus three exporters:
  :meth:`SemanticDiff.to_markdown` for PR comments,
  :meth:`SemanticDiff.to_html` for web review UIs, and
  :meth:`SemanticDiff.to_word_track_changes` for a best-effort
  third-document rendering with visible ``[INS]`` / ``[DEL]`` /
  ``[~MOD]`` markers (full ``w:ins`` / ``w:del`` track-changes
  authoring is out of scope for this exporter — the markers are
  human-readable text rather than reviewable Word revisions).
  Comparison uses :mod:`difflib` over normalised paragraph text, so
  whitespace-only re-flow is suppressed.
- **``Paragraph.add_link_to(target, ...)``** is a polymorphic wrapper
  for internal hyperlinks (#69). ``target`` accepts a |Bookmark|, a
  heading |Paragraph| (any paragraph whose style name starts with
  ``"Heading "``; a covering bookmark is auto-allocated when one
  doesn't already exist), or a bare bookmark-name string. Visible text
  defaults to the target's text; ``style=`` and a new ``tooltip=``
  keyword work the same as on :meth:`Paragraph.add_hyperlink`.
- **``Paragraph.add_hyperlink(..., tooltip=...)``** writes the
  ``w:hyperlink/@w:tooltip`` attribute (the popup text Word renders on
  hover). The new attribute round-trips through save/load. The
  matching read accessor :attr:`Hyperlink.tooltip` is read/write.
- **``Paragraph.add_url(url, ...)``** is the external-hyperlink
  counterpart (#70). Auto-prepends ``mailto:`` when ``url`` looks like
  an email but has no scheme, ``tel:`` for phone-shape strings, and
  ``http://`` for ``www.`` shortcuts. Visible text defaults to the
  caller's literal ``url`` argument so the displayed text matches what
  was passed in.
- **``Paragraph.add_text_with_links(text, ...)``** scans `text` for
  URLs and emails, splits the input into plain runs and hyperlinks,
  and appends them in document order (#70). Trailing sentence
  punctuation (``.``, ``,``, ``;``, ``!``, ``?``, ``)``, …) is stripped
  off URL matches and stays in the surrounding plain run.
- **``Document.outline()``** returns a hierarchical heading-tree
  snapshot of the body (#71) — the docx parallel of pptx's
  ``deck.summarize()`` / ``skeleton()``. Each
  :class:`docx.outline.OutlineNode` carries ``level``, ``text``,
  ``paragraph_index``, a stable 8-char ``id``, ``word_count`` for
  the section's body text, and a list of nested ``children``. The
  wrapper :class:`docx.outline.Outline` exposes ``walk()`` for
  depth-first traversal, ``find(heading)`` for text lookup, and
  ``to_dict()`` for JSON-serialisable output. Page numbers are
  intentionally omitted — python-docx has no layout engine — but the
  document-wide ``total_pages_estimated`` reads Word's cached
  ``docProps/app.xml`` ``<Pages>`` value when present.
  ``Document.slice(start, end)`` returns a new |Document| containing
  the paragraphs of one heading-bounded section, copied via
  :meth:`append_paragraph` so image / hyperlink / style references
  are rewired into the slice. ``[Added in 2026.05.7]``
- **``Paragraph.add_footnote(text="")`` / ``Paragraph.add_endnote(text="")``**
  ergonomic authoring API (#42) [Added in 2026.05.7]. Appends a reference
  run to the paragraph, allocates a fresh ``w:footnote`` / ``w:endnote``
  in the corresponding part, and seeds the new note with ``text``.
  Returns the |Footnote| / |Endnote| so the caller can append further
  paragraphs, runs, or tables. Refuses to nest a note inside another
  note. Companion shorthand setters
  ``Document.footnotes.numbering = "i, ii, iii"`` and
  ``Document.footnotes.restart = "section"`` (plus the matching
  ``Document.endnotes`` setters) accept friendly strings (``"arabic"``,
  ``"chicago"``, ``"*, dagger, double-dagger"``, …), |WD_NUMBER_FORMAT|
  / |WD_FOOTNOTE_RESTART| enum members, or raw OOXML tokens, and write
  through to ``w:footnotePr`` / ``w:endnotePr`` in the settings part.
  Round-trips cleanly with the existing read API.

2026.05.10 — Password-protected read + write
++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-10

python-docx now **reads and writes** password-protected ``.docx``
files (ECMA-376 Agile Encryption — the scheme Word uses when a user
sets a password in the desktop app). Previous releases only *detected*
encrypted input and raised ``EncryptedDocumentError`` pointing users at
an external tool. This release delegates actual AES key derivation and
CFBF (OLE2) compound-document parsing to the new optional
``python-ooxml-crypto`` dependency, mirroring the read/write surface
python-pptx already ships (closes #327 upstream in the sibling repo;
unlocks the same workflow for python-docx).

- **``Document(path, password=...)``** decrypts an encrypted ``.docx``
  on open. Supplying the kwarg with no encryption is a no-op; omitting
  it when the input is encrypted continues to raise
  ``EncryptedDocumentError`` with the now-updated message pointing
  callers at ``python-ooxml-crypto`` instead of the old
  ``msoffcrypto-tool`` recommendation.
- **``Document.save(path, password=...)``** encrypts the output using
  ECMA-376 Agile Encryption. ``flat_opc=True`` and ``password=`` are
  mutually exclusive (Flat-OPC is an XML document, not a zip).
  ``reproducible=True`` and ``password=`` compose normally — the
  fixed-timestamp zip is built first and then encrypted.
- **New ``docx.exceptions.RmsProtectedDocumentError``** (subclass of
  ``EncryptedDocumentError``) is raised when opening a file wrapped in
  Azure RMS / AIP / IRM protection. The payload is keyed to the user's
  Microsoft 365 identity rather than a password, so python-ooxml-crypto
  cannot decrypt it — delegate to Microsoft Office automation or the
  Microsoft Information Protection SDK before opening with python-docx.
- **New adapter module ``docx.opc._crypto``** with the public helpers
  ``is_encrypted_stream``, ``is_rms_protected_stream``,
  ``decrypt_stream``, and ``encrypt_bytes``. The adapter is the single
  point where the optional ``ooxml_crypto`` import is resolved; every
  error from that library is rewrapped as
  ``EncryptedDocumentError`` with an actionable message.
- **Optional install extra.** ``pip install 'python-docx[encryption]'``
  pulls in ``python-ooxml-crypto``. The library keeps zero new
  mandatory runtime dependencies; calling ``Document(path,
  password=...)`` (or ``Document.save(..., password=...)``) without
  the extra installed raises ``EncryptedDocumentError`` with the
  install instructions.

2026.05.9 — Audit bug-fix round
+++++++++++++++++++++++++++++++

Released: 2026-05-05

Small targeted fixes surfaced by the 2026-05-05 audit. No new
feature surface; existing behaviour either gets a regression test
or a crisper error type.

- **vt:date round-trip regression test.** The ``datetime.date``
  serialisation added in 2026.05.8 (commit ``c3edf01b``) now has a
  full ``Document`` → ``custom.xml`` → reload regression test
  (``tests/test_custom_properties.py::DescribeCustomProperties_RoundTrip``)
  so the GitHub issue #171 round-trip behaviour stays locked in.
- **Typed exception on missing ``[Content_Types].xml``** (closes
  #172). Loading a zip that happens to be a valid archive but lacks
  the mandatory OPC content-types part used to leak a bare
  ``KeyError('[Content_Types].xml')`` from ``zipfile.read``.
  ``docx.opc.pkgreader.PackageReader.from_file`` now wraps it in
  ``docx.opc.exceptions.PackageNotFoundError`` at the narrowest
  possible scope, matching the corpus manifest
  ``malformed-content-types-missing`` (whose ``forbidden_exception``
  clause explicitly rejected bare ``KeyError``).
- **Explicit ``__all__`` on 12 public submodules.** ``docx.table``,
  ``docx.section``, ``docx.bookmarks``, ``docx.blkcntnr``,
  ``docx.dml.color``, ``docx.drawing``, ``docx.equations``,
  ``docx.styles.styles``, ``docx.styles.style``,
  ``docx.text.paragraph``, ``docx.text.run``,
  ``docx.text.pagebreak`` now declare the public surface so
  internal ``CT_*`` / ``ST_*`` names can no longer be reached via
  ``from docx.<mod> import *``. Star-import only — existing explicit
  imports continue to work.

2026.05.8 — New authoring APIs
++++++++++++++++++++++++++++++

Released: 2026-05-05

Three independently-developed authoring feature branches landed in
this release, extending the fork's writer surface in areas previously
supported for *read* only (or not at all).

SmartArt
~~~~~~~~

- New ``Document.add_smart_art(layout_name)`` returns a ``SmartArt``
  proxy. Built-in layouts: ``"list"``, ``"cycle"``, ``"process"``.
  Each call provisions the full quartet of SmartArt parts
  (``diagrams/data{N}.xml``, ``layout{N}.xml``, ``quickStyle{N}.xml``,
  ``colors{N}.xml``) from the templates under
  ``src/docx/templates/smart_art/`` and wires the drawing into the
  document body at the current insertion point.
- New ``SmartArt.add_node(text)`` appends a data-point node into the
  underlying ``<dgm:dataModel>``/``<dgm:ptLst>`` with the text you
  supply, picking up the layout's default style so the rendered shape
  picks the right fill/line/font automatically.
- See ``FEATURES.md`` § "SmartArt" for the full snippet.

Bibliography and citations
~~~~~~~~~~~~~~~~~~~~~~~~~~

- New ``Document.bibliography`` property returns a ``Bibliography``
  proxy (read + write). On first access it lazily provisions
  ``/customXml/item{N}.xml`` (with a ``<b:Sources>`` root) plus the
  matching ``itemProps{N}.xml`` and relates both to the document part.
- New ``Document.add_citation(tag, source_type, ...)`` adds a
  ``<b:Source>`` entry to the bibliography. ``tag`` is the key that
  citation references resolve against.
- New ``Paragraph.add_citation_reference(tag)`` inserts an ``SDT``
  citation marker that Word reifies to ``(Author, Year)`` using the
  current bibliography style.
- The save-time custom-XML drop heuristic now preserves freshly-
  authored bibliography parts even without a ``w:dataBinding``
  (citations bind implicitly through matching ``<b:Tag>`` values).
- See ``FEATURES.md`` § "Bibliography and citations".

Field evaluation
~~~~~~~~~~~~~~~~

- New ``Field.evaluate(context)`` and
  ``Document.evaluate_fields(context)`` evaluate complex field codes
  against a supplied context dict. Supported codes:

  - ``MERGEFIELD FieldName`` — substitutes ``context["FieldName"]``.
  - ``IF cond op cond "then" "else"`` — boolean evaluation with
    nested ``{MERGEFIELD}`` allowed on either side of the comparator.
  - ``HYPERLINK "url"`` — resolves to the URL and updates the
    displayed run so the cached result matches.
  - ``= <expr>`` — arithmetic formula evaluator (``+``, ``-``, ``*``,
    ``/``, parentheses, numeric literals, and references to
    ``context`` keys).
  - ``PAGE`` / ``NUMPAGES`` / ``DATE`` / ``TIME`` — runtime-dynamic
    placeholders pulled from the context or from ``datetime.now()``.
- Deferred (raised as ``FieldEvalError``): string-function formulas
  (``=SUM()``, ``=AVERAGE()`` beyond arithmetic), nested ``IF``,
  ``QUOTE``, ``FILLIN``, and the full date-picture/numeric-format
  switch grammar.
- See ``FEATURES.md`` § "Complex-field evaluation".

2026.05.7 — Round-trip fidelity and performance fixes
+++++++++++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-05

Reproducible-save fidelity
~~~~~~~~~~~~~~~~~~~~~~~~~~

- ``Document.save(..., reproducible=True)`` no longer mints
  ``w:rsidR`` / ``w:rsidRDefault`` on paragraphs and runs that don't
  already carry them (#168). Those attributes are session-scoped
  churn markers; synthesising them from a constant-valued root made
  the output reproducible but *not* faithful — round-tripping
  ``bold-text.office.docx`` gained a spurious ``w:rsidR`` on its
  single ``<w:r>``. ``w14:paraId`` / ``w14:textId`` continue to be
  derived deterministically from paragraph content so repeated saves
  remain byte-identical.

Default template rebuild
~~~~~~~~~~~~~~~~~~~~~~~~

- ``src/docx/templates/default.docx`` has been rebuilt from the
  ``default-docx-template/`` source tree (#169) so a fresh
  ``Document()`` exposes the Word-2024 namespace set (``w15``,
  ``w16``, ``w16cex``, ``w16cid``, ``w16du``, ``w16sdtdh``,
  ``w16sdtfl``, ``w16se``, ``cx``–``cx8``, ``aink``, ``am3d``,
  ``oel``) plus the matching ``mc:Ignorable`` list. The unzipped
  tree was updated in 2026.05.2 but the zipped blob was not
  regenerated — ``Document()`` was still loading the pre-2026.05.2
  namespace set at runtime.
- New ``scripts/rebuild_default_template.py`` deterministically
  rebuilds the zipped blob from the source tree so future template
  edits cannot drift out of sync silently.

Narrow part-drop heuristics to preserve Word-authored data
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The 2026.05.4 "word-mimicry phase 3" release introduced aggressive
drop heuristics that silently destroyed optional parts from
Word-authored files on round-trip (#167). This release narrows the
policy so parts that shipped in the source package are preserved
verbatim — dropping happens only when python-docx itself created
the part.

- ``Unmarshaller._unmarshal_parts`` now flags every part it loads with
  ``_loaded_from_package = True``. Save-time heuristics consult this
  flag and preserve any part that shipped in the source package,
  regardless of whether python-docx can statically prove it is
  referenced.
- **``word/stylesWithEffects.xml``** — was dropped unconditionally.
  Now dropped only when python-docx created the part itself (it never
  does today, but the policy is symmetric with the others).
- **``customXml/*``** — was dropped whenever no ``<w:dataBinding>`` was
  present. That false-negatived on customXml used by Power BI,
  bibliography sources, and Office Add-in backing data. Now preserved
  whenever the source package shipped it.
- **``docProps/thumbnail.jpeg``** — was dropped unconditionally at
  the package level. Now preserved whenever the source package
  shipped it. Library-authored documents still skip the thumbnail.
- **``word/numbering.xml``** — the style-indirect heuristic now walks
  the ``w:basedOn`` chain when resolving which styles declare
  ``<w:numPr>``, catching user-defined styles rooted in a numbering
  style (the common "My Bullet → List Bullet" pattern). Dropped only
  when python-docx authored the part and the document uses no
  numbering at all.

Found by W5-A / W5-E / W6-A audits: every Word-authored corpus fixture
round-tripped through the 2026.05.4 drop heuristics lost at least one
of these four parts.

CustomProperties accepts datetime.date (vt:date)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

- ``CustomProperties`` now accepts ``datetime.date`` values (distinct
  from ``datetime.datetime``) and serialises them as ``vt:date``
  (ISO-8601 ``YYYY-MM-DD``) per ECMA-376 Part 1 §22.4.2.7 (#173).
  On read a ``vt:date`` element deserialises back to a plain
  ``datetime.date``; ``datetime.datetime`` values continue to
  round-trip as ``vt:filetime`` (ISO-8601 with trailing ``Z``).
- Surfaced by Wave 3-B: only ``python-xlsx`` previously mapped
  ``date`` to ``vt:date``; ``python-docx`` and ``python-pptx`` only
  recognised ``datetime``.

O(N^2) indexing on _Rows[i] and BlockItemContainer.paragraphs[i]
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

- ``_Rows.__getitem__`` no longer constructs a ``_Row`` proxy for
  every row in the table on each access (#170). It now reads the
  single requested ``<w:tr>`` out of ``self._tbl.tr_lst`` and wraps
  only that element, dropping a naive ``for i in range(N): rows[i]``
  loop from ~1.46 ms/access to ~0.54 ms/access at N = 2000.
- ``BlockItemContainer.paragraphs`` now returns a lightweight
  ``_ParagraphsView(Sequence[Paragraph])`` that memoises the
  underlying ``p_lst`` on first access and wraps only the ``<w:p>``
  the caller requests. The view supports ``len()``, indexed and
  sliced access, iteration, ``list(...)`` coercion, ``in``,
  ``.index(…)``, truthiness, and equality against a
  ``list[Paragraph]``.
- New ``tests/test_indexing_perf.py`` enforces a < 1 ms/access
  ceiling at N = 5000 (paragraphs) / N = 2000 (rows).

2026.05.6 — Section.vertical_alignment property
++++++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-05

- Add ``Section.vertical_alignment`` property + setter.
- Add ``WD_VERTICAL_ALIGNMENT`` enum (``TOP`` / ``CENTER`` / ``BOTH``
  / ``BOTTOM``) mapping to OOXML ``ST_VerticalJc``.
- Plumbed through ``CT_SectPr.vAlign``, following the existing
  ``Section.orientation`` pattern.
- 12 parametrised unit tests in ``tests/test_section.py``.

Surfaced by the ``docx/vertical-alignment`` parameterised family in
``loadfix/ooxml-reference-corpus`` — section-level cases previously
required ``OxmlElement("w:vAlign")`` fallback.

2026.05.5 — Document.add_comment accepts date=
++++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-04

- ``Document.add_comment()`` now forwards an optional
  ``date: datetime`` kwarg to the underlying comments collection,
  mirroring ``Comments.add_comment(date=...)``.

2026.05.4 — Word-mimicry phase 3: omit unused optional parts
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-04

python-docx now omits unused optional parts on save, matching Word's
"emit the minimum" behaviour for library-authored files. The default
template still carries these parts — they are pruned at save time
only when the document doesn't actually reference them.

- **`word/numbering.xml`** — dropped unless the document uses numbering
  directly (a paragraph with ``<w:numPr>``) or via a numbering-bearing
  style (``List Bullet``, ``List Number``, etc.). The check reads
  ``styles.xml`` to resolve style→numPr links.
- **`word/stylesWithEffects.xml`** — dropped unconditionally. This is
  a Word 2013-compat duplicate of ``styles.xml``; python-docx never
  produces effect-style content.
- **``customXml/``** items — dropped unless a content control's
  ``<w:dataBinding>`` references custom XML.
- **``docProps/thumbnail.jpeg``** — dropped unconditionally at the
  package level. python-docx has no renderer, so any thumbnail it
  ships would be stale.

Rel removal happens in the before_marshal hook (for document-rooted
parts) and at package save (for the package-level thumbnail rel),
which cascades automatically: ``[Content_Types].xml``, ``_rels/.rels``,
and ``word/_rels/document.xml.rels`` all rebuild from the pruned
rels graph without additional bookkeeping.

Concrete result on the corpus bold-text feature: the machine-generated
fixture now ships exactly the same 11 parts as the Word-authored
companion. The three-way diff's "only in machine" column is empty for
the simple-text feature pack; residual ``word/document.xml``
differences are only the locale-default page size / margins (A4 vs
US Letter, by design out of scope).

Full suite: 5004 pass / 6 skip. Corpus conformance: 5/5 pass.

2026.05.3 — Word-mimicry phase 2: paragraph-mark format mirror
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-04

- ``DocumentPart.before_marshal()`` now also mirrors a single-run
  paragraph's ``<w:rPr>`` formatting onto the paragraph mark via
  ``<w:pPr><w:rPr>``. This matches Word's "keep typing in bold"
  convention: when a paragraph ends in a bold/italic/coloured run,
  the paragraph mark inherits that formatting so text typed past the
  end continues in the same shape.
- Mirrored properties: b, bCs, i, iCs, u, strike, dstrike, caps,
  smallCaps, color, sz, szCs, rFonts, vertAlign. Explicitly excludes
  lang, spacing, border, shading — Word does not mirror these onto
  paragraph marks.
- Only applied to paragraphs that have exactly one direct ``<w:r>``
  child (the common one-run-per-paragraph case). Multi-run and
  hyperlinked paragraphs are left alone to avoid surprising behaviour.
- Existing ``<w:pPr><w:rPr>`` content is preserved; only missing
  mirror properties are added.

2026.05.2 — Word-mimicry phase 1: namespace decls, paraId, rsid
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-04

Narrow the XML python-docx emits toward the shape Microsoft Word itself
writes, so loadfix/ooxml-reference-corpus three-way diffs surface real
semantic differences instead of tooling-version noise.

- The default ``word/document.xml`` template now carries the full
  namespace set Word 2024 declares (cx, cx1-cx8, aink, am3d, oel, w15,
  w16, w16cid, w16cex, w16se, w16du, w16sdtdh, w16sdtfl) plus the
  matching ``mc:Ignorable`` list.
- New ``DocumentPart.before_marshal()`` hook stamps Word-style
  identifiers on every paragraph that lacks them just before
  serialization: ``w14:paraId``, ``w14:textId``, ``w:rsidR``,
  ``w:rsidRDefault``. Runs get ``w:rsidR``. A session-wide
  ``w:rsidRoot`` is generated per save call and recorded in
  ``word/settings.xml``'s ``<w:rsids>`` table via the new
  ``Settings.add_rsids()`` method.
- Existing identifiers are preserved on round-trip; only missing ones
  are minted.
- Reproducible-save mode (``Document.save(..., reproducible=True)``)
  derives identifiers deterministically from paragraph content, so
  repeated saves of the same document remain byte-identical.

Why: diffing python-docx output against Word-authored reference files
previously showed hundreds of lines of rsid/paraId/namespace churn
that obscured real bold/italic/layout differences. Post-fix, the noise
collapses and only behavioural divergences remain visible.

2026.05.1 — bCs/iCs correctness fix
+++++++++++++++++++++++++++++++++++

Released: 2026-05-04

- Fix: setting ``font.bold = True`` now also emits ``<w:bCs/>``
  (complex-script bold); setting ``font.italic = True`` emits
  ``<w:iCs/>``. Previously only ``<w:b/>`` / ``<w:i/>`` were emitted,
  which silently dropped bold/italic on Arabic, Hebrew, and Thai runs
  when Word reopened the file. Mirrors the behavior Word itself writes.
  Surfaced by the three-way comparison pipeline in
  ``loadfix/ooxml-reference-corpus/features/docx/bold-text.json``.

  The ``cs_bold`` / ``cs_italic`` properties continue to work
  independently; callers that need divergent values can still set them
  explicitly after setting bold/italic.

2026.05.0 — first release as independent fork
+++++++++++++++++++++++++++++++++++++++++++++

Released: 2026-05-02

This release marks the project's split from upstream
``python-openxml/python-docx``. Versioning switches to CalVer
(YYYY.MM.patch) from this point forward. The previous upstream line
stops at ``1.2.0`` (2025-06-16); everything below is new to this fork.

All 100+ features below shipped as part of this initial independent
release. Subsequent CalVer releases will have their own entries.

Phase A — Footnotes and endnotes
  - Add Document.footnotes and Footnotes / Footnote / FootnoteProperties (#1, #3, #17, #46, #48, #56, #82)
  - Add Document.endnotes mirror API (#17, #96)
  - Add Section.footnote_properties / endnote_properties (#17)

Phase B — Tracked changes
  - Add read of tracked insertions and deletions (#53)
  - Add accept / reject tracked changes (#7)
  - Add read of formatting changes (#8)
  - Add move revisions (w:moveFrom / w:moveTo) (#134)
  - Add cell and row-level tracked changes (#135)
  - Add revision_marks_text() for CLI previews (#163)

Phase C — Bookmarks and fields
  - Add bookmarks create / read / delete (#52)
  - Add simple and complex field codes (#10)
  - Add REF / PAGEREF cross-reference resolution (#115)

Phase D — Miscellaneous OOXML feature coverage
  - D.1  Hyperlink creation API (#97)
  - D.2  Comment replies (threaded) (#67)
  - D.3  Extended document settings + DocumentProtection (#66, #125)
  - D.4  Custom document properties (#14)
  - D.6  Cell shading and background color (#63)
  - D.7  Paragraph borders (#109)
  - D.9  Numbering style control (#22)
  - D.10 Search and replace with formatting preservation (#91)
  - D.13 Insert paragraph / table at arbitrary position (#26)
  - D.14 Content controls (SDTs) (#27)
  - D.15 Row.height setter (#28)
  - D.16 Row.allow_break_across_pages (#51)
  - D.17 Floating images with wp:anchor positioning (#30)
  - D.19 Multi-column section layout (#60)
  - D.20 Font.shading — run-level background color (#33)
  - D.22 SVG image support (#76)
  - D.23 Watermark support (text and image) (#36)
  - D.24 .docm macro-enabled file support (#65)
  - D.26 Table autofit and column-width control (#39)
  - D.27 DrawingML shapes and text-box content access (#75)

Other feature additions
  - Charts read + add_chart() (#111)
  - SmartArt detection and node text (#112)
  - Equation read + minimal create API (#113)
  - Add Run.add_symbol and Run.symbols (#114)
  - Add Section.page_borders (#121)
  - Add Section.line_numbering (#122)
  - Add Section.document_grid (#147)
  - Add Section.first_page / other_pages_paper_source (#146)
  - Add Section.text_direction / right_to_left (#148)
  - Add Section odd/even page header-footer (#149)
  - Add Font.border_* properties (#120)
  - Add Font.language / east_asian_language / bidi_language (#160)
  - Add East Asian typography (kinsoku, word_wrap, east_asian_layout) (#128)
  - Add RTL / bidi on Paragraph and Run (#127)
  - Add paragraph_format.frame for text frames (#126)
  - Add ParagraphBorders / Border (#109)
  - Add read-only ruby (#129)
  - Add read-only ink (#139)
  - Add read-only embedded OLE objects (#140)
  - Add read-only grouped shapes (#138)
  - Add read-only SmartArt (#112)
  - Add read-only Document.glossary (#132, #133)
  - Add read-only Document.theme (#117)
  - Add read-only Document.web_settings (#157)
  - Add Document.font_table (#119)
  - Add Document.background_color (#118)
  - Add Document.statistics (#161)
  - Add Document.search_regex / replace_regex / search_all / replace_all (#153, #154)
  - Add Document.add_table_of_contents (#116)
  - Add caption helpers (#141)
  - Add permission ranges (#124)
  - Add Settings.mail_merge (#130)
  - Add Settings.compat_flags / compat_settings (#156)
  - Add Settings.view (#164)
  - Add Style.link_style / next_style / is_redefined (#162)
  - Add Table.borders / _Cell.borders (#102)
  - Add Cell.margins (#143)
  - Add Table.style_flags (#144)
  - Add Cell.text_direction (#142)
  - Add Cell.is_merge_origin / merge_origin (#145)
  - Add _Row.is_header (#93)
  - Add Run.split (#94)
  - Add Paragraph.delete / Run.delete / Table.delete (#50)
  - Add alt_text / title on InlineShape and FloatingImage (#158)
  - Add stable_id on Paragraph / Run / Table / Cell (#155)
  - Add Paragraph.insert_paragraph_before arbitrary positioning (#26)
  - Add legacy form fields (#123)
  - Add heading-structure accessibility validator (#159)

Reliability / safety
  - Add recover=True mode for malformed .docx (#151)
  - Add EncryptedDocumentError for password-protected .docx (#152)
  - Add digital signature detection (#150)

Dev / tooling
  - Add py.typed, improve public types
  - Add AI-agent CI pipeline (Product / Develop / Review / Security / Revise
    / Merge / Debug / Watchdog)
  - Add interop-validate behave scenarios wiring loadfix/ooxml-validate as a round-trip fidelity check.

1.2.0 (2025-06-16)
++++++++++++++++++

- Add support for comments
- Drop support for Python 3.8, add testing for Python 3.13

1.1.2 (2024-05-01)
++++++++++++++++++

- Fix #1383 Revert lxml<=4.9.2 pin that breaks Python 3.12 install
- Fix #1385 Support use of Part._rels by python-docx-template
- Add support and testing for Python 3.12

1.1.1 (2024-04-29)
++++++++++++++++++

- Fix #531, #1146 Index error on table with misaligned borders
- Fix #1335 Tolerate invalid float value in bottom-margin
- Fix #1337 Do not require typing-extensions at runtime

1.1.0 (2023-11-03)
++++++++++++++++++

- Add BlockItemContainer.iter_inner_content()

1.0.1 (2023-10-12)
++++++++++++++++++

- Fix #1256: parse_xml() and OxmlElement moved.
- Add Hyperlink.fragment and .url

1.0.0 (2023-10-01)
+++++++++++++++++++

- Remove Python 2 support. Supported versions are 3.7+
- Fix #85:   Paragraph.text includes hyperlink text
- Add #1113: Hyperlink.address
- Add Hyperlink.contains_page_break
- Add Hyperlink.runs
- Add Hyperlink.text
- Add Paragraph.contains_page_break
- Add Paragraph.hyperlinks
- Add Paragraph.iter_inner_content()
- Add Paragraph.rendered_page_breaks
- Add RenderedPageBreak.following_paragraph_fragment
- Add RenderedPageBreak.preceding_paragraph_fragment
- Add Run.contains_page_break
- Add Run.iter_inner_content()
- Add Section.iter_inner_content()

0.8.11 (2021-05-15)
+++++++++++++++++++

- Small build changes and Python 3.8 version changes like collections.abc location.

0.8.10 (2019-01-08)
+++++++++++++++++++

- Revert use of expanded package directory for default.docx to work around setup.py
  problem with filenames containing square brackets.

0.8.9 (2019-01-08)
++++++++++++++++++

- Fix gap in MANIFEST.in that excluded default document template directory

0.8.8 (2019-01-07)
++++++++++++++++++

- Add support for headers and footers

0.8.7 (2018-08-18)
++++++++++++++++++

- Add _Row.height_rule
- Add _Row.height
- Add _Cell.vertical_alignment
- Fix #455: increment next_id, don't fill gaps
- Add #375: import docx failure on --OO optimization
- Add #254: remove default zoom percentage
- Add #266: miscellaneous documentation fixes
- Add #175: refine MANIFEST.ini
- Add #168: Unicode error on core-props in Python 2

0.8.6 (2016-06-22)
++++++++++++++++++

- Add #257: add Font.highlight_color
- Add #261: add ParagraphFormat.tab_stops
- Add #303: disallow XML entity expansion

0.8.5 (2015-02-21)
++++++++++++++++++

- Fix #149: KeyError on Document.add_table()
- Fix #78: feature: add_table() sets cell widths
- Add #106: feature: Table.direction (i.e. right-to-left)
- Add #102: feature: add CT_Row.trPr

0.8.4 (2015-02-20)
++++++++++++++++++

- Fix #151: tests won't run on PyPI distribution
- Fix #124: default to inches on no TIFF resolution unit

0.8.3 (2015-02-19)
++++++++++++++++++

- Add #121, #135, #139: feature: Font.color

0.8.2 (2015-02-16)
++++++++++++++++++

- Fix #94: picture prints at wrong size when scaled
- Extract `docx.document.Document` object from `DocumentPart`

  Refactor `docx.Document` from an object into a factory function for new
  `docx.document.Document object`. Extract methods from prior `docx.Document`
  and `docx.parts.document.DocumentPart` to form the new API class and retire
  `docx.Document` class.

- Migrate `Document.numbering_part` to `DocumentPart.numbering_part`. The
  `numbering_part` property is not part of the published API and is an
  interim internal feature to be replaced in a future release, perhaps with
  something like `Document.numbering_definitions`. In the meantime, it can
  now be accessed using ``Document.part.numbering_part``.

0.8.1 (2015-02-10)
++++++++++++++++++

- Fix #140: Warning triggered on Document.add_heading/table()

0.8.0 (2015-02-08)
++++++++++++++++++

- Add styles. Provides general capability to access and manipulate paragraph,
  character, and table styles.

- Add ParagraphFormat object, accessible on Paragraph.paragraph_format, and
  providing the following paragraph formatting properties:

  + paragraph alignment (justfification)
  + space before and after paragraph
  + line spacing
  + indentation
  + keep together, keep with next, page break before, and widow control

- Add Font object, accessible on Run.font, providing character-level
  formatting including:

  + typeface (e.g. 'Arial')
  + point size
  + underline
  + italic
  + bold
  + superscript and subscript

The following issues were retired:

- Add feature #56: superscript/subscript
- Add feature #67: lookup style by UI name
- Add feature #98: Paragraph indentation
- Add feature #120: Document.styles

**Backward incompatibilities**

Paragraph.style now returns a Style object. Previously it returned the style
name as a string. The name can now be retrieved using the Style.name
property, for example, `paragraph.style.name`.

0.7.6 (2014-12-14)
++++++++++++++++++

- Add feature #69: Table.alignment
- Add feature #29: Document.core_properties

0.7.5 (2014-11-29)
++++++++++++++++++

- Add feature #65: _Cell.merge()

0.7.4 (2014-07-18)
++++++++++++++++++

- Add feature #45: _Cell.add_table()
- Add feature #76: _Cell.add_paragraph()
- Add _Cell.tables property (read-only)

0.7.3 (2014-07-14)
++++++++++++++++++

- Add Table.autofit
- Add feature #46: _Cell.width

0.7.2 (2014-07-13)
++++++++++++++++++

- Fix: Word does not interpret <w:cr/> as line feed

0.7.1 (2014-07-11)
++++++++++++++++++

- Add feature #14: Run.add_picture()

0.7.0 (2014-06-27)
++++++++++++++++++

- Add feature #68: Paragraph.insert_paragraph_before()
- Add feature #51: Paragraph.alignment (read/write)
- Add feature #61: Paragraph.text setter
- Add feature #58: Run.add_tab()
- Add feature #70: Run.clear()
- Add feature #60: Run.text setter
- Add feature #39: Run.text and Paragraph.text interpret '\n' and '\t' chars

0.6.0 (2014-06-22)
++++++++++++++++++

- Add feature #15: section page size
- Add feature #66: add section
- Add page margins and page orientation properties on Section
- Major refactoring of oxml layer

0.5.3 (2014-05-10)
++++++++++++++++++

- Add feature #19: Run.underline property

0.5.2 (2014-05-06)
++++++++++++++++++

- Add feature #17: character style

0.5.1 (2014-04-02)
++++++++++++++++++

- Fix issue #23, `Document.add_picture()` raises ValueError when document
  contains VML drawing.

0.5.0 (2014-03-02)
++++++++++++++++++

- Add 20 tri-state properties on Run, including all-caps, double-strike,
  hidden, shadow, small-caps, and 15 others.

0.4.0 (2014-03-01)
++++++++++++++++++

- Advance from alpha to beta status.
- Add pure-python image header parsing; drop Pillow dependency

0.3.0a5 (2014-01-10)
++++++++++++++++++++++

- Hotfix: issue #4, Document.add_picture() fails on second and subsequent
  images.

0.3.0a4 (2014-01-07)
++++++++++++++++++++++

- Complete Python 3 support, tested on Python 3.3

0.3.0a3 (2014-01-06)
++++++++++++++++++++++

- Fix setup.py error on some Windows installs

0.3.0a1 (2014-01-05)
++++++++++++++++++++++

- Full object-oriented rewrite
- Feature-parity with prior version
- text: add paragraph, run, text, bold, italic
- table: add table, add row, add column
- styles: specify style for paragraph, table
- picture: add inline picture, auto-scaling
- breaks: add page break
- tests: full pytest and behave-based 2-layer test suite

0.3.0dev1 (2013-12-14)
++++++++++++++++++++++

- Round-trip .docx file, preserving all parts and relationships
- Load default "template" .docx on open with no filename
- Open from stream and save to stream (file-like object)
- Add paragraph at and of document
