.. _search_replace:

Searching and Replacing Text
============================

Word's *Find* and *Replace* dialogs let a user locate text in a document and, optionally,
swap it for something else. *python-docx* exposes a similar capability programmatically
through a small set of methods on the |Document| object and a helper class called
|SearchMatch|.

The feature is designed around three axes:

- **Plain text vs regular expression.** ``search`` / ``replace`` match a literal
  substring; ``search_regex`` / ``replace_regex`` accept a Python regular expression
  (either a string or a pre-compiled :class:`re.Pattern`).
- **Body-only vs every story.** The bare names (``search``, ``replace``,
  ``search_regex``, ``replace_regex``) look only at top-level paragraphs of the
  document body. The ``*_all`` variants (``search_all``, ``replace_all``,
  ``search_regex_all``, ``replace_regex_all``) additionally walk body tables, each
  section's non-inherited headers and footers, footnotes, endnotes, and comments.
- **Query vs mutate.** The ``search*`` methods are pure queries — they return a list
  of |SearchMatch| objects. The ``replace*`` methods return the number of replacements
  performed and mutate the document in place.


Plain text search
-----------------

Call :meth:`.Document.search` with a literal substring to scan the body paragraphs::

    >>> from docx import Document
    >>> document = Document("example.docx")
    >>> matches = document.search("Invoice")
    >>> len(matches)
    3
    >>> matches[0]
    <docx.search.SearchMatch object at 0x...>

By default, matching is case-sensitive and unanchored. The two optional flags
follow Word's Find dialog conventions::

    >>> document.search("invoice", case_sensitive=False)   # also matches "Invoice"
    >>> document.search("Total", whole_word=True)          # skips "SubTotal" etc.

Passing an empty string returns an empty list rather than raising.


Regular-expression search
-------------------------

Use :meth:`.Document.search_regex` when a literal substring is not expressive enough.
The ``pattern`` argument may be a string or an already-compiled :class:`re.Pattern`::

    >>> import re
    >>> document.search_regex(r"INV-\d+")
    [<...SearchMatch...>, <...SearchMatch...>]
    >>> document.search_regex(re.compile(r"\bID: [A-F0-9]+\b", re.IGNORECASE))

When ``pattern`` is a string, any ``flags`` you pass are applied at compile time. When
``pattern`` is already compiled, ``flags`` is silently ignored — the existing
compiled flags win, matching the behaviour of :func:`re.search`.

Zero-width matches (e.g. ``r"^"`` or lookarounds) are reported by ``search_regex``,
but they are *skipped* by ``replace_regex`` because there is no obvious run to host
an empty replacement.


The SearchMatch object
----------------------

Every hit is returned as a |SearchMatch| carrying:

- :attr:`~.SearchMatch.paragraph` — the |Paragraph| that contains the match.
- :attr:`~.SearchMatch.paragraph_index` — the paragraph's index within its *story*.
  For body-only searches, this is the index into ``document.paragraphs``. For
  cross-story searches, it is the index into the paragraph list of the specific
  story identified by :attr:`~.SearchMatch.location`.
- :attr:`~.SearchMatch.run_indices` — a sorted list of run indices that overlap
  the match. A match that lives in a single run reports ``[n]``; one that spans
  several has ``[n, n+1, ...]``.
- :attr:`~.SearchMatch.start` / :attr:`~.SearchMatch.end` — character offsets
  into ``paragraph.text`` (the reconstructed plain-text form), using Python
  half-open interval semantics. ``paragraph.text[match.start : match.end]``
  reproduces the matched text.
- :attr:`~.SearchMatch.location` — story identifier, populated by the ``_all``
  helpers and |None| for body-only searches. See below.


Matches that span several runs
------------------------------

Word stores runs of text with uniform formatting. A single paragraph that reads
"the **quick** brown fox" is three runs: ``"the "``, ``"quick"``, and
``" brown fox"``. A search term like ``"e qui"`` therefore spans two runs.

*python-docx* handles this transparently:

- During *search*, the match is reported once with
  :attr:`~.SearchMatch.run_indices` listing every run it crosses.
- During *replace*, the replacement text is written into the **first** run of the
  span (inheriting that run's formatting), and any matched characters in
  subsequent runs are trimmed away. Fully-consumed middle runs are left in
  place as empty runs so their formatting still exists for Word if needed.

That last point is important for preserving bold/italic/color applied inside the
match: whichever formatting the *first* matched character had is what the
replacement text will inherit.


Replacing text
--------------

:meth:`.Document.replace` mutates the document in place and returns the number of
replacements made::

    >>> n = document.replace("SpamCo", "EggCorp")
    >>> n
    4
    >>> document.save("example.docx")

The same ``case_sensitive`` and ``whole_word`` flags available on
:meth:`~.Document.search` are honoured here. Passing ``old_text=""`` returns
``0`` without touching the document.

:meth:`.Document.replace_regex` follows :func:`re.sub` semantics for the
``replacement`` argument: backreferences such as ``\1`` or ``\g<name>`` are
expanded per match::

    >>> document.replace_regex(r"INV-(\d+)", r"Invoice #\1")
    2


Searching every story with ``*_all``
------------------------------------

:meth:`~.Document.search` and :meth:`~.Document.replace` only consider the top-level
body paragraphs. Use the ``*_all`` variants to reach content that Word treats as
separate streams:

- the document body, tagged ``"body"``;
- paragraphs inside body-level tables, tagged
  ``"table:<t>:row:<r>:col:<c>"``;
- each section's *primary*, *even-page*, and *first-page* headers and footers
  (unless the section inherits the previous one's, in which case the inherited
  definition is visited only once), tagged
  ``"header:section<i>:primary"``, ``"footer:section<i>:even_page"``, etc.;
- footnote paragraphs, tagged ``"footnote:<id>"``;
- endnote paragraphs, tagged ``"endnote:<id>"``;
- comment paragraphs, tagged ``"comment:<id>"``.

::

    >>> matches = document.search_all("Confidential")
    >>> {m.location for m in matches}
    {'body', 'header:section0:primary', 'footnote:2'}

Tables nested inside *other* stories (a table inside a header, or a table inside
a body-table cell) are not recursively descended; the top-level cell text is
still searchable but doubly-nested tables are skipped. This matches the
invariant documented for :func:`docx.search._iter_all_paragraphs`.

The regex and replace variants behave identically with respect to story
coverage::

    >>> document.search_regex_all(r"\bTODO\b")
    >>> document.replace_all("v1.0", "v2.0")
    >>> document.replace_regex_all(r"\bDRAFT\b", r"FINAL")

The replace variants return the *total* number of replacements performed across
every story.


Working directly on a paragraph list
------------------------------------

The four :mod:`docx.search` module-level functions —
:func:`~docx.search.search_paragraphs`,
:func:`~docx.search.search_paragraphs_regex`,
:func:`~docx.search.replace_in_paragraphs`, and
:func:`~docx.search.replace_in_paragraphs_regex` — take a ``list[Paragraph]``
directly. They are handy when you already have a subset of paragraphs (for
example, the paragraphs inside one table cell) and only want to operate on
those::

    >>> from docx.search import search_paragraphs
    >>> cell = document.tables[0].cell(1, 1)
    >>> search_paragraphs(list(cell.paragraphs), "TBD")
    [<...SearchMatch...>]
