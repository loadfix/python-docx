"""``docx.kit`` — opinionated authoring helpers built on top of python-docx.

The :mod:`docx.kit` submodule is a curated collection of high-level
"recipe" helpers that compose the existing python-docx public API into
common document-authoring patterns. Each helper is a thin wrapper that
adds a styled, shaped chunk of content to a |Document| in one call —
the kind of boilerplate every report / book / proposal author writes
by hand the first time and copy-pastes thereafter. The kit is **opt-in**
via the ``[kit]`` extras flag (``pip install python-docx[kit]``) so
callers who use only the core authoring API pay nothing for it.

The kit lives **inside** :mod:`docx` (not as a sibling package) per the
project's Wave-4 scoping memo: kit content is Python helpers tightly
coupled to a single parent's public API with no cross-parent reuse
story, so an in-parent submodule under a ``[kit]`` extras flag is the
right home (the ``[kit]`` extras list is currently empty — kit modules
are pure-Python compositions of existing python-docx surface and add no
new runtime dependencies; the flag exists as a versioning hook so
future kit issues can declare deps without touching the core dependency
list).

The first kit submodule, :mod:`docx.kit.front_matter`, exposes seven
helpers for the conventional front-matter sections of a long-form
document:

* :func:`~docx.kit.front_matter.add_title_page` — title / subtitle /
  author / date, each in its own styled paragraph, followed by a page
  break.
* :func:`~docx.kit.front_matter.add_copyright_page` — copyright holder
  + year + edition + optional rights notice.
* :func:`~docx.kit.front_matter.add_dedication` — centred italic
  dedication paragraph.
* :func:`~docx.kit.front_matter.add_preface` — heading + body
  paragraphs; accepts ``body`` as a string (split on blank lines) or a
  list of paragraph strings.
* :func:`~docx.kit.front_matter.add_table_of_contents` — wraps
  :meth:`docx.document.Document.add_table_of_contents` with an optional
  preceding heading.
* :func:`~docx.kit.front_matter.add_list_of_figures` — TOC field
  filtered to ``Figure`` SEQ entries.
* :func:`~docx.kit.front_matter.add_list_of_tables` — TOC field
  filtered to ``Table`` SEQ entries.

Each helper appends a **section** of paragraphs at the end of the
document body and returns the list of paragraphs it created so the
caller can post-process them (e.g. tweak alignment, attach bookmarks).
"""

from __future__ import annotations

from docx.kit import front_matter

__all__ = ["front_matter"]
