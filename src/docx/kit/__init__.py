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

Available kit submodules:

* :mod:`docx.kit.front_matter` — title / copyright / dedication /
  preface / TOC / list-of-figures / list-of-tables helpers.
* :mod:`docx.kit.chapter` — chapter opener pages (large title +
  decorative image + drop cap).

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

from docx.kit import chapter, front_matter

__all__ = ["chapter", "front_matter"]
