"""Re-export of :mod:`ooxml_opc.constants`.

The ``CONTENT_TYPE`` / ``RELATIONSHIP_TYPE`` / ``NAMESPACE`` /
``RELATIONSHIP_TARGET_MODE`` registries now live in the shared
:mod:`ooxml_opc` package. Keeps the ``docx.opc.constants.*`` import
paths working for every existing caller.

A handful of docx-local extensions are attached on top of the shared
registries for content-types that live **below** the shared-package
promotion bar (e.g. Word 2013+ ``commentsExtended.xml``). These are
set via ``setattr`` so they appear as attributes on the shared
``CONTENT_TYPE`` / ``RELATIONSHIP_TYPE`` namespace classes without
forking the shared constants module.
"""

from __future__ import annotations

from ooxml_opc.constants import (
    CONTENT_TYPE,
    NAMESPACE,
    RELATIONSHIP_TARGET_MODE,
    RELATIONSHIP_TYPE,
)

# -- Word 2013+ ``commentsExtended.xml`` content-type and relationship. --
# -- These live in the Microsoft ``officeDocument/2011/`` namespace; they --
# -- aren't yet promoted to the shared ``ooxml_opc`` registry because only --
# -- Word uses them (no pptx/xlsx parity). Attach them here so callers --
# -- continue to use ``CT.WML_COMMENTS_EXTENDED`` / ``RT.COMMENTS_EXTENDED`` --
# -- uniformly even though the underlying strings are docx-owned. --
setattr(
    CONTENT_TYPE,
    "WML_COMMENTS_EXTENDED",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
)
setattr(
    RELATIONSHIP_TYPE,
    "COMMENTS_EXTENDED",
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
)

# -- Word 2016+ ``commentsIds.xml`` (``w16cid:commentsIds``) and Word 2018+
# -- ``commentsExtensible.xml`` (``w16cex:commentsExtensible``) content types
# -- and relationship types. Sourced from the shared ``ooxml_comments`` 0.4
# -- constants module so the docx ``CT.*`` / ``RT.*`` namespaces stay in
# -- sync with the canonical strings. Attached via ``setattr`` for the same
# -- reason as ``WML_COMMENTS_EXTENDED`` above — no pptx/xlsx parity yet,
# -- so not yet promoted to the shared ``ooxml_opc`` registry.
from ooxml_comments import (
    CONTENT_TYPE_WML_COMMENTS_EXTENSIBLE as _CT_WML_COMMENTS_EXTENSIBLE,
    CONTENT_TYPE_WML_COMMENTS_IDS as _CT_WML_COMMENTS_IDS,
    RELATIONSHIP_TYPE_COMMENTS_EXTENSIBLE as _RT_COMMENTS_EXTENSIBLE,
    RELATIONSHIP_TYPE_COMMENTS_IDS as _RT_COMMENTS_IDS,
)

setattr(CONTENT_TYPE, "WML_COMMENTS_IDS", _CT_WML_COMMENTS_IDS)
setattr(CONTENT_TYPE, "WML_COMMENTS_EXTENSIBLE", _CT_WML_COMMENTS_EXTENSIBLE)
setattr(RELATIONSHIP_TYPE, "COMMENTS_IDS", _RT_COMMENTS_IDS)
setattr(RELATIONSHIP_TYPE, "COMMENTS_EXTENSIBLE", _RT_COMMENTS_EXTENSIBLE)

del (
    _CT_WML_COMMENTS_IDS,
    _CT_WML_COMMENTS_EXTENSIBLE,
    _RT_COMMENTS_IDS,
    _RT_COMMENTS_EXTENSIBLE,
)

__all__ = [
    "CONTENT_TYPE",
    "NAMESPACE",
    "RELATIONSHIP_TARGET_MODE",
    "RELATIONSHIP_TYPE",
]
