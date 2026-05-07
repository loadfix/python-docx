"""Re-export of legacy WordprocessingML comment element classes from :mod:`ooxml_comments`.

Historically ``docx.oxml.comments`` defined the ``CT_Comments`` (``<w:comments>``)
and ``CT_Comment`` (``<w:comment>``) element classes inline. As of 2026.05 the
canonical implementations live in the shared :mod:`ooxml_comments.oxml`
package; this module keeps the historical import paths working for
downstream consumers.

.. versionchanged:: 2026.05.0
    Implementation relocated to ``python-ooxml-comments``. The behaviour is
    byte-for-byte preserved — same descriptors, same id-allocation rules,
    same ``w16cid:paraId`` generation.
"""

from __future__ import annotations

from ooxml_comments.oxml.comments import CT_Comment, CT_Comments

__all__ = [
    "CT_Comment",
    "CT_Comments",
]
