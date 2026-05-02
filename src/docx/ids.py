"""Stable identifier helpers for document elements.

OOXML lacks a per-element stable UUID. This module provides a pragmatic
"mostly-stable" identifier for paragraphs, runs, and tables by combining a
revision-save ID (``w:rsidR``) when present with a hash of the element's
structural position and immediate text content.

The resulting ID is a 16-character lowercase hexadecimal string. It is
recomputed on every access; it is deterministic with respect to its inputs
but not persisted on the XML itself.

Stability semantics
-------------------
- IDs are stable across save/load when the element retains the same position
  within its parent and the same immediate text.
- IDs change if the element is reordered (moved to a different index within
  its parent), if its text is edited, or if its parent's tag changes.
- When the element carries a ``w:rsidR`` attribute, that attribute is folded
  into the hash so two structurally identical elements created in different
  editing sessions still receive distinct IDs.
- These IDs are not cryptographic fingerprints and should not be used for
  authentication; their purpose is to let tools correlate elements across a
  save/reload cycle in a single editing session.
"""

from __future__ import annotations

import hashlib
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.oxml.xmlchemy import BaseOxmlElement


_STABLE_ID_HEX_LEN = 16


def _position_key(element: BaseOxmlElement) -> str:
    """Return a short string describing the element's position within its parent.

    The key combines the parent's tag and the element's zero-based index among
    its siblings. Elements with no parent (detached) use ``"<root>/0"``.
    """
    parent = element.getparent()
    if parent is None:
        return "<root>/0"
    # --- lxml's Element.index() raises if element is not a child; our caller
    # --- guarantees it is, but we still compute defensively via iteration.
    try:
        idx = list(parent).index(element)
    except ValueError:
        idx = -1
    return f"{str(parent.tag)}/{idx}"


def compute_stable_id(
    element: BaseOxmlElement, text: str, rsid: str | None = None
) -> str:
    """Return a 16-character hex stable-ID for `element`.

    Combines `rsid` (when truthy), the element's tag, its sibling-position key
    (parent's tag + index within parent), and `text` into a SHA-1 digest and
    returns the first 16 hex characters of that digest.

    The function is pure — it does not modify `element` — and deterministic:
    the same inputs always yield the same output string.

    .. versionadded:: 1.3.0.dev0
    """
    components: list[str] = [
        rsid or "",
        str(element.tag),
        _position_key(element),
        text,
    ]
    payload = "\x1f".join(components).encode("utf-8")
    return hashlib.sha1(payload).hexdigest()[:_STABLE_ID_HEX_LEN]
