"""Stable, addressable paragraph IDs via Word's ``w14:paraId`` attribute.

Closes #301.

Many real-world authoring workflows — CMS pipelines, collaborative
editors, AI-assisted draft tools, review systems — need to refer to a
*specific* paragraph across edits. The native ``Paragraph`` object has
no stable identity: its index shifts when paragraphs are inserted or
deleted ahead of it, and its in-memory id is invalidated on every
load.

Word itself solves this with the ``w14:paraId`` attribute on
``<w:p>``: an 8-hex-digit token that survives save / load round-trips
and tracks the paragraph through edits. Word's modern-comments and
threaded-replies machinery already keys off it; this module exposes
the same identifier as a public addressing API::

    from docx import Document
    from docx.kit import stable_paragraph_ids

    doc = Document("input.docx")

    stable_paragraph_ids.ensure(doc)            # idempotent stamp
    para = stable_paragraph_ids.get(doc, "A3F12B4C")
    for pid, para in stable_paragraph_ids.iter_with_ids(doc):
        print(pid, para.text[:40])
    stable_paragraph_ids.set_id(doc.paragraphs[0], "intro")
    stable_paragraph_ids.id_of(doc.paragraphs[0])
    doc.save("out.docx")

ID format
---------

Auto-minted ids are 8 uppercase hex characters — Word's own shape.
Caller-supplied ids via :func:`set_id` are validated against a
deliberately broader grammar — 1-32 chars of ``[A-Za-z0-9_]`` — so
callers may use human-readable tokens (``"intro"``, ``"section_1"``)
for cross-version stable references. The 8-hex shape is a strict
subset of that grammar so Word-authored ids always pass through.
Anything outside (whitespace, punctuation, non-ASCII) raises
:class:`ValueError`.

Design note — the kit's "compose, don't reach down" rule
---------------------------------------------------------

Every other :mod:`docx.kit` module sticks to python-docx's *public*
API. This module is the **one exception**. ``w14:paraId`` is a
Word-2010+ attribute python-docx does not expose on its public
``Paragraph`` surface — it is minted silently inside
:meth:`DocumentPart.before_marshal` at save time and is otherwise
reachable only through ``_element``. To stay useful as an addressing
primitive, this module reads and writes the attribute directly via
the private :func:`_paraId_attr` helper. The reach is contained to
this single file; the kit's overall rule remains valid elsewhere. A
future python-docx release that adds a public ``Paragraph.para_id``
property would let this module's internals collapse to a one-liner.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import re
import secrets
from typing import TYPE_CHECKING, Iterator, Optional, Tuple, Union

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.text.paragraph import Paragraph


# -- Word's own ``w14:paraId`` attribute, in lxml Clark notation. The
# -- ``w14`` prefix maps to ``http://schemas.microsoft.com/office/word/2010/wordml``
# -- and is registered in :mod:`docx.oxml.ns`.
_PARA_ID_QN = qn("w14:paraId")

# -- Length of an auto-minted id, in hex characters. Matches Word's own
# -- behaviour: every ``w14:paraId`` Word writes is exactly 8 hex chars
# -- (32 random bits). Per :class:`ST_LongHexNumber` in the XSD.
_AUTO_ID_HEX_LEN = 8

# -- Caller-supplied id grammar. ASCII letters, digits, and underscore;
# -- 1-32 characters. The lower bound rules out the empty string; the
# -- upper bound caps the length at something Word and downstream tools
# -- handle without truncation. The grammar is a deliberate superset of
# -- the strict 8-hex-char shape so callers can use human-readable
# -- tokens (``"intro"``, ``"section_1"``) when stable cross-version
# -- references matter more than schema purity.
_ID_PATTERN = re.compile(r"^[A-Za-z0-9_]{1,32}$")


def _paraId_attr(paragraph_or_p):
    # type: (Union[Paragraph, object]) -> object
    """Return the ``CT_P`` element backing ``paragraph_or_p``.

    Accepts either a public :class:`~docx.text.paragraph.Paragraph`
    proxy (the common case) or a raw ``CT_P`` element (used internally
    by :func:`ensure` while iterating ``root.iter(qn("w:p"))``).

    This is the single point of contact with the ``_element`` /
    ``oxml`` layer in this module. Every other helper goes through it
    so the kit's "compose, don't reach down" rule is contained to one
    well-named function (see the module docstring for the full design
    note).
    """
    # -- duck-type rather than isinstance-check so a raw CT_P (which is
    # -- what ``root.iter(qn("w:p"))`` yields) and a Paragraph proxy both
    # -- work. The ``_p`` attribute is the public-but-undocumented
    # -- accessor every Paragraph carries; ``_element`` is the alias
    # -- BlockItemContainer / BaseOxmlElement also expose.
    elm = getattr(paragraph_or_p, "_p", None)
    if elm is None:
        elm = getattr(paragraph_or_p, "_element", paragraph_or_p)
    return elm


def _new_paraId():
    # type: () -> str
    """Mint a fresh 8-hex-uppercase ``w14:paraId`` token.

    Matches Word's own ``paraId`` shape (32 random bits, uppercased)
    so output is indistinguishable from a Word-authored document at
    the byte level. Uses :mod:`secrets` rather than :mod:`random` so
    the token-space remains uniformly distributed even if a caller
    relies on the tokens for non-cryptographic addressing (no API
    contract, but it costs nothing to use the better RNG).
    """
    return secrets.token_hex(_AUTO_ID_HEX_LEN // 2).upper()


def _validate_id(id_str):
    # type: (str) -> None
    """Raise :class:`ValueError` when ``id_str`` is not a valid paraId.

    Validates the caller-supplied id grammar described in the module
    docstring: 1-32 ASCII letters, digits, or underscores. Word's own
    8-hex-char tokens are a strict subset and always pass.
    """
    if not isinstance(id_str, str):
        raise ValueError(
            "paraId must be a str; got %s" % type(id_str).__name__
        )
    if not _ID_PATTERN.match(id_str):
        raise ValueError(
            "paraId %r is invalid; expected 1-32 chars of "
            "[A-Za-z0-9_] (Word writes 8-hex; this module accepts a "
            "broader human-readable grammar)" % id_str
        )


def _iter_p_elements(document):
    # type: (Document) -> Iterator[object]
    """Yield every ``<w:p>`` element under ``document``'s body, in order.

    Reaches into the body element and uses ``lxml``'s ``iter`` to
    descend into table cells (and nested tables) — every paragraph
    that lives inside the document body, regardless of how deeply
    nested. Headers, footers, footnotes, and comments are *not*
    walked: those are separate parts of the OPC package, addressed
    via the document's relationships, and addressing them by
    ``paraId`` requires the caller to look up each part explicitly.
    A future iteration may extend the walk; the current scope
    matches the issue's "every paragraph in the body" contract.
    """
    body = document.element.body
    for p in body.iter(qn("w:p")):
        yield p


def ensure(document):
    # type: (Document) -> int
    """Stamp a stable ``w14:paraId`` on every body paragraph that lacks one.

    Walks every ``<w:p>`` under the document body — including those
    inside table cells and nested tables — and assigns a fresh
    8-hex-uppercase id to any paragraph that does not already carry
    one. Paragraphs that already have an id are left untouched, so the
    function is **idempotent**: a second call is a no-op (assuming no
    paragraphs were appended in between).

    Parameters
    ----------
    document
        The :class:`Document` to mutate.

    Returns
    -------
    int
        The number of paragraphs newly stamped (zero on a fully
        already-stamped document).
    """
    stamped = 0
    for p in _iter_p_elements(document):
        if not p.get(_PARA_ID_QN):
            p.set(_PARA_ID_QN, _new_paraId())
            stamped += 1
    return stamped


def id_of(paragraph):
    # type: (Paragraph) -> Optional[str]
    """Return the ``w14:paraId`` of ``paragraph``, or ``None`` if absent.

    Read-only — does not stamp a new id. Pair with :func:`ensure` if
    you want every paragraph to be addressable.
    """
    elm = _paraId_attr(paragraph)
    value = elm.get(_PARA_ID_QN)
    return value if value else None


def set_id(paragraph, id_str):
    # type: (Paragraph, str) -> None
    """Set the ``w14:paraId`` on ``paragraph`` to ``id_str``.

    Validates the id format up front (see the module docstring for the
    accepted grammar) so callers get a clean error rather than a
    silently-malformed file that Word will refuse to open.

    Idempotent — re-setting the same id is a no-op.

    Parameters
    ----------
    paragraph
        The :class:`~docx.text.paragraph.Paragraph` to stamp.
    id_str
        The id to set. Must be 1-32 characters of ASCII letters,
        digits, or underscores.

    Raises
    ------
    ValueError
        If ``id_str`` does not match the accepted grammar.
    """
    _validate_id(id_str)
    elm = _paraId_attr(paragraph)
    elm.set(_PARA_ID_QN, id_str)


def get(document, id_str):
    # type: (Document, str) -> Optional[Paragraph]
    """Return the body paragraph whose ``w14:paraId`` equals ``id_str``.

    Returns ``None`` when no paragraph in the body carries that id.
    The lookup is **exact** — case-sensitive, no whitespace
    normalisation. When two paragraphs share an id (which Word's own
    invariants forbid but a hand-edited document can produce), the
    *first* paragraph in document order wins.

    Parameters
    ----------
    document
        The :class:`Document` to search.
    id_str
        The id to look up. Validated against the accepted grammar so
        a typo raises rather than silently returning ``None``.

    Returns
    -------
    Paragraph or None
        The matching paragraph, or ``None`` when no paragraph in the
        body carries that id.

    Raises
    ------
    ValueError
        If ``id_str`` does not match the accepted grammar.
    """
    _validate_id(id_str)
    # Defer the import to runtime to keep the kit module lazy and to
    # avoid a circular import via ``docx.text.paragraph`` -> ``docx`` ->
    # ``docx.kit`` (the kit is re-exported from ``docx``).
    from docx.text.paragraph import Paragraph

    body = document.element.body
    for p in body.iter(qn("w:p")):
        if p.get(_PARA_ID_QN) == id_str:
            return Paragraph(p, body)
    return None


def iter_with_ids(document):
    # type: (Document) -> Iterator[Tuple[str, Paragraph]]
    """Yield ``(id, Paragraph)`` for every body paragraph that carries an id.

    Iterates in document order. Paragraphs without an id are skipped,
    so call :func:`ensure` first if you want to see every paragraph.

    Parameters
    ----------
    document
        The :class:`Document` to walk.

    Yields
    ------
    tuple of (str, Paragraph)
        The id and the wrapped :class:`~docx.text.paragraph.Paragraph`.
    """
    from docx.text.paragraph import Paragraph

    body = document.element.body
    for p in body.iter(qn("w:p")):
        pid = p.get(_PARA_ID_QN)
        if pid:
            yield pid, Paragraph(p, body)


__all__ = [
    "ensure",
    "get",
    "id_of",
    "iter_with_ids",
    "set_id",
]
