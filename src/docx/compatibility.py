# pyright: reportPrivateUsage=false

"""Compatibility-mode helpers for :meth:`docx.document.Document.save`.

Word recognises a ``<w:compat>`` element in ``settings.xml`` carrying a
``<w:compatSetting w:name="compatibilityMode" w:val="N"/>`` child where
``N`` is the major Word version that authored the document
(``11`` = Word 2003, ``12`` = Word 2007, ``14`` = Word 2010, ``15`` =
Word 2013, ``16`` = Word 2016+). When the value is set to an older
version, modern Word versions enter "compatibility mode" and disable
features the older release could not author.

This module exposes:

- :data:`COMPATIBILITY_LEVELS` — the canonical name → ``compatibilityMode``
  integer mapping.
- :func:`resolve_compatibility` — normalise a caller-supplied
  ``compatibility=`` argument (string label, ``int``, or |None|) to the
  matching integer.
- :func:`apply_compatibility` — stamp the ``compatibilityMode`` setting
  on a |Document| and best-effort filter known-incompatible features
  for the chosen target.

Saving with ``compatibility=`` is **not** a guarantee Word will render
the resulting document pixel-perfectly. The flag tells Word to *open*
the file as if it had been authored under the older version and to hide
the modern UI affordances; the actual content remaining in the package
is what determines the visual outcome. This module strips the
machine-readable parts that older Word releases would error on (the
modern threaded-comments parts when targeting Word 2003 / 2007), but
it does not, and cannot, perform a full feature audit. Documents that
rely on Word 2010+ features (SmartArt, content controls, OMML
equations, …) may render with placeholders or empty boxes when opened
in the older client.

.. versionadded:: 2026.05.dev0
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Mapping, Union

if TYPE_CHECKING:  # pragma: no cover -- type-checking only
    from docx.document import Document


# -- Canonical name → compatibilityMode integer map. ``int`` keys are
# -- silently accepted as well so callers can pass ``compatibility=14``
# -- when they know the raw value. The list reflects the values
# -- Microsoft documents in [MS-DOCX] / Open XML SDK and matches what
# -- Word 365 / 2024 emits when a user picks "Word 2003" / "Word 2007"
# -- / "Word 2010" / "Word 2013" / "Word 2016" from the
# -- File → Save-As → Compatibility dropdown.
COMPATIBILITY_LEVELS: Mapping[str, int] = {
    "Word 2003": 11,
    "Word 2007": 12,
    "Word 2010": 14,
    "Word 2013": 15,
    "Word 2016": 16,
}


# -- The set of valid integer compatibilityMode values. Mirrors the
# -- mapping above; kept as a separate frozen set so :func:`resolve_compatibility`
# -- can validate raw integers without iterating the dict each call.
_VALID_LEVELS = frozenset(COMPATIBILITY_LEVELS.values())


def resolve_compatibility(value: Union[str, int, None]) -> Union[int, None]:
    """Return the ``compatibilityMode`` integer for `value` (or |None|).

    `value` may be:

    * |None| — the no-op case; returns |None| unchanged.
    * A label from :data:`COMPATIBILITY_LEVELS` (e.g. ``"Word 2003"``).
      Comparison is exact; case and whitespace are not normalised.
    * An ``int`` already in the canonical set (``11``, ``12``, ``14``,
      ``15``, ``16``).

    Raises :class:`ValueError` for anything else, including booleans
    (which are technically ``int`` in Python but never a valid mode).

    .. versionadded:: 2026.05.dev0
    """
    if value is None:
        return None
    if isinstance(value, bool):
        # -- Reject ``True``/``False`` even though they are nominally
        # -- ``int``. Otherwise ``compatibility=True`` would resolve to
        # -- 1 and silently produce an invalid document.
        raise ValueError(
            "compatibility= must be a label like 'Word 2003' or an int "
            "in (11, 12, 14, 15, 16); got %r" % value
        )
    if isinstance(value, int):
        if value not in _VALID_LEVELS:
            raise ValueError(
                "compatibility= must be one of %s; got %r"
                % (sorted(_VALID_LEVELS), value)
            )
        return value
    if isinstance(value, str):
        try:
            return COMPATIBILITY_LEVELS[value]
        except KeyError:
            raise ValueError(
                "compatibility= must be one of %s or a valid int; got %r"
                % (sorted(COMPATIBILITY_LEVELS), value)
            )
    raise TypeError(
        "compatibility= must be a str label, int, or None; got %r"
        % type(value).__name__
    )


def apply_compatibility(document: "Document", level: int) -> None:
    """Stamp ``compatibilityMode=level`` and filter incompatible features.

    Writes the ``<w:compat>/<w:compatSetting w:name="compatibilityMode"
    w:val="<level>"/>`` setting on the document's settings part and,
    for older targets, drops parts and elements the older Word client
    cannot render:

    * **Word 2003 (11)** — strips the modern threaded-comments parts
      (``commentsIds.xml`` / ``commentsExtensible.xml``) and the
      Word 2013+ extended-comments part (``commentsExtended.xml``).
      Word 2003 only understood the legacy ``comments.xml`` ECMA
      schema; the modern threaded extensions trip its parser.
    * **Word 2007 (12)** — same threaded-comment strip as Word 2003.
      Word 2007 introduced the legacy ``comments.xml`` round-trip
      semantics but predates ``w15:`` / ``w16:`` extensions.
    * **Word 2010 / 2013 / 2016 (14 / 15 / 16)** — only the
      ``compatibilityMode`` setting is written. The modern threaded
      comments format dates from Word 2016, but Word 2010/2013 read it
      gracefully via the markup-compatibility ``mc:Ignorable``
      machinery, so we leave the parts in place.

    Filtering is best-effort by contract — a feature not yet on the
    drop list above (e.g. SmartArt, ``w:sdt`` content controls, OMML
    equations) is *kept* in the package. Word 2003 / 2007 will render
    it as a placeholder rather than fail to open the file.

    .. versionadded:: 2026.05.dev0
    """
    if level not in _VALID_LEVELS:
        raise ValueError(
            "compatibility level must be one of %s; got %r"
            % (sorted(_VALID_LEVELS), level)
        )

    # -- Always write the compatibilityMode setting first; that is the
    # -- single bit Word actually consults to enable compat-mode UI. --
    settings = document.settings
    settings.compatibility_mode = level

    # -- Strip threaded comments when targeting Word 2003 or Word 2007. --
    if level <= 12:
        _strip_threaded_comments(document)


def _strip_threaded_comments(document: "Document") -> None:
    """Drop the modern threaded-comments parts from `document`.

    The legacy ``word/comments.xml`` part (ECMA schema) is preserved
    so existing comments still round-trip through Word 2003 / 2007;
    only the ``commentsIds.xml`` (Word 2016 ``w16cid:``),
    ``commentsExtensible.xml`` (Word 2018 ``w16cex:``), and
    ``commentsExtended.xml`` (Word 2013 threaded-replies) parts are
    removed because they carry namespaces the older client does not
    declare in its ``mc:Ignorable`` allow-list and so trip its parser.

    The corresponding relationships hanging off the comments part are
    dropped via :meth:`docx.opc.part.Part.drop_rel` so the package
    serialises cleanly with no dangling rels.
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    document_part = document._part

    # -- Locate the comments part if and only if one is already
    # -- related; we do NOT want to materialise a fresh comments part
    # -- just to strip its modern children.
    try:
        comments_part = document_part.part_related_by(RT.COMMENTS)
    except KeyError:
        return

    # -- Each modern relationship type to drop. Keys are the relationship
    # -- type strings; values are the human-readable labels used in the
    # -- inline comments below for ``git blame`` clarity.
    modern_rels = (
        getattr(RT, "COMMENTS_IDS", None),  # w16cid (Word 2016)
        getattr(RT, "COMMENTS_EXTENSIBLE", None),  # w16cex (Word 2018)
        getattr(RT, "COMMENTS_EXTENDED", None),  # threaded replies (Word 2013)
    )

    rels = comments_part.rels
    for rId, rel in list(rels.items()):
        if rel.reltype in modern_rels and rel.reltype is not None:
            comments_part.drop_rel(rId)
