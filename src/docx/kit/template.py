"""Load a corporate ``.dotx`` template, substitute ``[KEY]`` placeholders, return |Document|.

Closes #305.

The :func:`from_template_dotx` helper composes the existing
``Document.from_template`` content-type swap with a straightforward
``[KEY]`` -> ``value`` placeholder pass over every run text in every
paragraph (body and table cells)::

    from docx.kit import from_template_dotx

    doc = from_template_dotx(
        "corporate_template.dotx",
        placeholders={
            "[CLIENT]": "ACME Corp",
            "[DATE]":   "2026-05-29",
            "[AUTHOR]": "Jane Smith",
        },
    )
    doc.save("client-letter.docx")

The substitution is intentionally simple — a flat string-replace on
each run's text. This preserves run-level formatting (bold, italic,
font, colour) because the run boundary is not crossed; the cost is
that placeholders that *span* run boundaries (Word inserts splits at
spell-check / autocorrect edits) are not rewritten. The kit ethos is
"works for hand-authored templates"; for run-spanning placeholders
the smart-placeholder machinery in :mod:`docx.bind_tokens` is the
right escalation.

The output is a regular ``.docx`` (not ``.dotx``) — the template's
content type is swapped to the document variant via the existing
:func:`docx.from_template` helper before substitution. Callers save
wherever they like via :meth:`Document.save`.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import io
import os
from typing import IO, TYPE_CHECKING, Mapping, Optional, Union

from docx import from_template as _from_template

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls


__all__ = ["from_template_dotx"]


# -- The template path / stream type accepted by :func:`from_template_dotx`.
# -- Mirrors the shape :func:`docx.from_template` already accepts. --
TemplatePath = Union[str, "os.PathLike[str]", IO[bytes]]


def from_template_dotx(
    path: TemplatePath,
    *,
    placeholders: Optional[Mapping[str, str]] = None,
    encoding: str = "utf-8",
) -> "DocumentCls":
    """Load a ``.dotx`` template, fill placeholders, return a |Document|.

    The template is loaded via :func:`docx.from_template` so the
    returned |Document|'s main-document content type is the regular
    document variant (``.docx``) rather than the template variant
    (``.dotx``) — saving the result produces a normal Word document.

    `placeholders` is a mapping of literal placeholder strings to
    their replacement values. The default convention is
    bracket-delimited keys (e.g. ``"[CLIENT]"``) but any non-empty
    string works — the substitution is a flat
    :meth:`str.replace` per run text. Run-level formatting is
    preserved because each replacement happens *inside* one run; for
    placeholders that span run boundaries see the smart-placeholder
    machinery in :mod:`docx.bind_tokens`.

    Substitution covers every run in:

    * every body paragraph (``document.paragraphs``),
    * every cell in every table (``document.tables`` recursing
      through nested tables).

    Headers, footers, footnotes, endnotes, and text boxes are *not*
    walked — kit consumers who need those should escalate to
    :mod:`docx.bind_tokens`. The kit aim is the 80% common case:
    body and table-cell placeholders.

    Parameters
    ----------
    path
        Filesystem path (``str`` / :class:`os.PathLike`) or
        binary file-like object pointing at a ``.dotx`` Word
        template package.
    placeholders
        Mapping of placeholder string -> replacement string. When
        omitted or empty, the template is loaded and returned with
        no substitution.
    encoding
        Reserved for future extension (e.g. byte-level template
        loading). Currently unused — placeholders and replacements
        are always Python ``str``. Defaults to ``"utf-8"``.

    Returns
    -------
    Document
        A freshly-loaded |Document| with the template content type
        already swapped to the document variant. Save with
        :meth:`Document.save` (any path / extension).

    Raises
    ------
    ValueError
        When `path` does not point to a Word template package
        (propagated from :func:`docx.from_template`).
    TypeError
        When any placeholder key or value is not a ``str``.

    Examples
    --------

    Load a corporate template, fill three placeholders, save::

        from docx.kit import from_template_dotx

        doc = from_template_dotx(
            "corporate_template.dotx",
            placeholders={
                "[CLIENT]": "ACME Corp",
                "[DATE]":   "2026-05-29",
                "[AUTHOR]": "Jane Smith",
            },
        )
        doc.save("acme-engagement.docx")

    Load with no substitution (effectively
    :func:`docx.from_template`)::

        doc = from_template_dotx("template.dotx")

    .. versionadded:: 2026.05.29
    """
    # -- ``encoding`` is currently a no-op but reserved on the
    # -- public signature so a future "read template bytes with this
    # -- codec before substitution" extension doesn't break callers.
    del encoding

    if placeholders is not None:
        _validate_placeholders(placeholders)

    document = _from_template(path)

    if placeholders:
        _apply_placeholders(document, placeholders)

    return document


# -- ---------------------------------------------------------------
# -- internals
# -- ---------------------------------------------------------------


def _validate_placeholders(placeholders: Mapping[str, str]) -> None:
    """Raise :class:`TypeError` when any key or value is not a ``str``.

    The substitution layer is a per-run :meth:`str.replace`; non-str
    keys / values would either raise deep inside lxml or produce
    silently broken output. Failing fast at the API boundary surfaces
    the bug at the caller's stack frame.
    """
    for key, value in placeholders.items():
        if not isinstance(key, str):
            raise TypeError(
                "placeholder keys must be str, got %s for key %r"
                % (type(key).__name__, key)
            )
        if not isinstance(value, str):
            raise TypeError(
                "placeholder values must be str, got %s for key %r -> %r"
                % (type(value).__name__, key, value)
            )


def _apply_placeholders(
    document: "DocumentCls", placeholders: Mapping[str, str]
) -> None:
    """Walk every run in body paragraphs and table cells; substitute placeholders.

    Run-level formatting (bold, italic, font, colour, …) is
    preserved because each replacement happens *inside* one run's
    ``w:t`` text rather than rebuilding the run.
    """
    # -- Body paragraphs --
    for paragraph in document.paragraphs:
        _substitute_paragraph(paragraph, placeholders)

    # -- Tables (recurse through nested tables) --
    for table in document.tables:
        _substitute_table(table, placeholders)


def _substitute_paragraph(paragraph, placeholders: Mapping[str, str]) -> None:
    """Replace placeholders within every run text on `paragraph`."""
    for run in paragraph.runs:
        original = run.text
        if not original:
            continue
        replaced = _replace_all(original, placeholders)
        if replaced != original:
            run.text = replaced


def _substitute_table(table, placeholders: Mapping[str, str]) -> None:
    """Recurse into every cell of `table`; substitute paragraphs and nested tables."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _substitute_paragraph(paragraph, placeholders)
            for nested in cell.tables:
                _substitute_table(nested, placeholders)


def _replace_all(text: str, placeholders: Mapping[str, str]) -> str:
    """Return `text` with every key in `placeholders` replaced by its value.

    Iteration order follows the mapping's iteration order, which is
    insertion-ordered in Python 3.7+. Callers worried about ordering
    edge cases (one placeholder being a substring of another) can
    pre-order their mapping accordingly.
    """
    out = text
    for key, value in placeholders.items():
        if key and key in out:
            out = out.replace(key, value)
    return out


# -- ``io`` is imported for symmetry with the rest of the kit even
# -- though :func:`docx.from_template` already handles file-like
# -- inputs internally; keeping the import preserves the option to
# -- switch to a snapshot pattern (cf. :mod:`docx.kit.mail_merge`)
# -- without re-touching the import block. --
_ = io
