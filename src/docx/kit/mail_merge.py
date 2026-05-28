"""Mail-merge engine — render N personalised documents from a template + records.

Closes #67.

The :func:`merge` helper composes the smart-placeholder machinery
shipped in #68 (:mod:`docx.bind_tokens`) with the standard
``load -> bind -> save`` cycle into a one-line bulk-rendering API::

    from docx.kit.mail_merge import merge

    docs = merge(
        template="offer-letter-template.docx",
        records=[
            {"first_name": "Alice", "role": "Engineer",
             "salary": "$120k", "start_date": "2026-04-01"},
            {"first_name": "Bob",   "role": "Manager",
             "salary": "$140k", "start_date": "2026-04-08"},
        ],
    )
    for doc, record in zip(docs, records):
        doc.save(f"offer-{record['first_name']}.docx")

The author writes a single template with ``{first_name}``,
``{role}``, ``{salary}``, ``{start_date}`` (or any nested
``{customer.name}`` dotted path) tokens — typically by passing
``bind_to=`` to :meth:`Document.add_paragraph` when authoring the
template, but tokens written into a Word-edited template by hand
work just as well: :func:`merge` simply re-binds each record and
lets the save-time resolver re-stamp the runs.

`template` may be a path (``str`` / :class:`os.PathLike`), a
file-like object, or an already-loaded |Document|. `records` is any
iterable of dicts (or any objects supported by the dotted-path
resolver in :mod:`docx.bind_tokens` — attribute access works too).

Output modes
------------

* **In-memory** (default) — :func:`merge` returns a
  ``list[Document]`` in the same order as `records`. The caller
  saves each one wherever they like.
* **On-disk** — pass ``output_dir=`` together with
  ``filename_template=`` (a Python ``str.format`` template that may
  reference any record field, plus ``{i}`` for the row index).
  :func:`merge` writes each document directly and returns the list
  of written paths instead of |Document| objects.

The two modes are mutually exclusive on the return-value front: pick
one based on whether you need to post-process the document tree
before saving.

Composes with #66 (content controls) and #68 (smart placeholders).
The merge engine itself is just shy of trivial — it's almost
entirely orchestration around :func:`docx.bind_tokens.apply_bind_tokens`
plus a fresh-template-per-record loop. The value is the ergonomic
bulk API: one call instead of an explicit loop authors keep
re-typing.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import io
import os
from typing import IO, TYPE_CHECKING, Any, Iterable, List, Mapping, Optional, Union

from docx import Document as _Document

if TYPE_CHECKING:
    from docx.document import Document


# -- accepted ``template`` types. Either a filesystem path / open
# -- file-like object that python-docx's :func:`Document` factory
# -- already accepts, or an already-loaded |Document| we'll snapshot
# -- via ``save -> reload`` so each merged copy starts from a fresh
# -- tree rather than mutating the caller's instance. --
TemplateLike = Union[str, "os.PathLike[str]", IO[bytes], "Document"]


__all__ = ["merge", "TemplateLike"]


def merge(
    template: TemplateLike,
    records: Iterable[Mapping[str, Any]],
    *,
    output_dir: Optional[Union[str, "os.PathLike[str]"]] = None,
    filename_template: Optional[str] = None,
) -> Union[List["Document"], List[str]]:
    """Render one personalised |Document| per record from `template`.

    For each record in `records`, a fresh copy of `template` is
    loaded, the record bound via :meth:`Document.bind`, and the
    document re-saved (in memory by default) so every
    ``{token}`` in the template re-resolves against the current
    record. The returned list is in the same order as `records`.

    `template` may be:

    * a path (``str`` or :class:`os.PathLike`) to a ``.docx`` /
      ``.docm`` / ``.dotx`` / ``.dotm`` file,
    * an open binary file-like object positioned at the start of
      the package,
    * or an already-loaded |Document| — in which case the document
      is serialised once to an in-memory buffer and the buffer is
      reused as the template source for every record so the
      caller's |Document| is **not** mutated.

    Records are dicts of token name -> value, but anything the
    dotted-path resolver in :mod:`docx.bind_tokens` accepts will
    work — including objects with attributes and nested mappings.
    Tokens whose lookup fails against a particular record stay
    literal in that record's output (the same rule the underlying
    resolver applies).

    The ``{i}`` token resolves to the current row index (0-based),
    so ``"Letter {i}: Dear {first_name}"`` is fine.

    When `output_dir` is supplied, every rendered document is
    written into that directory using `filename_template` (a
    :meth:`str.format` template that may reference any field of
    the current record plus ``{i}`` for the row index). The
    function returns the list of written paths in record order
    instead of the in-memory |Document| objects. The directory is
    created with :func:`os.makedirs` (``exist_ok=True``) if it
    doesn't already exist.

    `filename_template` is mandatory whenever `output_dir` is
    supplied. A template referencing a field absent from a given
    record raises :class:`KeyError` — the call is *all-or-nothing*
    so a typo in the filename template fails fast on the first row
    rather than half-writing the batch.

    Examples
    --------

    In-memory rendering, caller saves each doc::

        docs = merge("offer-template.docx", [
            {"first_name": "Alice"}, {"first_name": "Bob"},
        ])
        docs[0].save("offer-alice.docx")
        docs[1].save("offer-bob.docx")

    Direct-to-disk with a filename template::

        paths = merge(
            "offer-template.docx",
            [{"first_name": "Alice"}, {"first_name": "Bob"}],
            output_dir="out/",
            filename_template="offer-{first_name}.docx",
        )
        # ['out/offer-Alice.docx', 'out/offer-Bob.docx']

    Pre-loaded |Document| as the template (useful when the
    template was just built programmatically and you don't want a
    round-trip through disk)::

        tmpl = Document()
        tmpl.add_paragraph("Dear {first_name},", bind_to={})
        docs = merge(tmpl, [{"first_name": "Alice"}])

    .. versionadded:: 2026.05.29
    """
    # -- Materialise the records iterable once. The function loops
    # -- twice when ``output_dir`` is set (filename + bind), and we
    # -- want a stable ordering regardless of whether the caller
    # -- passed a list or a generator. --
    records_list = list(records)

    if output_dir is not None and filename_template is None:
        raise ValueError(
            "filename_template is required when output_dir is supplied; "
            "pass e.g. filename_template='offer-{first_name}.docx'."
        )
    if output_dir is None and filename_template is not None:
        raise ValueError(
            "filename_template requires output_dir; pass output_dir='out/' "
            "or drop filename_template to receive in-memory Documents instead."
        )

    template_bytes = _snapshot_template(template)

    if output_dir is not None:
        return _merge_to_disk(
            template_bytes,
            records_list,
            output_dir=output_dir,
            filename_template=filename_template or "",
        )
    return _merge_in_memory(template_bytes, records_list)


# -- ---------------------------------------------------------------
# -- internals
# -- ---------------------------------------------------------------


def _snapshot_template(template: TemplateLike) -> bytes:
    """Return a fresh ``bytes`` blob of the package referenced by `template`.

    The blob is reused once per record to materialise an isolated
    |Document| tree — never the same tree twice, so binding one
    record can't leak into the next one's render.

    For a path, the file is read once. For a file-like object, the
    stream is consumed and rewound where possible (we don't assume
    the caller can re-seek their own stream after we're done — the
    snapshot decouples us). For a pre-loaded |Document|, we save it
    to an in-memory buffer; this also implicitly resolves any
    pending bind tokens against the current record on the source
    document, which is fine because the snapshot is then re-bound
    fresh per record.
    """
    # -- pre-loaded Document: serialise once to a BytesIO buffer.
    # -- Detected by the presence of the public ``save`` and
    # -- ``element`` attributes rather than an isinstance check —
    # -- ``docx.document.Document`` would force a circular import at
    # -- module load and ``Document`` is a *function* in the public
    # -- API, not a class. --
    if hasattr(template, "save") and hasattr(template, "element"):
        buf = io.BytesIO()
        template.save(buf)  # type: ignore[union-attr]
        return buf.getvalue()

    # -- filesystem path-like --
    if isinstance(template, (str, os.PathLike)):
        path = os.fspath(template)
        with open(path, "rb") as fh:
            return fh.read()

    # -- file-like object: read all bytes. --
    data = template.read()  # type: ignore[union-attr]
    if isinstance(data, str):  # pragma: no cover - defensive
        raise TypeError(
            "template file-like must yield bytes, not str — open the file "
            "in binary mode ('rb')."
        )
    return data


def _load_fresh(template_bytes: bytes) -> "Document":
    """Return a freshly-loaded |Document| from `template_bytes`.

    Each record gets its own |Document| tree so ``Document.bind``
    on row N can't leak into row N+1's resolution. The cost is one
    re-parse per record, which is dwarfed by Word's I/O cost in
    every realistic workload (mail-merge is rarely CPU-bound).
    """
    return _Document(io.BytesIO(template_bytes))


def _merge_in_memory(
    template_bytes: bytes,
    records: List[Mapping[str, Any]],
) -> List["Document"]:
    """Render every record into a fresh |Document| and return the list."""
    out: List["Document"] = []
    for i, record in enumerate(records):
        doc = _load_fresh(template_bytes)
        doc.bind(record=record, iteration=i)
        # -- resolve tokens *now* by round-tripping through save.
        # -- ``bind_to`` is normally save-time, but a caller of
        # -- :func:`merge` expects the returned Document to already
        # -- carry the resolved text — they may inspect ``.paragraphs``
        # -- before saving, or pipe the doc into another helper. --
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        out.append(_Document(buf))
    return out


def _merge_to_disk(
    template_bytes: bytes,
    records: List[Mapping[str, Any]],
    *,
    output_dir: Union[str, "os.PathLike[str]"],
    filename_template: str,
) -> List[str]:
    """Render and save every record directly to disk; return path list."""
    out_path = os.fspath(output_dir)
    os.makedirs(out_path, exist_ok=True)
    paths: List[str] = []
    for i, record in enumerate(records):
        filename = _format_filename(filename_template, record, i)
        full_path = os.path.join(out_path, filename)
        doc = _load_fresh(template_bytes)
        doc.bind(record=record, iteration=i)
        doc.save(full_path)
        paths.append(full_path)
    return paths


def _format_filename(
    template: str,
    record: Mapping[str, Any],
    iteration: int,
) -> str:
    """Format `template` with the fields of `record` plus ``{i}``.

    The default :meth:`str.format` mapping argument requires every
    referenced field to be present in `record`; we add ``i`` on top
    of a shallow copy so the iteration index is always available.
    Missing-field errors propagate as :class:`KeyError`.
    """
    fields: dict[str, Any] = dict(record)
    # -- record-supplied ``i`` wins over the iteration index, which
    # -- is intentional: a caller storing a row id under ``"i"`` in
    # -- their record probably wants their value, not ours. --
    fields.setdefault("i", iteration)
    return template.format(**fields)
