"""RFC-6902 JSON-Patch over the python-docx structural model.

Closes #302.

:func:`apply` accepts a |Document| and a list of RFC-6902 patch
operations and applies them to a stable addressable subset of the
document's model — a flat *paragraphs* list, the *tables* grid, and
the *sections* list. Every op is validated against an in-memory
*shadow* of the document in a pre-flight pass before any mutation is
written to the live tree, so a failing op never leaves the document
half-mutated::

    from docx import Document
    from docx.kit import patch

    doc = Document("input.docx")
    patch.apply(doc, [
        {"op": "replace", "path": "/paragraphs/0/text", "value": "New title"},
        {"op": "add", "path": "/paragraphs/-",
         "value": {"text": "Final paragraph", "style": "Normal"}},
        {"op": "remove", "path": "/paragraphs/2"},
        {"op": "test", "path": "/paragraphs/0/style", "value": "Title"},
    ])
    doc.save("out.docx")

Path scheme (RFC-6901 with ``~0`` -> ``~`` and ``~1`` -> ``/``):

* ``/paragraphs/N[/text|/style]`` — N is 0-indexed; negative N counts
  from the end; ``-`` (RFC-6901 end-of-array) only valid as an
  ``add`` target.
* ``/by_id/<paraId>[/text|/style]`` — keyed on ``w14:paraId``
  (the stable ID Word stamps on every paragraph in 2010+ files).
* ``/tables/N/rows/M/cells/K/text`` — cell text content.
* ``/sections/N/page_orientation`` — ``"portrait"`` or ``"landscape"``.

Supported ops: ``add``, ``remove``, ``replace``, ``move``, ``copy``,
``test``. ``move`` / ``copy`` source via ``"from"``. ``test`` raises
:class:`PatchTestFailed` on inequality.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import copy
from typing import (
    TYPE_CHECKING,
    Any,
    Iterable,
    List,
    Mapping,
    Optional,
    Tuple,
)

from docx.enum.section import WD_ORIENTATION
from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls


__all__ = [
    "apply",
    "PatchError",
    "PathNotFound",
    "PatchTestFailed",
    "InvalidOp",
]


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


class PatchError(Exception):
    """Base error raised when :func:`apply` cannot apply an op.

    Subclasses distinguish *why*: :class:`PathNotFound` (the path
    doesn't resolve), :class:`PatchTestFailed` (a ``test`` op did not
    hold), or :class:`InvalidOp` (the op shape is malformed).
    """


class PathNotFound(PatchError):
    """The JSON-Pointer path does not resolve in the current document."""


class PatchTestFailed(PatchError):
    """A ``test`` op's value did not equal the document's current value."""


class InvalidOp(PatchError):
    """The patch op is malformed (unknown op, missing key, bad value)."""


# ---------------------------------------------------------------------------
# JSON Pointer (RFC 6901) helpers
# ---------------------------------------------------------------------------


def _unescape_token(token: str) -> str:
    """Decode an RFC-6901 reference-token (``~1`` -> ``/``, ``~0`` -> ``~``)."""
    # -- ~1 first then ~0 so ~01 decodes to ~1, not / --
    return token.replace("~1", "/").replace("~0", "~")


def _split_pointer(pointer: str) -> List[str]:
    """Return the decoded reference-tokens for `pointer` (RFC-6901)."""
    if pointer == "":
        return []
    if not pointer.startswith("/"):
        raise InvalidOp(f"JSON Pointer must start with '/': {pointer!r}")
    return [_unescape_token(t) for t in pointer[1:].split("/")]


def _parse_index(token: str, length: int) -> int:
    """Coerce a JSON-Pointer array token to a Python list index.

    Supports negative indices (``-1`` = last) on top of the RFC.
    Raises :class:`PathNotFound` for out-of-range or non-integer.
    """
    if token == "-":
        raise PathNotFound("'-' (end-of-array) only valid as an 'add' target")
    try:
        idx = int(token)
    except ValueError as exc:
        raise PathNotFound(f"expected integer index, got {token!r}") from exc
    if idx < 0:
        idx = length + idx
    if idx < 0 or idx >= length:
        raise PathNotFound(
            f"index {token} out of range for sequence of length {length}"
        )
    return idx


# ---------------------------------------------------------------------------
# Shadow model — a plain dict/list mirror of the addressable surface
# ---------------------------------------------------------------------------


def _para_id(paragraph: Any) -> Optional[str]:
    """Return the ``w14:paraId`` of `paragraph` if Word has assigned one.

    Reads ``w14:paraId`` off the underlying ``w:p`` element. The kit's
    "compose, don't reach down" rule applies to *mutations*; here we
    only *read* a stable ID that has no public-API surface yet.
    """
    elem = getattr(paragraph, "_p", None)
    if elem is None:
        return None
    return elem.get(qn("w14:paraId"))


def _orientation_str(value: Any) -> str:
    """Return the lowercase string form of an orientation."""
    if isinstance(value, str):
        return value.lower()
    if isinstance(value, WD_ORIENTATION):
        return value.xml_value or value.name.lower()
    return str(value).lower()


def _build_shadow(doc: "DocumentCls") -> dict:
    """Return a plain dict mirror of the patchable surface of `doc`."""
    shadow: dict = {"paragraphs": [], "tables": [], "sections": []}
    by_id: dict = {}
    for idx, paragraph in enumerate(doc.paragraphs):
        style = paragraph.style
        entry = {
            "_index": idx,
            "text": paragraph.text,
            "style": style.name if style is not None else None,
        }
        shadow["paragraphs"].append(entry)
        pid = _para_id(paragraph)
        if pid is not None:
            by_id[pid] = entry
    shadow["by_id"] = by_id
    for table in doc.tables:
        rows = [
            {"cells": [{"text": cell.text} for cell in row.cells]}
            for row in table.rows
        ]
        shadow["tables"].append({"rows": rows})
    for section in doc.sections:
        shadow["sections"].append(
            {"page_orientation": _orientation_str(section.orientation)}
        )
    return shadow


# ---------------------------------------------------------------------------
# Resolved-path tuple (used by both shadow and live pass)
# ---------------------------------------------------------------------------

# A resolved path is a tuple of (kind, *args). Kinds:
#   ("paragraph", index, field)        — field in {"text","style", None}
#   ("paragraph_append",)              — /paragraphs/-, add-only
#   ("paragraph_by_id", paraId, field) — field in {"text","style"}
#   ("cell", t_idx, r_idx, c_idx)      — table cell text
#   ("section", s_idx, "page_orientation")


ResolvedPath = Tuple[Any, ...]


def _resolve_path(
    pointer: str,
    shadow: dict,
    *,
    for_add: bool = False,
) -> ResolvedPath:
    """Return the structural address `pointer` refers to.

    Raises :class:`PathNotFound` when the path does not address an
    existing surface (or when `for_add=False` and an ``add``-only path
    such as ``/paragraphs/-`` was supplied).
    """
    tokens = _split_pointer(pointer)
    if not tokens:
        raise PathNotFound("root pointer '' is not patchable")

    head, rest = tokens[0], tokens[1:]

    if head == "paragraphs":
        if not rest:
            raise PathNotFound("/paragraphs is the container; address a child")
        if rest[0] == "-":
            if rest[1:]:
                raise PathNotFound("trailing tokens after '-' sentinel")
            if not for_add:
                raise PathNotFound("/paragraphs/- only valid as 'add' target")
            return ("paragraph_append",)
        n = _parse_index(rest[0], len(shadow["paragraphs"]))
        if not rest[1:]:
            return ("paragraph", n, None)
        field = rest[1]
        if field not in ("text", "style") or rest[2:]:
            raise PathNotFound(
                f"unknown paragraph path /paragraphs/{n}/{'/'.join(rest[1:])}"
            )
        return ("paragraph", n, field)

    if head == "by_id":
        if not rest:
            raise PathNotFound("/by_id is the container; supply a paraId")
        paragraph_id = rest[0]
        if paragraph_id not in shadow["by_id"]:
            raise PathNotFound(f"no paragraph with paraId {paragraph_id!r}")
        if not rest[1:] or rest[1] not in ("text", "style") or rest[2:]:
            raise PathNotFound(
                f"/by_id/{paragraph_id} requires /text or /style suffix"
            )
        return ("paragraph_by_id", paragraph_id, rest[1])

    if head == "tables":
        if len(rest) < 5 or rest[1] != "rows" or rest[3] != "cells":
            raise PathNotFound(
                "/tables paths must look like /tables/N/rows/M/cells/K/text"
            )
        t_idx = _parse_index(rest[0], len(shadow["tables"]))
        rows = shadow["tables"][t_idx]["rows"]
        r_idx = _parse_index(rest[2], len(rows))
        cells = rows[r_idx]["cells"]
        c_idx = _parse_index(rest[4], len(cells))
        if rest[5:6] != ["text"] or rest[6:]:
            raise PathNotFound("/tables cell paths must terminate at /text")
        return ("cell", t_idx, r_idx, c_idx)

    if head == "sections":
        if len(rest) != 2 or rest[1] != "page_orientation":
            raise PathNotFound(
                "/sections paths must look like /sections/N/page_orientation"
            )
        return ("section", _parse_index(rest[0], len(shadow["sections"])),
                "page_orientation")

    raise PathNotFound(f"unknown top-level path segment: /{head}")


# ---------------------------------------------------------------------------
# Shadow pass — apply each op to the in-memory mirror
# ---------------------------------------------------------------------------


def _shadow_get(shadow: dict, path: ResolvedPath) -> Any:
    kind = path[0]
    if kind == "paragraph":
        _, idx, field = path
        entry = shadow["paragraphs"][idx]
        return entry if field is None else entry[field]
    if kind == "paragraph_by_id":
        _, paragraph_id, field = path
        return shadow["by_id"][paragraph_id][field]
    if kind == "cell":
        _, t, r, c = path
        return shadow["tables"][t]["rows"][r]["cells"][c]["text"]
    if kind == "section":
        _, s, _field = path
        return shadow["sections"][s]["page_orientation"]
    raise InvalidOp(f"path kind {kind!r} is not readable")


def _validate_paragraph_value(value: Any) -> dict:
    """Return a normalised ``{"text": str, "style": str|None}`` for an add."""
    if not isinstance(value, Mapping):
        raise InvalidOp(
            "paragraph add 'value' must be a mapping with 'text' (and optional 'style')"
        )
    if "text" not in value:
        raise InvalidOp("paragraph add 'value' is missing required key 'text'")
    text = value["text"]
    if not isinstance(text, str):
        raise InvalidOp("paragraph add 'value.text' must be a string")
    style = value.get("style")
    if style is not None and not isinstance(style, str):
        raise InvalidOp("paragraph add 'value.style' must be a string or omitted")
    return {"text": text, "style": style}


def _normalise_orientation_value(value: Any) -> str:
    """Coerce a ``page_orientation`` patch value to ``portrait`` / ``landscape``."""
    s = _orientation_str(value)
    if s not in ("portrait", "landscape"):
        raise InvalidOp(
            f"page_orientation must be 'portrait' or 'landscape', got {value!r}"
        )
    return s


def _shadow_apply_add(shadow: dict, path: ResolvedPath, value: Any) -> None:
    kind = path[0]
    if kind == "paragraph_append":
        entry = _validate_paragraph_value(value)
        new_idx = len(shadow["paragraphs"])
        shadow["paragraphs"].append(
            {"_index": new_idx, "text": entry["text"], "style": entry["style"]}
        )
        return
    if kind == "paragraph":
        _, idx, field = path
        if field is None:
            # -- adding a whole paragraph at index — equivalent to
            # -- inserting before. The kit doesn't offer a public
            # -- ``Document.insert_paragraph(index)``, so this is
            # -- intentionally not supported. --
            raise InvalidOp(
                "add to /paragraphs/N (no field) is not supported — "
                "use /paragraphs/- to append"
            )
        if not isinstance(value, str):
            raise InvalidOp(f"paragraph {field} 'value' must be a string")
        shadow["paragraphs"][idx][field] = value
        return
    if kind == "paragraph_by_id":
        _, paragraph_id, field = path
        if not isinstance(value, str):
            raise InvalidOp(f"paragraph {field} 'value' must be a string")
        shadow["by_id"][paragraph_id][field] = value
        return
    if kind == "cell":
        _, t, r, c = path
        if not isinstance(value, str):
            raise InvalidOp("cell text 'value' must be a string")
        shadow["tables"][t]["rows"][r]["cells"][c]["text"] = value
        return
    if kind == "section":
        _, s, _field = path
        shadow["sections"][s]["page_orientation"] = _normalise_orientation_value(
            value
        )
        return
    raise InvalidOp(f"add path kind {kind!r} is not supported")


def _shadow_apply_remove(shadow: dict, path: ResolvedPath) -> None:
    kind = path[0]
    if kind == "paragraph":
        _, idx, field = path
        if field is None:
            entry = shadow["paragraphs"].pop(idx)
            # -- drop the by_id entry and re-number _index --
            for pid, ent in list(shadow["by_id"].items()):
                if ent is entry:
                    del shadow["by_id"][pid]
            for i, e in enumerate(shadow["paragraphs"]):
                e["_index"] = i
            return
        # -- field-level remove — clear the text or style. The kit
        # -- treats "remove a paragraph property" as setting it to the
        # -- empty value. --
        if field == "text":
            shadow["paragraphs"][idx]["text"] = ""
        else:
            shadow["paragraphs"][idx]["style"] = None
        return
    if kind == "paragraph_by_id":
        _, paragraph_id, field = path
        if field == "text":
            shadow["by_id"][paragraph_id]["text"] = ""
        else:
            shadow["by_id"][paragraph_id]["style"] = None
        return
    if kind == "cell":
        _, t, r, c = path
        shadow["tables"][t]["rows"][r]["cells"][c]["text"] = ""
        return
    if kind == "section":
        # -- section orientation has no "remove" — the spec mandates
        # -- one of two values. Reject up-front. --
        raise InvalidOp(
            "remove on /sections/.../page_orientation is not supported "
            "— orientation must be 'portrait' or 'landscape'"
        )
    raise InvalidOp(f"remove path kind {kind!r} is not supported")


def _shadow_apply_replace(shadow: dict, path: ResolvedPath, value: Any) -> None:
    # -- replace == add for our addressable surface; both write a value
    # -- to an existing slot. The semantic difference (does the slot
    # -- have to exist?) is enforced in :func:`_resolve_path`. --
    if path[0] == "paragraph_append":
        raise InvalidOp("'-' (end-of-array) is only valid as an 'add' target")
    _shadow_apply_add(shadow, path, value)


def _shadow_apply_test(shadow: dict, path: ResolvedPath, value: Any) -> None:
    actual = _shadow_get(shadow, path)
    expected: Any
    if path[0] == "section":
        expected = _normalise_orientation_value(value)
    else:
        expected = value
    if actual != expected:
        raise PatchTestFailed(
            f"test failed at path: expected {expected!r}, got {actual!r}"
        )


# ---------------------------------------------------------------------------
# Op dispatcher (shadow pass)
# ---------------------------------------------------------------------------


def _shadow_apply_op(shadow: dict, op: Mapping[str, Any]) -> None:
    if not isinstance(op, Mapping):
        raise InvalidOp(f"op must be a mapping, got {type(op).__name__}")
    if "op" not in op or "path" not in op:
        raise InvalidOp("op must carry both 'op' and 'path' keys")
    name = op["op"]
    path_str = op["path"]
    if not isinstance(name, str) or not isinstance(path_str, str):
        raise InvalidOp("'op' and 'path' must be strings")

    if name == "add":
        if "value" not in op:
            raise InvalidOp("'add' op is missing 'value'")
        path = _resolve_path(path_str, shadow, for_add=True)
        _shadow_apply_add(shadow, path, op["value"])
        return

    if name == "remove":
        path = _resolve_path(path_str, shadow)
        _shadow_apply_remove(shadow, path)
        return

    if name == "replace":
        if "value" not in op:
            raise InvalidOp("'replace' op is missing 'value'")
        path = _resolve_path(path_str, shadow)
        _shadow_apply_replace(shadow, path, op["value"])
        return

    if name == "test":
        if "value" not in op:
            raise InvalidOp("'test' op is missing 'value'")
        path = _resolve_path(path_str, shadow)
        _shadow_apply_test(shadow, path, op["value"])
        return

    if name in ("move", "copy"):
        if "from" not in op:
            raise InvalidOp(f"'{name}' op is missing 'from'")
        from_path_str = op["from"]
        if not isinstance(from_path_str, str):
            raise InvalidOp("'from' must be a string")
        src_path = _resolve_path(from_path_str, shadow)
        src_value = _shadow_get(shadow, src_path)
        # -- normalise the snapshot to a *value* (the dict view of a
        # -- whole paragraph, or a primitive for fields) so subsequent
        # -- mutation can't share state with the source. --
        if isinstance(src_value, dict):
            snapshot: Any = {
                "text": src_value["text"],
                "style": src_value["style"],
            }
        else:
            snapshot = src_value
        if name == "move":
            _shadow_apply_remove(shadow, src_path)
            # -- after removing the source, dest indices may have
            # -- shifted; re-resolve against the post-remove shadow. --
            dest_path = _resolve_path(path_str, shadow, for_add=True)
        else:
            dest_path = _resolve_path(path_str, shadow, for_add=True)
        _shadow_apply_add(shadow, dest_path, snapshot)
        return

    raise InvalidOp(f"unknown op {name!r}")


# ---------------------------------------------------------------------------
# Live pass — replay each op against the actual python-docx Document
# ---------------------------------------------------------------------------


def _live_paragraph(doc: "DocumentCls", index: int):
    paragraphs = doc.paragraphs
    if index < 0:
        index = len(paragraphs) + index
    return paragraphs[index]


def _live_paragraph_by_id(doc: "DocumentCls", paragraph_id: str):
    for paragraph in doc.paragraphs:
        if _para_id(paragraph) == paragraph_id:
            return paragraph
    raise PathNotFound(  # pragma: no cover — pre-flight already caught this
        f"no paragraph with paraId {paragraph_id!r} after live pass"
    )


def _live_set_paragraph_field(paragraph: Any, field: str, value: Any) -> None:
    if field == "text":
        paragraph.text = value if value is not None else ""
    elif field == "style":
        paragraph.style = value  # accepts str | None
    else:  # pragma: no cover — guarded in pre-flight
        raise InvalidOp(f"unknown paragraph field {field!r}")


def _live_remove_paragraph(paragraph: Any) -> None:
    """Detach `paragraph` from its parent body element.

    Goes through ``Paragraph._p`` -> ``getparent`` since python-docx
    has no public ``Document.remove_paragraph()``; mirrors the same
    access path used by :mod:`docx.kit.layout`.
    """
    elem = paragraph._p
    parent = elem.getparent()
    if parent is None:  # pragma: no cover — body is always present
        raise PatchError("paragraph has no parent body element")
    parent.remove(elem)


def _live_apply_op(doc: "DocumentCls", op: Mapping[str, Any]) -> None:
    name = op["op"]
    path_str = op["path"]
    # -- re-build the shadow each op so indices reflect prior mutations --
    shadow = _build_shadow(doc)

    if name == "add":
        path = _resolve_path(path_str, shadow, for_add=True)
        _live_apply_add(doc, path, op["value"])
        return

    if name == "remove":
        path = _resolve_path(path_str, shadow)
        _live_apply_remove(doc, path)
        return

    if name == "replace":
        path = _resolve_path(path_str, shadow)
        _live_apply_replace(doc, path, op["value"])
        return

    if name == "test":
        # -- test was already validated in the shadow pre-flight --
        return

    if name in ("move", "copy"):
        src_path = _resolve_path(op["from"], shadow)
        snapshot = _live_snapshot(doc, src_path)
        if name == "move":
            _live_apply_remove(doc, src_path)
            shadow = _build_shadow(doc)
        dest_path = _resolve_path(path_str, shadow, for_add=True)
        _live_apply_add(doc, dest_path, snapshot)
        return

    raise InvalidOp(f"unknown op {name!r}")  # pragma: no cover — pre-flight catches


def _live_snapshot(doc: "DocumentCls", path: ResolvedPath) -> Any:
    """Return a serialisable copy of the value at `path` for move/copy."""
    kind = path[0]
    if kind == "paragraph":
        _, idx, field = path
        paragraph = _live_paragraph(doc, idx)
        if field is None:
            style = paragraph.style
            return {
                "text": paragraph.text,
                "style": style.name if style is not None else None,
            }
        if field == "text":
            return paragraph.text
        style = paragraph.style
        return style.name if style is not None else None
    if kind == "paragraph_by_id":
        _, paragraph_id, field = path
        paragraph = _live_paragraph_by_id(doc, paragraph_id)
        if field == "text":
            return paragraph.text
        style = paragraph.style
        return style.name if style is not None else None
    if kind == "cell":
        _, t, r, c = path
        return doc.tables[t].rows[r].cells[c].text
    if kind == "section":
        return _orientation_str(doc.sections[path[1]].orientation)
    raise InvalidOp(  # pragma: no cover
        f"snapshot path kind {kind!r} is not supported"
    )


def _live_apply_add(doc: "DocumentCls", path: ResolvedPath, value: Any) -> None:
    kind = path[0]
    if kind == "paragraph_append":
        entry = _validate_paragraph_value(value)
        doc.add_paragraph(entry["text"], style=entry["style"])
        return
    if kind == "paragraph":
        _, idx, field = path
        paragraph = _live_paragraph(doc, idx)
        _live_set_paragraph_field(paragraph, field, value)
        return
    if kind == "paragraph_by_id":
        _, paragraph_id, field = path
        paragraph = _live_paragraph_by_id(doc, paragraph_id)
        _live_set_paragraph_field(paragraph, field, value)
        return
    if kind == "cell":
        _, t, r, c = path
        doc.tables[t].rows[r].cells[c].text = value
        return
    if kind == "section":
        _, s, _field = path
        normalised = _normalise_orientation_value(value)
        target = (
            WD_ORIENTATION.LANDSCAPE
            if normalised == "landscape"
            else WD_ORIENTATION.PORTRAIT
        )
        doc.sections[s].orientation = target
        return
    raise InvalidOp(  # pragma: no cover
        f"live add path kind {kind!r} is not supported"
    )


def _live_apply_remove(doc: "DocumentCls", path: ResolvedPath) -> None:
    kind = path[0]
    if kind == "paragraph":
        _, idx, field = path
        paragraph = _live_paragraph(doc, idx)
        if field is None:
            _live_remove_paragraph(paragraph)
            return
        _live_set_paragraph_field(
            paragraph, field, "" if field == "text" else None
        )
        return
    if kind == "paragraph_by_id":
        _, paragraph_id, field = path
        paragraph = _live_paragraph_by_id(doc, paragraph_id)
        _live_set_paragraph_field(
            paragraph, field, "" if field == "text" else None
        )
        return
    if kind == "cell":
        _, t, r, c = path
        doc.tables[t].rows[r].cells[c].text = ""
        return
    raise InvalidOp(  # pragma: no cover — section remove is rejected up-front
        f"live remove path kind {kind!r} is not supported"
    )


def _live_apply_replace(doc: "DocumentCls", path: ResolvedPath, value: Any) -> None:
    if path[0] == "paragraph_append":  # pragma: no cover — pre-flight catches
        raise InvalidOp("'-' is only valid as an 'add' target")
    _live_apply_add(doc, path, value)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def apply(
    doc: "DocumentCls",
    ops: Iterable[Mapping[str, Any]],
) -> None:
    """Apply a list of RFC-6902 patch operations to `doc`.

    All ops are validated against an in-memory shadow of the document
    before any mutation reaches the live tree. If any op fails (an
    unresolved path, a failed ``test``, a bad ``value`` type), the
    document is untouched and a :class:`PatchError` subclass is raised.

    Returns ``None`` on success. Mutations are visible on `doc`
    immediately; call :meth:`Document.save` to persist them.
    """
    if doc is None:
        raise InvalidOp("doc must be a Document, got None")

    op_list = list(ops)
    for op in op_list:
        if not isinstance(op, Mapping):
            raise InvalidOp(
                f"each op must be a mapping, got {type(op).__name__}"
            )

    shadow_copy = copy.deepcopy(_build_shadow(doc))
    for op in op_list:
        _shadow_apply_op(shadow_copy, op)

    for op in op_list:
        _live_apply_op(doc, op)
