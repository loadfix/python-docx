"""Save-time smart-placeholder field tokens for python-docx.

This module implements the in-text templating language proposed in
issue #68 — the docx mirror of the page-token mechanism shipped on
python-pptx for issue #38. An author writes a single literal string
into ``Document.add_paragraph(text, bind_to=record)``, where ``text``
contains ``{token}`` references such as ``{customer.name}``,
``{date:short}`` or ``{property:Title}``; on every ``Document.save``
the tokens are resolved against the bound record and stamped into
``<w:t>``. The original token source string is preserved in a fork-
scoped ``<lfxbind:src>`` child of the run so that the next
``load -> bind -> save`` cycle re-resolves against the new record
instead of carrying a stale literal forward.

Supported tokens
----------------

================================  ====================================================
Token                              Resolves to
================================  ====================================================
``{customer.name}``                dotted-path lookup against the bound record
``{customer.address.line1}``       nested dotted-path lookup
``{date:short}``                   today's date, locale-agnostic short ISO-style fmt
``{date:medium}``                  today's date, ``"MMM d, yyyy"`` style
``{date:long}``                    today's date, ``"MMMM d, yyyy"`` style
``{date:iso}``                     today's date, ``YYYY-MM-DD``
``{date:'<fmt>'}``                 today's date with a custom strftime/Babel-style
                                   format (e.g. ``{date:'MMM d, yyyy'}``)
``{i}``                            the current iteration index (mail-merge context)
``{property:<Name>}``              ``Document.core_properties.<Name>`` lookup,
                                   falling back to the custom-properties collection
================================  ====================================================

Round-trip preservation
-----------------------

Whenever a token-bearing run is resolved, the original source string
is stamped into a ``<lfxbind:src>`` child element appended after the
run's ``<w:t>``. Word and every other OOXML consumer follow the
"preserve but ignore unknown children" convention, so the marker
survives untouched across edits performed by other tools. On the
next ``load -> bind -> save`` cycle, :func:`apply_bind_tokens`
discovers the marker, reseats the source string, and re-resolves to
the new record.

Public surface
--------------

Most users never call into this module directly:

* :meth:`docx.document.Document.add_paragraph` accepts ``bind_to=``
  which feeds the bound record into this resolver.
* :meth:`docx.document.Document.bind` rebinds a different record and
  marks the document dirty so the next save re-resolves.

Test helpers :func:`render`, :func:`has_token`, :func:`get_token_source`
and :func:`reseat_token_source` are exposed for unit-testing the
substitution rules.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import datetime as _dt
import re
from typing import TYPE_CHECKING, Any, Mapping, Optional, Union

from lxml import etree

from docx.oxml.ns import nsmap, qn

if TYPE_CHECKING:
    from docx.document import Document


# -- fork extension URI persisted in saved documents. Must not be
# -- changed; doing so silently strips token-source preservation from
# -- every previously-saved document. --
LFXBIND_NS = "https://loadfix.dev/docx/bind-tokens"

# -- qualified tag of the source-preservation marker child. --
_LFXBIND_SRC_TAG = qn("lfxbind:src")
_W_R_TAG = qn("w:r")
_W_T_TAG = qn("w:t")


# -- token pattern. Matches ``{name}`` or ``{name:fmt}`` where
# -- ``name`` is dotted (letters/digits/underscores) and the optional
# -- ``:fmt`` suffix is either a bare identifier or a single-quoted
# -- format string. Tight on purpose: ``{Foo bar}`` and ``{}`` are
# -- preserved literally, so brace-delimited prose in the document
# -- (e.g. code samples) is never silently swallowed. --
TOKEN_RE = re.compile(
    r"\{"
    r"(?P<name>[A-Za-z_][\w.]*)"
    r"(?::(?P<fmt>'[^']*'|[\w-]+))?"
    r"\}"
)


# -- short-name format aliases for ``{date:<alias>}``. ``short`` is
# -- intentionally locale-agnostic (matches ECMA-376's "short date"
# -- style closely enough for cross-platform output). --
_DATE_FORMATS = {
    "short": "%Y-%m-%d",
    "medium": "%b %-d, %Y",
    "long": "%B %-d, %Y",
    "iso": "%Y-%m-%d",
}


# -- minimal Babel-style -> strftime token map for the custom format
# -- string variant ``{date:'MMM d, yyyy'}``. Covers the widely-used
# -- subset; anything not in this table is passed through literally
# -- so e.g. punctuation, slashes, and ``"of"`` survive. --
_BABEL_TO_STRFTIME = (
    ("yyyy", "%Y"),
    ("yy", "%y"),
    ("MMMM", "%B"),
    ("MMM", "%b"),
    ("MM", "%m"),
    ("dd", "%d"),
    ("d", "%-d"),
    ("HH", "%H"),
    ("mm", "%M"),
    ("ss", "%S"),
)


def _babel_to_strftime(fmt: str) -> str:
    """Translate a Babel-style date format to a ``strftime`` format.

    The mapping is greedy left-to-right with the longest tokens first
    (``yyyy`` before ``yy``), so common patterns like ``"MMM d, yyyy"``
    convert cleanly to ``"%b %-d, %Y"`` without overlapping replacements.

    Unknown tokens pass through literally — the format ``"foo bar"``
    survives unchanged so quoted prose inside a date format is preserved.
    """
    out: list[str] = []
    i = 0
    n = len(fmt)
    while i < n:
        matched = False
        for token, replacement in _BABEL_TO_STRFTIME:
            if fmt.startswith(token, i):
                out.append(replacement)
                i += len(token)
                matched = True
                break
        if not matched:
            out.append(fmt[i])
            i += 1
    return "".join(out)


def _resolve_dotted(record: Any, path: str) -> Optional[Any]:
    """Walk ``path`` (dotted) on ``record`` returning the value or |None|.

    Supports both attribute access (``record.customer.name``) and
    item access (``record["customer"]["name"]``). Returns |None| when
    any segment is missing — callers leave the token literal in that
    case rather than raising, mirroring the pptx page-token rule.
    """
    cur: Any = record
    for segment in path.split("."):
        if cur is None:
            return None
        if isinstance(cur, Mapping):
            if segment in cur:
                cur = cur[segment]
                continue
            return None
        if hasattr(cur, segment):
            cur = getattr(cur, segment)
            continue
        return None
    return cur


def _resolve_date(fmt: Optional[str], today: _dt.date) -> str:
    """Resolve a ``{date:...}`` token against ``today``.

    ``fmt`` may be |None| (default ISO), a short alias from
    :data:`_DATE_FORMATS`, or a single-quoted Babel-style custom
    format (the leading/trailing single quotes are stripped before
    interpretation).
    """
    if fmt is None or fmt == "":
        return today.strftime(_DATE_FORMATS["iso"])
    if fmt.startswith("'") and fmt.endswith("'"):
        babel = fmt[1:-1]
        return today.strftime(_babel_to_strftime(babel))
    if fmt in _DATE_FORMATS:
        return today.strftime(_DATE_FORMATS[fmt])
    # -- unknown alias: treat as an explicit strftime spec --
    return today.strftime(fmt)


def _resolve_property(name: str, properties: Mapping[str, Any]) -> Optional[Any]:
    """Resolve ``{property:Name}`` against the ``properties`` mapping.

    ``properties`` is built by :func:`_build_property_map` and contains
    both core (Title, Author, Subject, Keywords, Comments, Category,
    Created, Modified, ...) and any custom-property names available
    on the document. Lookup is case-insensitive against the keys.
    """
    if name in properties:
        return properties[name]
    lower = name.lower()
    for key, value in properties.items():
        if key.lower() == lower:
            return value
    return None


def render(
    template: str,
    record: Any = None,
    *,
    properties: Optional[Mapping[str, Any]] = None,
    iteration: Optional[int] = None,
    today: Optional[_dt.date] = None,
) -> str:
    """Return ``template`` with every recognised ``{token}`` resolved.

    Tokens whose dotted-path lookup against ``record`` fails (or
    whose record is |None| for record-binding tokens) are left
    literal — a stray ``{customer.name}`` in user prose stays a
    literal ``{customer.name}`` rather than raising. This mirrors
    the pptx page-token rule and keeps the substitution from
    surprising callers who never opted into binding.

    ``properties`` is consulted for ``{property:Name}`` tokens,
    ``iteration`` for ``{i}``, and ``today`` (default
    :meth:`datetime.date.today`) for ``{date:...}``.

    Values are coerced via :class:`str`; |None| renders as the
    empty string so an unset core-property reference doesn't leak
    a stray ``"None"`` into the document.
    """
    today = today if today is not None else _dt.date.today()
    properties = properties or {}

    def _replace(match: "re.Match[str]") -> str:
        name = match.group("name")
        fmt = match.group("fmt")
        # -- {date:...} family --
        if name == "date":
            return _resolve_date(fmt, today)
        # -- {i} — iteration index (mail-merge) --
        if name == "i":
            if iteration is None:
                return match.group(0)
            return str(iteration)
        # -- {property:Name} --
        if name == "property":
            if fmt is None:
                return match.group(0)
            key = fmt[1:-1] if fmt.startswith("'") and fmt.endswith("'") else fmt
            value = _resolve_property(key, properties)
            if value is None:
                return match.group(0)
            return str(value)
        # -- record-bound dotted path --
        if record is None:
            return match.group(0)
        value = _resolve_dotted(record, name)
        if value is None:
            return match.group(0)
        return str(value)

    return TOKEN_RE.sub(_replace, template)


def has_token(text: str) -> bool:
    """Return |True| when ``text`` carries at least one recognised token shape."""
    return TOKEN_RE.search(text) is not None


def _read_source_marker(parent: etree._Element) -> Optional[str]:
    """Return the persisted token source string when one is attached, else |None|."""
    src = parent.find(_LFXBIND_SRC_TAG)
    if src is None:
        return None
    return src.text or ""


def _write_source_marker(parent: etree._Element, source: str) -> None:
    """Append (or update) a ``<lfxbind:src>`` child preserving ``source``.

    Any existing marker is overwritten in place so we don't accumulate
    duplicates across repeated saves.
    """
    src = parent.find(_LFXBIND_SRC_TAG)
    if src is None:
        src = etree.SubElement(  # pyright: ignore[reportUnknownMemberType]
            parent,
            _LFXBIND_SRC_TAG,
            nsmap={"lfxbind": nsmap["lfxbind"]},
        )
    src.text = source


def reseat_token_source(carrier: etree._Element, source: str) -> None:
    """Stamp ``source`` as the token-source marker on ``carrier``.

    Public test helper for asserting that a manually-constructed run
    will resolve correctly on the next save. End-user code does not
    need to call this; passing ``bind_to=`` to ``add_paragraph`` and
    letting :func:`apply_bind_tokens` run on save is sufficient.
    """
    _write_source_marker(carrier, source)


def get_token_source(carrier: etree._Element) -> Optional[str]:
    """Return the persisted token source for ``carrier``, or |None|.

    Public test helper to assert that the round-trip preservation
    marker landed where expected.
    """
    return _read_source_marker(carrier)


def _build_property_map(document: "Document") -> dict[str, Any]:
    """Build a flat ``name -> value`` map of every property addressable
    via ``{property:Name}``.

    Includes both Dublin-Core core properties (with their canonical
    capitalised names: ``Title``, ``Author``, ``Subject``, ...) and
    every custom property defined on the document. Custom-property
    lookup falls through to core when both define the same name.
    """
    out: dict[str, Any] = {}
    try:
        cp = document.core_properties
    except Exception:  # pragma: no cover - defensive
        cp = None
    if cp is not None:
        for attr_name, key in (
            ("title", "Title"),
            ("author", "Author"),
            ("subject", "Subject"),
            ("keywords", "Keywords"),
            ("comments", "Comments"),
            ("category", "Category"),
            ("content_status", "ContentStatus"),
            ("identifier", "Identifier"),
            ("language", "Language"),
            ("last_modified_by", "LastModifiedBy"),
            ("revision", "Revision"),
            ("version", "Version"),
        ):
            try:
                value = getattr(cp, attr_name, None)
            except Exception:  # pragma: no cover - defensive
                value = None
            if value not in (None, ""):
                out[key] = value
    try:
        custom = document.custom_properties
    except Exception:  # pragma: no cover - defensive
        custom = None
    if custom is not None:
        try:
            names = list(custom)
        except Exception:  # pragma: no cover - defensive
            names = []
        for name in names:
            try:
                out[str(name)] = custom[name]
            except Exception:  # pragma: no cover - defensive
                continue
    return out


def _iter_run_carriers(root: etree._Element):
    """Yield every ``<w:r>`` element under ``root`` in document order."""
    yield from root.iter(_W_R_TAG)


def _carrier_t(carrier: etree._Element) -> Optional[etree._Element]:
    """Return the carrier's first ``<w:t>`` child, or |None|.

    Multi-``w:t`` runs collapse into a single token-resolution slot —
    we read the concatenated text, resolve, and write the result back
    onto the first ``w:t`` while clearing any siblings. Authors who
    set the run text via the public ``run.text = ...`` setter never
    end up with multi-``w:t`` runs in the first place, so this only
    matters for runs constructed with raw OOXML.
    """
    return carrier.find(_W_T_TAG)


def _read_carrier_text(carrier: etree._Element) -> str:
    """Concatenate every ``<w:t>`` child of ``carrier`` into a single string."""
    parts: list[str] = []
    for t in carrier.findall(_W_T_TAG):
        parts.append(t.text or "")
    return "".join(parts)


def _write_carrier_text(carrier: etree._Element, text: str) -> None:
    """Write ``text`` to the carrier's first ``<w:t>``, dropping siblings."""
    ts = carrier.findall(_W_T_TAG)
    if not ts:
        # -- create a fresh w:t at the end of the run --
        t = etree.SubElement(  # pyright: ignore[reportUnknownMemberType]
            carrier, _W_T_TAG
        )
        t.text = text
        return
    ts[0].text = text
    # -- preserve leading/trailing whitespace --
    if text != text.strip():
        ts[0].set(qn("xml:space"), "preserve")
    for extra in ts[1:]:
        carrier.remove(extra)


def _document_root(document: "Document") -> Optional[etree._Element]:
    """Return the ``<w:document>`` element of ``document``."""
    try:
        return document.element
    except Exception:  # pragma: no cover - defensive
        return None


# -- the bound-record / iteration is stored on the Document via a
# -- private attribute set by ``Document.bind`` and ``add_paragraph``. --
_BOUND_RECORD_ATTR = "_bind_token_record"
_BOUND_ITERATION_ATTR = "_bind_token_iteration"


def get_bound_record(document: "Document") -> Any:
    """Return the record currently bound to ``document``, or |None|."""
    return getattr(document, _BOUND_RECORD_ATTR, None)


def set_bound_record(
    document: "Document",
    record: Any,
    iteration: Optional[int] = None,
) -> None:
    """Bind ``record`` (and optional ``iteration``) to ``document``.

    A subsequent :meth:`Document.save` runs :func:`apply_bind_tokens`
    against the new record so every token-bearing run is re-resolved.
    """
    setattr(document, _BOUND_RECORD_ATTR, record)
    if iteration is not None:
        setattr(document, _BOUND_ITERATION_ATTR, iteration)


def apply_bind_tokens(
    document: "Document",
    record: Any = None,
    iteration: Optional[int] = None,
    today: Optional[_dt.date] = None,
) -> None:
    """Resolve ``{token}`` strings in every text run of ``document``.

    Walks every ``<w:r>`` under the document root once; for every run
    whose displayed text either currently carries a token *or* carries
    a persisted ``<lfxbind:src>`` marker, re-resolves the source against
    the (record, properties, iteration) context and writes the resolved
    string back onto the run's ``<w:t>``. The marker is stamped /
    refreshed after every resolution so the round-trip is stable.

    No-op for runs that neither carry a live token nor a previously-
    persisted source marker. Idempotent: calling twice in a row
    against an unchanged document produces identical XML.

    Called from :meth:`Document.save` immediately before delegating
    to the part-level save. Failing this step must not block save —
    the function's contract is best-effort.
    """
    root = _document_root(document)
    if root is None:
        return
    if record is None:
        record = get_bound_record(document)
    if iteration is None:
        iteration = getattr(document, _BOUND_ITERATION_ATTR, None)
    properties = _build_property_map(document)

    for carrier in _iter_run_carriers(root):
        if _carrier_t(carrier) is None and _read_source_marker(carrier) is None:
            continue
        current = _read_carrier_text(carrier)
        persisted = _read_source_marker(carrier)
        # -- choose the token source. Persisted marker wins so a
        # -- previously-resolved literal (e.g. "Dear Acme,") gets
        # -- resolved against the source string rather than treated
        # -- as inert literal text. --
        source = persisted if persisted is not None else current
        if not has_token(source):
            continue
        resolved = render(
            source,
            record=record,
            properties=properties,
            iteration=iteration,
            today=today,
        )
        if resolved != current:
            _write_carrier_text(carrier, resolved)
        # -- stamp / refresh the source marker so re-saves pick up
        # -- the live source instead of the resolved literal. --
        _write_source_marker(carrier, source)


__all__ = [
    "LFXBIND_NS",
    "TOKEN_RE",
    "apply_bind_tokens",
    "get_bound_record",
    "get_token_source",
    "has_token",
    "render",
    "reseat_token_source",
    "set_bound_record",
]
