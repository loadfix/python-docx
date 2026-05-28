"""Best-effort recovery for damaged ``.docx`` packages (issue #92).

``Document.repair(path, strategy='best-effort')`` returns a ``(Document,
RepairReport)`` tuple. The :class:`RepairReport` records every fix that was
applied, every issue that could not be repaired, and every part that had to
be dropped from the package.

Three strategies are supported:

``'best-effort'`` (default)
    Drop unparseable parts, fix common XML defects (orphan bookmark tags,
    invalid encoding declarations, NUL/control characters in element
    content), and continue. Returns even if the document body part itself
    had to be reconstructed from a stub.

``'strict'``
    Behaves exactly like the existing :func:`docx.Document` factory —
    raises on the first damaged part. Provided so the same call site can
    opt in or out of recovery without a separate code path. The report's
    fix lists are always empty in this mode.

``'truncate'``
    Preserve everything that successfully parses, drop everything from
    the first damaged byte onwards. Useful for documents whose tail was
    lost to a transfer error; the head is kept intact while the tail is
    replaced with a minimal stub.

The :class:`RepairReport` value object is intentionally simple — three
``list[str]`` fields (:attr:`~RepairReport.repaired`,
:attr:`~RepairReport.unrecoverable`, :attr:`~RepairReport.parts_dropped`)
plus :attr:`~RepairReport.strategy`. Callers that want richer telemetry
can wrap or subclass freely.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import io
import os
import re
import zipfile
from dataclasses import dataclass, field
from typing import IO, TYPE_CHECKING, List, Tuple, Union, cast

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.exceptions import (
    MissingDocxFileError,
    NotADocxError,
    PackageNotFoundError,
)

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject


__all__ = [
    "RepairError",
    "RepairReport",
    "RepairStrategy",
    "repair",
]


RepairStrategy = str  # one of: "best-effort", "strict", "truncate"


_VALID_STRATEGIES = frozenset(("best-effort", "strict", "truncate"))


class RepairError(Exception):
    """Raised by :func:`repair` when no |Document| can be produced.

    Distinct from :class:`docx.opc.exceptions.PackageNotFoundError` so callers
    can opt to fall through to a separate code path when even best-effort
    recovery fails (for example, when the input is not a zip at all).
    """


@dataclass
class RepairReport:
    """Outcome of a :func:`repair` call.

    Attributes
    ----------
    strategy:
        One of ``"best-effort"`` / ``"strict"`` / ``"truncate"``.
    repaired:
        Human-readable descriptions of fixes the repair pass applied. Each
        entry is prefixed with the affected part name (or ``"package:"`` for
        package-level fixes) followed by a colon and the change.
    unrecoverable:
        Issues encountered that could not be fixed but did not prevent a
        |Document| from being returned. Empty when all damage was repaired
        cleanly.
    parts_dropped:
        Part names (``str``) that had to be discarded entirely because they
        were unparseable, encoded with an unknown content type, or otherwise
        beyond rescue. The dropped parts are removed from the package
        relationship graph so the surviving document still saves cleanly.
    """

    strategy: RepairStrategy
    repaired: List[str] = field(default_factory=list)
    unrecoverable: List[str] = field(default_factory=list)
    parts_dropped: List[str] = field(default_factory=list)

    @property
    def is_clean(self) -> bool:
        """``True`` when every issue was repaired and nothing had to be dropped."""
        return not (self.repaired or self.unrecoverable or self.parts_dropped)

    def __bool__(self) -> bool:  # pragma: no cover - cosmetic
        # -- Truthy when *any* repair activity occurred. --
        return bool(self.repaired or self.unrecoverable or self.parts_dropped)


# ---------------------------------------------------------------------------
# XML-level fixers.
#
# Each fixer takes a ``(partname, blob)`` pair and returns a possibly-mutated
# blob plus a list of human-readable fix descriptions. They are intentionally
# string-level and conservative: they fix the specific defects the issue
# tracker has seen in real corruption cases without touching well-formed
# documents.
# ---------------------------------------------------------------------------


_ENCODING_DECL_RE = re.compile(rb'(<\?xml[^?]*?encoding=)(["\'])([^"\']*?)\2', re.I)

# -- ``\xfffe``/``\xffff`` and the C0 control range minus tab/LF/CR are illegal
# -- in XML 1.0 element content; lxml's strict parser rejects them, the
# -- recovering parser drops the surrounding chunk silently. Strip them up
# -- front so we keep the surrounding text. The set was lifted from the
# -- W3C XML 1.0 §2.2 "Char" production. --
_ILLEGAL_XML_BYTES = bytes(
    b for b in range(0x20) if b not in (0x09, 0x0A, 0x0D)
)
_ILLEGAL_XML_RE = re.compile(b"[" + re.escape(_ILLEGAL_XML_BYTES) + b"]")


def _fix_encoding_declaration(blob: bytes) -> Tuple[bytes, List[str]]:
    """Normalise the XML prolog's ``encoding=`` attribute.

    OOXML mandates UTF-8; some third-party tools emit the BOM-less
    ``<?xml version='1.0' encoding='utf-16'?>`` which lxml then rejects with
    "encoding declaration ... but actual encoding is utf-8". Force the
    declared encoding to UTF-8 when the body decodes cleanly as UTF-8.
    """
    notes: List[str] = []
    match = _ENCODING_DECL_RE.search(blob[:200])
    if match is None:
        return blob, notes
    declared = match.group(3).lower()
    if declared in (b"utf-8", b"utf8"):
        return blob, notes
    # -- only rewrite when the body is actually decodable as UTF-8 --
    try:
        blob.decode("utf-8")
    except UnicodeDecodeError:
        return blob, notes
    new_blob = blob[: match.start(3)] + b"UTF-8" + blob[match.end(3):]
    notes.append(
        f"corrected XML encoding declaration {declared.decode('ascii', 'replace')!r} -> 'UTF-8'"
    )
    return new_blob, notes


def _strip_illegal_xml_bytes(blob: bytes) -> Tuple[bytes, List[str]]:
    """Remove control characters that XML 1.0 forbids in element content."""
    if not _ILLEGAL_XML_RE.search(blob):
        return blob, []
    cleaned = _ILLEGAL_XML_RE.sub(b"", blob)
    n_removed = len(blob) - len(cleaned)
    return cleaned, [f"stripped {n_removed} illegal XML control byte(s)"]


_BOOKMARK_START_RE = re.compile(
    rb"<w:bookmarkStart\b[^>]*\bw:id=\"(?P<id>\d+)\"[^>]*/?>"
)
_BOOKMARK_END_RE = re.compile(
    rb"<w:bookmarkEnd\b[^>]*\bw:id=\"(?P<id>\d+)\"[^>]*/?>"
)


def _close_orphan_bookmarks(blob: bytes) -> Tuple[bytes, List[str]]:
    """Append a ``w:bookmarkEnd`` for every unmatched ``w:bookmarkStart``.

    Word treats an unmatched ``bookmarkStart`` as a parse-time warning and
    discards the bookmark on save; the lxml parser tolerates them but the
    file fails ECMA-376 spec validation. We append a synthesised
    ``bookmarkEnd`` immediately before ``</w:body>`` (or end of buffer) for
    every orphan id, in document order.
    """
    starts = {m.group("id") for m in _BOOKMARK_START_RE.finditer(blob)}
    ends = {m.group("id") for m in _BOOKMARK_END_RE.finditer(blob)}
    orphans = sorted(starts - ends, key=lambda s: int(s))
    if not orphans:
        return blob, []
    closer = b"".join(
        b'<w:bookmarkEnd w:id="' + bid + b'"/>' for bid in orphans
    )
    body_close = blob.rfind(b"</w:body>")
    if body_close == -1:
        new_blob = blob + closer
    else:
        new_blob = blob[:body_close] + closer + blob[body_close:]
    return new_blob, [
        f"closed orphan w:bookmarkStart id={bid.decode('ascii')}" for bid in orphans
    ]


_XML_FIXERS = (
    _fix_encoding_declaration,
    _strip_illegal_xml_bytes,
    _close_orphan_bookmarks,
)


def _apply_xml_fixers(partname: str, blob: bytes) -> Tuple[bytes, List[str]]:
    """Run every registered fixer in order; return the patched blob + notes."""
    notes: List[str] = []
    for fixer in _XML_FIXERS:
        blob, fix_notes = fixer(blob)
        for note in fix_notes:
            notes.append(f"{partname}: {note}")
    return blob, notes


# ---------------------------------------------------------------------------
# Truncated-zip reconstruction.
#
# `zipfile.ZipFile` reads the central directory from the trailer of the
# archive. When the trailer is missing/truncated, opening fails with
# `BadZipFile`. We work around this by scanning the file for local file
# headers (`PK\x03\x04`) and synthesising a fresh archive from any complete
# entries.
# ---------------------------------------------------------------------------

_PK_LOCAL_HEADER = b"PK\x03\x04"
_PK_CENTRAL_DIR = b"PK\x01\x02"
_PK_EOCD = b"PK\x05\x06"


def _is_zip_truncated(blob: bytes) -> bool:
    """Heuristic check for zip files whose central directory is missing.

    The end-of-central-directory record is the last 22+ bytes of a complete
    archive. If we can locate at least one local-file header but no EOCD,
    the archive is truncated.
    """
    if _PK_LOCAL_HEADER not in blob[:4096]:
        return False
    return _PK_EOCD not in blob


def _try_reconstruct_truncated_zip(
    blob: bytes, report: RepairReport
) -> Union[bytes, None]:
    """Salvage entries from a zip whose trailer is missing.

    Walks the byte stream looking for ``PK\\x03\\x04`` local-file headers
    and re-archives every entry whose payload reads cleanly. Returns the
    rebuilt archive bytes, or |None| if no entries could be recovered.
    """
    src = io.BytesIO(blob)
    out_buf = io.BytesIO()
    recovered = 0
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zf:
        for header_pos in _iter_local_headers(blob):
            entry = _read_local_entry(src, header_pos, len(blob))
            if entry is None:
                continue
            name, data = entry
            try:
                out_zf.writestr(name, data)
            except (ValueError, OSError):
                continue
            recovered += 1
    if recovered == 0:
        return None
    report.repaired.append(
        f"package: reconstructed truncated zip — recovered {recovered} entries"
    )
    return out_buf.getvalue()


def _iter_local_headers(blob: bytes):
    """Yield byte offsets of every ``PK\\x03\\x04`` signature in `blob`."""
    pos = 0
    while True:
        pos = blob.find(_PK_LOCAL_HEADER, pos)
        if pos == -1:
            return
        yield pos
        pos += 4


def _read_local_entry(
    src: IO[bytes], header_pos: int, blob_len: int
) -> Union[Tuple[str, bytes], None]:
    """Decode the local-file-header at `header_pos` and return ``(name, data)``.

    Returns |None| if the header is malformed or the entry's payload runs
    past the end of the buffer (truncated mid-entry).
    """
    import struct

    src.seek(header_pos)
    header = src.read(30)
    if len(header) < 30 or header[:4] != _PK_LOCAL_HEADER:
        return None
    (
        _sig,
        _ver,
        _flags,
        method,
        _mtime,
        _mdate,
        _crc,
        comp_size,
        uncomp_size,
        name_len,
        extra_len,
    ) = struct.unpack("<IHHHHHIIIHH", header)
    name_bytes = src.read(name_len)
    if len(name_bytes) != name_len:
        return None
    src.read(extra_len)
    if comp_size == 0xFFFFFFFF or uncomp_size == 0xFFFFFFFF:
        # -- ZIP64 extra field — not worth the complexity for repair --
        return None
    if comp_size == 0 and uncomp_size == 0 and name_bytes.endswith(b"/"):
        # -- directory entry; harmless to skip --
        return None
    payload = src.read(comp_size)
    if len(payload) != comp_size:
        return None
    if header_pos + 30 + name_len + extra_len + comp_size > blob_len:
        return None
    if method == zipfile.ZIP_STORED:
        data = payload
    elif method == zipfile.ZIP_DEFLATED:
        import zlib

        try:
            data = zlib.decompress(payload, -zlib.MAX_WBITS)
        except zlib.error:
            return None
    else:
        return None
    try:
        name = name_bytes.decode("utf-8")
    except UnicodeDecodeError:
        try:
            name = name_bytes.decode("cp437")
        except UnicodeDecodeError:
            return None
    return name, data


# ---------------------------------------------------------------------------
# Per-part XML preflight (best-effort and truncate).
# ---------------------------------------------------------------------------


def _preflight_zip_parts(
    blob: bytes, strategy: RepairStrategy, report: RepairReport
) -> bytes:
    """Apply per-entry XML fixes to every member of `blob` (a zip blob).

    Reads the input zip, runs each XML entry through :func:`_apply_xml_fixers`,
    re-emits the archive. Entirely non-XML members (images, fonts, embedded
    objects) pass through unchanged. Members whose XML fails to parse even
    after fixing are dropped under ``best-effort`` and ``truncate`` strategies
    and recorded on :attr:`RepairReport.parts_dropped`.

    A second pass over the rebuilt archive prunes ``.rels`` entries whose
    ``Target=`` points at a part that is no longer present (either dropped
    upstream or never existed in the first place).
    """
    in_buf = io.BytesIO(blob)
    out_buf = io.BytesIO()
    truncated_after: Union[str, None] = None
    surviving_members: List[str] = []
    rels_entries: List[Tuple[zipfile.ZipInfo, bytes]] = []

    with zipfile.ZipFile(in_buf, "r") as zin, zipfile.ZipFile(
        out_buf, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            name = info.filename
            try:
                data = zin.read(name)
            except (zipfile.BadZipFile, OSError) as exc:
                report.parts_dropped.append(f"/{name}: zip read failed ({exc})")
                if strategy == "truncate":
                    truncated_after = name
                    break
                continue
            if _looks_like_xml(name, data):
                fixed, notes = _apply_xml_fixers("/" + name, data)
                report.repaired.extend(notes)
                # -- validate parseability with the strict parser; if still
                # -- broken, fall back to recovery; if even recovery yields
                # -- nothing, drop the part. --
                if _xml_parses(fixed):
                    data = fixed
                else:
                    if strategy == "truncate":
                        truncated_after = name
                        report.parts_dropped.append(
                            f"/{name}: malformed XML — truncated here"
                        )
                        break
                    if _xml_recovers(fixed):
                        # -- leave as-is; the recovery parser will pick it up
                        # -- when the package loads in recover=True mode. --
                        data = fixed
                        report.repaired.append(
                            f"/{name}: malformed XML — left for recovery parser"
                        )
                    else:
                        report.parts_dropped.append(
                            f"/{name}: unparseable XML — dropped"
                        )
                        # -- drop this part entirely; do not write it to zout --
                        continue
            # -- defer .rels writes until we know which targets survived --
            if name.endswith(".rels"):
                rels_entries.append((info, data))
                continue
            zout.writestr(info, data)
            surviving_members.append(name)

        # -- second pass: emit .rels with unresolvable targets stripped --
        present = set(surviving_members)
        for info, data in rels_entries:
            cleaned, dropped = _strip_dangling_rels(info.filename, data, present)
            for note in dropped:
                report.repaired.append(note)
            zout.writestr(info, cleaned)

    if truncated_after is not None:
        report.repaired.append(
            f"package: truncated at /{truncated_after}; later entries discarded"
        )
    return out_buf.getvalue()


_RELATIONSHIP_RE = re.compile(
    rb'<Relationship\b[^>]*>',
    re.I,
)
_TARGET_ATTR_RE = re.compile(rb'\bTarget=(["\'])(?P<target>[^"\']*)\1', re.I)
_TARGET_MODE_ATTR_RE = re.compile(rb'\bTargetMode=(["\'])(?P<mode>[^"\']*)\1', re.I)
_ID_ATTR_RE = re.compile(rb'\bId=(["\'])(?P<rid>[^"\']*)\1', re.I)


def _strip_dangling_rels(
    rels_partname: str, blob: bytes, present_members: set
) -> Tuple[bytes, List[str]]:
    """Drop ``<Relationship>`` rows whose internal target is not in `present_members`.

    `rels_partname` is the zip member name of the `.rels` part itself
    (e.g. ``"word/_rels/document.xml.rels"``). External relationships
    (``TargetMode="External"``) and in-document fragments (``#anchor``)
    are preserved verbatim.
    """
    base_dir = _rels_base_dir(rels_partname)
    notes: List[str] = []
    out = bytearray()
    last = 0
    for match in _RELATIONSHIP_RE.finditer(blob):
        target_match = _TARGET_ATTR_RE.search(match.group(0))
        mode_match = _TARGET_MODE_ATTR_RE.search(match.group(0))
        target = target_match.group("target") if target_match else b""
        mode = mode_match.group("mode") if mode_match else b""
        if mode == b"External" or target.startswith(b"#") or not target:
            continue
        # -- resolve relative target against base_dir; collapse `../` segments --
        resolved = _resolve_target(base_dir, target.decode("utf-8", "replace"))
        if resolved in present_members:
            continue
        # -- drop this Relationship row --
        rid_match = _ID_ATTR_RE.search(match.group(0))
        rid = (
            rid_match.group("rid").decode("ascii", "replace")
            if rid_match
            else "<unknown>"
        )
        out.extend(blob[last : match.start()])
        last = match.end()
        notes.append(
            f"/{rels_partname}: dropped dangling rel {rid} -> {target.decode('utf-8', 'replace')}"
        )
    out.extend(blob[last:])
    return bytes(out), notes


def _rels_base_dir(rels_partname: str) -> str:
    """Return the directory whose rels entry is `rels_partname`.

    ``"word/_rels/document.xml.rels"`` → ``"word/"``;
    ``"_rels/.rels"`` → ``""`` (package root).
    """
    parent = os.path.dirname(rels_partname)
    if parent.endswith("/_rels"):
        parent = parent[: -len("/_rels")]
    elif parent.endswith("_rels"):
        parent = parent[: -len("_rels")]
    if parent and not parent.endswith("/"):
        parent += "/"
    return parent


def _resolve_target(base_dir: str, target: str) -> str:
    """Resolve a relative ``Target=`` URI against `base_dir`.

    Mirrors :func:`PackURI.from_rel_ref` (joins, normalises, strips any
    leading ``/``). Kept intentionally small — we only need a string the
    surviving-members set can be checked against.
    """
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base_dir + target).split("/")
    stack: List[str] = []
    for piece in parts:
        if piece in ("", "."):
            continue
        if piece == "..":
            if stack:
                stack.pop()
            continue
        stack.append(piece)
    return "/".join(stack)


_XML_EXTS = frozenset({".xml", ".rels"})


def _looks_like_xml(name: str, blob: bytes) -> bool:
    ext = os.path.splitext(name)[1].lower()
    if ext in _XML_EXTS:
        return True
    return blob[:5] == b"<?xml"


def _xml_parses(blob: bytes) -> bool:
    from lxml import etree

    try:
        etree.fromstring(blob)
    except etree.XMLSyntaxError:
        return False
    except ValueError:
        return False
    return True


def _xml_recovers(blob: bytes) -> bool:
    from lxml import etree

    parser = etree.XMLParser(recover=True, resolve_entities=False, no_network=True)
    try:
        result = etree.fromstring(blob, parser)
    except etree.XMLSyntaxError:
        return False
    except ValueError:
        return False
    return result is not None


# ---------------------------------------------------------------------------
# Public entry point.
# ---------------------------------------------------------------------------


def repair(
    docx: Union[str, "os.PathLike[str]", IO[bytes]],
    strategy: RepairStrategy = "best-effort",
) -> Tuple["DocumentObject", RepairReport]:
    """Open `docx`, applying recovery fixes per `strategy`.

    Returns a ``(document, report)`` tuple. The document is always returned
    when `strategy` is ``'best-effort'`` or ``'truncate'`` and at least one
    part of the input could be salvaged; otherwise :class:`RepairError` is
    raised.

    `strategy` selects the recovery posture:

    ``'best-effort'``
        Apply every fix the repair pass knows about, drop unparseable
        parts, and continue. Default.
    ``'strict'``
        Parse with the strict XML parser; raise on the first defect.
        Equivalent to calling :func:`docx.Document` without ``recover=``,
        with the side-benefit of a typed return shape.
    ``'truncate'``
        Keep everything that parses; the moment a part fails, drop it
        and every subsequent part. Useful for documents whose tail was
        lost mid-transfer.

    Raises
    ------
    ValueError
        If `strategy` is not one of the three supported values.
    RepairError
        If recovery cannot produce a usable |Document| (no zip entries
        recovered, or the input is missing entirely).
    """
    if strategy not in _VALID_STRATEGIES:
        raise ValueError(
            f"unknown repair strategy {strategy!r}; "
            f"expected one of {sorted(_VALID_STRATEGIES)!r}"
        )

    # -- import lazily so a fresh `docx` import doesn't pull repair's
    # -- dependency surface (lxml etree fixers, zipfile) on the hot path. --
    from docx.api import Document as _DocumentFactory

    report = RepairReport(strategy=strategy)

    if strategy == "strict":
        document = _DocumentFactory(docx)
        return document, report

    blob = _read_input_blob(docx)

    # -- truncated-zip detection up front so the per-part preflight has a
    # -- proper archive to walk over. --
    if _is_zip_truncated(blob):
        rebuilt = _try_reconstruct_truncated_zip(blob, report)
        if rebuilt is None:
            raise RepairError(
                "input zip is truncated and no entries could be recovered"
            )
        blob = rebuilt

    # -- second pass: best-effort/truncate per-part XML fixes. --
    try:
        blob = _preflight_zip_parts(blob, strategy, report)
    except zipfile.BadZipFile as exc:
        raise RepairError(f"input is not a valid zip archive: {exc}") from exc

    # -- final load via the regular factory in recover mode so any residual
    # -- defects (orphan rel targets, missing required parts, recoverable
    # -- XML) flow through the documented `recover=True` path. --
    stream = io.BytesIO(blob)
    try:
        document = _DocumentFactory(stream, recover=True)
    except (PackageNotFoundError, NotADocxError, MissingDocxFileError) as exc:
        raise RepairError(f"package could not be loaded after repair: {exc}") from exc
    except Exception as exc:  # pragma: no cover - defensive
        raise RepairError(f"unexpected failure loading repaired package: {exc}") from exc

    # -- propagate residual lxml warnings (rels dropped by the existing
    # -- pkgreader, recoverable XML, etc.) into the report. --
    for warning in document.recovery_warnings:
        report.repaired.append(f"package: recovery parser - {warning}")

    return cast("DocumentObject", document), report


def _read_input_blob(docx: Union[str, "os.PathLike[str]", IO[bytes]]) -> bytes:
    """Read `docx` (path or stream) into a bytes blob.

    Accepts the same input shapes as :func:`docx.Document` so callers can
    use either signature interchangeably. Raises :class:`RepairError` when
    the input is missing or is not a readable byte stream.
    """
    if isinstance(docx, (str, os.PathLike)):
        path = os.fspath(docx)
        if not os.path.exists(path):
            raise RepairError(f"no file at {path!r}")
        with open(path, "rb") as f:
            return f.read()
    # -- assume a file-like object --
    try:
        if hasattr(docx, "seek"):
            docx.seek(0)
        return docx.read()
    except (AttributeError, OSError) as exc:
        raise RepairError(f"input stream is not readable: {exc}") from exc
