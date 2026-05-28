# pyright: reportPrivateUsage=false

"""Bounded-memory streaming reader for very large ``.docx`` packages.

The default :func:`docx.Document` factory parses ``word/document.xml`` into a
single ``lxml`` element tree at load time. That works well for typical
documents but the peak memory cost scales linearly with the body size — a
100 MB ``document.xml`` can occupy several hundred MB of RAM once promoted
to the in-memory tree.

:func:`Document.stream` is the read-only alternative built on
:func:`lxml.etree.iterparse`. It yields :class:`docx.text.paragraph.Paragraph`
objects one at a time and **clears each element after the consumer has
seen it**, so the peak working set stays bounded regardless of body
length::

    from docx import Document

    with Document.stream("big.docx") as doc:
        for para in doc.paragraphs:
            if para.style.name == "Heading 1":
                print(para.text)
            # paragraph element is dropped from memory once the loop body
            # finishes — holding references to ``para`` past iteration
            # breaks because the underlying CT_P has been cleared.

Decision tree — when to use which loader:

* Document body fits comfortably in memory (a few MB or less), or you
  need random access (``doc.paragraphs[7]``), mutation
  (``doc.add_paragraph(...)``), or :meth:`Document.save` →
  :func:`docx.Document` (eager).
* Body is hundreds of MB, you only need a single forward pass, and the
  workload is read-only (filtering, extraction, search) →
  :meth:`Document.stream`.
* Mixed needs (large body + occasional mutation) → load eagerly with
  ``Document(huge_tree=True)`` and accept the memory cost; streaming
  cannot mutate.

Trade-offs vs the eager :func:`docx.Document`:

* :attr:`StreamingDocument.paragraphs` is a **generator**, not a
  ``Sequence`` — single forward pass, no ``len()``, no indexing.
* :attr:`StreamingDocument.tables` is likewise a generator and yields
  only **top-level body tables** (matching :attr:`Document.tables`); a
  table nested inside a cell is reachable through its enclosing
  :class:`Table.rows[i].cells[j].tables`, but iterating those still
  requires fully parsing the parent table — streaming does not extend
  into nested stories.
* :meth:`StreamingDocument.save` raises
  :class:`StreamingNotMutableError`. Re-load eagerly if you need to
  write.
* Headers, footers, sections, and styles remain eager — those parts
  are small (kilobytes, not megabytes) so there is no memory benefit
  to streaming them. Sections are populated by a single targeted
  ``iterparse`` pass that retains only ``w:sectPr`` elements; the rest
  of the body is discarded as it streams past.

The streaming reader holds onto the source ``.docx`` bytes (or a path)
for the lifetime of the :class:`StreamingDocument` so multiple
disjoint generators (``paragraphs``, ``tables``, the section scan) can
re-read ``word/document.xml`` independently. Closing the context
manager releases that handle.

.. versionadded:: 2026.05.13
"""

from __future__ import annotations

import io
import os
import zipfile
from typing import IO, TYPE_CHECKING, Iterator, Optional, Union, cast

from lxml import etree

from docx.exceptions import PythonDocxError
from docx.oxml.ns import qn
from docx.oxml.parser import element_class_lookup

if TYPE_CHECKING:
    from docx.oxml.document import CT_Document
    from docx.oxml.section import CT_SectPr
    from docx.parts.document import DocumentPart
    from docx.section import Sections, _Footer, _Header
    from docx.styles.styles import Styles
    from docx.table import Table
    from docx.text.paragraph import Paragraph


_W_BODY = qn("w:body")
_W_P = qn("w:p")
_W_TBL = qn("w:tbl")
_W_SECTPR = qn("w:sectPr")
_W_DOCUMENT = qn("w:document")

_DOCUMENT_PART_NAME = "word/document.xml"

# -- minimal stub used to seed the eagerly-parsed document part. The
# -- streaming reader rewrites the document.xml entry to this stub before
# -- delegating to ``Package.open``, so the package builds a valid
# -- ``CT_Document`` without ever materialising the real (potentially
# -- multi-GB) body. The original bytes are kept in ``_zip_bytes`` for the
# -- generator passes to re-read.
_BODY_STUB = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:document xmlns:w="http://schemas.openxmlformats.org/'
    b'wordprocessingml/2006/main"><w:body/></w:document>'
)


class StreamingNotMutableError(PythonDocxError):
    """Raised on mutation attempts against a :class:`StreamingDocument`.

    The streaming reader is intentionally read-only — it never holds the
    full body tree, so methods that would emit a saved package
    (:meth:`StreamingDocument.save`) or mutate body content cannot
    function. Re-open the document via :func:`docx.Document` for an
    eager load if you need to write.

    .. versionadded:: 2026.05.13
    """


def _read_source_bytes(
    source: Union[str, "os.PathLike[str]", IO[bytes], bytes],
) -> bytes:
    """Return the raw ``.docx`` bytes for `source`.

    `source` may be a filesystem path, an :class:`os.PathLike`, a binary
    file-like object, or a ``bytes`` payload. The bytes are read fully
    into memory once so subsequent passes (paragraphs, tables, section
    scan) can re-open the zip without re-reading the underlying file
    descriptor — a single 100 MB ``.docx`` is dwarfed by even a modest
    Python process working set, and the alternative (re-seeking
    arbitrary file-likes) is fragile.
    """
    if isinstance(source, bytes):
        return source
    if isinstance(source, (str, os.PathLike)):
        with open(os.fspath(source), "rb") as fh:
            return fh.read()
    # -- file-like: rewind and read --
    pos = None
    try:
        pos = source.tell()
        source.seek(0)
    except (AttributeError, OSError):
        pos = None
    data = source.read()
    if pos is not None:
        try:
            source.seek(pos)
        except (AttributeError, OSError):
            pass
    return data


def _stub_document_xml(zip_bytes: bytes) -> bytes:
    """Return a fresh zip with ``word/document.xml`` swapped for an empty stub.

    Used to bootstrap the eager :class:`DocumentPart` (styles,
    relationships, headers/footers all wire up through ``Package.open``
    as normal) without materialising the original body. The body
    contents are recovered via :func:`iterparse` on the original zip
    when ``paragraphs`` / ``tables`` / ``sections`` are accessed.
    """
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == _DOCUMENT_PART_NAME:
                    zout.writestr(item, _BODY_STUB)
                else:
                    # -- preserve original compression / metadata where possible --
                    zout.writestr(item, zin.read(item.filename))
    return out.getvalue()


class _StreamingParent:
    """Minimal :class:`ProvidesStoryPart` adapter for streamed proxies.

    The :class:`Paragraph` and :class:`Table` proxies require a parent
    exposing ``.part`` so they can resolve styles, relationships, and
    image lookups. Pointing them at the eagerly-loaded
    :class:`DocumentPart` reuses every existing accessor without
    rebuilding the part graph. The CT_P / CT_Tbl elements yielded by
    iterparse are detached (no real ``w:body`` ancestor) so accessors
    like :attr:`Paragraph._get_body` would otherwise fail; we adopt
    them into the stub body just before yielding so ancestor walks
    terminate cleanly.
    """

    __slots__ = ("_part",)

    def __init__(self, part: "DocumentPart"):
        self._part = part

    @property
    def part(self) -> "DocumentPart":
        return self._part


class StreamingDocument:
    """Read-only, bounded-memory view of a WordprocessingML package.

    Constructed via :meth:`Document.stream`. The body of
    ``word/document.xml`` is **never** held as a single tree; instead,
    :attr:`paragraphs` and :attr:`tables` are forward-only generators
    that surface one element at a time and clear it from memory once
    the consumer has yielded.

    Read-only properties of the yielded :class:`Paragraph` /
    :class:`Table` proxies (``text``, ``style``, ``runs``,
    ``alignment``, etc.) work transparently — they resolve through the
    eagerly-loaded :class:`DocumentPart` for styles, relationships,
    and numbering.

    .. versionadded:: 2026.05.13
    """

    def __init__(
        self,
        source: Union[str, "os.PathLike[str]", IO[bytes], bytes],
    ):
        from docx.api import Document as _DocumentFn

        self._zip_bytes: Optional[bytes] = _read_source_bytes(source)
        # -- bootstrap a real DocumentPart over a stubbed body so styles,
        # -- relationships, headers, footers, theme, settings, comments,
        # -- numbering, etc. all wire up via the existing factory. The
        # -- stubbed body is replaced with iterparse passes for the
        # -- paragraph / table / section accessors. --
        stubbed = _stub_document_xml(self._zip_bytes)
        eager = _DocumentFn(io.BytesIO(stubbed))
        self._eager_doc = eager
        self._part = eager.part
        self._parent = _StreamingParent(self._part)
        self._closed: bool = False
        # -- lazy section-scan cache; populated on first ``sections`` access --
        self._sections_populated: bool = False

    # -- context manager / lifecycle --------------------------------------

    def __enter__(self) -> "StreamingDocument":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()

    def close(self) -> None:
        """Release the cached source bytes and the underlying eager package.

        Idempotent — subsequent calls are no-ops. After ``close()``, the
        generator properties yield an empty iterator; accessing
        :attr:`sections` raises :class:`ValueError`.
        """
        if self._closed:
            return
        self._closed = True
        self._zip_bytes = None
        # -- best-effort drop of the eager doc's cached state --
        try:
            self._eager_doc.close()
        except Exception:  # pragma: no cover - defensive
            pass

    # -- core streaming generators ---------------------------------------

    def _iter_body_children(
        self, tag: Union[str, tuple]
    ) -> Iterator[etree._Element]:  # pyright: ignore[reportPrivateUsage]
        """Yield body-level elements matching ``tag`` via ``iterparse``.

        Only direct children of ``<w:body>`` are surfaced — paragraphs
        nested inside table cells, headers, or sdt content do not match
        the parent test and are skipped. Each yielded element is
        cleared from its parent after the consumer has resumed, freeing
        the lxml memory regardless of body length.
        """
        if self._closed or self._zip_bytes is None:
            return
        with zipfile.ZipFile(io.BytesIO(self._zip_bytes)) as zf:
            try:
                source = zf.open(_DOCUMENT_PART_NAME)
            except KeyError:
                return
            try:
                ip = etree.iterparse(source, events=("end",), tag=tag)
                ip.set_element_class_lookup(element_class_lookup)
                for _event, elem in ip:
                    parent = elem.getparent()
                    if parent is None or parent.tag != _W_BODY:
                        # -- nested w:p / w:tbl: still clear so memory
                        # -- doesn't accumulate, but don't yield. --
                        elem.clear()
                        continue
                    yield elem
                    # -- consumer has finished with the element; drop it
                    # -- and any preceding siblings so lxml can reclaim
                    # -- the memory. ``clear()`` keeps the element in
                    # -- the tree but drops its children; the loop after
                    # -- removes it from its parent entirely. --
                    elem.clear()
                    while elem.getprevious() is not None:
                        del parent[0]
            finally:
                source.close()

    @property
    def paragraphs(self) -> Iterator["Paragraph"]:
        """Forward-only generator yielding body |Paragraph| proxies.

        Only top-level body paragraphs surface, matching the eager
        :attr:`Document.paragraphs` contract. Each ``w:p`` element is
        adopted into the stub body so ancestor walks (e.g.
        :meth:`Paragraph._get_body`) terminate cleanly, then cleared
        from memory once the consumer has finished with the proxy.

        Re-iterating this property re-streams ``word/document.xml`` from
        the cached source bytes — paragraph state is **not** retained
        across passes.

        .. versionadded:: 2026.05.13
        """
        from docx.text.paragraph import Paragraph

        body = self._stub_body()
        for p in self._iter_body_children(_W_P):
            # -- adopt the parsed CT_P into the stub body so
            # -- ``Paragraph._get_body`` succeeds and so the proxy's
            # -- relative xpath queries resolve against a w:body root. --
            body.append(p)
            try:
                yield Paragraph(p, self._parent)  # type: ignore[arg-type]
            finally:
                # -- detach so memory doesn't accumulate in the stub --
                if p.getparent() is body:
                    body.remove(p)

    @property
    def tables(self) -> Iterator["Table"]:
        """Forward-only generator yielding top-level body |Table| proxies.

        Tables nested inside table cells are not yielded — the streaming
        reader does not descend into the inner-cell story. The eager
        path :attr:`Document.tables` has the same contract.

        .. versionadded:: 2026.05.13
        """
        from docx.table import Table

        body = self._stub_body()
        for tbl in self._iter_body_children(_W_TBL):
            body.append(tbl)
            try:
                yield Table(tbl, self._parent)  # type: ignore[arg-type]
            finally:
                if tbl.getparent() is body:
                    body.remove(tbl)

    # -- eager small-part accessors --------------------------------------

    @property
    def sections(self) -> "Sections":
        """All |Section| objects in this document, in order.

        Eager — the first access scans ``word/document.xml`` once via
        ``iterparse`` collecting only ``w:sectPr`` elements (typically
        a few hundred bytes each, regardless of body size) and grafts
        them into the stub body so the existing :class:`Sections`
        sequence can navigate them.

        Subsequent accesses reuse the cached scan. The current
        implementation is **read-only with respect to the stream**:
        adding or removing sections via the returned :class:`Sections`
        object will not be persisted (calling :meth:`save` raises).

        .. versionadded:: 2026.05.13
        """
        from docx.section import Sections

        if self._closed:
            raise ValueError("StreamingDocument is closed")
        if not self._sections_populated:
            self._populate_sections()
            self._sections_populated = True
        document_elm = cast("CT_Document", self._part.element)
        return Sections(document_elm, self._part)

    def _populate_sections(self) -> None:
        """Graft every ``w:sectPr`` from the source into the stub body.

        Targeted iterparse: the parser sees the entire body but we only
        capture ``w:sectPr`` end events and dispose of every other
        element as soon as it closes. Memory peaks at the size of the
        largest ``w:p`` / ``w:tbl`` *while it is being parsed* — Word
        emits paragraphs serially and rarely produces a single
        paragraph above a few hundred KB, so the working set is
        bounded.
        """
        if self._zip_bytes is None:
            return
        body = self._stub_body()
        with zipfile.ZipFile(io.BytesIO(self._zip_bytes)) as zf:
            try:
                source = zf.open(_DOCUMENT_PART_NAME)
            except KeyError:
                return
            try:
                ip = etree.iterparse(source, events=("end",))
                ip.set_element_class_lookup(element_class_lookup)
                for _event, elem in ip:
                    if elem.tag == _W_SECTPR:
                        parent = elem.getparent()
                        # -- mid-document: w:p/w:pPr/w:sectPr — clone the
                        # -- ancestor chain so the section break stays
                        # -- attached to a paragraph (Section reads its
                        # -- start_type from this position).
                        if parent is not None and parent.tag == qn("w:pPr"):
                            from copy import deepcopy

                            grandparent = parent.getparent()  # w:p
                            if grandparent is not None:
                                # -- detach the w:p shell carrying this
                                # -- sectPr and append a deep copy under
                                # -- our stub body so it survives the
                                # -- iterparse cleanup. --
                                shell = etree.SubElement(body, qn("w:p"))
                                shell.append(deepcopy(parent))
                            else:
                                from copy import deepcopy as _dc

                                body.append(_dc(elem))
                        else:
                            # -- final body sectPr --
                            from copy import deepcopy as _dc

                            body.append(_dc(elem))
                    # -- non-sectPr elements: drop their content as
                    # -- soon as we see them close, so the working set
                    # -- stays bounded regardless of document size. --
                    if elem.tag != _W_SECTPR:
                        elem.clear()
                        # -- detach from parent to release memory; siblings
                        # -- of body itself we leave alone (the iterparse
                        # -- root is our stub document element, not a real
                        # -- one whose contents we must preserve). --
                        prev = elem.getprevious()
                        while prev is not None:
                            del prev.getparent()[0]
                            prev = elem.getprevious()
            finally:
                source.close()

    @property
    def headers(self) -> "list[_Header]":
        """All header proxies referenced by any section, in document order.

        Headers are stored in separate ``word/header*.xml`` parts; the
        streaming reader exposes them through the eagerly-loaded
        :class:`DocumentPart`. The list is deduplicated by underlying
        relationship id so a header shared across sections appears
        once.

        .. versionadded:: 2026.05.13
        """
        seen: set[int] = set()
        result: list = []
        for section in self.sections:
            for header in (
                section.header,
                section.even_page_header,
                section.first_page_header,
            ):
                if header is None:
                    continue
                key = id(getattr(header, "_element", header))
                if key in seen:
                    continue
                seen.add(key)
                result.append(header)
        return result

    @property
    def footers(self) -> "list[_Footer]":
        """All footer proxies referenced by any section, in document order.

        See :attr:`headers` for the deduplication contract.

        .. versionadded:: 2026.05.13
        """
        seen: set[int] = set()
        result: list = []
        for section in self.sections:
            for footer in (
                section.footer,
                section.even_page_footer,
                section.first_page_footer,
            ):
                if footer is None:
                    continue
                key = id(getattr(footer, "_element", footer))
                if key in seen:
                    continue
                seen.add(key)
                result.append(footer)
        return result

    @property
    def styles(self) -> "Styles":
        """The |Styles| collection from the eagerly-loaded styles part.

        Styles are typically a few KB and reading them eagerly costs
        nothing — every :class:`Paragraph` proxy yielded by the stream
        resolves its style id through this collection.

        .. versionadded:: 2026.05.13
        """
        return self._eager_doc.styles

    @property
    def part(self) -> "DocumentPart":
        """The eagerly-loaded :class:`DocumentPart` backing this stream.

        Exposed for parity with :attr:`Document.part`; allows callers
        to reach into headers, footers, settings, comments, and other
        package-level resources without re-implementing the lookup
        graph.

        .. versionadded:: 2026.05.13
        """
        return self._part

    # -- mutation guard ---------------------------------------------------

    def save(self, *args, **kwargs) -> None:
        """Always raises :class:`StreamingNotMutableError`.

        The streaming reader never materialises the full body, so a
        round-trip save would silently lose the unread tail. Re-open
        the document via :func:`docx.Document` for an eager load if
        you need to write.

        .. versionadded:: 2026.05.13
        """
        raise StreamingNotMutableError(
            "StreamingDocument is read-only; re-open via docx.Document(...) "
            "to mutate or save."
        )

    # -- helpers ----------------------------------------------------------

    def _stub_body(self) -> etree._Element:  # pyright: ignore[reportPrivateUsage]
        """Return the (mostly empty) ``<w:body>`` of the stubbed document part.

        The body acts as a temporary harness for streamed elements so
        that descendant xpath / ancestor lookups against ``Paragraph``
        / ``Table`` proxies resolve to a real ``w:body`` ancestor (some
        accessors walk up to the body to allocate ids). Elements are
        appended just before they're yielded and removed immediately
        after, so the body never accumulates more than one entry from
        the stream at a time.
        """
        document_elm = cast("CT_Document", self._part.element)
        return cast("etree._Element", document_elm.body)


def open_stream(
    source: Union[str, "os.PathLike[str]", IO[bytes], bytes],
) -> StreamingDocument:
    """Open `source` as a :class:`StreamingDocument`.

    Public shim under the :mod:`docx.streaming` namespace; the canonical
    entry point is :meth:`Document.stream`.

    .. versionadded:: 2026.05.13
    """
    return StreamingDocument(source)
