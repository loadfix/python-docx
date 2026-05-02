"""Cross-document append helpers for :class:`docx.document.Document`.

Implements the ``Document.append_document`` / ``append_body`` / ``append_paragraph``
surface described by upstream issues #1457, #558, #543, #437, #460, #44, and #709:
copy body content (paragraphs, tables, block-level SDTs) from another document into
this one while rewiring the accompanying relationships so images, embedded objects,
hyperlinks, and so on continue to work.

The copy is intentionally pragmatic rather than perfect:

- Block-level children of ``w:body`` (``w:p``, ``w:tbl``, ``w:sdt``) are deep-copied
  and inserted before the destination body's sentinel ``w:sectPr`` (so destination
  section settings win).
- Each ``r:id`` / ``r:embed`` / ``r:link`` attribute found on the copied subtree is
  resolved against the *source* document part's relationship map; the referenced
  target part (image, embedded object, etc.) is added to the destination package if
  not already present (images deduplicated by SHA-1, everything else by source URI)
  and the attribute is rewritten to the new rId.
- Paragraph-style references (``w:pPr/w:pStyle/@w:val``) and run-style references
  (``w:rPr/w:rStyle/@w:val``) are resolved against the source ``styles.xml``;
  missing styles are copied into the destination (including any linked / basedOn /
  next-style dependencies).
- ``w:numPr/w:numId/@w:val`` references are remapped onto copies of the source
  numbering definitions, copied into the destination numbering part.

.. versionadded:: 1.3.0.dev0
"""

from __future__ import annotations

import copy
import io
from typing import TYPE_CHECKING, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.document import Document
    from docx.opc.part import Part
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.text.paragraph import Paragraph


# -- lxml QName constants for the three "relationship-reference" attributes used by
# -- WordprocessingML / DrawingML content (image blips, hyperlinks, embedded
# -- objects, charts, headers/footers in sectPr, etc.). --
_REL_ATTRS = (qn("r:id"), qn("r:embed"), qn("r:link"))


def append_document(dest: Document, source: Document) -> int:
    """Append the entire body of `source` to `dest` and return the number of block
    elements copied.

    Images, embedded objects, styles, and numbering definitions referenced by the
    copied content are imported into `dest`'s package as needed. The destination's
    section properties are preserved — copied content is inserted before the
    destination's sentinel ``w:sectPr``.

    .. versionadded:: 1.3.0.dev0
    """
    src_body = source._element.body  # type: ignore[attr-defined]
    dst_body = dest._element.body  # type: ignore[attr-defined]
    src_part = source.part
    dst_part = dest.part

    # -- gather block-level children (paragraphs, tables, sdt) in document order --
    block_children = list(src_body.xpath("./w:p | ./w:tbl | ./w:sdt"))
    count = 0
    for child in block_children:
        _append_block_element(child, src_part, dst_body, dst_part)
        count += 1
    return count


def append_body(dest: Document, source: Document) -> int:
    """Alias for :func:`append_document` kept as a distinct entry-point.

    Historically python-docx users have asked for both names (upstream#1457,
    upstream#558); the behaviour is identical today.

    .. versionadded:: 1.3.0.dev0
    """
    return append_document(dest, source)


def append_paragraph(dest: Document, paragraph: Paragraph) -> Paragraph:
    """Copy `paragraph` from its owning document into `dest` and return the new
    :class:`Paragraph` instance.

    Relationships referenced by the paragraph (images, hyperlinks, embedded
    objects, style / numbering references) are rewired the same way as for
    :func:`append_document`.

    .. versionadded:: 1.3.0.dev0
    """
    from docx.text.paragraph import Paragraph as ParagraphCls

    src_part = cast("Part", paragraph.part)
    dst_body = dest._element.body  # type: ignore[attr-defined]
    dst_part = dest.part

    new_elm = _append_block_element(paragraph._p, src_part, dst_body, dst_part)  # type: ignore[attr-defined]
    return ParagraphCls(new_elm, dest._body)  # type: ignore[attr-defined]


# ---------------------------------------------------------------------------
# internal helpers
# ---------------------------------------------------------------------------


def _append_block_element(
    element: BaseOxmlElement,
    src_part: Part,
    dst_body: BaseOxmlElement,
    dst_part: Part,
) -> BaseOxmlElement:
    """Deep-copy `element` into `dst_body` (before any trailing sectPr) and return it.

    Performs rId remapping, style import, and numbering import on the clone.
    """
    clone = copy.deepcopy(element)
    _remap_rels_in_subtree(clone, src_part, dst_part)
    _import_referenced_styles(clone, src_part, dst_part)
    _import_referenced_numbering(clone, src_part, dst_part)

    # -- insert before trailing w:sectPr if present, else append --
    sectPr = dst_body.find(qn("w:sectPr"))
    if sectPr is not None:
        sectPr.addprevious(clone)
    else:
        dst_body.append(clone)
    return clone


def _remap_rels_in_subtree(
    subtree: BaseOxmlElement, src_part: Part, dst_part: Part
) -> None:
    """Rewrite every relationship-reference attribute in `subtree` from src->dst rId.

    Walks the subtree once and, for each `r:id` / `r:embed` / `r:link` attribute,
    resolves the source rId to a source part (or external URL), adds that target
    to the destination package if needed, and rewrites the attribute value in
    place.
    """
    src_rels = src_part.rels
    # -- Use plain iteration over descendants; the xpath cache will pick up this
    # -- expression after the first use. We include the element itself. --
    for el in subtree.iter():
        for attr in _REL_ATTRS:
            rId = el.get(attr)
            if not rId or rId not in src_rels:
                continue
            new_rId = _import_rel(rId, src_rels, dst_part)
            if new_rId != rId:
                el.set(attr, new_rId)


def _import_rel(rId: str, src_rels, dst_part: Part) -> str:
    """Clone the `rId` relationship from `src_rels` into `dst_part.rels` and return
    the destination rId (which may equal the source rId when it happens to be free).
    """
    rel = src_rels[rId]
    if rel.is_external:
        return dst_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
    target = rel.target_part
    imported = _import_part(target, dst_part)
    return dst_part.relate_to(imported, rel.reltype)


def _import_part(src_part: Part, dst_part: Part) -> Part:
    """Return a destination-side |Part| mirroring `src_part`.

    Images are deduplicated via the destination package's existing image-parts
    collection (hashed by SHA-1). Other parts (embedded objects, charts, etc.)
    are cloned into the destination package under a partname that doesn't
    collide with existing parts.
    """
    from docx.opc.constants import CONTENT_TYPE as CT
    from docx.opc.part import Part as PartCls

    dst_package = dst_part.package
    assert dst_package is not None

    # -- image: round-trip through the package's image_parts collection so the
    # -- SHA-1 dedup path can reuse an existing copy. --
    if src_part.content_type and src_part.content_type.startswith("image/"):
        # -- `get_or_add_image_part` is defined on WordprocessingML Package but
        # -- not on generic OpcPackage, so resolve via getattr for safety. --
        get_or_add_image_part = getattr(dst_package, "get_or_add_image_part", None)
        if get_or_add_image_part is not None:
            return get_or_add_image_part(io.BytesIO(src_part.blob))

    # -- anything else: clone under a fresh partname and return. If the same
    # -- source partname was previously imported, reuse that clone. --
    cache = _importer_cache(dst_part)
    src_key = str(src_part.partname)
    if src_key in cache:
        return cache[src_key]

    new_partname = _next_free_partname(dst_package, src_part.partname)
    # -- subclass `type(src_part)` so part-specific behaviour survives the copy --
    cls = type(src_part)
    try:
        new_part = cls(
            new_partname,
            src_part.content_type,
            src_part.blob,
            dst_package,
        )
    except TypeError:
        # -- XmlPart signature takes an element, not a blob --
        try:
            from docx.oxml.parser import parse_xml

            new_part = cls(
                new_partname,
                src_part.content_type,
                parse_xml(src_part.blob),
                dst_package,
            )
        except Exception:
            new_part = PartCls(
                new_partname,
                src_part.content_type,
                src_part.blob,
                dst_package,
            )
    cache[src_key] = new_part
    return new_part


def _importer_cache(dst_part: Part) -> dict[str, Part]:
    """Per-destination map of source-partname -> destination Part.

    Mutating lookup used by :func:`_import_part` to avoid creating duplicate
    destination parts when the same source part is referenced from multiple
    places (e.g. two paragraphs that both embed the same image).
    """
    cache = getattr(dst_part, "_append_import_cache", None)
    if cache is None:
        cache = {}
        dst_part._append_import_cache = cache  # type: ignore[attr-defined]
    return cache


def _next_free_partname(dst_package, template_partname: PackURI) -> PackURI:
    """Return a partname of the same shape as `template_partname` not used in dst."""
    existing = {str(p.partname) for p in dst_package.iter_parts()}
    # -- split "/word/media/image3.png" into ("/word/media/image", "3", ".png") --
    stem = template_partname.rsplit("/", 1)[-1]
    base, sep, ext = stem.rpartition(".")
    head = template_partname.rsplit("/", 1)[0]
    if not base:
        base = stem
        ext = ""
        sep = ""
    # -- drop trailing digits off base so we can renumber --
    core = base.rstrip("0123456789")
    if not core:
        core = base
    for n in range(1, len(existing) + 2):
        candidate = "%s/%s%d%s%s" % (head, core, n, sep, ext)
        if candidate not in existing:
            return PackURI(candidate)
    # -- fallback, should not normally reach here --
    return PackURI(str(template_partname))


# ---------------------------------------------------------------------------
# style + numbering import
# ---------------------------------------------------------------------------


def _import_referenced_styles(
    subtree: BaseOxmlElement, src_part: Part, dst_part: Part
) -> None:
    """Copy any style used by `subtree` from source styles.xml into destination.

    Looks at `w:pStyle` (paragraph), `w:rStyle` (run), `w:tblStyle` (table), and
    `w:numStyleLink` / `w:styleLink` descendants of style definitions.
    """
    src_styles_part = _safe_styles_part(src_part)
    dst_styles_part = _safe_styles_part(dst_part)
    if src_styles_part is None or dst_styles_part is None:
        return

    src_styles = src_styles_part.element
    dst_styles = dst_styles_part.element
    existing_dst = {
        s.get(qn("w:styleId")) for s in dst_styles.findall(qn("w:style"))
    }
    existing_dst.discard(None)

    # -- xpath for all style refs in the subtree --
    refs = subtree.xpath(
        ".//w:pStyle/@w:val | .//w:rStyle/@w:val | .//w:tblStyle/@w:val"
    )
    to_import = [r for r in refs if r and r not in existing_dst]
    # -- depth-first BFS so basedOn/next/link dependencies come along --
    seen: set[str] = set()
    queue = list(to_import)
    while queue:
        style_id = queue.pop(0)
        if style_id in seen or style_id in existing_dst:
            continue
        seen.add(style_id)
        src_style = _find_style(src_styles, style_id)
        if src_style is None:
            continue
        dst_styles.append(copy.deepcopy(src_style))
        existing_dst.add(style_id)
        # -- chase dependencies (basedOn, next, link, numId references skipped here) --
        for dep_tag in ("w:basedOn", "w:next", "w:link"):
            for dep in src_style.findall(qn(dep_tag)):
                dep_id = dep.get(qn("w:val"))
                if dep_id and dep_id not in existing_dst:
                    queue.append(dep_id)


def _find_style(styles_elm: BaseOxmlElement, style_id: str):
    for style in styles_elm.findall(qn("w:style")):
        if style.get(qn("w:styleId")) == style_id:
            return style
    return None


def _safe_styles_part(part: Part):
    """Return the StylesPart related to `part` or None if unavailable."""
    try:
        return part.part_related_by(RT.STYLES)
    except (KeyError, AttributeError):
        return None


def _import_referenced_numbering(
    subtree: BaseOxmlElement, src_part: Part, dst_part: Part
) -> None:
    """Ensure `w:numId` references in `subtree` resolve in the destination numbering.

    Each ``w:numId/@w:val`` in the copied subtree is looked up in the source
    numbering part; the referenced ``w:num`` and its linked ``w:abstractNum``
    are cloned into the destination numbering part under fresh ids, and the
    reference in `subtree` is rewritten.
    """
    num_val_attrs = subtree.xpath(".//w:numPr/w:numId[@w:val]")
    if not num_val_attrs:
        return

    src_numbering = _safe_numbering_element(src_part)
    dst_numbering = _safe_numbering_element(dst_part)
    if src_numbering is None or dst_numbering is None:
        return

    # -- cache per (src, dst) to avoid re-importing a numId we've already mapped --
    cache = _importer_cache(dst_part).setdefault("__numIds__", {})  # type: ignore[assignment]

    for numId_el in num_val_attrs:
        old = numId_el.get(qn("w:val"))
        if old is None:
            continue
        if old in cache:
            new = cache[old]
        else:
            new = _clone_num_definition(old, src_numbering, dst_numbering)
            cache[old] = new
        if new != old:
            numId_el.set(qn("w:val"), new)


def _safe_numbering_element(part: Part):
    try:
        numbering_part = part.part_related_by(RT.NUMBERING)
    except (KeyError, AttributeError):
        return None
    return getattr(numbering_part, "element", None)


def _clone_num_definition(
    src_numId: str, src_numbering: BaseOxmlElement, dst_numbering: BaseOxmlElement
) -> str:
    """Copy the ``w:num`` (and its ``w:abstractNum``) with numId=`src_numId` from
    `src_numbering` to `dst_numbering` under fresh ids and return the new numId."""
    src_num = None
    for num in src_numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == src_numId:
            src_num = num
            break
    if src_num is None:
        return src_numId

    # -- determine fresh numId / abstractNumId --
    existing_numIds = {
        n.get(qn("w:numId")) for n in dst_numbering.findall(qn("w:num"))
    }
    new_numId = _next_free_id(existing_numIds)
    existing_absIds = {
        a.get(qn("w:abstractNumId"))
        for a in dst_numbering.findall(qn("w:abstractNum"))
    }

    # -- find the src abstractNumId referenced by src_num --
    src_absId_el = src_num.find(qn("w:abstractNumId"))
    src_absId = src_absId_el.get(qn("w:val")) if src_absId_el is not None else None

    new_absId = _next_free_id(existing_absIds)
    # -- clone abstractNum (if present) --
    if src_absId is not None:
        src_abs = None
        for a in src_numbering.findall(qn("w:abstractNum")):
            if a.get(qn("w:abstractNumId")) == src_absId:
                src_abs = a
                break
        if src_abs is not None:
            abs_clone = copy.deepcopy(src_abs)
            abs_clone.set(qn("w:abstractNumId"), new_absId)
            # -- w:abstractNum must come before w:num in numbering.xml --
            first_num = dst_numbering.find(qn("w:num"))
            if first_num is not None:
                first_num.addprevious(abs_clone)
            else:
                dst_numbering.append(abs_clone)

    # -- clone num --
    num_clone = copy.deepcopy(src_num)
    num_clone.set(qn("w:numId"), new_numId)
    if src_absId is not None:
        new_abs_ref = num_clone.find(qn("w:abstractNumId"))
        if new_abs_ref is not None:
            new_abs_ref.set(qn("w:val"), new_absId)
    dst_numbering.append(num_clone)
    return new_numId


def _next_free_id(existing: set[str | None]) -> str:
    """Return the smallest positive integer (as string) not in `existing`."""
    used = {int(x) for x in existing if x and x.lstrip("-").isdigit()}
    n = 1
    while n in used:
        n += 1
    return str(n)
