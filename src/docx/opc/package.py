"""Objects that implement reading and writing OPC packages.

:class:`OpcPackage` is the docx-local top-level package class. It is kept
separate from the shared :class:`ooxml_opc.package.OpcPackage` because docx's
loader dispatches through the docx-shape :class:`Unmarshaller` +
:class:`~docx.opc.pkgreader.PackageReader` (from_file/iter_sparts/iter_srels)
while the shared loader is :class:`~ooxml_opc.package._PackageLoader`
(PackageReader Mapping lookups). Both converge on the same shared primitives:
:class:`~docx.opc.part.PartFactory` (docx-shape subclass of the shared
factory), :class:`~docx.opc.rel.Relationships` (docx-shape subclass of the
shared collection), and :class:`~docx.opc.pkgwriter.PackageWriter` on save.

Docx-specific save-time logic retained here:

* ``_drop_unused_package_rels`` — prune library-authored THUMBNAIL rels.
* ``_remap_clashing_cp_prefix`` — rebind the LibreOffice-mis-bound ``cp:``
  core-properties prefix.
* ``_validate_save_path`` — reject Windows-invalid filename characters.
* ``recover=True`` / ``huge_tree=True`` open-time plumbing.

.. versionchanged:: 2026.05.11
   Uses the shared runtime's part/rel/oxml primitives under the hood; the
   outward :class:`OpcPackage` API is unchanged.
"""

from __future__ import annotations

import copy
import os
import posixpath
from typing import IO, TYPE_CHECKING, cast
from collections.abc import Iterator

from lxml import etree

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PACKAGE_URI, PackURI
from docx.opc.part import PartFactory
from docx.opc.parts.coreprops import CorePropertiesPart
from docx.opc.pkgreader import PackageReader
from docx.opc.pkgwriter import PackageWriter
from docx.opc.rel import Relationships
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from typing_extensions import Self

    from docx.opc.coreprops import CoreProperties
    from docx.opc.part import Part
    from docx.opc.rel import _Relationship  # pyright: ignore[reportPrivateUsage]


#: Characters Windows disallows in file names (drive-letter colons are allowed
#: in the path prefix but not inside the filename itself). ``/`` and ``\`` are
#: directory separators and are intentionally omitted from this set.
_WINDOWS_INVALID_FILENAME_CHARS = frozenset('<>:"|?*')

#: Canonical URI bound to the ``cp:`` prefix inside core.xml (the Open Packaging
#: Conventions core-properties namespace). Anything else bound to ``cp:`` in an
#: incoming core-properties part collides with python-docx's own use of the
#: prefix on the write side.
_CORE_PROPS_URI = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"

#: URI LibreOffice mis-binds to ``cp:`` in some core.xml files — the
#: officeDocument custom-properties namespace. When detected we rebind it to
#: the ``custprops:`` prefix python-docx already uses elsewhere.
_CUSTOM_PROPS_URI = (
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
)


def _remap_clashing_cp_prefix(part: CorePropertiesPart) -> None:
    """Rebuild `part`'s root element when its nsmap binds ``cp:`` to a non-core URI.

    LibreOffice (and a handful of other producers) occasionally emit
    ``docProps/core.xml`` with an ``xmlns:cp="...custom-properties"``
    declaration — the ``cp:`` prefix is bound to the custom-properties
    namespace rather than the canonical core-properties one. python-docx's
    own code emits ``<cp:lastModifiedBy>`` and siblings using ``qn("cp:...")``,
    which expands to a *core-properties* Clark name. Serialising against the
    incoming (mis-bound) nsmap then produces two elements that both print as
    ``cp:lastModifiedBy`` but carry different URIs, which Word rejects
    (upstream#1037, extends PR#1436).

    This helper re-emits the tree with a clean nsmap: ``cp:`` firmly bound to
    the core-properties URI and the custom-properties URI moved to the
    ``custprops:`` prefix. Element Clark names are preserved — nothing is
    silently reassigned to a different namespace.
    """
    element = part.element
    # -- only intervene when the root's `cp` prefix points somewhere other
    # -- than the core-properties URI. The common case (a well-formed
    # -- core.xml or one with no `cp` binding at all) is a no-op.
    # -- We also guard against mock parts used in unit tests whose
    # -- ``element`` isn't a real lxml node. --
    if element is None or not isinstance(
        element, etree._Element  # pyright: ignore[reportPrivateUsage]
    ):
        return
    root_nsmap = element.nsmap
    clashing_uri = root_nsmap.get("cp")
    if clashing_uri is None or clashing_uri == _CORE_PROPS_URI:
        return
    # -- Preserve any existing well-behaved bindings; override the offender
    # -- and give the custom-properties URI a safe prefix. --
    replacement_nsmap: dict[str | None, str] = {}
    for pfx, uri in root_nsmap.items():
        if pfx == "cp":
            continue  # -- will rebind below --
        replacement_nsmap[pfx] = uri
    replacement_nsmap["cp"] = _CORE_PROPS_URI
    # -- Only add `custprops:` when it isn't already bound to a *different*
    # -- URI in the file; if it is, pick a non-clashing fallback prefix. --
    safe_prefix = "custprops"
    if replacement_nsmap.get(safe_prefix, _CUSTOM_PROPS_URI) != _CUSTOM_PROPS_URI:
        for candidate in ("custprops0", "custprops1", "custprops_alt"):
            if candidate not in replacement_nsmap:
                safe_prefix = candidate
                break
    replacement_nsmap[safe_prefix] = clashing_uri
    # -- Build the replacement root with the corrected nsmap and deep-copy
    # -- each child in. Children inherit the new root's nsmap so lxml
    # -- serialises each element with a prefix matching its actual URI. --
    new_root = etree.SubElement(
        etree.Element("tmp"), element.tag, attrib=dict(element.attrib), nsmap=replacement_nsmap
    )
    new_root.text = element.text
    new_root.tail = element.tail
    for child in list(element):
        new_root.append(copy.deepcopy(child))
    # -- Detach from the scratch parent and reparse through the registered
    # -- CT_CoreProperties element class so the part's ``element`` keeps the
    # -- same type. --
    from docx.oxml.parser import parse_xml

    part._element = parse_xml(etree.tostring(new_root))  # pyright: ignore[reportPrivateUsage]


def _validate_save_path(path: str) -> None:
    """Raise :class:`OSError` when `path`'s filename contains Windows-invalid chars.

    Only the basename portion is inspected — ``C:/foo/bar.docx`` is fine on
    every platform because the drive-letter colon is part of the path prefix,
    not the filename. Control characters (``\\x00``-``\\x1f``) are also
    rejected. The check runs on every platform so that scripts developed on
    POSIX surface the problem early rather than silently producing a file that
    Windows consumers can't open. Closes upstream#1111.
    """
    # -- split on both forward and backward slashes so Windows-style paths
    # -- work on POSIX and vice-versa. --
    basename = os.path.basename(path.replace("\\", "/"))
    basename = posixpath.basename(basename)
    if not basename:
        raise OSError(
            "invalid save path %r: no filename component" % path
        )
    bad = sorted({ch for ch in basename if ch in _WINDOWS_INVALID_FILENAME_CHARS})
    if bad:
        raise OSError(
            "invalid character(s) %s in filename %r (Windows-invalid)"
            % (", ".join(repr(c) for c in bad), basename)
        )
    ctrl = sorted({ch for ch in basename if ord(ch) < 0x20})
    if ctrl:
        raise OSError(
            "control character(s) in filename %r are not permitted" % basename
        )


class OpcPackage:
    """Main API class for |python-opc|.

    A new instance is constructed by calling the :meth:`open` class method with a path
    to a package file or file-like object containing one.
    """

    # -- parse warnings accumulated during `open(..., recover=True)`;
    # -- assigned on the package instance after unmarshalling and read via
    # -- the :attr:`recovery_warnings` property. --
    _recovery_warnings: list[str]

    def after_unmarshal(self):
        """Entry point for any post-unmarshaling processing.

        May be overridden by subclasses without forwarding call to super.
        """
        # don't place any code here, just catch call if not overridden by
        # subclass
        pass

    @property
    def core_properties(self) -> CoreProperties:
        """|CoreProperties| object providing read/write access to the Dublin Core
        properties for this document."""
        return self._core_properties_part.core_properties

    def iter_rels(self) -> Iterator[_Relationship]:
        """Generate exactly one reference to each relationship in the package by
        performing a depth-first traversal of the rels graph."""

        def walk_rels(
            source: OpcPackage | Part, visited: list[Part] | None = None
        ) -> Iterator[_Relationship]:
            visited = [] if visited is None else visited
            for rel in source.rels.values():
                yield rel
                if rel.is_external:
                    continue
                part = rel.target_part
                if part in visited:
                    continue
                visited.append(part)
                new_source = part
                for rel in walk_rels(new_source, visited):
                    yield rel

        for rel in walk_rels(self):
            yield rel

    def iter_parts(self) -> Iterator[Part]:
        """Generate exactly one reference to each of the parts in the package by
        performing a depth-first traversal of the rels graph."""

        def walk_parts(source, visited=[]):
            for rel in source.rels.values():
                if rel.is_external:
                    continue
                part = rel.target_part
                if part in visited:
                    continue
                visited.append(part)
                yield part
                new_source = part
                for part in walk_parts(new_source, visited):
                    yield part

        for part in walk_parts(self):
            yield part

    def load_rel(self, reltype: str, target: Part | str, rId: str, is_external: bool = False):
        """Return newly added |_Relationship| instance of `reltype` between this part
        and `target` with key `rId`.

        Target mode is set to ``RTM.EXTERNAL`` if `is_external` is |True|. Intended for
        use during load from a serialized package, where the rId is well known. Other
        methods exist for adding a new relationship to the package during processing.
        """
        return self.rels.add_relationship(reltype, target, rId, is_external)

    @property
    def main_document_part(self):
        """Return a reference to the main document part for this package.

        Examples include a document part for a WordprocessingML package, a presentation
        part for a PresentationML package, or a workbook part for a SpreadsheetML
        package.
        """
        return self.part_related_by(RT.OFFICE_DOCUMENT)

    def next_partname(self, template: str) -> PackURI:
        """Return a |PackURI| instance representing partname matching `template`.

        The returned part-name has the next available numeric suffix to distinguish it
        from other parts of its type. `template` is a printf (%)-style template string
        containing a single replacement item, a '%d' to be used to insert the integer
        portion of the partname. Example: "/word/header%d.xml"
        """
        partnames = {part.partname for part in self.iter_parts()}
        for n in range(1, len(partnames) + 2):
            candidate_partname = template % n
            if candidate_partname not in partnames:
                return PackURI(candidate_partname)

    @classmethod
    def open(
        cls,
        pkg_file: str | IO[bytes],
        recover: bool = False,
        huge_tree: bool = False,
        password: str | None = None,
    ) -> Self:
        """Return an |OpcPackage| instance loaded with the contents of `pkg_file`.

        When `recover` is True, XML parsing falls back to lxml's recovering
        parser when it encounters malformed content and any parse warnings are
        accumulated on ``package.recovery_warnings`` instead of raising
        :class:`lxml.etree.XMLSyntaxError`. Default behaviour (``recover=False``)
        is unchanged.

        When `huge_tree` is True, the lxml ``huge_tree=True`` parser variant is
        used for every part in the package. This lifts libxml2's default
        10 MB-per-AttValue and 256-deep nesting limits so extremely large
        documents can be parsed (upstream#1086). Only enable for trusted
        input — the default parser's XML-bomb protections no longer apply.

        When `password` is provided and `pkg_file` is an ECMA-376 Agile-
        Encryption (password-protected) ``.docx``, it is transparently
        decrypted before loading. Decryption requires the optional
        ``python-ooxml-crypto`` dependency.

        .. versionchanged:: 2026.05.10
           Added ``password`` parameter.
        """
        from docx.oxml.parser import huge_tree_mode, recovery_mode

        def _load() -> Self:
            pkg_reader = PackageReader.from_file(pkg_file, password=password)
            package = cls()
            Unmarshaller.unmarshal(pkg_reader, package, PartFactory)
            return package

        if recover and huge_tree:
            with huge_tree_mode(), recovery_mode() as warnings:
                package = _load()
            package._recovery_warnings = list(warnings)
            return package

        if recover:
            with recovery_mode() as warnings:
                package = _load()
            package._recovery_warnings = list(warnings)
            return package

        if huge_tree:
            with huge_tree_mode():
                return _load()

        return _load()

    @property
    def recovery_warnings(self) -> list[str]:
        """List of parse-warning strings collected while opening this package.

        Empty for packages opened without ``recover=True`` or for well-formed
        packages opened in recovery mode.
        """
        return list(getattr(self, "_recovery_warnings", []))

    def part_related_by(self, reltype: str) -> Part:
        """Return part to which this package has a relationship of `reltype`.

        Raises |KeyError| if no such relationship is found and |ValueError| if more than
        one such relationship is found.
        """
        return self.rels.part_with_reltype(reltype)

    @property
    def parts(self) -> list[Part]:
        """Return a list containing a reference to each of the parts in this package."""
        return list(self.iter_parts())

    def relate_to(self, part: Part, reltype: str):
        """Return rId key of new or existing relationship to `part`.

        If a relationship of `reltype` to `part` already exists, its rId is returned. Otherwise a
        new relationship is created and that rId is returned.
        """
        rel = self.rels.get_or_add(reltype, part)
        return rel.rId

    @lazyproperty
    def rels(self):
        """Return a reference to the |Relationships| instance holding the collection of
        relationships for this package."""
        return Relationships(PACKAGE_URI.baseURI)

    def save(
        self,
        pkg_file: str | IO[bytes],
        reproducible: bool = False,
        password: str | None = None,
    ):
        """Save this package to `pkg_file`.

        `pkg_file` can be either a file-path or a file-like object.

        When `pkg_file` is a string, the filename component is validated against
        the characters Windows disallows in filenames (``< > : " | ? *``). A
        mismatch raises :class:`OSError`. This avoids the silently-truncated /
        no-file-created failure mode reported in upstream#1111 when callers
        pass e.g. ``"my:file.docx"`` as the save target.

        When `reproducible` is True, the emitted zip archive uses fixed
        timestamps and sorted member names so repeated saves of the same content
        produce byte-identical output. Closes upstream#1042 / upstream-PR#810.

        When `password` is provided the saved ``.docx`` is password-protected
        using ECMA-376 Agile Encryption (the scheme Word uses). Encryption
        requires the optional ``python-ooxml-crypto`` dependency.
        ``reproducible`` and ``password`` are orthogonal — fixed timestamps
        stamp the inner (plaintext) zip members before the encryption wrapper
        is applied.

        .. versionadded:: 2026.05.0
           The `reproducible` parameter.
        .. versionadded:: 2026.05.10
           The `password` parameter.
        """
        if isinstance(pkg_file, str):
            _validate_save_path(pkg_file)
        self._drop_unused_package_rels()
        for part in self.parts:
            part.before_marshal(reproducible=reproducible)
        PackageWriter.write(
            pkg_file,
            self.rels,
            self.parts,
            reproducible=reproducible,
            password=password,
        )

    def _drop_unused_package_rels(self) -> None:
        """Drop package-level rels whose target parts python-docx doesn't author.

        Historically this pruned the ``RT.THUMBNAIL`` rel unconditionally
        on the theory that python-docx has no renderer, so any shipped
        thumbnail is stale the moment any content changes. That heuristic
        was too aggressive: files authored by Word ship a thumbnail part
        and users reasonably expect it to survive a round-trip. Dropping
        it silently destroys user data with no signal.

        The narrowed policy: drop the thumbnail **only** when python-docx
        created it itself (no ``_loaded_from_package`` flag). Thumbnails
        that shipped in the source package are preserved verbatim — they
        may be stale relative to new edits but keeping them is strictly
        less destructive than dropping them.
        """
        from docx.opc.constants import RELATIONSHIP_TYPE as RT

        to_drop: list[str] = []
        for rId, rel in list(self.rels.items()):
            if rel.is_external:
                continue
            if rel.reltype != RT.THUMBNAIL:
                continue
            # -- only drop library-authored thumbnails; preserve those
            # -- that shipped in the source package. --
            if getattr(rel.target_part, "_loaded_from_package", False):
                continue
            to_drop.append(rId)
        for rId in to_drop:
            del self.rels[rId]

    @property
    def _extended_properties_part(self):
        """Return the |ExtendedPropertiesPart| related to this package.

        Creates a default (empty) extended-properties part lazily when none is
        already related. Mirrors :attr:`_core_properties_part`; the
        extended-properties part is conventionally wired to the package root.

        .. versionadded:: 2026.05.0
        """
        from docx.parts.extended_properties import ExtendedPropertiesPart

        try:
            return cast(ExtendedPropertiesPart, self.part_related_by(RT.EXTENDED_PROPERTIES))
        except KeyError:
            part = ExtendedPropertiesPart.default(self)
            self.relate_to(part, RT.EXTENDED_PROPERTIES)
            return part

    @property
    def _core_properties_part(self) -> CorePropertiesPart:
        """|CorePropertiesPart| object related to this package.

        Creates a default core properties part if one is not present (not common).
        Honours both the canonical ``package/2006`` core-properties reltype and
        the ``officeDocument/2006`` alternate emitted by some producers
        (upstream-PR#1436) so a duplicate ``docProps/core.xml`` isn't created
        when the existing rel uses the alternate form.

        If the located part has the ``cp:`` prefix bound to a non-core URI
        (LibreOffice, upstream#1037), the tree is rebuilt with a safe prefix
        so subsequent writes don't emit duplicate ``cp:lastModifiedBy``
        elements.
        """
        part: CorePropertiesPart | None = None
        try:
            part = cast(CorePropertiesPart, self.part_related_by(RT.CORE_PROPERTIES))
        except KeyError:
            try:
                part = cast(
                    CorePropertiesPart, self.part_related_by(RT.CORE_PROPERTIES_ALT)
                )
            except KeyError:
                part = CorePropertiesPart.default(self)
                self.relate_to(part, RT.CORE_PROPERTIES)
        _remap_clashing_cp_prefix(part)
        return part


class Unmarshaller:
    """Hosts static methods for unmarshalling a package from a |PackageReader|."""

    @staticmethod
    def unmarshal(pkg_reader, package, part_factory):
        """Construct graph of parts and realized relationships based on the contents of
        `pkg_reader`, delegating construction of each part to `part_factory`.

        Package relationships are added to `pkg`.
        """
        parts = Unmarshaller._unmarshal_parts(pkg_reader, package, part_factory)
        Unmarshaller._unmarshal_relationships(pkg_reader, package, parts)
        for part in parts.values():
            part.after_unmarshal()
        package.after_unmarshal()

    @staticmethod
    def _unmarshal_parts(pkg_reader, package, part_factory):
        """Return a dictionary of |Part| instances unmarshalled from `pkg_reader`, keyed
        by partname.

        Side-effect is that each part in `pkg_reader` is constructed using
        `part_factory`. Each constructed part is flagged with
        ``_loaded_from_package = True`` so later save-time heuristics can
        distinguish parts that shipped in the source package from parts the
        library itself created on demand. The distinction matters because
        optional parts (``stylesWithEffects.xml``, ``customXml/*``,
        ``thumbnail.jpeg``) must survive round-tripping even when python-docx
        cannot statically prove they are referenced — dropping them destroys
        user data that Word-authored files depend on.
        """
        parts = {}
        for partname, content_type, reltype, blob in pkg_reader.iter_sparts():
            part = part_factory(partname, content_type, reltype, blob, package)
            part._loaded_from_package = True  # pyright: ignore[reportPrivateUsage]
            parts[partname] = part
        return parts

    @staticmethod
    def _unmarshal_relationships(pkg_reader, package, parts):
        """Add a relationship to the source object corresponding to each of the
        relationships in `pkg_reader` with its target_part set to the actual target part
        in `parts`."""
        for source_uri, srel in pkg_reader.iter_srels():
            source = package if source_uri == "/" else parts[source_uri]
            target = srel.target_ref if srel.is_external else parts[srel.target_partname]
            source.load_rel(srel.reltype, target, srel.rId, srel.is_external)
