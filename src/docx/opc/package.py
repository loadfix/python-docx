"""Objects that implement reading and writing OPC packages."""

from __future__ import annotations

import os
import posixpath
from typing import IO, TYPE_CHECKING, cast
from collections.abc import Iterator

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
        """
        from docx.oxml.parser import huge_tree_mode, recovery_mode

        def _load() -> Self:
            pkg_reader = PackageReader.from_file(pkg_file)
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

    def save(self, pkg_file: str | IO[bytes]):
        """Save this package to `pkg_file`.

        `pkg_file` can be either a file-path or a file-like object.

        When `pkg_file` is a string, the filename component is validated against
        the characters Windows disallows in filenames (``< > : " | ? *``). A
        mismatch raises :class:`OSError`. This avoids the silently-truncated /
        no-file-created failure mode reported in upstream#1111 when callers
        pass e.g. ``"my:file.docx"`` as the save target.
        """
        if isinstance(pkg_file, str):
            _validate_save_path(pkg_file)
        for part in self.parts:
            part.before_marshal()
        PackageWriter.write(pkg_file, self.rels, self.parts)

    @property
    def _core_properties_part(self) -> CorePropertiesPart:
        """|CorePropertiesPart| object related to this package.

        Creates a default core properties part if one is not present (not common).
        Honours both the canonical ``package/2006`` core-properties reltype and
        the ``officeDocument/2006`` alternate emitted by some producers
        (upstream-PR#1436) so a duplicate ``docProps/core.xml`` isn't created
        when the existing rel uses the alternate form.
        """
        try:
            return cast(CorePropertiesPart, self.part_related_by(RT.CORE_PROPERTIES))
        except KeyError:
            pass
        try:
            return cast(CorePropertiesPart, self.part_related_by(RT.CORE_PROPERTIES_ALT))
        except KeyError:
            core_properties_part = CorePropertiesPart.default(self)
            self.relate_to(core_properties_part, RT.CORE_PROPERTIES)
            return core_properties_part


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
        `part_factory`.
        """
        parts = {}
        for partname, content_type, reltype, blob in pkg_reader.iter_sparts():
            parts[partname] = part_factory(partname, content_type, reltype, blob, package)
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
