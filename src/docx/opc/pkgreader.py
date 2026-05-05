"""Low-level, read-only API to a serialized Open Packaging Convention (OPC) package."""

from docx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from docx.opc.exceptions import PackageNotFoundError
from docx.opc.oxml import parse_xml
from docx.opc.packuri import PACKAGE_URI, PackURI
from docx.opc.phys_pkg import PhysPkgReader, _looks_like_strict_package
from docx.opc.shared import CaseInsensitiveDict


class PackageReader:
    """Provides access to the contents of a zip-format OPC package via its
    :attr:`serialized_parts` and :attr:`pkg_srels` attributes."""

    def __init__(self, content_types, pkg_srels, sparts):
        super().__init__()
        self._pkg_srels = pkg_srels
        self._sparts = sparts

    @staticmethod
    def from_file(pkg_file, password: str | None = None):
        """Return a |PackageReader| instance loaded with contents of `pkg_file`.

        If `pkg_file` is a Flat-OPC ``<pkg:package>`` XML file, it is expanded
        to an in-memory zip first so the normal reader path handles it.
        Strict-OOXML packages are transparently translated to Transitional
        as blobs flow through the physical reader. Closes upstream#892,
        upstream#1520, upstream#693.

        When `password` is provided and `pkg_file` is an ECMA-376 Agile-
        Encryption container (CFBF / OLE2), it is transparently decrypted via
        the optional ``python-ooxml-crypto`` dependency before being read.

        .. versionchanged:: 2026.05.10
           Added ``password`` parameter.
        """
        from docx.opc.flat_opc import (
            expand_flat_opc_to_zip_stream,
            looks_like_flat_opc,
        )
        from docx.opc.phys_pkg import _StrictTranslatingPkgReader

        if looks_like_flat_opc(pkg_file):
            pkg_file = expand_flat_opc_to_zip_stream(pkg_file)
        phys_reader = PhysPkgReader(pkg_file, password=password)
        if _looks_like_strict_package(phys_reader):
            phys_reader = _StrictTranslatingPkgReader(phys_reader)
        # -- `[Content_Types].xml` is mandatory per OPC §9.2; a zip that lacks it
        # -- is not a valid OOXML package. `ZipFile.read` raises a bare `KeyError`
        # -- whose message is the missing member name, which leaked through to
        # -- callers as an opaque `KeyError('[Content_Types].xml')`. Wrap it in
        # -- the typed `PackageNotFoundError` so upstream handlers can catch the
        # -- OPC-level failure without matching on exception message. Closes #172.
        try:
            content_types_blob = phys_reader.content_types_xml
        except KeyError as e:
            raise PackageNotFoundError(
                "Package is not a valid OPC file: missing [Content_Types].xml"
            ) from e
        content_types = _ContentTypeMap.from_xml(content_types_blob)
        pkg_srels = PackageReader._srels_for(phys_reader, PACKAGE_URI)
        sparts = PackageReader._load_serialized_parts(phys_reader, pkg_srels, content_types)
        phys_reader.close()
        return PackageReader(content_types, pkg_srels, sparts)

    def iter_sparts(self):
        """Generate a 4-tuple `(partname, content_type, reltype, blob)` for each of the
        serialized parts in the package."""
        for s in self._sparts:
            yield (s.partname, s.content_type, s.reltype, s.blob)

    def iter_srels(self):
        """Generate a 2-tuple `(source_uri, srel)` for each of the relationships in the
        package."""
        for srel in self._pkg_srels:
            yield (PACKAGE_URI, srel)
        for spart in self._sparts:
            for srel in spart.srels:
                yield (spart.partname, srel)

    @staticmethod
    def _load_serialized_parts(phys_reader, pkg_srels, content_types):
        """Return a list of |_SerializedPart| instances corresponding to the parts in
        `phys_reader` accessible by walking the relationship graph starting with
        `pkg_srels`."""
        sparts = []
        part_walker = PackageReader._walk_phys_parts(phys_reader, pkg_srels)
        for partname, blob, reltype, srels in part_walker:
            content_type = content_types[partname]
            spart = _SerializedPart(partname, content_type, reltype, blob, srels)
            sparts.append(spart)
        return tuple(sparts)

    @staticmethod
    def _srels_for(phys_reader, source_uri):
        """Return |_SerializedRelationships| instance populated with relationships for
        source identified by `source_uri`."""
        rels_xml = phys_reader.rels_xml_for(source_uri)
        return _SerializedRelationships.load_from_xml(source_uri.baseURI, rels_xml)

    @staticmethod
    def _walk_phys_parts(phys_reader, srels, visited_partnames=None):
        """Generate a 4-tuple `(partname, blob, reltype, srels)` for each of the parts
        in `phys_reader` by walking the relationship graph rooted at srels."""
        if visited_partnames is None:
            visited_partnames = []
        for srel in srels:
            if srel.is_external:
                continue
            # -- Skip relationships whose target is a pure in-document fragment
            # -- (e.g. "#bookmark1") or a NULL/empty target. Such relationships
            # -- describe internal bookmark hyperlinks and don't refer to any
            # -- package part. upstream#902, #1349, #678, PR#1498, PR#1518, PR#1350.
            target_ref = srel.target_ref
            if not target_ref or target_ref.startswith("#"):
                continue
            partname = srel.target_partname
            if partname in visited_partnames:
                continue
            visited_partnames.append(partname)
            reltype = srel.reltype
            part_srels = PackageReader._srels_for(phys_reader, partname)
            # -- Tolerate rels that point at a part which is not present in the
            # -- package. Word itself emits such "dangling" rels in some loose
            # -- / partially-repaired documents; we skip them so the rest of
            # -- the package still loads. upstream-PR#1219.
            try:
                blob = phys_reader.blob_for(partname)
            except KeyError:
                continue
            yield (partname, blob, reltype, part_srels)
            next_walker = PackageReader._walk_phys_parts(phys_reader, part_srels, visited_partnames)
            for partname, blob, reltype, srels in next_walker:
                yield (partname, blob, reltype, srels)


class _ContentTypeMap:
    """Value type providing dictionary semantics for looking up content type by part
    name, e.g. ``content_type = cti['/ppt/presentation.xml']``."""

    def __init__(self):
        super().__init__()
        self._overrides = CaseInsensitiveDict()
        self._defaults = CaseInsensitiveDict()

    def __getitem__(self, partname):
        """Return content type for part identified by `partname`."""
        if not isinstance(partname, PackURI):
            tmpl = "_ContentTypeMap key must be <type 'PackURI'>, got %s"
            raise KeyError(tmpl % type(partname))
        if partname in self._overrides:
            return self._overrides[partname]
        if partname.ext in self._defaults:
            return self._defaults[partname.ext]
        tmpl = "no content type for partname '%s' in [Content_Types].xml"
        raise KeyError(tmpl % partname)

    @staticmethod
    def from_xml(content_types_xml):
        """Return a new |_ContentTypeMap| instance populated with the contents of
        `content_types_xml`."""
        types_elm = parse_xml(content_types_xml)
        ct_map = _ContentTypeMap()
        for o in types_elm.overrides:
            ct_map._add_override(o.partname, o.content_type)
        for d in types_elm.defaults:
            ct_map._add_default(d.extension, d.content_type)
        return ct_map

    def _add_default(self, extension, content_type):
        """Add the default mapping of `extension` to `content_type` to this content type
        mapping."""
        self._defaults[extension] = content_type

    def _add_override(self, partname, content_type):
        """Add the default mapping of `partname` to `content_type` to this content type
        mapping."""
        self._overrides[partname] = content_type


class _SerializedPart:
    """Value object for an OPC package part.

    Provides access to the partname, content type, blob, and serialized relationships
    for the part.
    """

    def __init__(self, partname, content_type, reltype, blob, srels):
        super().__init__()
        self._partname = partname
        self._content_type = content_type
        self._reltype = reltype
        self._blob = blob
        self._srels = srels

    @property
    def partname(self):
        return self._partname

    @property
    def content_type(self):
        return self._content_type

    @property
    def blob(self):
        return self._blob

    @property
    def reltype(self):
        """The referring relationship type of this part."""
        return self._reltype

    @property
    def srels(self):
        return self._srels


class _SerializedRelationship:
    """Value object representing a serialized relationship in an OPC package.

    Serialized, in this case, means any target part is referred to via its partname
    rather than a direct link to an in-memory |Part| object.
    """

    def __init__(self, baseURI, rel_elm):
        super().__init__()
        self._baseURI = baseURI
        self._rId = rel_elm.rId
        self._reltype = rel_elm.reltype
        self._target_mode = rel_elm.target_mode
        # -- Normalise Windows-style backslashes to forward slashes. Some
        # -- third-party DOCX producers write `Target="media\image1.png"`
        # -- which, left as-is, breaks `PackURI.from_rel_ref()`'s posixpath
        # -- join and yields a bogus part-name. upstream-PR#1205.
        target_ref = rel_elm.target_ref
        if target_ref is not None and "\\" in target_ref:
            target_ref = target_ref.replace("\\", "/")
        self._target_ref = target_ref

    @property
    def is_external(self):
        """True if target_mode is ``RTM.EXTERNAL``"""
        return self._target_mode == RTM.EXTERNAL

    @property
    def reltype(self):
        """Relationship type, like ``RT.OFFICE_DOCUMENT``"""
        return self._reltype

    @property
    def rId(self):
        """Relationship id, like 'rId9', corresponds to the ``Id`` attribute on the
        ``CT_Relationship`` element."""
        return self._rId

    @property
    def target_mode(self):
        """String in ``TargetMode`` attribute of ``CT_Relationship`` element, one of
        ``RTM.INTERNAL`` or ``RTM.EXTERNAL``."""
        return self._target_mode

    @property
    def target_ref(self):
        """String in ``Target`` attribute of ``CT_Relationship`` element, a relative
        part reference for internal target mode or an arbitrary URI, e.g. an HTTP URL,
        for external target mode."""
        return self._target_ref

    @property
    def target_partname(self):
        """|PackURI| instance containing partname targeted by this relationship.

        Raises ``ValueError`` on reference if target_mode is ``'External'``. Use
        :attr:`target_mode` to check before referencing.
        """
        if self.is_external:
            msg = (
                "target_partname attribute on Relationship is undefined w"
                'here TargetMode == "External"'
            )
            raise ValueError(msg)
        # lazy-load _target_partname attribute
        if not hasattr(self, "_target_partname"):
            self._target_partname = PackURI.from_rel_ref(self._baseURI, self.target_ref)
        return self._target_partname


class _SerializedRelationships:
    """Read-only sequence of |_SerializedRelationship| instances corresponding to the
    relationships item XML passed to constructor."""

    def __init__(self):
        super().__init__()
        self._srels = []

    def __iter__(self):
        """Support iteration, e.g. 'for x in srels:'."""
        return self._srels.__iter__()

    @staticmethod
    def load_from_xml(baseURI, rels_item_xml):
        """Return |_SerializedRelationships| instance loaded with the relationships
        contained in `rels_item_xml`.

        Returns an empty collection if `rels_item_xml` is |None|.
        """
        srels = _SerializedRelationships()
        if rels_item_xml is not None:
            rels_elm = parse_xml(rels_item_xml)
            for rel_elm in rels_elm.Relationship_lst:
                srels._srels.append(_SerializedRelationship(baseURI, rel_elm))
        return srels
