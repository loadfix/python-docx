"""Provides low-level, write-only API to serialized (OPC) package.

OPC stands for Open Packaging Convention. This is essentially an implementation
of :meth:`OpcPackage.save`.

Kept docx-local rather than re-exporting :mod:`ooxml_opc.pkgwriter` because the
two are not API-compatible: docx's :class:`PackageWriter` exposes static
methods (``write`` / ``_write_content_types_stream`` / ``_write_pkg_rels`` /
``_write_parts``) that are individually unit-tested, while the shared
runtime's :class:`~ooxml_opc.pkgwriter.PackageWriter` is instance-based with
a ``zip_date_time`` knob. Both paths depend on the same shared primitives.

Adoption is effective via the transitive graph: :class:`~docx.opc.oxml.CT_Types`
and :func:`~docx.opc.oxml.serialize_part_xml` come from :mod:`ooxml_opc.oxml`;
:class:`~docx.opc.phys_pkg.PhysPkgWriter` is the docx-shape wrapper around
:class:`ooxml_opc.phys_pkg`'s reproducible-zip primitives.

.. versionchanged:: 2026.05.11
   Uses the shared runtime's xml serialisation + physical-writer primitives;
   the outward docx :class:`PackageWriter` static-method API is unchanged.
"""

from __future__ import annotations

import io
from typing import TYPE_CHECKING
from collections.abc import Iterable

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.oxml import CT_Types, serialize_part_xml
from docx.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI
from docx.opc.phys_pkg import PhysPkgWriter
from docx.opc.shared import CaseInsensitiveDict
from docx.opc.spec import default_content_types

if TYPE_CHECKING:
    from docx.opc.part import Part


class PackageWriter:
    """Writes a zip-format OPC package to `pkg_file`, where `pkg_file` can be either a
    path to a zip file (a string) or a file-like object.

    Its single API method, :meth:`write`, is static, so this class is not intended to be
    instantiated.
    """

    @staticmethod
    def write(
        pkg_file,
        pkg_rels,
        parts,
        reproducible: bool = False,
        password: str | None = None,
    ):
        """Write a physical package (.docx file) to `pkg_file` containing `pkg_rels` and
        `parts` and a content types stream based on the content types of the parts.

        When `reproducible` is True, the underlying zip writer emits fixed
        timestamps and sorted member names so repeated saves of the same input
        produce byte-identical output. Closes upstream#1042 / upstream-PR#810.

        When `password` is provided the resulting .docx is password-protected
        using ECMA-376 Agile Encryption; this requires the optional
        ``python-ooxml-crypto`` dependency. ``reproducible`` and ``password``
        are orthogonal — fixed timestamps stamp the inner (plaintext) zip
        members before the encryption wrapper is applied.

        .. versionadded:: 2026.05.0
           The `reproducible` parameter.
        .. versionadded:: 2026.05.10
           The `password` parameter.
        """
        if password is None:
            phys_writer = PhysPkgWriter(pkg_file, reproducible=reproducible)
            PackageWriter._write_content_types_stream(phys_writer, parts)
            PackageWriter._write_pkg_rels(phys_writer, pkg_rels)
            PackageWriter._write_parts(phys_writer, parts, reproducible=reproducible)
            phys_writer.close()
            return

        # -- build the plain zip into an in-memory buffer, then encrypt it --
        from docx.opc._crypto import encrypt_bytes

        plain_buffer = io.BytesIO()
        phys_writer = PhysPkgWriter(plain_buffer, reproducible=reproducible)
        PackageWriter._write_content_types_stream(phys_writer, parts)
        PackageWriter._write_pkg_rels(phys_writer, pkg_rels)
        PackageWriter._write_parts(phys_writer, parts, reproducible=reproducible)
        phys_writer.close()

        encrypted = encrypt_bytes(plain_buffer.getvalue(), password)

        if isinstance(pkg_file, str):
            with open(pkg_file, "wb") as f:
                f.write(encrypted)
        else:
            pkg_file.write(encrypted)

    @staticmethod
    def _write_content_types_stream(phys_writer, parts):
        """Write ``[Content_Types].xml`` part to the physical package with an
        appropriate content type lookup target for each part in `parts`."""
        cti = _ContentTypesItem.from_parts(parts)
        phys_writer.write(CONTENT_TYPES_URI, cti.blob)

    @staticmethod
    def _write_parts(
        phys_writer: PhysPkgWriter,
        parts: Iterable[Part],
        reproducible: bool = False,
    ):
        """Write the blob of each part in `parts` to the package, along with a rels item
        for its relationships if and only if it has any.

        When `reproducible` is True, parts are emitted in sorted partname order.
        The underlying reproducible writer also sorts internally before flushing,
        so this is mostly for tidy debugging of write order.
        """
        if reproducible:
            parts = sorted(parts, key=lambda p: p.partname)
        for part in parts:
            phys_writer.write(part.partname, part.blob)
            if len(part.rels):
                phys_writer.write(part.partname.rels_uri, part.rels.xml)

    @staticmethod
    def _write_pkg_rels(phys_writer, pkg_rels):
        """Write the XML rels item for `pkg_rels` ('/_rels/.rels') to the package."""
        phys_writer.write(PACKAGE_URI.rels_uri, pkg_rels.xml)


class _ContentTypesItem:
    """Service class that composes a content types item ([Content_Types].xml) based on a
    list of parts.

    Not meant to be instantiated directly, its single interface method is xml_for(),
    e.g. ``_ContentTypesItem.xml_for(parts)``.
    """

    def __init__(self):
        self._defaults = CaseInsensitiveDict()
        self._overrides = {}

    @property
    def blob(self):
        """Return XML form of this content types item, suitable for storage as
        ``[Content_Types].xml`` in an OPC package."""
        return serialize_part_xml(self._element)

    @classmethod
    def from_parts(cls, parts):
        """Return content types XML mapping each part in `parts` to the appropriate
        content type and suitable for storage as ``[Content_Types].xml`` in an OPC
        package."""
        cti = cls()
        cti._defaults["rels"] = CT.OPC_RELATIONSHIPS
        cti._defaults["xml"] = CT.XML
        for part in parts:
            cti._add_content_type(part.partname, part.content_type)
        return cti

    def _add_content_type(self, partname, content_type):
        """Add a content type for the part with `partname` and `content_type`, using a
        default or override as appropriate."""
        ext = partname.ext
        if (ext.lower(), content_type) in default_content_types:
            self._defaults[ext] = content_type
        else:
            self._overrides[partname] = content_type

    @property
    def _element(self):
        """Return XML form of this content types item, suitable for storage as
        ``[Content_Types].xml`` in an OPC package.

        Although the sequence of elements is not strictly significant, as an aid to
        testing and readability Default elements are sorted by extension and Override
        elements are sorted by partname.
        """
        _types_elm = CT_Types.new()
        for ext in sorted(self._defaults.keys()):
            _types_elm.add_default(ext, self._defaults[ext])
        for partname in sorted(self._overrides.keys()):
            _types_elm.add_override(partname, self._overrides[partname])
        return _types_elm
