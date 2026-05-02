"""WordprocessingML Package class and related objects."""

from __future__ import annotations

from typing import IO, cast

from docx.image.image import Image
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.package import OpcPackage
from docx.opc.packuri import PackURI
from docx.parts.image import ImagePart
from docx.shared import lazyproperty
from docx.signatures import SignatureInfo


class Package(OpcPackage):
    """Customizations specific to a WordprocessingML package."""

    def after_unmarshal(self):
        """Called by loading code after all parts and relationships have been loaded.

        This method affords the opportunity for any required post-processing.
        """
        self._gather_image_parts()

    def get_or_add_image_part(self, image_descriptor: str | IO[bytes]) -> ImagePart:
        """Return |ImagePart| containing image specified by `image_descriptor`.

        The image-part is newly created if a matching one is not already present in the
        collection.
        """
        return self.image_parts.get_or_add_image_part(image_descriptor)

    @lazyproperty
    def image_parts(self) -> ImageParts:
        """|ImageParts| collection object for this package."""
        return ImageParts()

    @property
    def is_signed(self) -> bool:
        """True when the package contains at least one digital-signature part.

        Specifically, when a package-level relationship of type
        ``.../digital-signature/origin`` or ``.../digital-signature/signature`` is
        present. python-docx does not verify signatures; this only reports whether
        they are present in the package.

        .. versionadded:: 1.3.0.dev0
        """
        for rel in self.rels.values():
            if rel.is_external:
                continue
            if rel.reltype in (RT.DIGITAL_SIGNATURE_ORIGIN, RT.DIGITAL_SIGNATURE):
                return True
        return False

    @property
    def signatures(self) -> list[SignatureInfo]:
        """List of |SignatureInfo| for each digital signature in the package.

        Returns an empty list for unsigned packages. Signatures are discovered by
        walking from the package-level ``digital-signature/origin`` relationship to
        each ``digital-signature/signature`` relationship on the origin part. If no
        origin part is present, package-level ``digital-signature/signature``
        relationships are used directly as a fallback.

        python-docx does not verify signatures; callers receive read-only metadata
        parsed from the signature XML.

        .. versionadded:: 1.3.0.dev0
        """
        from docx.opc.rel import Relationships as _Relationships

        signatures: list[SignatureInfo] = []
        seen_partnames: set[str] = set()

        def _collect_from(source_rels: _Relationships) -> None:
            for rel in source_rels.values():
                if rel.is_external:
                    continue
                if rel.reltype != RT.DIGITAL_SIGNATURE:
                    continue
                try:
                    target_part = rel.target_part
                except ValueError:
                    continue
                partname_str = str(target_part.partname)
                if partname_str in seen_partnames:
                    continue
                seen_partnames.add(partname_str)
                signatures.append(SignatureInfo(target_part))

        # -- normal path: follow the origin part and enumerate its signature rels --
        for rel in self.rels.values():
            if rel.is_external:
                continue
            if rel.reltype != RT.DIGITAL_SIGNATURE_ORIGIN:
                continue
            origin_part = rel.target_part
            _collect_from(origin_part.rels)

        # -- fallback: packages that declare signature rels at the package level --
        _collect_from(self.rels)

        return signatures

    def _gather_image_parts(self):
        """Load the image part collection with all the image parts in package."""
        for rel in self.iter_rels():
            if rel.is_external:
                continue
            if rel.reltype != RT.IMAGE:
                continue
            if rel.target_part in self.image_parts:
                continue
            self.image_parts.append(cast("ImagePart", rel.target_part))


class ImageParts:
    """Collection of |ImagePart| objects corresponding to images in the package."""

    def __init__(self):
        self._image_parts: list[ImagePart] = []

    def __contains__(self, item: object):
        return self._image_parts.__contains__(item)

    def __iter__(self):
        return self._image_parts.__iter__()

    def __len__(self):
        return self._image_parts.__len__()

    def append(self, item: ImagePart):
        self._image_parts.append(item)

    def get_or_add_image_part(self, image_descriptor: str | IO[bytes]) -> ImagePart:
        """Return |ImagePart| object containing image identified by `image_descriptor`.

        The image-part is newly created if a matching one is not present in the
        collection.
        """
        image = Image.from_file(image_descriptor)
        matching_image_part = self._get_by_sha1(image.sha1)
        if matching_image_part is not None:
            return matching_image_part
        return self._add_image_part(image)

    def _add_image_part(self, image: Image):
        """Return |ImagePart| instance newly created from `image` and appended to the collection."""
        partname = self._next_image_partname(image.ext)
        image_part = ImagePart.from_image(image, partname)
        self.append(image_part)
        return image_part

    def _get_by_sha1(self, sha1: str) -> ImagePart | None:
        """Return the image part in this collection having a SHA1 hash matching `sha1`,
        or |None| if not found."""
        for image_part in self._image_parts:
            if image_part.sha1 == sha1:
                return image_part
        return None

    def _next_image_partname(self, ext: str) -> PackURI:
        """The next available image partname, starting from ``/word/media/image1.{ext}``
        where unused numbers are reused.

        The partname is unique by number, without regard to the extension. `ext` does
        not include the leading period.
        """

        def image_partname(n: int) -> PackURI:
            return PackURI("/word/media/image%d.%s" % (n, ext))

        used_numbers = [image_part.partname.idx for image_part in self]
        for n in range(1, len(self) + 1):
            if n not in used_numbers:
                return image_partname(n)
        return image_partname(len(self) + 1)
