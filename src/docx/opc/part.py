# pyright: reportImportCycles=false

"""Re-export of :mod:`ooxml_opc.part` with docx-shape arg-order adapters.

The shared :class:`~ooxml_opc.part.Part` / :class:`~ooxml_opc.part.XmlPart` /
:class:`~ooxml_opc.part.PartFactory` constructor shapes follow the pptx
convention ``(partname, content_type, package, blob)``. docx historically
passes ``(partname, content_type, blob, package)`` so every subclass under
``docx.parts.*`` would break without an adapter layer.

This shim subclasses the shared runtime classes and re-declares their
``__init__`` / ``load`` / ``__new__`` signatures in docx shape. Where docx
callers or test fixtures patch module-local names
(``docx.opc.part.parse_xml``, ``docx.opc.part.serialize_part_xml``,
``docx.opc.part.cls_method_fn``, ``docx.opc.part.Relationships``), the
relevant calls go through those names so :func:`function_mock` and
:func:`class_mock` continue to work unchanged.

.. versionchanged:: 2026.05.11
   Re-exported from :mod:`ooxml_opc.part`; docx-shape arg order preserved.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, cast

from ooxml_opc.part import Part as _SharedPart
from ooxml_opc.part import PartFactory as _SharedPartFactory
from ooxml_opc.part import PartRelationshipCloner  # noqa: F401 -- re-export
from ooxml_opc.part import XmlPart as _SharedXmlPart

from docx.opc.oxml import serialize_part_xml
from docx.opc.packuri import PackURI
from docx.opc.rel import Relationships
from docx.opc.shared import cls_method_fn
from docx.oxml.parser import parse_xml
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.package import Package

__all__ = ["Part", "PartFactory", "PartRelationshipCloner", "XmlPart"]


class Part(_SharedPart):
    """docx-shape wrapper around :class:`ooxml_opc.part.Part`.

    docx constructor shape is ``(partname, content_type, blob=None, package=None)``.
    The shared base uses ``(partname, content_type, package, blob=None)``. The
    shim's :meth:`__init__` swaps the last two positional args before forwarding.
    """

    #: Flag set to ``True`` by the package loader when the part was parsed from
    #: the source package. Library-authored parts leave this at ``False``.
    #:
    #: .. versionadded:: 2026.05.7
    _loaded_from_package: bool = False

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        blob: bytes | None = None,
        package: Package | None = None,
    ):
        # -- docx arg order → shared arg order --
        super().__init__(partname, content_type, package, blob=blob)  # type: ignore[arg-type]

    def after_unmarshal(self) -> None:
        """Post-unmarshal hook (override without forwarding to super)."""
        return

    def before_marshal(self, reproducible: bool = False) -> None:
        """Pre-serialisation hook (override without forwarding to super)."""
        return

    @property
    def blob(self) -> bytes:
        """Contents of this part as bytes.

        Overridden by :class:`XmlPart` — this implementation returns the
        load-time blob (or ``b""`` when there is none).
        """
        return self._blob or b""

    def drop_rel(self, rId: str) -> None:
        """Remove the relationship `rId` if its reference count is < 2."""
        if self._rel_ref_count(rId) < 2:
            del self.rels[rId]

    @classmethod
    def load(
        cls,
        partname: PackURI,
        content_type: str,
        blob: bytes,
        package: Package,
    ) -> Part:
        """Return a new instance of `cls`.

        docx-shape ``(partname, content_type, blob, package)``. The shared
        :meth:`ooxml_opc.part.Part.load` takes a ``(partname, content_type,
        package, blob)`` tuple; the shim adapts the call.
        """
        return cls(partname, content_type, blob, package)

    def load_rel(
        self,
        reltype: str,
        target: Part | str,
        rId: str,
        is_external: bool = False,
    ):
        """Mint a relationship with a caller-supplied `rId`."""
        return self.rels.add_relationship(reltype, target, rId, is_external)

    @property
    def partname(self) -> PackURI:
        """``PackURI`` identifying this part."""
        return self._partname

    @partname.setter
    def partname(self, partname: PackURI) -> None:
        if not isinstance(partname, PackURI):
            raise TypeError(
                f"partname must be instance of PackURI, got "
                f"'{type(partname).__name__}'"
            )
        self._partname = partname

    def part_related_by(self, reltype: str) -> Part:
        """Return the part this part has a `reltype` relationship to."""
        return self.rels.part_with_reltype(reltype)

    def relate_to(
        self,
        target: Part | str,
        reltype: str,
        is_external: bool = False,
    ) -> str:
        """Return the rId of a relationship of `reltype` to `target`."""
        if is_external:
            return self.rels.get_or_add_ext_rel(reltype, cast(str, target))
        rel = self.rels.get_or_add(reltype, cast(Part, target))
        return rel.rId

    @property
    def related_parts(self) -> dict[str, Part]:
        """Dict mapping rIds to target parts for explicit internal rels."""
        return self.rels.related_parts

    @lazyproperty
    def rels(self):  # type: ignore[override]
        """:class:`Relationships` collection for this part.

        Reference `Relationships` via the module-level name so tests that
        patch ``docx.opc.part.Relationships`` via :func:`class_mock` see the
        patched constructor.
        """
        return Relationships(self._partname.baseURI)

    def target_ref(self, rId: str) -> str:
        """Return the URL contained in the target ref of relationship `rId`."""
        rel = self.rels[rId]
        return rel.target_ref

    def _rel_ref_count(self, rId: str) -> int:
        """Return the count of references in this part to rel `rId`.

        Non-XML parts cannot contain references; the generic implementation
        returns 0. :class:`XmlPart` overrides.
        """
        return 0


class PartFactory(_SharedPartFactory):
    """docx-shape wrapper around :class:`ooxml_opc.part.PartFactory`.

    docx constructor shape is
    ``(partname, content_type, reltype, blob, package)``; the shared factory
    uses ``(partname, content_type, package, blob, reltype=None)``.

    Keeps class-level ``part_class_selector`` / ``part_type_for`` /
    ``default_part_type`` for registration at import time.
    """

    part_class_selector: Callable[[str, str], type[Part] | None] | None = None  # type: ignore[assignment]
    part_type_for: dict[str, type[Part]] = {}  # type: ignore[assignment]
    default_part_type: type[Part] = Part  # type: ignore[assignment]

    def __new__(  # type: ignore[override]
        cls,
        partname: PackURI,
        content_type: str,
        reltype: str,
        blob: bytes,
        package: Package,
    ) -> Part:
        PartClass: type[Part] | None = None
        if cls.part_class_selector is not None:
            # -- Lookup ``cls_method_fn`` via the module namespace so tests can
            # -- patch ``docx.opc.part.cls_method_fn``. --
            selector = cls_method_fn(cls, "part_class_selector")
            PartClass = selector(content_type, reltype)
        if PartClass is None:
            PartClass = cls._part_cls_for(content_type)
        return PartClass.load(partname, content_type, blob, package)

    @classmethod
    def _part_cls_for(cls, content_type: str) -> type[Part]:
        """Return the Part subclass registered for `content_type`."""
        if content_type in cls.part_type_for:
            return cls.part_type_for[content_type]
        return cls.default_part_type


class XmlPart(_SharedXmlPart, Part):
    """docx-shape wrapper around :class:`ooxml_opc.part.XmlPart`.

    docx constructor shape is
    ``(partname, content_type, element, package)``; the shared base uses
    ``(partname, content_type, package, element)``. The shim adapts.
    """

    def __init__(  # type: ignore[override]
        self,
        partname: PackURI,
        content_type: str,
        element: BaseOxmlElement,
        package: Package,
    ):
        # -- bypass both parents' __init__; wire up state directly so the
        # -- (deliberately inconsistent) arg orders don't collide. --
        self._partname = partname
        self._content_type = content_type
        self._blob = None
        self._package = package
        self._element = element

    @property
    def blob(self) -> bytes:  # type: ignore[override]
        """Re-serialise ``self._element`` to bytes on demand."""
        return serialize_part_xml(self._element)

    @property
    def element(self) -> BaseOxmlElement:
        """The root XML element of this XML part."""
        return self._element

    @classmethod
    def load(  # type: ignore[override]
        cls,
        partname: PackURI,
        content_type: str,
        blob: bytes,
        package: Package,
    ) -> XmlPart:
        """Return an :class:`XmlPart` with its blob parsed into an element tree.

        ``parse_xml`` may return ``None`` under recovery mode when lxml
        couldn't recover anything from `blob`; fall back to a minimal stub
        element so downstream code always has a valid tree.
        """
        element = cast("BaseOxmlElement | None", parse_xml(blob))
        if element is None:
            element = cls._recovery_stub_element(content_type)
        return cls(partname, content_type, element, package)

    @staticmethod
    def _recovery_stub_element(content_type: str) -> BaseOxmlElement:
        """Return a minimal valid root element for the given `content_type`.

        Used as a last-resort fallback when recovery parsing cannot produce any
        usable element. Subclasses may override to emit a more specific stub.
        """
        from docx.opc.constants import CONTENT_TYPE as CT

        if content_type in (
            CT.WML_DOCUMENT_MAIN,
            CT.WML_DOCUMENT_MACRO,
            CT.WML_TEMPLATE_MAIN,
            CT.WML_TEMPLATE_MACRO,
        ):
            stub = (
                b'<w:document xmlns:w="http://schemas.openxmlformats.org/'
                b'wordprocessingml/2006/main"><w:body/></w:document>'
            )
            return parse_xml(stub)
        return parse_xml(b"<root/>")

    @property
    def part(self) -> XmlPart:
        """Self-reference for the parent-protocol chain.

        Children of an XML part ask their parent for the containing part;
        delegation terminates here.
        """
        return self

    def _rel_ref_count(self, rId: str) -> int:
        """Count references in this part's XML to rel `rId`."""
        rIds = cast("list[str]", self._element.xpath("//@r:id"))
        return len([_rId for _rId in rIds if _rId == rId])
