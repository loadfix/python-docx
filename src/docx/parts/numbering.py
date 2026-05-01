"""|NumberingPart| and closely related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from ..opc.constants import CONTENT_TYPE as CT
from ..opc.packuri import PackURI
from ..opc.part import XmlPart
from ..oxml.ns import nsdecls
from ..oxml.parser import parse_xml
from ..shared import lazyproperty

if TYPE_CHECKING:
    from ..oxml.numbering import CT_Numbering
    from ..package import Package


class NumberingPart(XmlPart):
    """Proxy for the numbering.xml part containing numbering definitions for a document
    or glossary."""

    @classmethod
    def new(cls) -> "NumberingPart":
        """Newly created numbering part, containing only the root ``<w:numbering>`` element."""
        # -- preserved for backwards compatibility with callers that invoke
        # -- `NumberingPart.new()` without a package. Creates an orphan part; use
        # -- :meth:`default` when a package is available.
        partname = PackURI("/word/numbering.xml")
        content_type = CT.WML_NUMBERING
        element = cast("CT_Numbering", parse_xml(cls._default_numbering_xml()))
        return cls(partname, content_type, element, None)  # type: ignore[arg-type]

    @classmethod
    def default(cls, package: "Package") -> Self:
        """A newly created numbering part containing an empty ``w:numbering`` root.

        This follows the same lazy-creation pattern used by the footnotes and
        comments parts.
        """
        partname = PackURI("/word/numbering.xml")
        content_type = CT.WML_NUMBERING
        element = cast("CT_Numbering", parse_xml(cls._default_numbering_xml()))
        return cls(partname, content_type, element, package)

    @staticmethod
    def _default_numbering_xml() -> bytes:
        """Minimal ``w:numbering`` XML used as the body of a freshly created part."""
        return f'<w:numbering {nsdecls("w")}/>'.encode("utf-8")

    @property
    def numbering(self):
        """Return the |Numbering| proxy object for this part."""
        from ..numbering import Numbering

        return Numbering(cast("CT_Numbering", self._element), self)

    @property
    def numbering_element(self) -> "CT_Numbering":
        """The root ``w:numbering`` element for this part."""
        return cast("CT_Numbering", self._element)

    @lazyproperty
    def numbering_definitions(self):
        """The |_NumberingDefinitions| instance containing the numbering definitions
        (<w:num> element proxies) for this numbering part."""
        return _NumberingDefinitions(self._element)


class _NumberingDefinitions:
    """Collection of |_NumberingDefinition| instances corresponding to the ``<w:num>``
    elements in a numbering part."""

    def __init__(self, numbering_elm):
        super().__init__()
        self._numbering = numbering_elm

    def __len__(self):
        return len(self._numbering.num_lst)
