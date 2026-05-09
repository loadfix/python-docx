"""Custom element classes for mail-merge / ODSO (Office Data Source Object).

Mail-merge configuration is stored in ``w:settings/w:mailMerge``. The ``CT_MailMerge``
shell already lives in :mod:`docx.oxml.settings`; this module contributes the ODSO
(Office Data Source Object) substructure plus the small val-wrapper CT classes that
the main-merge shell references by type.

Scope: read/write round-trip of stored configuration. python-docx does not execute a
mail-merge; it simply preserves the authored metadata so Word-side re-open of the
document remains byte-faithful.

.. versionadded:: 2026.05.10
"""

from __future__ import annotations

from typing import Iterator, Optional, Union

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.simpletypes import (
    ST_DecimalNumber,
    ST_OnOff,
    ST_RelationshipId,
    ST_String,
    XsdString,
)
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


# ---------------------------------------------------------------------------
# Shared CT helpers
# ---------------------------------------------------------------------------


class CT_Base64Binary(BaseOxmlElement):
    """Container for a single base64-encoded binary `@w:val` payload.

    Used for ``w:uniqueTag`` inside ``w:recipientData``. Word stores the
    unique identifier for a data-source record here.
    """

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


# ---------------------------------------------------------------------------
# Val-wrapper CTs that the CT_MailMerge shell references
# ---------------------------------------------------------------------------


class CT_MailMergeDocType(BaseOxmlElement):
    """`w:mainDocumentType` — val is one of ``catalog``, ``envelopes``,
    ``mailingLabels``, ``formLetters``, ``email``, ``fax``."""

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_MailMergeDataType(BaseOxmlElement):
    """`w:dataType` — val is one of ``textFile``, ``database``, ``spreadsheet``,
    ``query``, ``odbc``, ``native``."""

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_MailMergeDest(BaseOxmlElement):
    """`w:destination` — val is one of ``newDocument``, ``printer``, ``email``,
    ``fax``."""

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_MailMergeSourceType(BaseOxmlElement):
    """`w:odso/w:type` — val is one of ``database``, ``addressBook``,
    ``document1``, ``document2``, ``text``, ``email``, ``native``, ``legacy``,
    ``master``."""

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_MailMergeOdsoFMDFieldType(BaseOxmlElement):
    """`w:odso/w:fieldMapData/w:type` — val is one of ``null``, ``dbColumn``."""

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


# ---------------------------------------------------------------------------
# Generic data-source / relationship reference
# ---------------------------------------------------------------------------


class CT_DataSourceObject(BaseOxmlElement):
    """Generic relationship-shaped data source reference.

    Used by ``w:mailMerge/w:dataSource`` and ``w:mailMerge/w:headerSource`` (and
    by ``w:odso/w:src`` + ``w:odso/w:recipientData``). The underlying XSD type
    is ``CT_Rel`` — a single ``r:id`` attribute pointing at a package
    relationship whose target holds the ODSO payload.

    Named ``CT_DataSourceObject`` rather than ``CT_Rel`` because docx already
    has several relationship-reference flavors; this name keeps the mail-merge
    intent visible at the import site.
    """

    rId: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:id", ST_RelationshipId
    )


# ---------------------------------------------------------------------------
# ODSO substructure
# ---------------------------------------------------------------------------


class CT_OdsoFieldMapData(BaseOxmlElement):
    """`w:fieldMapData` — one field-mapping record inside ``w:odso``.

    Maps a merge-field name onto a column in the external data source.
    """

    _tag_seq = (
        "w:type",
        "w:name",
        "w:mappedName",
        "w:column",
        "w:lid",
        "w:dynamicAddress",
    )

    type: "CT_MailMergeOdsoFMDFieldType | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:type", successors=_tag_seq[1:]
    )
    name: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:name", successors=_tag_seq[2:]
    )
    mappedName: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:mappedName", successors=_tag_seq[3:]
    )
    column: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:column", successors=_tag_seq[4:]
    )
    lid: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:lid", successors=_tag_seq[5:]
    )
    dynamicAddress: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:dynamicAddress", successors=()
    )

    del _tag_seq


class CT_Odso(BaseOxmlElement):
    """`w:odso` — Office Data Source Object container.

    Describes an OLE DB data source (UDL, table, source, column delimiter,
    source-type enum, optional header flag, an unbounded list of
    field-map-data records, and an unbounded list of recipient-data
    references).
    """

    _tag_seq = (
        "w:udl",
        "w:table",
        "w:src",
        "w:colDelim",
        "w:type",
        "w:fHdr",
        "w:fieldMapData",
        "w:recipientData",
    )

    udl: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:udl", successors=_tag_seq[1:]
    )
    table: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:table", successors=_tag_seq[2:]
    )
    src: "CT_DataSourceObject | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:src", successors=_tag_seq[3:]
    )
    colDelim: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:colDelim", successors=_tag_seq[4:]
    )
    type: "CT_MailMergeSourceType | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:type", successors=_tag_seq[5:]
    )
    fHdr: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:fHdr", successors=_tag_seq[6:]
    )
    fieldMapData = ZeroOrMore("w:fieldMapData", successors=("w:recipientData",))
    recipientData = ZeroOrMore("w:recipientData", successors=())

    del _tag_seq

    # ------------------------------------------------------------------
    # Generic val-child helpers.
    #
    # The child elements registered globally (for example ``w:type``
    # resolves to ``CT_SectType``) can't be relied on for ``.val`` access
    # inside ODSO because their registered type is a different CT. These
    # helpers read and write ``@w:val`` directly on whatever element the
    # descriptor returns — no enum coercion, no side-effects on siblings.
    # ------------------------------------------------------------------

    def _val_child_read(self, tag_local: str) -> Optional[str]:
        el = self.find(qn(f"w:{tag_local}"))
        if el is None:
            return None
        return el.get(qn("w:val"))

    def _val_child_write(
        self,
        tag_local: str,
        value: Optional[Union[str, int]],
    ) -> None:
        if value is None:
            remover = getattr(self, f"_remove_{tag_local}", None)
            if remover is not None:
                remover()
            else:
                el = self.find(qn(f"w:{tag_local}"))
                if el is not None:
                    self.remove(el)
            return
        adder = getattr(self, f"get_or_add_{tag_local}")
        el = adder()
        el.set(qn("w:val"), str(value))

    # ------------------------------------------------------------------
    # Field-map helpers — convert the unbounded ``w:fieldMapData`` list
    # into the dict-shaped view the proxy layer wants.
    # ------------------------------------------------------------------

    def iter_field_map(self) -> Iterator["tuple[str, str]"]:
        """Yield ``(name, mappedName)`` pairs for each ``w:fieldMapData`` child.

        Records missing a ``w:name`` or ``w:mappedName`` grand-child are
        skipped — the mapping is only meaningful when both sides are
        present.
        """
        for fmd in self.findall(qn("w:fieldMapData")):
            name_el = fmd.find(qn("w:name"))
            mapped_el = fmd.find(qn("w:mappedName"))
            if name_el is None or mapped_el is None:
                continue
            name = name_el.get(qn("w:val"))
            mapped = mapped_el.get(qn("w:val"))
            if name is None or mapped is None:
                continue
            yield name, mapped

    def set_field_map(self, mapping: "dict[str, str]") -> None:
        """Replace the current ``w:fieldMapData`` list with `mapping`.

        Existing ``w:fieldMapData`` children are removed and a fresh child is
        emitted for each `(name, mappedName)` pair. Order of iteration in
        `mapping` is preserved in the resulting XML.
        """
        for fmd in list(self.findall(qn("w:fieldMapData"))):
            self.remove(fmd)
        for name, mapped in mapping.items():
            fmd = self.add_fieldMapData()
            name_el = OxmlElement("w:name")
            name_el.set(qn("w:val"), name)
            mapped_el = OxmlElement("w:mappedName")
            mapped_el.set(qn("w:val"), mapped)
            fmd.append(name_el)
            fmd.append(mapped_el)


# ---------------------------------------------------------------------------
# Recipient subset — top-level `w:recipients` part
# ---------------------------------------------------------------------------


class CT_RecipientData(BaseOxmlElement):
    """`w:recipientData` — a single recipient row inside ``w:recipients``.

    Records the column index, the unique tag, and whether the recipient is
    active in the merge.
    """

    _tag_seq = ("w:active", "w:column", "w:uniqueTag")

    active: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:active", successors=_tag_seq[1:]
    )
    column: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:column", successors=_tag_seq[2:]
    )
    uniqueTag: "CT_Base64Binary | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:uniqueTag", successors=()
    )

    del _tag_seq


class CT_OdsoRecipientData(BaseOxmlElement):
    """`w:recipients` — top-level ODSO recipient-subset part.

    Holds one or more ``w:recipientData`` children. Named
    ``CT_OdsoRecipientData`` to mirror the mail-merge roadmap terminology; the
    XSD calls this ``CT_Recipients``.
    """

    recipientData = ZeroOrMore("w:recipientData", successors=())


# ---------------------------------------------------------------------------
# Target screen size (not strictly ODSO, but groups here with mail-merge
# settings family — `w:targetScreenSz` is a settings-root child)
# ---------------------------------------------------------------------------


class CT_TargetScreenSz(BaseOxmlElement):
    """`w:targetScreenSz` — preferred target screen resolution for printing.

    Val is one of ``544x376``, ``640x480``, ``720x512``, ``800x600``,
    ``1024x768``, ``1152x882``, ``1152x900``, ``1280x1024``, ``1600x1200``,
    ``1800x1440``, ``1920x1200``.
    """

    val: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", XsdString
    )
