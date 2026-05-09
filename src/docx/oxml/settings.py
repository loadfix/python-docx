"""Custom element classes related to document settings."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator
from collections.abc import Callable

from lxml import etree

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.simpletypes import ST_DecimalNumber, ST_OnOff, ST_String, ST_TwipsMeasure
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.endnotes import CT_EdnDocProps
    from docx.oxml.footnotes import CT_FtnDocProps
    from docx.oxml.mail_merge import CT_DataSourceObject, CT_Odso
    from docx.oxml.shared import CT_OnOff
    from docx.shared import Length


class CT_View(BaseOxmlElement):
    """`w:view` element, specifying the preferred view mode for the document."""

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_Zoom(BaseOxmlElement):
    """`w:zoom` element, specifying the magnification level for the document."""

    percent: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:percent", ST_DecimalNumber
    )


class CT_MailMerge(BaseOxmlElement):
    """`w:mailMerge` element, container for mail-merge configuration metadata.

    Describes the main document type, data source, query, destination, and
    related fields. python-docx does not execute the merge; it only exposes the
    stored configuration for read/write.
    """

    _tag_seq = (
        "w:mainDocumentType",
        "w:linkToQuery",
        "w:dataType",
        "w:connectString",
        "w:query",
        "w:dataSource",
        "w:headerSource",
        "w:doNotSuppressBlankLines",
        "w:destination",
        "w:addressFieldName",
        "w:mailSubject",
        "w:mailAsAttachment",
        "w:viewMergedData",
        "w:activeRecord",
        "w:checkErrors",
        "w:odso",
    )

    mainDocumentType: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:mainDocumentType", successors=_tag_seq[1:]
    )
    linkToQuery: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:linkToQuery", successors=_tag_seq[2:]
    )
    dataType: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:dataType", successors=_tag_seq[3:]
    )
    connectString: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:connectString", successors=_tag_seq[4:]
    )
    query: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:query", successors=_tag_seq[5:]
    )
    dataSource: "CT_DataSourceObject | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:dataSource", successors=_tag_seq[6:]
    )
    headerSource: "CT_DataSourceObject | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:headerSource", successors=_tag_seq[7:]
    )
    doNotSuppressBlankLines: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:doNotSuppressBlankLines", successors=_tag_seq[8:]
    )
    destination: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:destination", successors=_tag_seq[9:]
    )
    addressFieldName: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:addressFieldName", successors=_tag_seq[10:]
    )
    mailSubject: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:mailSubject", successors=_tag_seq[11:]
    )
    mailAsAttachment: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:mailAsAttachment", successors=_tag_seq[12:]
    )
    viewMergedData: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:viewMergedData", successors=_tag_seq[13:]
    )
    activeRecord: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:activeRecord", successors=_tag_seq[14:]
    )
    checkErrors: "_CT_MMVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:checkErrors", successors=_tag_seq[15:]
    )
    odso: "CT_Odso | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:odso", successors=()
    )

    del _tag_seq


class _CT_MMVal(BaseOxmlElement):
    """Shared CT for the many `w:mailMerge` children that only carry a `w:val` attribute.

    Covers `w:mainDocumentType`, `w:dataType`, `w:connectString`, `w:query`,
    `w:destination`, `w:addressFieldName`, `w:mailSubject`, `w:activeRecord`,
    `w:checkErrors`.
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_WriteProtection(BaseOxmlElement):
    """`w:writeProtection` element, specifying the password-to-modify marker.

    Corresponds to ECMA-376 ``CT_WriteProtection``. When present with
    ``@w:recommended="1"`` Word surfaces the "Read-only recommended" banner when
    the document is opened. When the password attributes are populated, Word
    requires that password in order to save back to the same file. This is
    independent of :class:`CT_DocProtect` (``w:documentProtection``), which
    restricts *editing modes* rather than save access.

    .. versionadded:: 2026.05.10
    """

    recommended: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:recommended", ST_OnOff, default=False
    )
    cryptProviderType: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptProviderType", ST_String
    )
    cryptAlgorithmClass: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptAlgorithmClass", ST_String
    )
    cryptAlgorithmType: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptAlgorithmType", ST_String
    )
    cryptAlgorithmSid: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptAlgorithmSid", ST_DecimalNumber
    )
    cryptSpinCount: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptSpinCount", ST_DecimalNumber
    )
    hash: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:hash", ST_String
    )
    salt: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:salt", ST_String
    )


class CT_DocProtect(BaseOxmlElement):
    """`w:documentProtection` element, specifying document editing restrictions."""

    edit: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:edit", ST_String
    )
    formatting: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:formatting", ST_OnOff, default=False
    )
    enforcement: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:enforcement", ST_OnOff, default=False
    )
    cryptProviderType: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptProviderType", ST_String
    )
    cryptAlgorithmClass: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptAlgorithmClass", ST_String
    )
    cryptAlgorithmType: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptAlgorithmType", ST_String
    )
    cryptAlgorithmSid: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptAlgorithmSid", ST_DecimalNumber
    )
    cryptSpinCount: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:cryptSpinCount", ST_DecimalNumber
    )
    hash: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:hash", ST_String
    )
    salt: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:salt", ST_String
    )


class CT_CompatSetting(BaseOxmlElement):
    """`w:compatSetting` element, a single compatibility setting name/value pair."""

    name: str = RequiredAttribute("w:name", ST_String)  # pyright: ignore[reportAssignmentType]
    uri: str = RequiredAttribute("w:uri", ST_String)  # pyright: ignore[reportAssignmentType]
    val: str = RequiredAttribute("w:val", ST_String)  # pyright: ignore[reportAssignmentType]


_COMPAT_SETTING_DEFAULT_URI = "http://schemas.microsoft.com/office/word"


class CT_Compat(BaseOxmlElement):
    """`w:compat` element, containing document compatibility settings."""

    _add_compatSetting: Callable[..., CT_CompatSetting]
    compatSetting_lst: list[CT_CompatSetting]

    compatSetting = ZeroOrMore("w:compatSetting", successors=())

    # --- compatSetting dict-style accessors --------------------------------

    def get_compat_setting(self, name: str) -> str | None:
        """Return the value of the ``w:compatSetting`` with ``@w:name == name``.

        Returns |None| if no such compatSetting is present.
        """
        for setting in self.compatSetting_lst:
            if setting.name == name:
                return setting.val
        return None

    def set_compat_setting(
        self, name: str, value: str, uri: str = _COMPAT_SETTING_DEFAULT_URI
    ) -> None:
        """Set the value of the ``w:compatSetting`` matching `name`.

        If a ``w:compatSetting`` with ``@w:name == name`` exists, its ``@w:val`` is
        updated in place (its ``@w:uri`` is left unchanged). Otherwise a new
        compatSetting is appended using `uri` as the URI.
        """
        for setting in self.compatSetting_lst:
            if setting.name == name:
                setting.val = value
                return
        self._add_compatSetting(name=name, uri=uri, val=value)

    def remove_compat_setting(self, name: str) -> bool:
        """Remove the ``w:compatSetting`` with ``@w:name == name``.

        Returns |True| if a matching element was removed, |False| otherwise.
        """
        removed = False
        for setting in list(self.compatSetting_lst):
            if setting.name == name:
                self.remove(setting)
                removed = True
        return removed

    def iter_compat_setting_names(self) -> Iterator[str]:
        """Yield the ``@w:name`` of each ``w:compatSetting`` child in document order."""
        for setting in self.compatSetting_lst:
            yield setting.name

    # --- direct-child flag accessors ---------------------------------------

    def has_flag(self, name: str) -> bool:
        """Return |True| if direct child element ``w:{name}`` is present.

        ``name`` is the local name (without the ``w:`` prefix). A flag is considered
        "present" regardless of any ``w:val`` attribute value -- merely occurring as a
        child element is enough (the ``w:val`` semantics follow Word's default
        on-when-present convention used for :class:`CT_OnOff` elements).
        """
        return self.find(qn("w:%s" % name)) is not None

    def set_flag(self, name: str, value: bool) -> None:
        """Ensure direct child ``w:{name}`` is present if `value` else absent.

        When `value` is |True|, an empty ``w:{name}`` element is appended if no such
        child is already present (existing children are left unchanged, including any
        ``w:val`` attribute). When `value` is |False|, every direct child matching
        ``w:{name}`` is removed.
        """
        tag = qn("w:%s" % name)
        if value:
            if self.find(tag) is None:
                self.append(OxmlElement("w:%s" % name))
            return
        for child in list(self.findall(tag)):
            self.remove(child)

    def iter_flag_names(self) -> Iterator[str]:
        """Yield the local name of each direct child that is not a ``w:compatSetting``.

        Children are yielded in document order. The ``w:`` prefix is stripped.
        """
        compatSetting_tag = qn("w:compatSetting")
        for child in self:
            if child.tag == compatSetting_tag:
                continue
            yield etree.QName(child.tag).localname  # pyright: ignore[reportArgumentType]

    def clear_flags(self) -> None:
        """Remove every direct child that is not a ``w:compatSetting``."""
        compatSetting_tag = qn("w:compatSetting")
        for child in list(self):
            if child.tag == compatSetting_tag:
                continue
            self.remove(child)


class CT_DefaultTabStop(BaseOxmlElement):
    """`w:defaultTabStop` element, specifying default tab-stop interval."""

    val: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_TwipsMeasure
    )


class CT_LongHexNumber(BaseOxmlElement):
    """`w:rsidRoot` or `w:rsid` element, carrying an 8-char hex RSID in ``@w:val``.

    Both elements share the same content-type in the schema -- a single
    ``@w:val`` attribute holding the 8-digit uppercase hex string identifying a
    Word editing session.
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )


class CT_Rsids(BaseOxmlElement):
    """`w:rsids` element, containing the set of revision-save IDs for the document.

    Contains at most one ``w:rsidRoot`` (the first RSID ever assigned) and zero
    or more ``w:rsid`` children (every RSID used across editing sessions).
    """

    rsidRoot: CT_LongHexNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rsidRoot", successors=("w:rsid",)
    )
    rsid = ZeroOrMore("w:rsid", successors=())
    rsid_lst: list[CT_LongHexNumber]

    @property
    def rsidRoot_val(self) -> str | None:
        """Value of `w:rsidRoot/@w:val` or |None| if absent."""
        rsidRoot = self.rsidRoot
        if rsidRoot is None:
            return None
        return rsidRoot.val

    @property
    def rsid_vals(self) -> list[str]:
        """``@w:val`` of each `w:rsid` child in document order.

        Children whose ``@w:val`` attribute is missing are skipped.
        """
        return [rsid.val for rsid in self.rsid_lst if rsid.val is not None]


class CT_DecimalNumberWithVal(BaseOxmlElement):
    """Generic ``@w:val``-only decimal-number element used for several settings.

    Used for ``w:consecutiveHyphenLimit`` (integer cap on consecutive hyphens),
    and similar decimal-valued settings children.

    .. versionadded:: 2026.05.0
    """

    val: int = RequiredAttribute("w:val", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]


class CT_Language(BaseOxmlElement):
    """`w:themeFontLang` element, specifying the default language tags for the theme fonts.

    Carries optional ``@w:val`` (Latin), ``@w:eastAsia`` (East-Asian), and
    ``@w:bidi`` (bidirectional / complex-script) language-tag attributes.

    .. versionadded:: 2026.05.0
    """

    val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_String
    )
    eastAsia: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:eastAsia", ST_String
    )
    bidi: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:bidi", ST_String
    )


class CT_DocVar(BaseOxmlElement):
    """`w:docVar` element, a single document variable name/value pair.

    .. versionadded:: 2026.05.0
    """

    name: str = RequiredAttribute("w:name", ST_String)  # pyright: ignore[reportAssignmentType]
    val: str = RequiredAttribute("w:val", ST_String)  # pyright: ignore[reportAssignmentType]


class CT_DocVars(BaseOxmlElement):
    """`w:docVars` element, container for ``w:docVar`` document-variable entries.

    .. versionadded:: 2026.05.0
    """

    _add_docVar: Callable[..., CT_DocVar]
    docVar_lst: list[CT_DocVar]

    docVar = ZeroOrMore("w:docVar", successors=())

    def get_var(self, name: str) -> str | None:
        """Return the ``@w:val`` of the ``w:docVar`` whose ``@w:name == name``."""
        for dv in self.docVar_lst:
            if dv.name == name:
                return dv.val
        return None

    def set_var(self, name: str, value: str) -> None:
        """Create or update the ``w:docVar`` with ``@w:name == name``."""
        for dv in self.docVar_lst:
            if dv.name == name:
                dv.val = value
                return
        self._add_docVar(name=name, val=value)

    def remove_var(self, name: str) -> bool:
        """Remove the ``w:docVar`` matching ``name``; return True if found."""
        for dv in list(self.docVar_lst):
            if dv.name == name:
                self.remove(dv)
                return True
        return False


class CT_DocId(BaseOxmlElement):
    """`w15:docId` or `w14:docId` element, carrying a document-identifier GUID.

    Word 2013+ stamps ``<w15:docId w15:val="{GUID}"/>`` inside ``settings.xml``
    on every new document so its revision-tracking and "same document?"
    heuristics have a stable identifier. ``w14:docId`` is the legacy 2010
    sibling. Each namespace exposes its own ``@val`` attribute.

    .. versionadded:: 2026.05.3
    """

    w15_val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w15:val", ST_String
    )
    w14_val: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w14:val", ST_String
    )


class CT_Settings(BaseOxmlElement):
    """`w:settings` element, root element for the settings part."""

    get_or_add_writeProtection: Callable[[], CT_WriteProtection]
    _remove_writeProtection: Callable[[], None]
    get_or_add_view: Callable[[], CT_View]
    _remove_view: Callable[[], None]
    get_or_add_zoom: Callable[[], CT_Zoom]
    _remove_zoom: Callable[[], None]
    get_or_add_trackRevisions: Callable[[], CT_OnOff]
    _remove_trackRevisions: Callable[[], None]
    get_or_add_documentProtection: Callable[[], CT_DocProtect]
    _remove_documentProtection: Callable[[], None]
    get_or_add_defaultTabStop: Callable[[], CT_DefaultTabStop]
    _remove_defaultTabStop: Callable[[], None]
    get_or_add_evenAndOddHeaders: Callable[[], CT_OnOff]
    _remove_evenAndOddHeaders: Callable[[], None]
    get_or_add_updateFields: Callable[[], CT_OnOff]
    _remove_updateFields: Callable[[], None]
    get_or_add_footnotePr: Callable[[], "CT_FtnDocProps"]
    _remove_footnotePr: Callable[[], None]
    get_or_add_endnotePr: Callable[[], "CT_EdnDocProps"]
    _remove_endnotePr: Callable[[], None]
    get_or_add_compat: Callable[[], CT_Compat]
    _remove_compat: Callable[[], None]
    get_or_add_hideSpellingErrors: Callable[[], CT_OnOff]
    _remove_hideSpellingErrors: Callable[[], None]
    get_or_add_hideGrammaticalErrors: Callable[[], CT_OnOff]
    _remove_hideGrammaticalErrors: Callable[[], None]
    get_or_add_autoHyphenation: Callable[[], CT_OnOff]
    _remove_autoHyphenation: Callable[[], None]
    get_or_add_doNotHyphenateCaps: Callable[[], CT_OnOff]
    _remove_doNotHyphenateCaps: Callable[[], None]
    get_or_add_consecutiveHyphenLimit: Callable[[], "CT_DecimalNumberWithVal"]
    _remove_consecutiveHyphenLimit: Callable[[], None]
    get_or_add_hyphenationZone: Callable[[], CT_DefaultTabStop]
    _remove_hyphenationZone: Callable[[], None]
    get_or_add_themeFontLang: Callable[[], CT_Language]
    _remove_themeFontLang: Callable[[], None]
    get_or_add_docVars: Callable[[], CT_DocVars]
    _remove_docVars: Callable[[], None]

    _tag_seq = (
        "w:writeProtection",
        "w:view",
        "w:zoom",
        "w:removePersonalInformation",
        "w:removeDateAndTime",
        "w:doNotDisplayPageBoundaries",
        "w:displayBackgroundShape",
        "w:printPostScriptOverText",
        "w:printFractionalCharacterWidth",
        "w:printFormsData",
        "w:embedTrueTypeFonts",
        "w:embedSystemFonts",
        "w:saveSubsetFonts",
        "w:saveFormsData",
        "w:mirrorMargins",
        "w:alignBordersAndEdges",
        "w:bordersDoNotSurroundHeader",
        "w:bordersDoNotSurroundFooter",
        "w:gutterAtTop",
        "w:hideSpellingErrors",
        "w:hideGrammaticalErrors",
        "w:activeWritingStyle",
        "w:proofState",
        "w:formsDesign",
        "w:attachedTemplate",
        "w:linkStyles",
        "w:stylePaneFormatFilter",
        "w:stylePaneSortMethod",
        "w:documentType",
        "w:mailMerge",
        "w:revisionView",
        "w:trackRevisions",
        "w:doNotTrackMoves",
        "w:doNotTrackFormatting",
        "w:documentProtection",
        "w:autoFormatOverride",
        "w:styleLockTheme",
        "w:styleLockQFSet",
        "w:defaultTabStop",
        "w:autoHyphenation",
        "w:consecutiveHyphenLimit",
        "w:hyphenationZone",
        "w:doNotHyphenateCaps",
        "w:showEnvelope",
        "w:summaryLength",
        "w:clickAndTypeStyle",
        "w:defaultTableStyle",
        "w:evenAndOddHeaders",
        "w:bookFoldRevPrinting",
        "w:bookFoldPrinting",
        "w:bookFoldPrintingSheets",
        "w:drawingGridHorizontalSpacing",
        "w:drawingGridVerticalSpacing",
        "w:displayHorizontalDrawingGridEvery",
        "w:displayVerticalDrawingGridEvery",
        "w:doNotUseMarginsForDrawingGridOrigin",
        "w:drawingGridHorizontalOrigin",
        "w:drawingGridVerticalOrigin",
        "w:doNotShadeFormData",
        "w:noPunctuationKerning",
        "w:characterSpacingControl",
        "w:printTwoOnOne",
        "w:strictFirstAndLastChars",
        "w:noLineBreaksAfter",
        "w:noLineBreaksBefore",
        "w:savePreviewPicture",
        "w:doNotValidateAgainstSchema",
        "w:saveInvalidXml",
        "w:ignoreMixedContent",
        "w:alwaysShowPlaceholderText",
        "w:doNotDemarcateInvalidXml",
        "w:saveXmlDataOnly",
        "w:useXSLTWhenSaving",
        "w:saveThroughXslt",
        "w:showXMLTags",
        "w:alwaysMergeEmptyNamespace",
        "w:updateFields",
        "w:hdrShapeDefaults",
        "w:footnotePr",
        "w:endnotePr",
        "w:compat",
        "w:docVars",
        "w:rsids",
        "m:mathPr",
        "w:attachedSchema",
        "w:themeFontLang",
        "w:clrSchemeMapping",
        "w:doNotIncludeSubdocsInStats",
        "w:doNotAutoCompressPictures",
        "w:forceUpgrade",
        "w:captions",
        "w:readModeInkLockDown",
        "w:smartTagType",
        "sl:schemaLibrary",
        "w:shapeDefaults",
        "w:doNotEmbedSmartTags",
        "w:decimalSymbol",
        "w:listSeparator",
        # -- Microsoft extension children (appear at the tail in every
        # -- Office-authored settings.xml, gated via mc:Ignorable) --
        "w14:docId",
        "w15:chartTrackingRefBased",
        "w15:docId",
    )
    writeProtection: CT_WriteProtection | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:writeProtection", successors=_tag_seq[1:]
    )
    view: CT_View | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:view", successors=_tag_seq[2:]
    )
    zoom: CT_Zoom | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:zoom", successors=_tag_seq[3:]
    )
    hideSpellingErrors: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:hideSpellingErrors", successors=_tag_seq[20:]
    )
    hideGrammaticalErrors: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:hideGrammaticalErrors", successors=_tag_seq[21:]
    )
    mailMerge: CT_MailMerge | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:mailMerge", successors=_tag_seq[30:]
    )
    trackRevisions: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:trackRevisions", successors=_tag_seq[32:]
    )
    documentProtection: CT_DocProtect | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:documentProtection", successors=_tag_seq[35:]
    )
    defaultTabStop: CT_DefaultTabStop | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:defaultTabStop", successors=_tag_seq[39:]
    )
    autoHyphenation: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:autoHyphenation", successors=_tag_seq[40:]
    )
    consecutiveHyphenLimit: "CT_DecimalNumberWithVal | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:consecutiveHyphenLimit", successors=_tag_seq[41:]
    )
    hyphenationZone: CT_DefaultTabStop | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:hyphenationZone", successors=_tag_seq[42:]
    )
    doNotHyphenateCaps: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:doNotHyphenateCaps", successors=_tag_seq[43:]
    )
    evenAndOddHeaders: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:evenAndOddHeaders", successors=_tag_seq[48:]
    )
    updateFields: "CT_OnOff | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:updateFields", successors=_tag_seq[77:]
    )
    footnotePr: "CT_FtnDocProps | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:footnotePr", successors=_tag_seq[79:]
    )
    endnotePr: "CT_EdnDocProps | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:endnotePr", successors=_tag_seq[80:]
    )
    compat: CT_Compat | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:compat", successors=_tag_seq[81:]
    )
    docVars: "CT_DocVars | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docVars", successors=_tag_seq[82:]
    )
    rsids: "CT_Rsids | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rsids", successors=_tag_seq[83:]
    )
    themeFontLang: "CT_Language | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:themeFontLang", successors=_tag_seq[85:]
    )
    w14_docId: "CT_DocId | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w14:docId", successors=_tag_seq[-2:]
    )
    chartTrackingRefBased: "BaseOxmlElement | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w15:chartTrackingRefBased", successors=_tag_seq[-1:]
    )
    w15_docId: "CT_DocId | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w15:docId", successors=()
    )
    del _tag_seq

    @property
    def evenAndOddHeaders_val(self) -> bool:
        """Value of `w:evenAndOddHeaders/@w:val` or False if not present."""
        evenAndOddHeaders = self.evenAndOddHeaders
        if evenAndOddHeaders is None:
            return False
        return evenAndOddHeaders.val

    @evenAndOddHeaders_val.setter
    def evenAndOddHeaders_val(self, value: bool | None):
        if value is None or value is False:
            self._remove_evenAndOddHeaders()
            return
        self.get_or_add_evenAndOddHeaders().val = value

    @property
    def view_val(self) -> str | None:
        """Value of `w:view/@w:val` or None if not present."""
        view = self.view
        if view is None:
            return None
        return view.val

    @view_val.setter
    def view_val(self, value: str | None):
        if value is None:
            self._remove_view()
            return
        self.get_or_add_view().val = value

    @property
    def zoom_percent(self) -> int | None:
        """Value of `w:zoom/@w:percent` or None if not present."""
        zoom = self.zoom
        if zoom is None:
            return None
        return zoom.percent

    @zoom_percent.setter
    def zoom_percent(self, value: int | None):
        if value is None:
            self._remove_zoom()
            return
        self.get_or_add_zoom().percent = value

    @property
    def trackRevisions_val(self) -> bool:
        """True if `w:trackRevisions` is present and enabled."""
        trackRevisions = self.trackRevisions
        if trackRevisions is None:
            return False
        return trackRevisions.val

    @trackRevisions_val.setter
    def trackRevisions_val(self, value: bool | None):
        if value is None or value is False:
            self._remove_trackRevisions()
            return
        self.get_or_add_trackRevisions().val = value

    @property
    def updateFields_val(self) -> bool:
        """True if `w:updateFields` is present and enabled."""
        updateFields = self.updateFields
        if updateFields is None:
            return False
        return updateFields.val

    @updateFields_val.setter
    def updateFields_val(self, value: bool | None):
        if value is None or value is False:
            self._remove_updateFields()
            return
        self.get_or_add_updateFields().val = value

    @property
    def defaultTabStop_val(self) -> Length | None:
        """Value of `w:defaultTabStop/@w:val` as a Length or None if not present."""
        defaultTabStop = self.defaultTabStop
        if defaultTabStop is None:
            return None
        return defaultTabStop.val

    @defaultTabStop_val.setter
    def defaultTabStop_val(self, value: int | Length | None):
        if value is None:
            self._remove_defaultTabStop()
            return
        self.get_or_add_defaultTabStop().val = value

    @property
    def documentProtection_edit(self) -> str | None:
        """Value of `w:documentProtection/@w:edit` or None if not present."""
        dp = self.documentProtection
        if dp is None:
            return None
        return dp.edit

    @documentProtection_edit.setter
    def documentProtection_edit(self, value: str | None):
        if value is None:
            dp = self.documentProtection
            if dp is not None:
                dp.edit = None
            return
        self.get_or_add_documentProtection().edit = value

    @property
    def documentProtection_enforcement(self) -> bool:
        """True if `w:documentProtection/@w:enforcement` is enabled."""
        dp = self.documentProtection
        if dp is None:
            return False
        return dp.enforcement

    @documentProtection_enforcement.setter
    def documentProtection_enforcement(self, value: bool):
        self.get_or_add_documentProtection().enforcement = bool(value)

    @property
    def compatibilityMode(self) -> int | None:
        """The compatibility-mode value from `w:compat/w:compatSetting` or None."""
        compat = self.compat
        if compat is None:
            return None
        for setting in compat.compatSetting_lst:
            if setting.name == "compatibilityMode":
                return int(setting.val)
        return None

    @compatibilityMode.setter
    def compatibilityMode(self, value: int | None):
        if value is None:
            compat = self.compat
            if compat is not None:
                for setting in list(compat.compatSetting_lst):
                    if setting.name == "compatibilityMode":
                        compat.remove(setting)
                if len(compat.compatSetting_lst) == 0:
                    self._remove_compat()
            return
        compat = self.get_or_add_compat()
        for setting in compat.compatSetting_lst:
            if setting.name == "compatibilityMode":
                setting.val = str(value)
                return
        compat._add_compatSetting(
            name="compatibilityMode",
            uri="http://schemas.microsoft.com/office/word",
            val=str(value),
        )
