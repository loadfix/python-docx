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


class CT_DocProtect(BaseOxmlElement):
    """`w:documentProtection` element, specifying document editing restrictions."""

    edit: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:edit", ST_String
    )
    enforcement: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:enforcement", ST_OnOff, default=False
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


class CT_Settings(BaseOxmlElement):
    """`w:settings` element, root element for the settings part."""

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
    get_or_add_footnotePr: Callable[[], "CT_FtnDocProps"]
    _remove_footnotePr: Callable[[], None]
    get_or_add_endnotePr: Callable[[], "CT_EdnDocProps"]
    _remove_endnotePr: Callable[[], None]
    get_or_add_compat: Callable[[], CT_Compat]
    _remove_compat: Callable[[], None]

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
    )
    view: CT_View | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:view", successors=_tag_seq[2:]
    )
    zoom: CT_Zoom | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:zoom", successors=_tag_seq[3:]
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
    evenAndOddHeaders: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:evenAndOddHeaders", successors=_tag_seq[48:]
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
    rsids: "CT_Rsids | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rsids", successors=_tag_seq[83:]
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

    @property
    def documentProtection_enforcement(self) -> bool:
        """True if `w:documentProtection/@w:enforcement` is enabled."""
        dp = self.documentProtection
        if dp is None:
            return False
        return dp.enforcement

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
