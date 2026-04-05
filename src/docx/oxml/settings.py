"""Custom element classes related to document settings."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List

from docx.oxml.simpletypes import ST_DecimalNumber, ST_OnOff, ST_String, ST_TwipsMeasure
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.shared import CT_OnOff
    from docx.shared import Length


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


class CT_Compat(BaseOxmlElement):
    """`w:compat` element, containing document compatibility settings."""

    _add_compatSetting: Callable[..., CT_CompatSetting]
    compatSetting_lst: List[CT_CompatSetting]

    compatSetting = ZeroOrMore("w:compatSetting", successors=())


class CT_DefaultTabStop(BaseOxmlElement):
    """`w:defaultTabStop` element, specifying default tab-stop interval."""

    val: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", ST_TwipsMeasure
    )


class CT_Settings(BaseOxmlElement):
    """`w:settings` element, root element for the settings part."""

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
    compat: CT_Compat | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:compat", successors=_tag_seq[81:]
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
