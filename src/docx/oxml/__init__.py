# ruff: noqa: E402, I001

"""Initializes oxml sub-package.

This including registering custom element classes corresponding to Open XML elements.
"""

from __future__ import annotations

from docx.oxml.drawing import (
    CT_Drawing,
    CT_GroupShape,
    CT_NonVisualGroupShapeProperties,
    CT_TextBox,
    CT_TxbxContent,
    CT_WordprocessingCanvas,
    CT_WordprocessingShape,
)
from docx.oxml.parser import OxmlElement, parse_xml, register_element_cls
from docx.oxml.shape import (
    CT_Alpha,
    CT_AlphaModulateFixedEffect,
    CT_Anchor,
    CT_Blip,
    CT_BlipFillProperties,
    CT_EffectList,
    CT_GraphicalObject,
    CT_GraphicalObjectData,
    CT_Inline,
    CT_LineProperties,
    CT_NonVisualDrawingProps,
    CT_NonVisualPictureProperties,
    CT_OuterShadow,
    CT_Picture,
    CT_PictureLocking,
    CT_PictureNonVisual,
    CT_Point2D,
    CT_PosOffset,
    CT_PositionH,
    CT_PositionV,
    CT_PositiveSize2D,
    CT_RelativeRect,
    CT_ShapeProperties,
    CT_SolidColorFill,
    CT_Transform2D,
    CT_WrapNone,
    CT_WrapPolygon,
    CT_WrapPolygonPoint,
    CT_WrapSquare,
    CT_WrapThrough,
    CT_WrapTight,
    CT_WrapTopBottom,
)
from docx.oxml.shared import CT_DecimalNumber, CT_OnOff, CT_ProofErr, CT_String
from docx.oxml.text.hyperlink import CT_Hyperlink
from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
from docx.oxml.text.run import (
    CT_R,
    CT_Br,
    CT_Cr,
    CT_NoBreakHyphen,
    CT_PTab,
    CT_Sym,
    CT_Text,
)

# -- `OxmlElement` and `parse_xml()` are not used in this module but several downstream
# -- "extension" packages expect to find them here and there's no compelling reason
# -- not to republish them here so those keep working.
__all__ = ["OxmlElement", "parse_xml"]

# ---------------------------------------------------------------------------
# DrawingML-related elements

register_element_cls("a:alpha", CT_Alpha)
register_element_cls("a:alphaModFix", CT_AlphaModulateFixedEffect)
register_element_cls("a:blip", CT_Blip)
register_element_cls("a:effectLst", CT_EffectList)
register_element_cls("a:ext", CT_PositiveSize2D)
register_element_cls("a:graphic", CT_GraphicalObject)
register_element_cls("a:graphicData", CT_GraphicalObjectData)
register_element_cls("a:ln", CT_LineProperties)
register_element_cls("a:off", CT_Point2D)
register_element_cls("a:outerShdw", CT_OuterShadow)
register_element_cls("a:picLocks", CT_PictureLocking)
register_element_cls("a:solidFill", CT_SolidColorFill)
register_element_cls("a:srcRect", CT_RelativeRect)
register_element_cls("a:xfrm", CT_Transform2D)

# --- DrawingML gradient-fill CT_* classes from python-ooxml-shared-drawingml ---
# Registering these in docx's element_class_lookup lets the shared
# ``ooxml_chart.FormatFill`` proxy author / read gradient fills on
# chart ``c:spPr`` children through docx's parser — R17-3 adoption of
# ooxml-chart 0.5 gradient-fill accessors.
from ooxml_shared_drawingml.fill import (
    CT_GradientFillProperties as _ShrdCT_GradientFillProperties,
    CT_GradientStop as _ShrdCT_GradientStop,
    CT_GradientStopList as _ShrdCT_GradientStopList,
    CT_LinearShadeProperties as _ShrdCT_LinearShadeProperties,
)

register_element_cls("a:gradFill", _ShrdCT_GradientFillProperties)
register_element_cls("a:gs", _ShrdCT_GradientStop)
register_element_cls("a:gsLst", _ShrdCT_GradientStopList)
register_element_cls("a:lin", _ShrdCT_LinearShadeProperties)
register_element_cls("pic:blipFill", CT_BlipFillProperties)
register_element_cls("pic:cNvPicPr", CT_NonVisualPictureProperties)
register_element_cls("pic:cNvPr", CT_NonVisualDrawingProps)
register_element_cls("pic:nvPicPr", CT_PictureNonVisual)
register_element_cls("pic:pic", CT_Picture)
register_element_cls("pic:spPr", CT_ShapeProperties)
register_element_cls("w:drawing", CT_Drawing)
register_element_cls("w:txbxContent", CT_TxbxContent)
register_element_cls("wp:anchor", CT_Anchor)
register_element_cls("wp:docPr", CT_NonVisualDrawingProps)
register_element_cls("wp:extent", CT_PositiveSize2D)
register_element_cls("wp:inline", CT_Inline)
register_element_cls("wp:posOffset", CT_PosOffset)
register_element_cls("wp:positionH", CT_PositionH)
register_element_cls("wp:positionV", CT_PositionV)
register_element_cls("wp:lineTo", CT_WrapPolygonPoint)
register_element_cls("wp:start", CT_WrapPolygonPoint)
register_element_cls("wp:wrapNone", CT_WrapNone)
register_element_cls("wp:wrapPolygon", CT_WrapPolygon)
register_element_cls("wp:wrapSquare", CT_WrapSquare)
register_element_cls("wp:wrapThrough", CT_WrapThrough)
register_element_cls("wp:wrapTight", CT_WrapTight)
register_element_cls("wp:wrapTopAndBottom", CT_WrapTopBottom)
register_element_cls("wpc:wpc", CT_WordprocessingCanvas)
register_element_cls("wpg:grpSp", CT_GroupShape)
register_element_cls("wpg:wgp", CT_GroupShape)
register_element_cls("wpg:nvGrpSpPr", CT_NonVisualGroupShapeProperties)
register_element_cls("wpg:cNvPr", CT_NonVisualDrawingProps)
register_element_cls("wps:cNvPr", CT_NonVisualDrawingProps)
register_element_cls("wps:txbx", CT_TextBox)
register_element_cls("wps:wsp", CT_WordprocessingShape)

# ---------------------------------------------------------------------------
# hyperlink-related elements

register_element_cls("w:hyperlink", CT_Hyperlink)

# ---------------------------------------------------------------------------
# text-related elements

register_element_cls("w:br", CT_Br)
register_element_cls("w:cr", CT_Cr)
register_element_cls("w:lastRenderedPageBreak", CT_LastRenderedPageBreak)
register_element_cls("w:noBreakHyphen", CT_NoBreakHyphen)
register_element_cls("w:ptab", CT_PTab)
register_element_cls("w:proofErr", CT_ProofErr)
register_element_cls("w:r", CT_R)
register_element_cls("w:sym", CT_Sym)
register_element_cls("w:t", CT_Text)

# ---------------------------------------------------------------------------
# header/footer-related mappings

register_element_cls("w:bidi", CT_OnOff)
register_element_cls("w:evenAndOddHeaders", CT_OnOff)
register_element_cls("w:rtlGutter", CT_OnOff)
register_element_cls("w:titlePg", CT_OnOff)

# ---------------------------------------------------------------------------
# other custom element class mappings

from .bookmarks import CT_BookmarkEnd, CT_BookmarkStart, CT_MarkupRange, CT_MoveBookmark

register_element_cls("w:bookmarkEnd", CT_BookmarkEnd)
register_element_cls("w:bookmarkStart", CT_BookmarkStart)
# -- ECMA-376 move-range and comment-range markers share the CT_MarkupRange /
#    CT_MoveBookmark shape; register them with typed element classes so
#    @w:id / @w:name / @w:author / @w:date are accessible as attributes. --
register_element_cls("w:moveFromRangeStart", CT_MoveBookmark)
register_element_cls("w:moveFromRangeEnd", CT_MarkupRange)
register_element_cls("w:moveToRangeStart", CT_MoveBookmark)
register_element_cls("w:moveToRangeEnd", CT_MarkupRange)

from .comments import CT_Comments, CT_Comment

register_element_cls("w:comments", CT_Comments)
register_element_cls("w:comment", CT_Comment)

# -- Word 2013+ ``w15:commentsEx`` extended-comments part (resolve/reopen
# -- state + threaded reply parent linkage + co-author presence info). The
# -- element classes carry the ``@w15:paraId`` / ``@w15:done`` /
# -- ``@w15:paraIdParent`` descriptors and the ``<w15:presenceInfo>``
# -- provider/user pair. Registering here so ``parse_xml()`` resolves
# -- these tags to the typed classes rather than plain lxml elements. --
from .comments_extended import (
    CT_CommentExtended,
    CT_CommentExtendedList,
    CT_PresenceInfo,
)

register_element_cls("w15:commentsEx", CT_CommentExtendedList)
register_element_cls("w15:commentEx", CT_CommentExtended)
register_element_cls("w15:presenceInfo", CT_PresenceInfo)

# -- Word 2016+ ``w16cid:commentsIds`` and Word 2018+ ``w16cex:commentsExtensible``
# -- auxiliary parts. The commentsIds part maps legacy ``w:comment/@w:id``
# -- integers to stable paragraph ids (``w16cid:paraId``) used by Office's
# -- threaded-reply feature. The commentsExtensible part attaches durable
# -- GUID identifiers to each legacy comment so Office 365 clients don't
# -- renumber them across edit sessions. Element classes live in the shared
# -- ``ooxml_comments`` package; we re-export and register them here.
from .comments_ids import (
    CT_CommentExtensible,
    CT_CommentExtensibleList,
    CT_CommentId,
    CT_CommentIdList,
)

register_element_cls("w16cid:commentsIds", CT_CommentIdList)
register_element_cls("w16cid:commentId", CT_CommentId)
register_element_cls("w16cex:commentsExtensible", CT_CommentExtensibleList)
register_element_cls("w16cex:commentExtensible", CT_CommentExtensible)

from .content_controls import (
    CT_DataBinding,
    CT_Lock,
    CT_Sdt,
    CT_SdtComboBox,
    CT_SdtContent,
    CT_SdtContentBlock,
    CT_SdtContentCell,
    CT_SdtContentRow,
    CT_SdtContentRun,
    CT_SdtContentRunRuby,
    CT_SdtDate,
    CT_SdtDateMappingType,
    CT_SdtDocPart,
    CT_SdtDropDownList,
    CT_SdtEndPr,
    CT_SdtListItem,
    CT_SdtPr,
    CT_SdtRepeatedSection,
    CT_SdtRepeatedSectionItem,
    CT_SdtText,
)

register_element_cls("w:dataBinding", CT_DataBinding)
register_element_cls("w:lock", CT_Lock)
register_element_cls("w:sdt", CT_Sdt)
register_element_cls("w:sdtContent", CT_SdtContent)
register_element_cls("w:sdtPr", CT_SdtPr)
# -- SDT property-value types (ECMA-376 wml.xsd) --
register_element_cls("w:comboBox", CT_SdtComboBox)
register_element_cls("w:date", CT_SdtDate)
register_element_cls("w:docPartList", CT_SdtDocPart)
register_element_cls("w:docPartObj", CT_SdtDocPart)
register_element_cls("w:dropDownList", CT_SdtDropDownList)
register_element_cls("w:listItem", CT_SdtListItem)
register_element_cls("w:sdtEndPr", CT_SdtEndPr)
register_element_cls("w:storeMappedDataAs", CT_SdtDateMappingType)
register_element_cls("w:text", CT_SdtText)
# -- SDT content-container types (ECMA-376 wml.xsd) --
#    These share the ``w:sdtContent`` tag with the default ``CT_SdtContent``
#    registration above; callers opt into a typed container (block / cell /
#    row / run / ruby) by constructing it directly when they need the
#    container-specific accessors.  The tuple keeps the classes reachable
#    as module-level exports so downstream code and typing tools can
#    reference them.
_SDT_CONTENT_CONTAINER_CLASSES = (
    CT_SdtContentBlock,
    CT_SdtContentCell,
    CT_SdtContentRow,
    CT_SdtContentRun,
    CT_SdtContentRunRuby,
)
# -- w15 repeated-section SDT markers (MS Word 2013+ extension) --
register_element_cls("w15:repeatingSection", CT_SdtRepeatedSection)
register_element_cls("w15:repeatingSectionItem", CT_SdtRepeatedSectionItem)

from .coreprops import CT_CoreProperties

register_element_cls("cp:coreProperties", CT_CoreProperties)

from .custom_properties import CT_CustomProperties, CT_CustomProperty

register_element_cls("custprops:Properties", CT_CustomProperties)
register_element_cls("custprops:property", CT_CustomProperty)

from .bibliography import CT_Source, CT_Sources

register_element_cls("b:Sources", CT_Sources)
register_element_cls("b:Source", CT_Source)

from .extended_properties import CT_ExtendedProperties

register_element_cls("extprops:Properties", CT_ExtendedProperties)

from .endnotes import CT_EdnDocProps, CT_Endnote, CT_Endnotes

register_element_cls("w:endnote", CT_Endnote)
register_element_cls("w:endnotes", CT_Endnotes)
register_element_cls("w:endnotePr", CT_EdnDocProps)

from .glossary import (
    CT_DocPart,
    CT_DocPartBehavior,
    CT_DocPartBehaviors,
    CT_DocPartBody,
    CT_DocPartCategory,
    CT_DocPartGallery,
    CT_DocPartPr,
    CT_DocPartType,
    CT_DocPartTypes,
    CT_DocParts,
    CT_GlossaryDocument,
)

register_element_cls("w:docPart", CT_DocPart)
register_element_cls("w:docPartBody", CT_DocPartBody)
register_element_cls("w:docParts", CT_DocParts)
register_element_cls("w:docPartPr", CT_DocPartPr)
register_element_cls("w:category", CT_DocPartCategory)
register_element_cls("w:gallery", CT_DocPartGallery)
register_element_cls("w:glossaryDocument", CT_GlossaryDocument)
# -- ``w:type`` and ``w:behavior`` (the per-entry children) are not --
# -- registered globally: ``w:type`` is already bound to ``CT_SectType`` --
# -- via :mod:`docx.oxml.section`. Access the inner ``w:val`` attribute --
# -- via the ``values`` property on CT_DocPartTypes / CT_DocPartBehaviors --
# -- which reads the attribute directly rather than relying on a typed --
# -- child class. The wrapper elements ``w:types`` and ``w:behaviors`` --
# -- are docPart-specific and safe to bind. --
register_element_cls("w:behaviors", CT_DocPartBehaviors)
register_element_cls("w:types", CT_DocPartTypes)

from .fields import CT_FldChar, CT_FldSimple, CT_InstrText

register_element_cls("w:fldChar", CT_FldChar)
register_element_cls("w:fldSimple", CT_FldSimple)
register_element_cls("w:instrText", CT_InstrText)

from .form_fields import CT_FFCheckBox, CT_FFData, CT_FFDDList, CT_FFTextInput

register_element_cls("w:ffData", CT_FFData)
register_element_cls("w:textInput", CT_FFTextInput)
register_element_cls("w:checkBox", CT_FFCheckBox)
register_element_cls("w:ddList", CT_FFDDList)

from .document import CT_AltChunk, CT_AltChunkPr, CT_Background, CT_Body, CT_Document

register_element_cls("w:altChunk", CT_AltChunk)
register_element_cls("w:altChunkPr", CT_AltChunkPr)
register_element_cls("w:matchSrc", CT_OnOff)
register_element_cls("w:background", CT_Background)
register_element_cls("w:body", CT_Body)
register_element_cls("w:document", CT_Document)

from .font_table import (
    CT_Charset,
    CT_Font,
    CT_FontFamily,
    CT_FontName,
    CT_FontRel,
    CT_Fonts as CT_FontTable,
    CT_Panose,
    CT_Pitch,
)

register_element_cls("w:fonts", CT_FontTable)
register_element_cls("w:font", CT_Font)
register_element_cls("w:altName", CT_FontName)
register_element_cls("w:charset", CT_Charset)
register_element_cls("w:family", CT_FontFamily)
register_element_cls("w:panose1", CT_Panose)
register_element_cls("w:pitch", CT_Pitch)
register_element_cls("w:embedRegular", CT_FontRel)
register_element_cls("w:embedBold", CT_FontRel)
register_element_cls("w:embedItalic", CT_FontRel)
register_element_cls("w:embedBoldItalic", CT_FontRel)

from .footnotes import (
    CT_FtnDocProps,
    CT_FtnEdnPos,
    CT_Footnote,
    CT_Footnotes,
    CT_NumFmt,
    CT_NumRestart,
    CT_NumStart,
)

register_element_cls("w:footnote", CT_Footnote)
register_element_cls("w:footnotes", CT_Footnotes)
register_element_cls("w:footnotePr", CT_FtnDocProps)
register_element_cls("w:pos", CT_FtnEdnPos)
register_element_cls("w:numFmt", CT_NumFmt)
register_element_cls("w:numStart", CT_NumStart)
register_element_cls("w:numRestart", CT_NumRestart)

from .math import CT_MathR, CT_MathT, CT_OMath, CT_OMathPara

register_element_cls("m:oMath", CT_OMath)
register_element_cls("m:oMathPara", CT_OMathPara)
register_element_cls("m:r", CT_MathR)
register_element_cls("m:t", CT_MathT)

from .permissions import CT_PermEnd, CT_PermStart

register_element_cls("w:permEnd", CT_PermEnd)
register_element_cls("w:permStart", CT_PermStart)

from .numbering import (
    CT_AbstractNum,
    CT_Lvl,
    CT_LvlText,
    CT_Num,
    CT_Numbering,
    CT_NumLvl,
    CT_NumPicBullet,
    CT_NumPr,
)

register_element_cls("w:abstractNum", CT_AbstractNum)
register_element_cls("w:abstractNumId", CT_DecimalNumber)
register_element_cls("w:ilvl", CT_DecimalNumber)
register_element_cls("w:lvl", CT_Lvl)
register_element_cls("w:lvlOverride", CT_NumLvl)
register_element_cls("w:lvlText", CT_LvlText)
register_element_cls("w:num", CT_Num)
register_element_cls("w:numId", CT_DecimalNumber)
register_element_cls("w:numPicBullet", CT_NumPicBullet)
register_element_cls("w:numPicBulletId", CT_DecimalNumber)
register_element_cls("w:numPr", CT_NumPr)
register_element_cls("w:numbering", CT_Numbering)
register_element_cls("w:start", CT_DecimalNumber)
register_element_cls("w:startOverride", CT_DecimalNumber)

from .ruby import (
    CT_Ruby,
    CT_RubyAlign,
    CT_RubyContent,
    CT_RubyHps,
    CT_RubyLang,
    CT_RubyPr,
)

register_element_cls("w:ruby", CT_Ruby)
register_element_cls("w:rubyPr", CT_RubyPr)
register_element_cls("w:rt", CT_RubyContent)
register_element_cls("w:rubyBase", CT_RubyContent)
register_element_cls("w:rubyAlign", CT_RubyAlign)
register_element_cls("w:hps", CT_RubyHps)
register_element_cls("w:hpsRaise", CT_RubyHps)
register_element_cls("w:hpsBaseText", CT_RubyHps)
register_element_cls("w:lid", CT_RubyLang)

from .section import (
    CT_Col,
    CT_Cols,
    CT_DocGrid,
    CT_HdrFtr,
    CT_HdrFtrRef,
    CT_LineNumber,
    CT_PageMar,
    CT_PageNumber,
    CT_PageSz,
    CT_PaperSource,
    CT_PgBorders,
    CT_SectPr,
    CT_SectType,
)

register_element_cls("w:col", CT_Col)
register_element_cls("w:cols", CT_Cols)
register_element_cls("w:docGrid", CT_DocGrid)
register_element_cls("w:footerReference", CT_HdrFtrRef)
register_element_cls("w:ftr", CT_HdrFtr)
register_element_cls("w:hdr", CT_HdrFtr)
register_element_cls("w:headerReference", CT_HdrFtrRef)
register_element_cls("w:lnNumType", CT_LineNumber)
register_element_cls("w:paperSrc", CT_PaperSource)
register_element_cls("w:pgBorders", CT_PgBorders)
register_element_cls("w:pgMar", CT_PageMar)
register_element_cls("w:pgNumType", CT_PageNumber)
register_element_cls("w:pgSz", CT_PageSz)
register_element_cls("w:sectPr", CT_SectPr)
register_element_cls("w:type", CT_SectType)

from .settings import (
    CT_Compat,
    CT_CompatSetting,
    CT_DecimalNumberWithVal,
    CT_DefaultTabStop,
    CT_DocId,
    CT_DocProtect,
    CT_DocVar,
    CT_DocVars,
    CT_Language,
    CT_LongHexNumber,
    CT_MailMerge,
    CT_Rsids,
    CT_Settings,
    CT_View,
    CT_WriteProtection,
    CT_Zoom,
    _CT_MMVal,
)

register_element_cls("w:compat", CT_Compat)
register_element_cls("w:compatSetting", CT_CompatSetting)
register_element_cls("w:defaultTabStop", CT_DefaultTabStop)
register_element_cls("w:documentProtection", CT_DocProtect)
register_element_cls("w:writeProtection", CT_WriteProtection)
register_element_cls("w:mailMerge", CT_MailMerge)
register_element_cls("w:mainDocumentType", _CT_MMVal)
register_element_cls("w:dataType", _CT_MMVal)
register_element_cls("w:connectString", _CT_MMVal)
register_element_cls("w:query", _CT_MMVal)
register_element_cls("w:destination", _CT_MMVal)
register_element_cls("w:addressFieldName", _CT_MMVal)
register_element_cls("w:mailSubject", _CT_MMVal)
register_element_cls("w:activeRecord", _CT_MMVal)
register_element_cls("w:checkErrors", _CT_MMVal)
register_element_cls("w:linkToQuery", CT_OnOff)
register_element_cls("w:doNotSuppressBlankLines", CT_OnOff)
register_element_cls("w:mailAsAttachment", CT_OnOff)
register_element_cls("w:viewMergedData", CT_OnOff)
register_element_cls("w:rsid", CT_LongHexNumber)
register_element_cls("w:rsidRoot", CT_LongHexNumber)
register_element_cls("w:rsids", CT_Rsids)
register_element_cls("w:settings", CT_Settings)
register_element_cls("w:hideSpellingErrors", CT_OnOff)
register_element_cls("w:hideGrammaticalErrors", CT_OnOff)
register_element_cls("w:autoHyphenation", CT_OnOff)
register_element_cls("w:doNotHyphenateCaps", CT_OnOff)
register_element_cls("w:consecutiveHyphenLimit", CT_DecimalNumberWithVal)
register_element_cls("w:hyphenationZone", CT_DefaultTabStop)
register_element_cls("w:themeFontLang", CT_Language)
register_element_cls("w:docVars", CT_DocVars)
register_element_cls("w:docVar", CT_DocVar)
register_element_cls("w:trackRevisions", CT_OnOff)
register_element_cls("w:updateFields", CT_OnOff)
register_element_cls("w:removePersonalInformation", CT_OnOff)
register_element_cls("w:removeDateAndTime", CT_OnOff)
register_element_cls("w:charactersWithNumbersWidth", CT_OnOff)
register_element_cls("w:view", CT_View)
register_element_cls("w:zoom", CT_Zoom)
register_element_cls("w14:docId", CT_DocId)
register_element_cls("w15:docId", CT_DocId)

from .styles import (
    CT_DocDefaults,
    CT_LatentStyles,
    CT_LsdException,
    CT_RPrDefault,
    CT_Style,
    CT_Styles,
)

register_element_cls("w:autoRedefine", CT_OnOff)
register_element_cls("w:basedOn", CT_String)
register_element_cls("w:docDefaults", CT_DocDefaults)
register_element_cls("w:latentStyles", CT_LatentStyles)
register_element_cls("w:link", CT_String)
register_element_cls("w:locked", CT_OnOff)
register_element_cls("w:lsdException", CT_LsdException)
register_element_cls("w:name", CT_String)
register_element_cls("w:next", CT_String)
register_element_cls("w:qFormat", CT_OnOff)
register_element_cls("w:rPrDefault", CT_RPrDefault)
register_element_cls("w:semiHidden", CT_OnOff)
register_element_cls("w:style", CT_Style)
register_element_cls("w:styles", CT_Styles)
register_element_cls("w:uiPriority", CT_DecimalNumber)
register_element_cls("w:unhideWhenUsed", CT_OnOff)

from .table import (
    CT_Border,
    CT_Height,
    CT_Row,
    CT_Shd,
    CT_Tbl,
    CT_TblBorders,
    CT_TblCellMar,
    CT_TblGrid,
    CT_TblGridCol,
    CT_TblLayoutType,
    CT_TblLook,
    CT_TblPr,
    CT_TblPrEx,
    CT_TblWidth,
    CT_Tc,
    CT_TcBorders,
    CT_TcMar,
    CT_TcPr,
    CT_TextDirection,
    CT_TrPr,
    CT_VMerge,
    CT_VerticalJc,
)

register_element_cls("w:bidiVisual", CT_OnOff)
register_element_cls("w:tblBorders", CT_TblBorders)
register_element_cls("w:tcBorders", CT_TcBorders)
# -- `w:top`/`w:bottom`/`w:left`/`w:right`/`w:insideH`/`w:insideV` register to --
# -- CT_Border in the parfmt block below; the shared class handles table, cell, --
# -- paragraph and page border usage (#165). --
register_element_cls("w:cantSplit", CT_OnOff)
register_element_cls("w:gridAfter", CT_DecimalNumber)
register_element_cls("w:gridBefore", CT_DecimalNumber)
register_element_cls("w:gridCol", CT_TblGridCol)
register_element_cls("w:gridSpan", CT_DecimalNumber)
register_element_cls("w:shd", CT_Shd)
register_element_cls("w:tbl", CT_Tbl)
register_element_cls("w:tblCaption", CT_String)
register_element_cls("w:tblCellMar", CT_TblCellMar)
register_element_cls("w:tblDescription", CT_String)
register_element_cls("w:tblGrid", CT_TblGrid)
register_element_cls("w:tblHeader", CT_OnOff)
register_element_cls("w:tblInd", CT_TblWidth)
register_element_cls("w:tblLayout", CT_TblLayoutType)
register_element_cls("w:tblLook", CT_TblLook)
register_element_cls("w:tblPr", CT_TblPr)
register_element_cls("w:tblPrEx", CT_TblPrEx)
register_element_cls("w:tblStyle", CT_String)
register_element_cls("w:tblW", CT_TblWidth)
register_element_cls("w:tc", CT_Tc)
register_element_cls("w:tcMar", CT_TcMar)
register_element_cls("w:tcPr", CT_TcPr)
register_element_cls("w:tcW", CT_TblWidth)
register_element_cls("w:textDirection", CT_TextDirection)
register_element_cls("w:tr", CT_Row)
register_element_cls("w:trHeight", CT_Height)
register_element_cls("w:trPr", CT_TrPr)
register_element_cls("w:vAlign", CT_VerticalJc)
register_element_cls("w:vMerge", CT_VMerge)

from .text.font import (
    CT_Color,
    CT_EastAsianLayout,
    CT_Fonts,
    CT_Highlight,
    CT_HpsMeasure,
    CT_Language,
    CT_Ligatures,
    CT_RPr,
    CT_TextScale,
    CT_Underline,
    CT_VerticalAlignRun,
)

register_element_cls("w:b", CT_OnOff)
register_element_cls("w:bCs", CT_OnOff)
register_element_cls("w:caps", CT_OnOff)
register_element_cls("w:color", CT_Color)
register_element_cls("w:cs", CT_OnOff)
register_element_cls("w:dstrike", CT_OnOff)
register_element_cls("w:eastAsianLayout", CT_EastAsianLayout)
register_element_cls("w:emboss", CT_OnOff)
register_element_cls("w:highlight", CT_Highlight)
register_element_cls("w:i", CT_OnOff)
register_element_cls("w:iCs", CT_OnOff)
register_element_cls("w:imprint", CT_OnOff)
register_element_cls("w:kern", CT_HpsMeasure)
register_element_cls("w:lang", CT_Language)
register_element_cls("w:noProof", CT_OnOff)
register_element_cls("w:oMath", CT_OnOff)
register_element_cls("w:outline", CT_OnOff)
register_element_cls("w:rFonts", CT_Fonts)
register_element_cls("w:rPr", CT_RPr)
register_element_cls("w:rStyle", CT_String)
register_element_cls("w:rtl", CT_OnOff)
register_element_cls("w:shadow", CT_OnOff)
register_element_cls("w:smallCaps", CT_OnOff)
register_element_cls("w:snapToGrid", CT_OnOff)
register_element_cls("w:specVanish", CT_OnOff)
register_element_cls("w:strike", CT_OnOff)
register_element_cls("w:sz", CT_HpsMeasure)
register_element_cls("w:szCs", CT_HpsMeasure)
register_element_cls("w:u", CT_Underline)
register_element_cls("w:vanish", CT_OnOff)
register_element_cls("w:vertAlign", CT_VerticalAlignRun)
register_element_cls("w:w", CT_TextScale)
register_element_cls("w:webHidden", CT_OnOff)
register_element_cls("w14:ligatures", CT_Ligatures)

from .text.paragraph import CT_P

register_element_cls("w:p", CT_P)

from .web_settings import (
    CT_Encoding,
    CT_Frame,
    CT_Frameset,
    CT_OptimizeForBrowser,
    CT_WebSettings,
)

register_element_cls("w:webSettings", CT_WebSettings)
register_element_cls("w:encoding", CT_Encoding)
register_element_cls("w:optimizeForBrowser", CT_OptimizeForBrowser)
register_element_cls("w:frameset", CT_Frameset)
register_element_cls("w:frame", CT_Frame)
register_element_cls("w:relyOnVML", CT_OnOff)
register_element_cls("w:allowPNG", CT_OnOff)
register_element_cls("w:doNotRelyOnCSS", CT_OnOff)
register_element_cls("w:doNotSaveAsSingleFile", CT_OnOff)

from .theme import (
    CT_ClrScheme,
    CT_ColorChoice,
    CT_FontCollection,
    CT_FontScheme,
    CT_SRgbColor,
    CT_SysColor,
    CT_TextFont,
    CT_Theme,
    CT_ThemeElements,
)

register_element_cls("a:theme", CT_Theme)
register_element_cls("a:themeElements", CT_ThemeElements)
register_element_cls("a:clrScheme", CT_ClrScheme)
register_element_cls("a:fontScheme", CT_FontScheme)
register_element_cls("a:majorFont", CT_FontCollection)
register_element_cls("a:minorFont", CT_FontCollection)
register_element_cls("a:latin", CT_TextFont)
register_element_cls("a:ea", CT_TextFont)
register_element_cls("a:cs", CT_TextFont)
register_element_cls("a:srgbClr", CT_SRgbColor)
register_element_cls("a:sysClr", CT_SysColor)
register_element_cls("a:dk1", CT_ColorChoice)
register_element_cls("a:lt1", CT_ColorChoice)
register_element_cls("a:dk2", CT_ColorChoice)
register_element_cls("a:lt2", CT_ColorChoice)
register_element_cls("a:accent1", CT_ColorChoice)
register_element_cls("a:accent2", CT_ColorChoice)
register_element_cls("a:accent3", CT_ColorChoice)
register_element_cls("a:accent4", CT_ColorChoice)
register_element_cls("a:accent5", CT_ColorChoice)
register_element_cls("a:accent6", CT_ColorChoice)
register_element_cls("a:hlink", CT_ColorChoice)
register_element_cls("a:folHlink", CT_ColorChoice)

from .tracked_changes import (
    CT_CellDel,
    CT_CellIns,
    CT_Del,
    CT_DelText,
    CT_Ins,
    CT_MoveFrom,
    CT_MoveTo,
    CT_PPrChange,
    CT_RPrChange,
    CT_SectPrChange,
    CT_TblPrChange,
    CT_TcPrChange,
    CT_TrPrChange,
)

register_element_cls("w:cellDel", CT_CellDel)
register_element_cls("w:cellIns", CT_CellIns)
register_element_cls("w:del", CT_Del)
register_element_cls("w:delText", CT_DelText)
register_element_cls("w:ins", CT_Ins)
register_element_cls("w:moveFrom", CT_MoveFrom)
register_element_cls("w:moveTo", CT_MoveTo)
register_element_cls("w:pPrChange", CT_PPrChange)
register_element_cls("w:rPrChange", CT_RPrChange)
register_element_cls("w:sectPrChange", CT_SectPrChange)
register_element_cls("w:tblPrChange", CT_TblPrChange)
register_element_cls("w:tcPrChange", CT_TcPrChange)
register_element_cls("w:trPrChange", CT_TrPrChange)

from .text.parfmt import (
    CT_Border,
    CT_FramePr,
    CT_Ind,
    CT_Jc,
    CT_PBdr,
    CT_PPr,
    CT_Spacing,
    CT_TabStop,
    CT_TabStops,
)

register_element_cls("w:autoSpaceDE", CT_OnOff)
register_element_cls("w:autoSpaceDN", CT_OnOff)
register_element_cls("w:bar", CT_Border)
register_element_cls("w:bdr", CT_Border)
register_element_cls("w:between", CT_Border)
register_element_cls("w:bottom", CT_Border)
register_element_cls("w:contextualSpacing", CT_OnOff)
register_element_cls("w:framePr", CT_FramePr)
register_element_cls("w:ind", CT_Ind)
register_element_cls("w:insideH", CT_Border)
register_element_cls("w:insideV", CT_Border)
register_element_cls("w:jc", CT_Jc)
register_element_cls("w:keepLines", CT_OnOff)
register_element_cls("w:keepNext", CT_OnOff)
register_element_cls("w:kinsoku", CT_OnOff)
register_element_cls("w:outlineLvl", CT_DecimalNumber)
register_element_cls("w:pageBreakBefore", CT_OnOff)
register_element_cls("w:left", CT_Border)
register_element_cls("w:pBdr", CT_PBdr)
register_element_cls("w:pPr", CT_PPr)
register_element_cls("w:pStyle", CT_String)
register_element_cls("w:right", CT_Border)
register_element_cls("w:spacing", CT_Spacing)
register_element_cls("w:tab", CT_TabStop)
register_element_cls("w:tabs", CT_TabStops)
register_element_cls("w:top", CT_Border)
register_element_cls("w:widowControl", CT_OnOff)
register_element_cls("w:wordWrap", CT_OnOff)

# ---------------------------------------------------------------------------
# Annotation reference elements — used in comments/footnotes markup but do not
# need custom behaviour beyond what BaseOxmlElement provides.  Registering them
# ensures they are recognised by the parser's element-class lookup.

from docx.oxml.xmlchemy import BaseOxmlElement as _Base

register_element_cls("w:annotationRef", _Base)
register_element_cls("w:commentRangeEnd", CT_MarkupRange)
register_element_cls("w:commentRangeStart", CT_MarkupRange)
register_element_cls("w:commentReference", _Base)
register_element_cls("w:contentPart", _Base)
register_element_cls("w:footnoteRef", _Base)
register_element_cls("w:footnoteReference", _Base)
register_element_cls("w:object", _Base)
register_element_cls("o:OLEObject", _Base)

# ---------------------------------------------------------------------------
# SmartArt (DrawingML diagram) elements
#
# Families A (``dgm:relIds``) and B (``dgm:dataModel``) have been covered
# since 0.1.0 of the shared ``ooxml_smartart`` package. 0.2.0 adds
# families D (``dgm:styleDef`` style-label catalogue) and E
# (``dgm:colorsDef`` colour transforms). Register each new tag against
# docx's element-class lookup so ``docx.oxml.parser.parse_xml`` resolves
# the diagram quickStyle and colors parts to the typed classes instead of
# returning plain ``lxml.etree._Element`` stubs.

from .smart_art import (
    CT_ColorTransform,
    CT_ColorTransformHeader,
    CT_ColorTransformHeaderLst,
    CT_Colors,
    CT_Cxn,
    CT_DataModel,
    CT_Pt,
    CT_PtLst,
    CT_RelIds,
    CT_StyleDefinition,
    CT_StyleDefinitionHeader,
    CT_StyleDefinitionHeaderLst,
    CT_StyleLabel,
    CT_TextProps,
)

# -- Family A + B (relIds + dataModel) --
register_element_cls("dgm:cxn", CT_Cxn)
register_element_cls("dgm:dataModel", CT_DataModel)
register_element_cls("dgm:pt", CT_Pt)
register_element_cls("dgm:ptLst", CT_PtLst)
register_element_cls("dgm:relIds", CT_RelIds)

# -- Family D (styleDef) — new in ``ooxml_smartart`` 0.2.0. --
register_element_cls("dgm:styleDef", CT_StyleDefinition)
register_element_cls("dgm:styleDefHdr", CT_StyleDefinitionHeader)
register_element_cls("dgm:styleDefHdrLst", CT_StyleDefinitionHeaderLst)
register_element_cls("dgm:styleLbl", CT_StyleLabel)
register_element_cls("dgm:txPr", CT_TextProps)

# -- Family E (colorsDef) — new in ``ooxml_smartart`` 0.2.0. The six
# -- colour-list tags all map to :class:`CT_Colors` per the shared
# -- package's registration. --
register_element_cls("dgm:colorsDef", CT_ColorTransform)
register_element_cls("dgm:colorsDefHdr", CT_ColorTransformHeader)
register_element_cls("dgm:colorsDefHdrLst", CT_ColorTransformHeaderLst)
register_element_cls("dgm:fillClrLst", CT_Colors)
register_element_cls("dgm:linClrLst", CT_Colors)
register_element_cls("dgm:effectClrLst", CT_Colors)
register_element_cls("dgm:txLinClrLst", CT_Colors)
register_element_cls("dgm:txFillClrLst", CT_Colors)
register_element_cls("dgm:txEffectClrLst", CT_Colors)

# ---------------------------------------------------------------------------
# Chart-related elements

from .chart import (
    CT_AreaChart,
    CT_BandFmt,
    CT_BandFmts,
    CT_Bar3DChart,
    CT_BarChart,
    CT_Chart,
    CT_ChartSpace,
    CT_DepthPercent,
    CT_DoughnutChart,
    CT_HPercent,
    CT_HeaderFooter,
    CT_Line3DChart,
    CT_LineChart,
    CT_MultiLvlStrData,
    CT_MultiLvlStrRef,
    CT_PageMargins,
    CT_PageSetup,
    CT_Perspective,
    CT_PieChart,
    CT_PivotFmt,
    CT_PivotFmts,
    CT_PivotSource,
    CT_PlotArea,
    CT_PrintSettings,
    CT_RelId,
    CT_RotX,
    CT_RotY,
    CT_ScatterChart,
    CT_Ser,
    CT_Shape,
    CT_Surface,
    CT_Surface3DChart,
    CT_SurfaceChart,
    CT_Thickness,
    CT_UpDownBar,
    CT_UpDownBars,
    CT_View3D,
)

register_element_cls("c:chartSpace", CT_ChartSpace)
register_element_cls("c:chart", CT_Chart)
register_element_cls("c:plotArea", CT_PlotArea)
register_element_cls("c:barChart", CT_BarChart)
register_element_cls("c:lineChart", CT_LineChart)
register_element_cls("c:pieChart", CT_PieChart)
register_element_cls("c:doughnutChart", CT_DoughnutChart)
register_element_cls("c:scatterChart", CT_ScatterChart)
register_element_cls("c:areaChart", CT_AreaChart)
register_element_cls("c:ser", CT_Ser)

# -- chart 0.2.0 additions: 3D view + dedicated 3D/surface chart kinds +
# -- indicators + pivot-chart + print-settings + multi-level axis. All
# -- plain re-exports from ``ooxml_chart.oxml.*`` (no docx-local subclass
# -- overrides needed) — registering them in docx's own element-class
# -- lookup so the docx-side parser / ``OxmlElement`` factory resolves
# -- these tags to their typed CT_* classes rather than plain
# -- ``lxml.etree._Element`` (matching the 0.1.0 behaviour for the
# -- classic chart kinds above).
register_element_cls("c:view3D", CT_View3D)
register_element_cls("c:rotX", CT_RotX)
register_element_cls("c:rotY", CT_RotY)
register_element_cls("c:perspective", CT_Perspective)
register_element_cls("c:hPercent", CT_HPercent)
register_element_cls("c:depthPercent", CT_DepthPercent)
register_element_cls("c:floor", CT_Surface)
register_element_cls("c:sideWall", CT_Surface)
register_element_cls("c:backWall", CT_Surface)
register_element_cls("c:thickness", CT_Thickness)
register_element_cls("c:bar3DChart", CT_Bar3DChart)
register_element_cls("c:line3DChart", CT_Line3DChart)
register_element_cls("c:surfaceChart", CT_SurfaceChart)
register_element_cls("c:surface3DChart", CT_Surface3DChart)
register_element_cls("c:shape", CT_Shape)
register_element_cls("c:bandFmt", CT_BandFmt)
register_element_cls("c:bandFmts", CT_BandFmts)
register_element_cls("c:upDownBars", CT_UpDownBars)
register_element_cls("c:upBars", CT_UpDownBar)
register_element_cls("c:downBars", CT_UpDownBar)
register_element_cls("c:pivotSource", CT_PivotSource)
register_element_cls("c:pivotFmts", CT_PivotFmts)
register_element_cls("c:pivotFmt", CT_PivotFmt)
register_element_cls("c:printSettings", CT_PrintSettings)
register_element_cls("c:headerFooter", CT_HeaderFooter)
register_element_cls("c:pageMargins", CT_PageMargins)
register_element_cls("c:pageSetup", CT_PageSetup)
register_element_cls("c:multiLvlStrRef", CT_MultiLvlStrRef)
register_element_cls("c:multiLvlStrCache", CT_MultiLvlStrData)
register_element_cls("c:userShapes", CT_RelId)

# ---------------------------------------------------------------------------
# VML watermark-related elements

from .watermark import CT_Pict, CT_VmlFill, CT_VmlImageData, CT_VmlShape, CT_VmlTextpath

register_element_cls("w:pict", CT_Pict)
register_element_cls("v:fill", CT_VmlFill)
register_element_cls("v:imagedata", CT_VmlImageData)
register_element_cls("v:shape", CT_VmlShape)
register_element_cls("v:textpath", CT_VmlTextpath)

# ---------------------------------------------------------------------------
# Mail-merge / ODSO (Office Data Source Object) elements

from .mail_merge import (
    CT_Base64Binary,
    CT_DataSourceObject,
    CT_MailMergeDataType,
    CT_MailMergeDest,
    CT_MailMergeDocType,
    CT_MailMergeOdsoFMDFieldType,
    CT_MailMergeSourceType,
    CT_Odso,
    CT_OdsoFieldMapData,
    CT_OdsoRecipientData,
    CT_RecipientData,
    CT_TargetScreenSz,
)

# w:mailMerge/<val-wrapper> children — override the generic _CT_MMVal bindings
# for the three spec-typed slots so the correct CT class lights up at parse time.
register_element_cls("w:mainDocumentType", CT_MailMergeDocType)
register_element_cls("w:dataType", CT_MailMergeDataType)
register_element_cls("w:destination", CT_MailMergeDest)

# ODSO substructure
register_element_cls("w:odso", CT_Odso)
register_element_cls("w:fieldMapData", CT_OdsoFieldMapData)
register_element_cls("w:dataSource", CT_DataSourceObject)
register_element_cls("w:headerSource", CT_DataSourceObject)
register_element_cls("w:src", CT_DataSourceObject)

# w:odso/w:type — reuse CT_MailMergeSourceType
# (Note: w:odso/w:fieldMapData/w:type shares the `w:type` QName; lxml registers
# a single class per QName. The source-type form is the one actually occurring
# at document root; the FMD form is only reachable inside `w:fieldMapData` and
# is parse-compatible because both carry a required `@w:val` string attribute.)

# w:recipients top-level part
register_element_cls("w:recipients", CT_OdsoRecipientData)
register_element_cls("w:recipientData", CT_RecipientData)
register_element_cls("w:uniqueTag", CT_Base64Binary)

# w:targetScreenSz — settings-root child
register_element_cls("w:targetScreenSz", CT_TargetScreenSz)

# Silence unused-import warnings for CT classes exported for tests / consumers
_ = (CT_MailMergeSourceType, CT_MailMergeOdsoFMDFieldType)

# ---------------------------------------------------------------------------
# Inline CustomXml container elements

from .custom_xml import (
    CT_Attr,
    CT_CustomXmlBlock,
    CT_CustomXmlCell,
    CT_CustomXmlPr,
    CT_CustomXmlRow,
    CT_CustomXmlRun,
)

# w:customXml is a single QName used in block / row / cell / run positions;
# register the block flavor as the default — its grammar (customXmlPr +
# block-content) is the most permissive and parses the other three flavors
# faithfully. CT_CustomXmlRow / CT_CustomXmlCell / CT_CustomXmlRun remain
# available for programmatic construction.
register_element_cls("w:customXml", CT_CustomXmlBlock)
register_element_cls("w:customXmlPr", CT_CustomXmlPr)
register_element_cls("w:attr", CT_Attr)

_ = (CT_CustomXmlRow, CT_CustomXmlCell, CT_CustomXmlRun)


# ---------------------------------------------------------------------------
# DrawingML table-style vocabulary (``a:tblStyleLst`` + descendants, 22 tags).
#
# The grammar lives in ``ooxml_shared_drawingml.tblstyle``; docx re-exports
# via ``docx.oxml.tblstyle`` (pure shim) and wires the 22 namespaced tags
# into docx's element-class registry here so the parser instantiates the
# rich shared CT_* classes when reading DrawingML table-style content
# (DrawingML-framework table styles accompanying inline ``w:drawing``
# graphic-frame tables). Not to be confused with WordprocessingML's own
# table-style family (``w:tblStyle`` / ``w:tblStylePr`` in styles.xml —
# a distinct vocabulary modelled locally in ``docx.oxml.styles``).

from .tblstyle import (
    CT_Cell3D,
    CT_TableBackgroundStyle,
    CT_TableCellBorderStyle,
    CT_TablePartStyle,
    CT_TableStyle,
    CT_TableStyleCellStyle,
    CT_TableStyleList,
    CT_TableStyleTextStyle,
    CT_ThemeableLineStyle,
)
from ooxml_shared_drawingml.scene3d import CT_Bevel as _SDML_CT_Bevel

register_element_cls("a:tblStyleLst", CT_TableStyleList)
register_element_cls("a:tblStyle", CT_TableStyle)
register_element_cls("a:tblBg", CT_TableBackgroundStyle)
# -- 13 part-style slots all share CT_TablePartStyle --
register_element_cls("a:wholeTbl", CT_TablePartStyle)
register_element_cls("a:band1H", CT_TablePartStyle)
register_element_cls("a:band2H", CT_TablePartStyle)
register_element_cls("a:band1V", CT_TablePartStyle)
register_element_cls("a:band2V", CT_TablePartStyle)
register_element_cls("a:firstRow", CT_TablePartStyle)
register_element_cls("a:lastRow", CT_TablePartStyle)
register_element_cls("a:firstCol", CT_TablePartStyle)
register_element_cls("a:lastCol", CT_TablePartStyle)
register_element_cls("a:nwCell", CT_TablePartStyle)
register_element_cls("a:neCell", CT_TablePartStyle)
register_element_cls("a:swCell", CT_TablePartStyle)
register_element_cls("a:seCell", CT_TablePartStyle)
register_element_cls("a:tcTxStyle", CT_TableStyleTextStyle)
register_element_cls("a:tcStyle", CT_TableStyleCellStyle)
register_element_cls("a:tcBdr", CT_TableCellBorderStyle)
register_element_cls("a:cell3D", CT_Cell3D)
# -- ``a:bevel`` (singular): the mandatory child of CT_Cell3D. CT_Bevel
# -- has only optional attributes, so the same class services the empty
# -- ``a:bevel`` marker that appears as a line-join style inside
# -- CT_LineProperties.
register_element_cls("a:bevel", _SDML_CT_Bevel)
# -- eight border slots of CT_TableCellBorderStyle (CT_ThemeableLineStyle) --
register_element_cls("a:left", CT_ThemeableLineStyle)
register_element_cls("a:right", CT_ThemeableLineStyle)
register_element_cls("a:top", CT_ThemeableLineStyle)
register_element_cls("a:bottom", CT_ThemeableLineStyle)
register_element_cls("a:insideH", CT_ThemeableLineStyle)
register_element_cls("a:insideV", CT_ThemeableLineStyle)
register_element_cls("a:tl2br", CT_ThemeableLineStyle)
register_element_cls("a:tr2bl", CT_ThemeableLineStyle)
