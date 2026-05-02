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
    CT_WordprocessingShape,
)
from docx.oxml.parser import OxmlElement, parse_xml, register_element_cls
from docx.oxml.shape import (
    CT_Anchor,
    CT_Blip,
    CT_BlipFillProperties,
    CT_GraphicalObject,
    CT_GraphicalObjectData,
    CT_Inline,
    CT_NonVisualDrawingProps,
    CT_Picture,
    CT_PictureNonVisual,
    CT_Point2D,
    CT_PosOffset,
    CT_PositionH,
    CT_PositionV,
    CT_PositiveSize2D,
    CT_ShapeProperties,
    CT_Transform2D,
    CT_WrapNone,
    CT_WrapSquare,
    CT_WrapThrough,
    CT_WrapTight,
    CT_WrapTopBottom,
)
from docx.oxml.shared import CT_DecimalNumber, CT_OnOff, CT_String
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

register_element_cls("a:blip", CT_Blip)
register_element_cls("a:ext", CT_PositiveSize2D)
register_element_cls("a:graphic", CT_GraphicalObject)
register_element_cls("a:graphicData", CT_GraphicalObjectData)
register_element_cls("a:off", CT_Point2D)
register_element_cls("a:xfrm", CT_Transform2D)
register_element_cls("pic:blipFill", CT_BlipFillProperties)
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
register_element_cls("wp:wrapNone", CT_WrapNone)
register_element_cls("wp:wrapSquare", CT_WrapSquare)
register_element_cls("wp:wrapThrough", CT_WrapThrough)
register_element_cls("wp:wrapTight", CT_WrapTight)
register_element_cls("wp:wrapTopAndBottom", CT_WrapTopBottom)
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
register_element_cls("w:r", CT_R)
register_element_cls("w:sym", CT_Sym)
register_element_cls("w:t", CT_Text)

# ---------------------------------------------------------------------------
# header/footer-related mappings

register_element_cls("w:bidi", CT_OnOff)
register_element_cls("w:evenAndOddHeaders", CT_OnOff)
register_element_cls("w:titlePg", CT_OnOff)

# ---------------------------------------------------------------------------
# other custom element class mappings

from .bookmarks import CT_BookmarkEnd, CT_BookmarkStart

register_element_cls("w:bookmarkEnd", CT_BookmarkEnd)
register_element_cls("w:bookmarkStart", CT_BookmarkStart)

from .comments import CT_Comments, CT_Comment

register_element_cls("w:comments", CT_Comments)
register_element_cls("w:comment", CT_Comment)

from .content_controls import CT_DataBinding, CT_Sdt, CT_SdtContent, CT_SdtPr

register_element_cls("w:dataBinding", CT_DataBinding)
register_element_cls("w:sdt", CT_Sdt)
register_element_cls("w:sdtContent", CT_SdtContent)
register_element_cls("w:sdtPr", CT_SdtPr)

from .coreprops import CT_CoreProperties

register_element_cls("cp:coreProperties", CT_CoreProperties)

from .custom_properties import CT_CustomProperties, CT_CustomProperty

register_element_cls("custprops:Properties", CT_CustomProperties)
register_element_cls("custprops:property", CT_CustomProperty)

from .endnotes import CT_EdnDocProps, CT_Endnote, CT_Endnotes

register_element_cls("w:endnote", CT_Endnote)
register_element_cls("w:endnotes", CT_Endnotes)
register_element_cls("w:endnotePr", CT_EdnDocProps)

from .glossary import (
    CT_DocPart,
    CT_DocPartBody,
    CT_DocPartCategory,
    CT_DocPartGallery,
    CT_DocPartPr,
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

from .fields import CT_FldChar, CT_FldSimple, CT_InstrText

register_element_cls("w:fldChar", CT_FldChar)
register_element_cls("w:fldSimple", CT_FldSimple)
register_element_cls("w:instrText", CT_InstrText)

from .form_fields import CT_FFCheckBox, CT_FFData, CT_FFDDList, CT_FFTextInput

register_element_cls("w:ffData", CT_FFData)
register_element_cls("w:textInput", CT_FFTextInput)
register_element_cls("w:checkBox", CT_FFCheckBox)
register_element_cls("w:ddList", CT_FFDDList)

from .document import CT_Background, CT_Body, CT_Document

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
register_element_cls("w:pgSz", CT_PageSz)
register_element_cls("w:sectPr", CT_SectPr)
register_element_cls("w:type", CT_SectType)

from .settings import (
    CT_Compat,
    CT_CompatSetting,
    CT_DefaultTabStop,
    CT_DocProtect,
    CT_LongHexNumber,
    CT_MailMerge,
    CT_Rsids,
    CT_Settings,
    CT_View,
    CT_Zoom,
    _CT_MMVal,
)

register_element_cls("w:compat", CT_Compat)
register_element_cls("w:compatSetting", CT_CompatSetting)
register_element_cls("w:defaultTabStop", CT_DefaultTabStop)
register_element_cls("w:documentProtection", CT_DocProtect)
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
register_element_cls("w:trackRevisions", CT_OnOff)
register_element_cls("w:view", CT_View)
register_element_cls("w:zoom", CT_Zoom)

from .styles import CT_LatentStyles, CT_LsdException, CT_Style, CT_Styles

register_element_cls("w:autoRedefine", CT_OnOff)
register_element_cls("w:basedOn", CT_String)
register_element_cls("w:latentStyles", CT_LatentStyles)
register_element_cls("w:link", CT_String)
register_element_cls("w:locked", CT_OnOff)
register_element_cls("w:lsdException", CT_LsdException)
register_element_cls("w:name", CT_String)
register_element_cls("w:next", CT_String)
register_element_cls("w:qFormat", CT_OnOff)
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
register_element_cls("w:tblGrid", CT_TblGrid)
register_element_cls("w:tblHeader", CT_OnOff)
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
    CT_RPr,
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
register_element_cls("w:webHidden", CT_OnOff)

from .text.paragraph import CT_P

register_element_cls("w:p", CT_P)

from .web_settings import CT_Encoding, CT_OptimizeForBrowser, CT_WebSettings

register_element_cls("w:webSettings", CT_WebSettings)
register_element_cls("w:encoding", CT_Encoding)
register_element_cls("w:optimizeForBrowser", CT_OptimizeForBrowser)
register_element_cls("w:relyOnVML", CT_OnOff)
register_element_cls("w:allowPNG", CT_OnOff)
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

register_element_cls("w:bar", CT_Border)
register_element_cls("w:bdr", CT_Border)
register_element_cls("w:between", CT_Border)
register_element_cls("w:bottom", CT_Border)
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
register_element_cls("w:commentRangeEnd", _Base)
register_element_cls("w:commentRangeStart", _Base)
register_element_cls("w:commentReference", _Base)
register_element_cls("w:contentPart", _Base)
register_element_cls("w:footnoteRef", _Base)
register_element_cls("w:footnoteReference", _Base)
register_element_cls("w:object", _Base)
register_element_cls("o:OLEObject", _Base)

# ---------------------------------------------------------------------------
# SmartArt (DrawingML diagram) elements

from .smart_art import CT_Cxn, CT_DataModel, CT_Pt, CT_PtLst, CT_RelIds

register_element_cls("dgm:cxn", CT_Cxn)
register_element_cls("dgm:dataModel", CT_DataModel)
register_element_cls("dgm:pt", CT_Pt)
register_element_cls("dgm:ptLst", CT_PtLst)
register_element_cls("dgm:relIds", CT_RelIds)

# ---------------------------------------------------------------------------
# Chart-related elements

from .chart import (
    CT_AreaChart,
    CT_BarChart,
    CT_Chart,
    CT_ChartSpace,
    CT_DoughnutChart,
    CT_LineChart,
    CT_PieChart,
    CT_PlotArea,
    CT_ScatterChart,
    CT_Ser,
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

# ---------------------------------------------------------------------------
# VML watermark-related elements

from .watermark import CT_Pict, CT_VmlFill, CT_VmlImageData, CT_VmlShape, CT_VmlTextpath

register_element_cls("w:pict", CT_Pict)
register_element_cls("v:fill", CT_VmlFill)
register_element_cls("v:imagedata", CT_VmlImageData)
register_element_cls("v:shape", CT_VmlShape)
register_element_cls("v:textpath", CT_VmlTextpath)
