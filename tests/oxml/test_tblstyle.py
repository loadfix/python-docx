"""Smoke-test suite for `docx.oxml.tblstyle` re-export shim.

Verifies the 9 shared-drawingml 0.4.0 DrawingML table-style CT_*
classes are:
- importable from the historical ``docx.oxml.tblstyle`` module path,
- wired into docx's element-class registry for the 22 namespaced tags
  that compose the ``a:tblStyleLst`` subtree, and
- round-trip faithfully through docx's hardened parser.

Note this exercises the DrawingML ``a:tblStyleLst`` vocabulary —
distinct from WordprocessingML's ``w:tblStyle`` / ``w:tblStylePr``
family, which docx continues to model locally in ``docx.oxml.styles``.
"""

from __future__ import annotations

from ooxml_shared_drawingml.scene3d import CT_Bevel
from docx.oxml import parse_xml
from docx.oxml.tblstyle import (
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


class DescribeTableStyleShim:
    """The shim re-exports 9 shared-drawingml classes unchanged."""

    def it_re_exports_the_nine_shared_drawingml_classes(self):
        from ooxml_shared_drawingml import tblstyle as shared

        assert CT_Cell3D is shared.CT_Cell3D
        assert CT_TableBackgroundStyle is shared.CT_TableBackgroundStyle
        assert CT_TableCellBorderStyle is shared.CT_TableCellBorderStyle
        assert CT_TablePartStyle is shared.CT_TablePartStyle
        assert CT_TableStyle is shared.CT_TableStyle
        assert CT_TableStyleCellStyle is shared.CT_TableStyleCellStyle
        assert CT_TableStyleList is shared.CT_TableStyleList
        assert CT_TableStyleTextStyle is shared.CT_TableStyleTextStyle
        assert CT_ThemeableLineStyle is shared.CT_ThemeableLineStyle


class DescribeTableStyleRegistration:
    """Parsing ``a:tblStyleLst`` XML through docx's parser yields rich CT_* instances."""

    def it_parses_tblStyleLst_as_CT_TableStyleList(self):
        xml = (
            '<a:tblStyleLst'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'
        )
        element = parse_xml(xml)

        assert isinstance(element, CT_TableStyleList)
        assert element.def_ == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

    def it_parses_the_full_subtree_as_rich_CT_instances(self):
        xml = (
            '<a:tblStyleLst'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' def="{00000000-0000-0000-0000-000000000001}">'
            '  <a:tblStyle'
            '   styleId="{00000000-0000-0000-0000-000000000002}"'
            '   styleName="Demo">'
            '    <a:tblBg/>'
            '    <a:wholeTbl>'
            '      <a:tcTxStyle/>'
            '      <a:tcStyle>'
            '        <a:tcBdr>'
            '          <a:left/><a:right/><a:top/><a:bottom/>'
            '          <a:insideH/><a:insideV/>'
            '          <a:tl2br/><a:tr2bl/>'
            '        </a:tcBdr>'
            '        <a:cell3D><a:bevel/></a:cell3D>'
            '      </a:tcStyle>'
            '    </a:wholeTbl>'
            '    <a:band1H/>'
            '  </a:tblStyle>'
            '</a:tblStyleLst>'
        )
        root = parse_xml(xml)

        assert isinstance(root, CT_TableStyleList)
        tbl_styles = root.tblStyle_lst
        assert len(tbl_styles) == 1
        ts = tbl_styles[0]
        assert isinstance(ts, CT_TableStyle)
        assert ts.styleId == "{00000000-0000-0000-0000-000000000002}"
        assert ts.styleName == "Demo"

        assert isinstance(ts.tblBg, CT_TableBackgroundStyle)
        assert isinstance(ts.wholeTbl, CT_TablePartStyle)
        assert isinstance(ts.band1H, CT_TablePartStyle)

        tc_style = ts.wholeTbl.tcStyle
        assert isinstance(tc_style, CT_TableStyleCellStyle)
        assert isinstance(tc_style.tcBdr, CT_TableCellBorderStyle)
        assert isinstance(tc_style.cell3D, CT_Cell3D)
        bevel = tc_style.cell3D.bevel
        assert isinstance(bevel, CT_Bevel)

        bdr = tc_style.tcBdr
        for slot in ("left", "right", "top", "bottom",
                     "insideH", "insideV", "tl2br", "tr2bl"):
            assert isinstance(getattr(bdr, slot), CT_ThemeableLineStyle)

        tc_tx = ts.wholeTbl.tcTxStyle
        assert isinstance(tc_tx, CT_TableStyleTextStyle)


class DescribeTableStyleRoundTrip:
    """Parse -> serialise preserves the full subtree bit-for-bit."""

    def it_round_trips_a_minimal_style_list(self):
        xml = (
            '<a:tblStyleLst'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' def="{AAAA1111-0000-0000-0000-000000000000}">'
            '<a:tblStyle'
            ' styleId="{BBBB2222-0000-0000-0000-000000000000}"'
            ' styleName="Round-trip"/>'
            '</a:tblStyleLst>'
        )
        root = parse_xml(xml)
        assert len(root.tblStyle_lst) == 1

        from lxml import etree

        out = etree.tostring(root, encoding="unicode")
        assert 'def="{AAAA1111-0000-0000-0000-000000000000}"' in out
        assert 'styleId="{BBBB2222-0000-0000-0000-000000000000}"' in out
        assert 'styleName="Round-trip"' in out

    def it_can_programmatically_build_a_tblStyle_entry(self):
        xml = (
            '<a:tblStyleLst'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' def="{00000000-0000-0000-0000-000000000000}"/>'
        )
        root = parse_xml(xml)

        ts = root.add_tblStyle()
        ts.styleId = "{11111111-0000-0000-0000-000000000000}"
        ts.styleName = "Programmatic"
        assert isinstance(ts, CT_TableStyle)

        whole = ts.get_or_add_wholeTbl()
        assert isinstance(whole, CT_TablePartStyle)

        from lxml import etree

        out = etree.tostring(root, encoding="unicode")
        assert 'styleId="{11111111-0000-0000-0000-000000000000}"' in out
        assert "<a:wholeTbl" in out
