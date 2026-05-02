# pyright: reportPrivateUsage=false

"""Test suite for the docx.oxml.text module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.table import (
    WD_BORDER_STYLE,
    WD_ROW_HEIGHT_RULE,
    WD_SHADING_PATTERN,
    WD_TEXT_DIRECTION,
)
from docx.exceptions import InvalidSpanError
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.oxml.table import (
    CT_Border,
    CT_Row,
    CT_Shd,
    CT_Tbl,
    CT_TblBorders,
    CT_TblLook,
    CT_TblPr,
    CT_TblWidth,
    CT_Tc,
    CT_TcBorders,
    CT_TcMar,
    CT_TcPr,
)
from docx.oxml.text.paragraph import CT_P
from docx.shared import Emu, Inches, Length, Pt, RGBColor, Twips

from ..unitutil.cxml import element, xml
from ..unitutil.file import snippet_seq
from ..unitutil.mock import FixtureRequest, Mock, call, instance_mock, method_mock, property_mock


class DescribeCT_Border:
    """Unit-test suite for `docx.oxml.table.CT_Border` objects."""

    @pytest.mark.parametrize(
        ("border_cxml", "expected_val"),
        [
            ("w:top", None),
            ("w:top{w:val=single}", WD_BORDER_STYLE.SINGLE),
            ("w:top{w:val=double}", WD_BORDER_STYLE.DOUBLE),
            ("w:top{w:val=none}", WD_BORDER_STYLE.NONE),
        ],
    )
    def it_can_get_the_val_attribute(
        self, border_cxml: str, expected_val: WD_BORDER_STYLE | None
    ):
        border = cast(CT_Border, element(border_cxml))
        assert border.val == expected_val

    @pytest.mark.parametrize(
        ("border_cxml", "expected_sz"),
        [
            ("w:top", None),
            # `sz` is eighth-points, exposed as Length (EMU). 4 eighth-points = Pt(0.5).
            ("w:top{w:sz=4}", Pt(0.5)),
            ("w:top{w:sz=12}", Pt(1.5)),
        ],
    )
    def it_can_get_the_sz_attribute(
        self, border_cxml: str, expected_sz: Length | None
    ):
        border = cast(CT_Border, element(border_cxml))
        assert border.sz == expected_sz

    @pytest.mark.parametrize(
        ("border_cxml", "expected_color"),
        [
            ("w:top", None),
            ("w:top{w:color=FF0000}", RGBColor(0xFF, 0x00, 0x00)),
            ("w:top{w:color=auto}", "auto"),
        ],
    )
    def it_can_get_the_color_attribute(
        self, border_cxml: str, expected_color: RGBColor | str | None
    ):
        border = cast(CT_Border, element(border_cxml))
        assert border.color == expected_color

    @pytest.mark.parametrize(
        ("border_cxml", "expected_space"),
        [
            ("w:top", None),
            # `space` is in whole points, exposed as Length (EMU).
            ("w:top{w:space=0}", Pt(0)),
            ("w:top{w:space=4}", Pt(4)),
        ],
    )
    def it_can_get_the_space_attribute(
        self, border_cxml: str, expected_space: Length | None
    ):
        border = cast(CT_Border, element(border_cxml))
        assert border.space == expected_space


class DescribeCT_TblBorders:
    """Unit-test suite for `docx.oxml.table.CT_TblBorders` objects."""

    def it_can_get_and_add_border_children(self):
        tblBorders = cast(CT_TblBorders, element("w:tblBorders"))
        assert tblBorders.top is None
        top = tblBorders.get_or_add_top()
        assert isinstance(top, CT_Border)
        assert tblBorders.top is top

    def it_inserts_borders_in_the_right_order(self):
        tblBorders = cast(CT_TblBorders, element("w:tblBorders"))
        tblBorders.get_or_add_insideV()
        tblBorders.get_or_add_top()
        expected = xml("w:tblBorders/(w:top,w:insideV)")
        assert tblBorders.xml == expected

    @pytest.mark.parametrize("attr", ["top", "left", "bottom", "right", "insideH", "insideV"])
    def it_can_remove_each_border(self, attr: str):
        tblBorders = cast(CT_TblBorders, element("w:tblBorders"))
        get_or_add = getattr(tblBorders, f"get_or_add_{attr}")
        remove = getattr(tblBorders, f"_remove_{attr}")
        get_or_add()
        assert getattr(tblBorders, attr) is not None
        remove()
        assert getattr(tblBorders, attr) is None


class DescribeCT_TcBorders:
    """Unit-test suite for `docx.oxml.table.CT_TcBorders` objects."""

    def it_can_get_and_add_border_children(self):
        tcBorders = cast(CT_TcBorders, element("w:tcBorders"))
        assert tcBorders.top is None
        top = tcBorders.get_or_add_top()
        assert isinstance(top, CT_Border)
        assert tcBorders.top is top

    def it_inserts_borders_in_the_right_order(self):
        tcBorders = cast(CT_TcBorders, element("w:tcBorders"))
        tcBorders.get_or_add_right()
        tcBorders.get_or_add_top()
        expected = xml("w:tcBorders/(w:top,w:right)")
        assert tcBorders.xml == expected


class DescribeCT_TblPr_borders:
    """Unit-test suite for border-related features of CT_TblPr."""

    def it_can_get_the_tblBorders_child(self):
        tblPr = cast(CT_TblPr, element("w:tblPr"))
        assert tblPr.tblBorders is None

    def it_can_add_tblBorders(self):
        tblPr = cast(CT_TblPr, element("w:tblPr"))
        tblBorders = tblPr.get_or_add_tblBorders()
        assert isinstance(tblBorders, CT_TblBorders)
        assert tblPr.tblBorders is tblBorders

    def it_inserts_tblBorders_in_the_right_position(self):
        tblPr = cast(CT_TblPr, element("w:tblPr/(w:tblStyle,w:tblLayout)"))
        tblPr.get_or_add_tblBorders()
        expected = xml("w:tblPr/(w:tblStyle,w:tblBorders,w:tblLayout)")
        assert tblPr.xml == expected


class DescribeCT_TblPr_tblLook:
    """Unit-test suite for `w:tblLook` access via CT_TblPr."""

    def it_is_None_when_no_tblLook_child_is_present(self):
        tblPr = cast(CT_TblPr, element("w:tblPr"))
        assert tblPr.tblLook is None

    def it_can_add_tblLook(self):
        tblPr = cast(CT_TblPr, element("w:tblPr"))
        tblLook = tblPr.get_or_add_tblLook()
        assert isinstance(tblLook, CT_TblLook)
        assert tblPr.tblLook is tblLook

    def it_inserts_tblLook_after_tblCellMar(self):
        tblPr = cast(CT_TblPr, element("w:tblPr/(w:tblStyle,w:tblLayout)"))
        tblPr.get_or_add_tblLook()
        expected = xml("w:tblPr/(w:tblStyle,w:tblLayout,w:tblLook)")
        assert tblPr.xml == expected


class DescribeCT_TblLook:
    """Unit-test suite for `docx.oxml.table.CT_TblLook` objects."""

    @pytest.mark.parametrize(
        ("name", "cxml", "expected"),
        [
            ("firstRow", "w:tblLook", False),
            ("firstRow", 'w:tblLook{w:firstRow=1}', True),
            ("firstRow", 'w:tblLook{w:firstRow=0}', False),
            ("firstRow", 'w:tblLook{w:firstRow=true}', True),
            ("lastRow", 'w:tblLook{w:lastRow=1}', True),
            ("firstColumn", 'w:tblLook{w:firstColumn=1}', True),
            ("lastColumn", 'w:tblLook{w:lastColumn=1}', True),
            ("noHBand", 'w:tblLook{w:noHBand=1}', True),
            ("noVBand", 'w:tblLook{w:noVBand=1}', True),
        ],
    )
    def it_reads_individual_flag_attrs(self, name: str, cxml: str, expected: bool):
        tblLook = cast(CT_TblLook, element(cxml))
        assert tblLook.get_flag(name) is expected

    def it_falls_back_to_the_legacy_val_bitmask_when_flag_attr_absent(self):
        # 0x04A0 = firstRow(0x0020) | firstColumn(0x0080) | noVBand(0x0400)
        tblLook = cast(CT_TblLook, element("w:tblLook{w:val=04A0}"))
        assert tblLook.get_flag("firstRow") is True
        assert tblLook.get_flag("firstColumn") is True
        assert tblLook.get_flag("noVBand") is True
        assert tblLook.get_flag("lastRow") is False
        assert tblLook.get_flag("lastColumn") is False
        assert tblLook.get_flag("noHBand") is False

    def it_prefers_individual_flag_attr_over_legacy_val(self):
        # val bitmask says firstRow=1, but explicit attr says 0
        tblLook = cast(
            CT_TblLook, element("w:tblLook{w:val=04A0,w:firstRow=0}")
        )
        assert tblLook.get_flag("firstRow") is False

    def it_ignores_malformed_val_bitmask(self):
        tblLook = cast(CT_TblLook, element("w:tblLook{w:val=notahex}"))
        assert tblLook.get_flag("firstRow") is False

    def it_writes_True_as_1(self):
        tblLook = cast(CT_TblLook, element("w:tblLook"))
        tblLook.set_flag("firstRow", True)
        assert tblLook.xml == xml('w:tblLook{w:firstRow=1}')

    def it_writes_False_as_0(self):
        tblLook = cast(CT_TblLook, element("w:tblLook"))
        tblLook.set_flag("firstRow", False)
        assert tblLook.xml == xml('w:tblLook{w:firstRow=0}')

    def it_overwrites_an_existing_flag(self):
        tblLook = cast(CT_TblLook, element('w:tblLook{w:firstRow=1}'))
        tblLook.set_flag("firstRow", False)
        assert tblLook.get_flag("firstRow") is False

    def it_round_trips_each_flag(self):
        tblLook = cast(CT_TblLook, element("w:tblLook"))
        names = ("firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand")
        for name in names:
            tblLook.set_flag(name, True)
            assert tblLook.get_flag(name) is True
            tblLook.set_flag(name, False)
            assert tblLook.get_flag(name) is False


class DescribeCT_TblPr_width:
    """Unit-test suite for `w:tblW` features of CT_TblPr."""

    def it_is_None_when_no_tblW_child_is_present(self):
        tblPr = cast(CT_TblPr, element("w:tblPr"))
        assert tblPr.tblW is None
        assert tblPr.preferred_width is None

    def it_returns_None_for_non_dxa_tblW(self):
        tblPr = cast(CT_TblPr, element("w:tblPr/w:tblW{w:type=pct,w:w=5000}"))
        assert tblPr.preferred_width is None

    def it_returns_EMU_for_dxa_tblW(self):
        tblPr = cast(CT_TblPr, element("w:tblPr/w:tblW{w:type=dxa,w:w=1440}"))
        assert tblPr.preferred_width == Inches(1)

    def it_can_set_a_preferred_width(self):
        tblPr = cast(CT_TblPr, element("w:tblPr"))
        tblPr.preferred_width = Inches(1)
        assert tblPr.xml == xml("w:tblPr/w:tblW{w:type=dxa,w:w=1440}")

    def it_can_clear_a_preferred_width(self):
        tblPr = cast(CT_TblPr, element("w:tblPr/w:tblW{w:type=dxa,w:w=1440}"))
        tblPr.preferred_width = None
        assert tblPr.xml == xml("w:tblPr")

    def it_inserts_tblW_in_the_right_position(self):
        tblPr = cast(CT_TblPr, element("w:tblPr/(w:tblStyle,w:tblLayout)"))
        tblPr.set_tblW(1440, "dxa")
        expected = xml(
            "w:tblPr/(w:tblStyle,w:tblW{w:type=dxa,w:w=1440},w:tblLayout)"
        )
        assert tblPr.xml == expected

    def it_can_update_an_existing_tblW(self):
        tblPr = cast(CT_TblPr, element("w:tblPr/w:tblW{w:type=auto,w:w=0}"))
        tblPr.set_tblW(5000, "pct")
        assert tblPr.xml == xml("w:tblPr/w:tblW{w:type=pct,w:w=5000}")


class DescribeCT_TblWidth:
    """Unit-test suite for `docx.oxml.table.CT_TblWidth` objects."""

    def it_returns_None_when_type_is_not_dxa(self):
        tblW = cast(CT_TblWidth, element("w:tblW{w:type=auto,w:w=0}"))
        assert tblW.width is None

    def it_returns_EMU_length_for_dxa(self):
        tblW = cast(CT_TblWidth, element("w:tblW{w:type=dxa,w:w=1440}"))
        assert tblW.width == Inches(1)

    def it_switches_to_dxa_when_width_is_set(self):
        tblW = cast(CT_TblWidth, element("w:tblW{w:type=pct,w:w=5000}"))
        tblW.width = Emu(914400)
        assert tblW.type == "dxa"
        assert tblW.w == 1440


class DescribeCT_TcPr_borders:
    """Unit-test suite for border-related features of CT_TcPr."""

    def it_can_get_the_tcBorders_child(self):
        tcPr = cast(CT_TcPr, element("w:tcPr"))
        assert tcPr.tcBorders is None

    def it_can_add_tcBorders(self):
        tcPr = cast(CT_TcPr, element("w:tcPr"))
        tcBorders = tcPr.get_or_add_tcBorders()
        assert isinstance(tcBorders, CT_TcBorders)
        assert tcPr.tcBorders is tcBorders

    def it_inserts_tcBorders_in_the_right_position(self):
        tcPr = cast(CT_TcPr, element("w:tcPr/(w:tcW,w:shd)"))
        tcPr.get_or_add_tcBorders()
        expected = xml("w:tcPr/(w:tcW,w:tcBorders,w:shd)")
        assert tcPr.xml == expected


class DescribeCT_TcMar:
    """Unit-test suite for `docx.oxml.table.CT_TcMar` objects."""

    def it_returns_None_for_all_edges_when_empty(self):
        tcMar = cast(CT_TcMar, element("w:tcMar"))
        assert tcMar.get_margin("top") is None
        assert tcMar.get_margin("bottom") is None
        assert tcMar.get_margin("start") is None
        assert tcMar.get_margin("end") is None

    @pytest.mark.parametrize(
        ("edge", "tag"),
        [("top", "w:top"), ("bottom", "w:bottom"), ("start", "w:start"), ("end", "w:end")],
    )
    def it_can_read_an_edge_value(self, edge: str, tag: str):
        tcMar = cast(CT_TcMar, element("w:tcMar/%s{w:w=144,w:type=dxa}" % tag))
        assert tcMar.get_margin(edge) == Twips(144)

    def it_reads_start_from_legacy_w_left(self):
        tcMar = cast(CT_TcMar, element("w:tcMar/w:left{w:w=240,w:type=dxa}"))
        assert tcMar.get_margin("start") == Twips(240)

    def it_reads_end_from_legacy_w_right(self):
        tcMar = cast(CT_TcMar, element("w:tcMar/w:right{w:w=360,w:type=dxa}"))
        assert tcMar.get_margin("end") == Twips(360)

    @pytest.mark.parametrize(
        ("edge", "value"),
        [
            ("top", Inches(0.1)),
            ("bottom", Pt(6)),
            ("start", Twips(100)),
            ("end", Inches(0.25)),
        ],
    )
    def it_can_round_trip_a_margin_value(self, edge: str, value):
        tcMar = cast(CT_TcMar, element("w:tcMar"))
        tcMar.set_margin(edge, value)
        assert tcMar.get_margin(edge) == value

    def it_writes_start_as_w_start_even_when_legacy_left_is_present(self):
        tcMar = cast(CT_TcMar, element("w:tcMar/w:left{w:w=100,w:type=dxa}"))
        tcMar.set_margin("start", Twips(200))
        # -- legacy w:left should be replaced by w:start --
        assert tcMar.get_margin("start") == Twips(200)
        assert tcMar.find(qn("w:left")) is None
        assert tcMar.find(qn("w:start")) is not None

    def it_can_remove_a_margin_edge(self):
        tcMar = cast(
            CT_TcMar,
            element("w:tcMar/(w:top{w:w=100,w:type=dxa},w:bottom{w:w=200,w:type=dxa})"),
        )
        tcMar.remove_margin("top")
        assert tcMar.get_margin("top") is None
        assert tcMar.get_margin("bottom") == Twips(200)

    def it_removes_the_legacy_tag_when_asked_to_remove_start_or_end(self):
        tcMar = cast(
            CT_TcMar,
            element("w:tcMar/(w:left{w:w=100,w:type=dxa},w:right{w:w=200,w:type=dxa})"),
        )
        tcMar.remove_margin("start")
        tcMar.remove_margin("end")
        assert tcMar.find(qn("w:left")) is None
        assert tcMar.find(qn("w:right")) is None

    def it_keeps_children_in_schema_order(self):
        tcMar = cast(CT_TcMar, element("w:tcMar"))
        tcMar.set_margin("end", Twips(40))
        tcMar.set_margin("top", Twips(10))
        tcMar.set_margin("bottom", Twips(30))
        tcMar.set_margin("start", Twips(20))
        expected = xml(
            "w:tcMar/(w:top{w:w=10,w:type=dxa},w:start{w:w=20,w:type=dxa},"
            "w:bottom{w:w=30,w:type=dxa},w:end{w:w=40,w:type=dxa})"
        )
        assert tcMar.xml == expected

    def it_raises_on_unknown_edge_name(self):
        tcMar = cast(CT_TcMar, element("w:tcMar"))
        with pytest.raises(ValueError):
            tcMar.get_margin("middle")


class DescribeCT_TcPr_margins:
    """Unit-test suite for `w:tcMar` features of CT_TcPr."""

    def it_is_None_when_no_tcMar_child_is_present(self):
        tcPr = cast(CT_TcPr, element("w:tcPr"))
        assert tcPr.tcMar is None

    def it_can_add_a_tcMar_child(self):
        tcPr = cast(CT_TcPr, element("w:tcPr"))
        tcMar = tcPr.get_or_add_tcMar()
        assert isinstance(tcMar, CT_TcMar)
        assert tcPr.tcMar is tcMar

    def it_inserts_tcMar_in_the_right_position(self):
        tcPr = cast(CT_TcPr, element("w:tcPr/(w:tcW,w:vAlign{w:val=center})"))
        tcPr.get_or_add_tcMar()
        # -- tcMar should appear between tcW (earlier) and vAlign (later) --
        expected = xml("w:tcPr/(w:tcW,w:tcMar,w:vAlign{w:val=center})")
        assert tcPr.xml == expected

    def it_can_remove_tcMar(self):
        tcPr = cast(
            CT_TcPr, element("w:tcPr/w:tcMar/w:top{w:w=100,w:type=dxa}")
        )
        tcPr._remove_tcMar()
        assert tcPr.tcMar is None
        assert tcPr.xml == xml("w:tcPr")


class DescribeCT_Shd:
    """Unit-test suite for `docx.oxml.table.CT_Shd` objects."""

    @pytest.mark.parametrize(
        ("shd_cxml", "expected_fill"),
        [
            ("w:shd", None),
            ("w:shd{w:fill=D9E2F3}", RGBColor(0xD9, 0xE2, 0xF3)),
            ("w:shd{w:fill=auto}", "auto"),
        ],
    )
    def it_can_get_the_fill_attribute(self, shd_cxml: str, expected_fill: RGBColor | str | None):
        shd = cast(CT_Shd, element(shd_cxml))
        assert shd.fill == expected_fill

    @pytest.mark.parametrize(
        ("shd_cxml", "expected_val"),
        [
            ("w:shd", None),
            ("w:shd{w:val=clear}", WD_SHADING_PATTERN.CLEAR),
            ("w:shd{w:val=solid}", WD_SHADING_PATTERN.SOLID),
            ("w:shd{w:val=pct10}", WD_SHADING_PATTERN.PCT_10),
        ],
    )
    def it_can_get_the_val_attribute(
        self, shd_cxml: str, expected_val: WD_SHADING_PATTERN | None
    ):
        shd = cast(CT_Shd, element(shd_cxml))
        assert shd.val == expected_val


class DescribeCT_TcPr:
    """Unit-test suite for `docx.oxml.table.CT_TcPr` objects."""

    @pytest.mark.parametrize(
        ("tcPr_cxml", "expected_shd_present"),
        [
            ("w:tcPr", False),
            ("w:tcPr/w:shd{w:val=clear,w:fill=D9E2F3}", True),
        ],
    )
    def it_can_get_the_shd_child(self, tcPr_cxml: str, expected_shd_present: bool):
        tcPr = cast(CT_TcPr, element(tcPr_cxml))
        if expected_shd_present:
            assert tcPr.shd is not None
            assert isinstance(tcPr.shd, CT_Shd)
        else:
            assert tcPr.shd is None

    def it_can_add_a_shd_child(self):
        tcPr = cast(CT_TcPr, element("w:tcPr"))
        shd = tcPr.get_or_add_shd()
        assert isinstance(shd, CT_Shd)
        assert tcPr.shd is shd

    def it_inserts_shd_in_the_right_position(self):
        tcPr = cast(CT_TcPr, element("w:tcPr/(w:tcW,w:vAlign{w:val=center})"))
        shd = tcPr.get_or_add_shd()
        assert isinstance(shd, CT_Shd)
        # shd should appear between tcW and vAlign
        expected_xml = xml("w:tcPr/(w:tcW,w:shd,w:vAlign{w:val=center})")
        assert tcPr.xml == expected_xml

    @pytest.mark.parametrize(
        ("tcPr_cxml", "expected_value"),
        [
            ("w:tcPr", None),
            ("w:tcPr/w:textDirection{w:val=lrTb}", WD_TEXT_DIRECTION.LR_TB),
            ("w:tcPr/w:textDirection{w:val=tbRl}", WD_TEXT_DIRECTION.TB_RL),
            ("w:tcPr/w:textDirection{w:val=btLr}", WD_TEXT_DIRECTION.BT_LR),
            ("w:tcPr/w:textDirection{w:val=lrTbV}", WD_TEXT_DIRECTION.LR_TB_V),
            ("w:tcPr/w:textDirection{w:val=tbRlV}", WD_TEXT_DIRECTION.TB_RL_V),
            ("w:tcPr/w:textDirection{w:val=tbLrV}", WD_TEXT_DIRECTION.TB_LR_V),
        ],
    )
    def it_knows_its_text_direction(
        self, tcPr_cxml: str, expected_value: WD_TEXT_DIRECTION | None
    ):
        tcPr = cast(CT_TcPr, element(tcPr_cxml))
        assert tcPr.text_direction == expected_value

    @pytest.mark.parametrize(
        ("tcPr_cxml", "new_value", "expected_cxml"),
        [
            (
                "w:tcPr",
                WD_TEXT_DIRECTION.TB_RL,
                "w:tcPr/w:textDirection{w:val=tbRl}",
            ),
            (
                "w:tcPr/w:textDirection{w:val=tbRl}",
                WD_TEXT_DIRECTION.BT_LR,
                "w:tcPr/w:textDirection{w:val=btLr}",
            ),
            ("w:tcPr/w:textDirection{w:val=tbRl}", None, "w:tcPr"),
            ("w:tcPr", None, "w:tcPr"),
        ],
    )
    def it_can_change_its_text_direction(
        self, tcPr_cxml: str, new_value: WD_TEXT_DIRECTION | None, expected_cxml: str
    ):
        tcPr = cast(CT_TcPr, element(tcPr_cxml))
        tcPr.text_direction = new_value
        assert tcPr.xml == xml(expected_cxml)

    def it_inserts_textDirection_in_the_right_position(self):
        tcPr = cast(CT_TcPr, element("w:tcPr/(w:tcW,w:vAlign{w:val=center})"))
        tcPr.text_direction = WD_TEXT_DIRECTION.BT_LR
        # textDirection should appear between tcW and vAlign
        expected_xml = xml(
            "w:tcPr/(w:tcW,w:textDirection{w:val=btLr},w:vAlign{w:val=center})"
        )
        assert tcPr.xml == expected_xml


class DescribeCT_Row:
    @pytest.mark.parametrize(
        ("tr_cxml", "expected_cxml"),
        [
            ("w:tr", "w:tr/w:trPr"),
            ("w:tr/w:tblPrEx", "w:tr/(w:tblPrEx,w:trPr)"),
            ("w:tr/w:tc", "w:tr/(w:trPr,w:tc)"),
            ("w:tr/(w:sdt,w:del,w:tc)", "w:tr/(w:trPr,w:sdt,w:del,w:tc)"),
        ],
    )
    def it_can_add_a_trPr(self, tr_cxml: str, expected_cxml: str):
        tr = cast(CT_Row, element(tr_cxml))
        tr._add_trPr()
        assert tr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("tr_cxml", "expected_value"),
        [
            ("w:tr", True),
            ("w:tr/w:trPr", True),
            ("w:tr/w:trPr/w:cantSplit", False),
            ("w:tr/w:trPr/w:cantSplit{w:val=true}", False),
            ("w:tr/w:trPr/w:cantSplit{w:val=false}", True),
        ],
    )
    def it_knows_whether_it_allows_break_across_pages(
        self, tr_cxml: str, expected_value: bool
    ):
        tr = cast(CT_Row, element(tr_cxml))
        assert tr.allow_break_across_pages is expected_value

    @pytest.mark.parametrize(
        ("tr_cxml", "new_value", "expected_cxml"),
        [
            ("w:tr", False, "w:tr/w:trPr/w:cantSplit"),
            ("w:tr/w:trPr", False, "w:tr/w:trPr/w:cantSplit"),
            ("w:tr/w:trPr/w:cantSplit", True, "w:tr/w:trPr"),
            ("w:tr/w:trPr/w:cantSplit", None, "w:tr/w:trPr"),
            ("w:tr", True, "w:tr/w:trPr"),
        ],
    )
    def it_can_change_whether_it_allows_break_across_pages(
        self, tr_cxml: str, new_value: bool | None, expected_cxml: str
    ):
        tr = cast(CT_Row, element(tr_cxml))
        tr.allow_break_across_pages = new_value
        assert tr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("tr_cxml", "expected_value"),
        [
            ("w:tr", False),
            ("w:tr/w:trPr", False),
            ("w:tr/w:trPr/w:tblHeader", True),
            ("w:tr/w:trPr/w:tblHeader{w:val=true}", True),
            ("w:tr/w:trPr/w:tblHeader{w:val=false}", False),
        ],
    )
    def it_knows_whether_it_is_a_header_row(self, tr_cxml: str, expected_value: bool):
        tr = cast(CT_Row, element(tr_cxml))
        assert tr.is_header is expected_value

    @pytest.mark.parametrize(
        ("tr_cxml", "new_value", "expected_cxml"),
        [
            ("w:tr", True, "w:tr/w:trPr/w:tblHeader"),
            ("w:tr/w:trPr", True, "w:tr/w:trPr/w:tblHeader"),
            ("w:tr/w:trPr/w:tblHeader", False, "w:tr/w:trPr"),
            ("w:tr/w:trPr/w:tblHeader", None, "w:tr/w:trPr"),
            ("w:tr", False, "w:tr/w:trPr"),
        ],
    )
    def it_can_change_whether_it_is_a_header_row(
        self, tr_cxml: str, new_value: bool | None, expected_cxml: str
    ):
        tr = cast(CT_Row, element(tr_cxml))
        tr.is_header = new_value
        assert tr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("tr_cxml", "expected_value"),
        [
            ("w:tr", None),
            ("w:tr/w:trPr", None),
            ("w:tr/w:trPr/w:trHeight", None),
            ("w:tr/w:trPr/w:trHeight{w:val=0}", 0),
            ("w:tr/w:trPr/w:trHeight{w:val=1440}", 914400),
        ],
    )
    def it_knows_its_trHeight_val(self, tr_cxml: str, expected_value: int | None):
        tr = cast(CT_Row, element(tr_cxml))
        assert tr.trHeight_val == expected_value

    @pytest.mark.parametrize(
        ("tr_cxml", "new_value", "expected_cxml"),
        [
            ("w:tr", Inches(1), "w:tr/w:trPr/w:trHeight{w:val=1440}"),
            ("w:tr/w:trPr", Inches(1), "w:tr/w:trPr/w:trHeight{w:val=1440}"),
            ("w:tr/w:trPr/w:trHeight", Inches(1), "w:tr/w:trPr/w:trHeight{w:val=1440}"),
            (
                "w:tr/w:trPr/w:trHeight{w:val=1440}",
                Inches(2),
                "w:tr/w:trPr/w:trHeight{w:val=2880}",
            ),
            ("w:tr/w:trPr/w:trHeight{w:val=2880}", None, "w:tr/w:trPr/w:trHeight"),
            ("w:tr", None, "w:tr/w:trPr"),
            ("w:tr/w:trPr", None, "w:tr/w:trPr"),
        ],
    )
    def it_can_change_its_trHeight_val(
        self, tr_cxml: str, new_value: Length | None, expected_cxml: str
    ):
        tr = cast(CT_Row, element(tr_cxml))
        tr.trHeight_val = new_value
        assert tr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("tr_cxml", "expected_value"),
        [
            ("w:tr", None),
            ("w:tr/w:trPr", None),
            ("w:tr/w:trPr/w:trHeight", None),
            ("w:tr/w:trPr/w:trHeight{w:hRule=auto}", WD_ROW_HEIGHT_RULE.AUTO),
            (
                "w:tr/w:trPr/w:trHeight{w:val=1440, w:hRule=atLeast}",
                WD_ROW_HEIGHT_RULE.AT_LEAST,
            ),
            (
                "w:tr/w:trPr/w:trHeight{w:val=2880, w:hRule=exact}",
                WD_ROW_HEIGHT_RULE.EXACTLY,
            ),
        ],
    )
    def it_knows_its_trHeight_hRule(
        self, tr_cxml: str, expected_value: WD_ROW_HEIGHT_RULE | None
    ):
        tr = cast(CT_Row, element(tr_cxml))
        assert tr.trHeight_hRule == expected_value

    @pytest.mark.parametrize(
        ("tr_cxml", "new_value", "expected_cxml"),
        [
            ("w:tr", WD_ROW_HEIGHT_RULE.AUTO, "w:tr/w:trPr/w:trHeight{w:hRule=auto}"),
            (
                "w:tr/w:trPr",
                WD_ROW_HEIGHT_RULE.AT_LEAST,
                "w:tr/w:trPr/w:trHeight{w:hRule=atLeast}",
            ),
            (
                "w:tr/w:trPr/w:trHeight",
                WD_ROW_HEIGHT_RULE.EXACTLY,
                "w:tr/w:trPr/w:trHeight{w:hRule=exact}",
            ),
            (
                "w:tr/w:trPr/w:trHeight{w:val=1440, w:hRule=exact}",
                WD_ROW_HEIGHT_RULE.AUTO,
                "w:tr/w:trPr/w:trHeight{w:val=1440, w:hRule=auto}",
            ),
            (
                "w:tr/w:trPr/w:trHeight{w:val=1440, w:hRule=auto}",
                None,
                "w:tr/w:trPr/w:trHeight{w:val=1440}",
            ),
            ("w:tr", None, "w:tr/w:trPr"),
            ("w:tr/w:trPr", None, "w:tr/w:trPr"),
        ],
    )
    def it_can_change_its_trHeight_hRule(
        self,
        tr_cxml: str,
        new_value: WD_ROW_HEIGHT_RULE | None,
        expected_cxml: str,
    ):
        tr = cast(CT_Row, element(tr_cxml))
        tr.trHeight_hRule = new_value
        assert tr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("snippet_idx", "row_idx", "col_idx"),
        [
            # -- grid_offset beyond last tc (snippet 0 is 3x3 uniform) --
            (0, 0, 3),
            # -- negative grid_offset is out of range regardless of spans --
            (0, 0, -1),
        ],
    )
    def it_raises_on_tc_at_grid_col(self, snippet_idx: int, row_idx: int, col_idx: int):
        tr = cast(CT_Tbl, parse_xml(snippet_seq("tbl-cells")[snippet_idx])).tr_lst[row_idx]
        with pytest.raises(ValueError, match=f"no `tc` element at grid_offset={col_idx}"):
            tr.tc_at_grid_offset(col_idx)

    @pytest.mark.parametrize(
        # -- regression for upstream#1458: tc_at_grid_offset must match by
        # -- range, not by exact starting offset. A horizontally-spanning
        # -- w:tc (gridSpan > 1) "covers" every grid column within its span,
        # -- and w:gridBefore pushes the first tc's starting offset rightward.
        ("tr_cxml", "grid_offset", "expected_tc_idx"),
        [
            # -- gridSpan=2 covers both grid_offset 0 and 1 --
            ("w:tr/(w:tc/w:tcPr/w:gridSpan{w:val=2},w:tc)", 0, 0),
            ("w:tr/(w:tc/w:tcPr/w:gridSpan{w:val=2},w:tc)", 1, 0),
            ("w:tr/(w:tc/w:tcPr/w:gridSpan{w:val=2},w:tc)", 2, 1),
            # -- gridBefore=2 means first tc starts at grid_offset 2 --
            ("w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc,w:tc)", 2, 0),
            ("w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc,w:tc)", 3, 1),
            # -- gridBefore plus a spanned cell: tc_0 covers 2 and 3 --
            (
                "w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc/w:tcPr/w:gridSpan{w:val=2},w:tc)",
                2,
                0,
            ),
            (
                "w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc/w:tcPr/w:gridSpan{w:val=2},w:tc)",
                3,
                0,
            ),
            (
                "w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc/w:tcPr/w:gridSpan{w:val=2},w:tc)",
                4,
                1,
            ),
        ],
    )
    def it_matches_tc_at_grid_offset_by_range(
        self, tr_cxml: str, grid_offset: int, expected_tc_idx: int
    ):
        tr = cast(CT_Row, element(tr_cxml))

        tc = tr.tc_at_grid_offset(grid_offset)

        assert tc is tr.tc_lst[expected_tc_idx]

    @pytest.mark.parametrize(
        # -- offsets inside the w:gridBefore run have no covering tc --
        ("tr_cxml", "grid_offset"),
        [
            ("w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc,w:tc)", 0),
            ("w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc,w:tc)", 1),
            ("w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc,w:tc)", 4),
        ],
    )
    def it_raises_on_tc_at_grid_offset_in_gridBefore_or_beyond(
        self, tr_cxml: str, grid_offset: int
    ):
        tr = cast(CT_Row, element(tr_cxml))
        with pytest.raises(ValueError, match=f"no `tc` element at grid_offset={grid_offset}"):
            tr.tc_at_grid_offset(grid_offset)


class DescribeCT_Tc:
    """Unit-test suite for `docx.oxml.table.CT_Tc` objects."""

    @pytest.mark.parametrize(
        ("tr_cxml", "tc_idx", "expected_value"),
        [
            ("w:tr/(w:tc/w:p,w:tc/w:p)", 0, 0),
            ("w:tr/(w:tc/w:p,w:tc/w:p)", 1, 1),
            ("w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc/w:p,w:tc/w:p)", 0, 2),
            ("w:tr/(w:trPr/w:gridBefore{w:val=2},w:tc/w:p,w:tc/w:p)", 1, 3),
            ("w:tr/(w:trPr/w:gridBefore{w:val=4},w:tc/w:p,w:tc/w:p,w:tc/w:p,w:tc/w:p)", 2, 6),
        ],
    )
    def it_knows_its_grid_offset(self, tr_cxml: str, tc_idx: int, expected_value: int):
        tr = cast(CT_Row, element(tr_cxml))
        tc = tr.tc_lst[tc_idx]

        assert tc.grid_offset == expected_value

    def it_can_merge_to_another_tc(
        self, tr_: Mock, _span_dimensions_: Mock, _tbl_: Mock, _grow_to_: Mock, top_tc_: Mock
    ):
        top_tr_ = tr_
        tc, other_tc = cast(CT_Tc, element("w:tc")), cast(CT_Tc, element("w:tc"))
        top, left, height, width = 0, 1, 2, 3
        _span_dimensions_.return_value = top, left, height, width
        _tbl_.return_value.tr_lst = [tr_]
        tr_.tc_at_grid_offset.return_value = top_tc_

        merged_tc = tc.merge(other_tc)

        _span_dimensions_.assert_called_once_with(tc, other_tc)
        top_tr_.tc_at_grid_offset.assert_called_once_with(left)
        top_tc_._grow_to.assert_called_once_with(width, height)
        assert merged_tc is top_tc_

    @pytest.mark.parametrize(
        ("snippet_idx", "row", "col", "attr_name", "expected_value"),
        [
            (0, 0, 0, "top", 0),
            (2, 0, 1, "top", 0),
            (2, 1, 1, "top", 0),
            (4, 2, 1, "top", 1),
            (0, 0, 0, "left", 0),
            (1, 0, 1, "left", 2),
            (3, 1, 0, "left", 0),
            (3, 1, 1, "left", 2),
            (0, 0, 0, "bottom", 1),
            (1, 0, 0, "bottom", 1),
            (2, 0, 1, "bottom", 2),
            (4, 1, 1, "bottom", 3),
            (0, 0, 0, "right", 1),
            (1, 0, 0, "right", 2),
            (4, 2, 1, "right", 3),
        ],
    )
    def it_knows_its_extents_to_help(
        self, snippet_idx: int, row: int, col: int, attr_name: str, expected_value: int
    ):
        tbl = self._snippet_tbl(snippet_idx)
        tc = tbl.tr_lst[row].tc_lst[col]

        extent = getattr(tc, attr_name)

        assert extent == expected_value

    @pytest.mark.parametrize(
        ("snippet_idx", "row", "col", "row_2", "col_2", "expected_value"),
        [
            (0, 0, 0, 0, 1, (0, 0, 1, 2)),
            (0, 0, 1, 2, 1, (0, 1, 3, 1)),
            (0, 2, 2, 1, 1, (1, 1, 2, 2)),
            (0, 1, 2, 1, 0, (1, 0, 1, 3)),
            (1, 0, 0, 1, 1, (0, 0, 2, 2)),
            (1, 0, 1, 0, 0, (0, 0, 1, 3)),
            (2, 0, 1, 2, 1, (0, 1, 3, 1)),
            (2, 0, 1, 1, 0, (0, 0, 2, 2)),
            (2, 1, 2, 0, 1, (0, 1, 2, 2)),
            (4, 0, 1, 0, 0, (0, 0, 1, 3)),
        ],
    )
    def it_calculates_the_dimensions_of_a_span_to_help(
        self,
        snippet_idx: int,
        row: int,
        col: int,
        row_2: int,
        col_2: int,
        expected_value: tuple[int, int, int, int],
    ):
        tbl = self._snippet_tbl(snippet_idx)
        tc = tbl.tr_lst[row].tc_lst[col]
        other_tc = tbl.tr_lst[row_2].tc_lst[col_2]

        dimensions = tc._span_dimensions(other_tc)

        assert dimensions == expected_value

    @pytest.mark.parametrize(
        ("snippet_idx", "row", "col", "row_2", "col_2"),
        [
            (1, 0, 0, 1, 0),  # inverted-L horz
            (1, 1, 0, 0, 0),  # same in opposite order
            (2, 0, 2, 0, 1),  # inverted-L vert
            (5, 0, 1, 1, 0),  # tee-shape horz bar
            (5, 1, 0, 2, 1),  # same, opposite side
            (6, 1, 0, 0, 1),  # tee-shape vert bar
            (6, 0, 1, 1, 2),  # same, opposite side
        ],
    )
    def it_raises_on_invalid_span(
        self, snippet_idx: int, row: int, col: int, row_2: int, col_2: int
    ):
        tbl = self._snippet_tbl(snippet_idx)
        tc = tbl.tr_lst[row].tc_lst[col]
        other_tc = tbl.tr_lst[row_2].tc_lst[col_2]

        with pytest.raises(InvalidSpanError):
            tc._span_dimensions(other_tc)

    @pytest.mark.parametrize(
        ("snippet_idx", "row", "col", "width", "height"),
        [
            (0, 0, 0, 2, 1),
            (0, 0, 1, 1, 2),
            (0, 1, 1, 2, 2),
            (1, 0, 0, 2, 2),
            (2, 0, 0, 2, 2),
            (2, 1, 2, 1, 2),
        ],
    )
    def it_can_grow_itself_to_help_merge(
        self, snippet_idx: int, row: int, col: int, width: int, height: int, _span_to_width_: Mock
    ):
        tbl = self._snippet_tbl(snippet_idx)
        tc = tbl.tr_lst[row].tc_lst[col]
        start = 0 if height == 1 else 1
        end = start + height

        tc._grow_to(width, height, None)

        assert (
            _span_to_width_.call_args_list
            == [
                call(width, tc, None),
                call(width, tc, "restart"),
                call(width, tc, "continue"),
                call(width, tc, "continue"),
            ][start:end]
        )

    def it_can_extend_its_horz_span_to_help_merge(
        self, top_tc_: Mock, grid_span_: Mock, _move_content_to_: Mock, _swallow_next_tc_: Mock
    ):
        grid_span_.side_effect = [1, 3, 4]
        grid_width, vMerge = 4, "continue"
        tc = cast(CT_Tc, element("w:tc"))

        tc._span_to_width(grid_width, top_tc_, vMerge)

        _move_content_to_.assert_called_once_with(tc, top_tc_)
        assert _swallow_next_tc_.call_args_list == [
            call(tc, grid_width, top_tc_),
            call(tc, grid_width, top_tc_),
        ]
        assert tc.vMerge == vMerge

    def it_knows_its_inner_content_block_item_elements(self):
        tc = cast(CT_Tc, element("w:tc/(w:p,w:tbl,w:p)"))
        assert [type(e) for e in tc.inner_content_elements] == [CT_P, CT_Tbl, CT_P]

    @pytest.mark.parametrize(
        ("tr_cxml", "tc_idx", "grid_width", "expected_cxml"),
        [
            (
                "w:tr/(w:tc/w:p,w:tc/w:p)",
                0,
                2,
                "w:tr/(w:tc/(w:tcPr/w:gridSpan{w:val=2},w:p))",
            ),
            (
                "w:tr/(w:tc/w:p,w:tc/w:p,w:tc/w:p)",
                1,
                2,
                "w:tr/(w:tc/w:p,w:tc/(w:tcPr/w:gridSpan{w:val=2},w:p))",
            ),
            (
                'w:tr/(w:tc/w:p/w:r/w:t"a",w:tc/w:p/w:r/w:t"b")',
                0,
                2,
                'w:tr/(w:tc/(w:tcPr/w:gridSpan{w:val=2},w:p/w:r/w:t"a",w:p/w:r/w:t"b"))',
            ),
            (
                "w:tr/(w:tc/(w:tcPr/w:gridSpan{w:val=2},w:p),w:tc/w:p)",
                0,
                3,
                "w:tr/(w:tc/(w:tcPr/w:gridSpan{w:val=3},w:p))",
            ),
            (
                "w:tr/(w:tc/w:p,w:tc/(w:tcPr/w:gridSpan{w:val=2},w:p))",
                0,
                3,
                "w:tr/(w:tc/(w:tcPr/w:gridSpan{w:val=3},w:p))",
            ),
        ],
    )
    def it_can_swallow_the_next_tc_help_merge(
        self, tr_cxml: str, tc_idx: int, grid_width: int, expected_cxml: str
    ):
        tr = cast(CT_Row, element(tr_cxml))
        tc = top_tc = tr.tc_lst[tc_idx]

        tc._swallow_next_tc(grid_width, top_tc)

        assert tr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("tr_cxml", "tc_idx", "grid_width", "expected_cxml"),
        [
            # both cells have a width
            (
                "w:tr/(w:tc/(w:tcPr/w:tcW{w:w=1440,w:type=dxa},w:p),"
                "w:tc/(w:tcPr/w:tcW{w:w=1440,w:type=dxa},w:p))",
                0,
                2,
                "w:tr/(w:tc/(w:tcPr/(w:tcW{w:w=2880,w:type=dxa},w:gridSpan{w:val=2}),w:p))",
            ),
            # neither have a width
            (
                "w:tr/(w:tc/w:p,w:tc/w:p)",
                0,
                2,
                "w:tr/(w:tc/(w:tcPr/w:gridSpan{w:val=2},w:p))",
            ),
            # only second one has a width
            (
                "w:tr/(w:tc/w:p,w:tc/(w:tcPr/w:tcW{w:w=1440,w:type=dxa},w:p))",
                0,
                2,
                "w:tr/(w:tc/(w:tcPr/w:gridSpan{w:val=2},w:p))",
            ),
            # only first one has a width
            (
                "w:tr/(w:tc/(w:tcPr/w:tcW{w:w=1440,w:type=dxa},w:p),w:tc/w:p)",
                0,
                2,
                "w:tr/(w:tc/(w:tcPr/(w:tcW{w:w=1440,w:type=dxa},w:gridSpan{w:val=2}),w:p))",
            ),
        ],
    )
    def it_adds_cell_widths_on_swallow(
        self, tr_cxml: str, tc_idx: int, grid_width: int, expected_cxml: str
    ):
        tr = cast(CT_Row, element(tr_cxml))
        tc = top_tc = tr.tc_lst[tc_idx]
        tc._swallow_next_tc(grid_width, top_tc)
        assert tr.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("tr_cxml", "tc_idx", "grid_width"),
        [
            ("w:tr/w:tc/w:p", 0, 2),
            ("w:tr/(w:tc/w:p,w:tc/(w:tcPr/w:gridSpan{w:val=2},w:p))", 0, 2),
        ],
    )
    def it_raises_on_invalid_swallow(self, tr_cxml: str, tc_idx: int, grid_width: int):
        tr = cast(CT_Row, element(tr_cxml))
        tc = top_tc = tr.tc_lst[tc_idx]

        with pytest.raises(InvalidSpanError):
            tc._swallow_next_tc(grid_width, top_tc)

    @pytest.mark.parametrize(
        ("tc_cxml", "tc_2_cxml", "expected_tc_cxml", "expected_tc_2_cxml"),
        [
            ("w:tc/w:p", "w:tc/w:p", "w:tc/w:p", "w:tc/w:p"),
            ("w:tc/w:p", "w:tc/w:p/w:r", "w:tc/w:p", "w:tc/w:p/w:r"),
            ("w:tc/w:p/w:r", "w:tc/w:p", "w:tc/w:p", "w:tc/w:p/w:r"),
            ("w:tc/(w:p/w:r,w:sdt)", "w:tc/w:p", "w:tc/w:p", "w:tc/(w:p/w:r,w:sdt)"),
            (
                "w:tc/(w:p/w:r,w:sdt)",
                "w:tc/(w:tbl,w:p)",
                "w:tc/w:p",
                "w:tc/(w:tbl,w:p/w:r,w:sdt)",
            ),
        ],
    )
    def it_can_move_its_content_to_help_merge(
        self, tc_cxml: str, tc_2_cxml: str, expected_tc_cxml: str, expected_tc_2_cxml: str
    ):
        tc, tc_2 = cast(CT_Tc, element(tc_cxml)), cast(CT_Tc, element(tc_2_cxml))

        tc._move_content_to(tc_2)

        assert tc.xml == xml(expected_tc_cxml)
        assert tc_2.xml == xml(expected_tc_2_cxml)

    @pytest.mark.parametrize(("snippet_idx", "row_idx", "col_idx"), [(0, 0, 0), (4, 0, 0)])
    def it_raises_on_tr_above(self, snippet_idx: int, row_idx: int, col_idx: int):
        tbl = cast(CT_Tbl, parse_xml(snippet_seq("tbl-cells")[snippet_idx]))
        tc = tbl.tr_lst[row_idx].tc_lst[col_idx]

        with pytest.raises(ValueError, match="no tr above topmost tr"):
            tc._tr_above

    @pytest.mark.parametrize(
        ("tc_cxml", "expected"),
        [
            # -- absent gridSpan defaults to 1 --
            ("w:tc", 1),
            ("w:tc/w:tcPr", 1),
            # -- explicit value of 1 --
            ("w:tc/w:tcPr/w:gridSpan{w:val=1}", 1),
            # -- normal span values --
            ("w:tc/w:tcPr/w:gridSpan{w:val=2}", 2),
            ("w:tc/w:tcPr/w:gridSpan{w:val=5}", 5),
            # -- malformed: gridSpan=0 is coerced to 1 (read robustness) --
            ("w:tc/w:tcPr/w:gridSpan{w:val=0}", 1),
        ],
    )
    def it_coerces_malformed_grid_span_to_one(self, tc_cxml: str, expected: int):
        tc = cast(CT_Tc, element(tc_cxml))
        assert tc.grid_span == expected

    def it_traces_bottom_through_multiple_vMerge_continuations(self):
        """Restart followed by 3 continuations should report bottom==4."""
        tbl = cast(
            CT_Tbl,
            parse_xml(
                '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:tblGrid><w:gridCol/></w:tblGrid>"
                "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p/></w:tc></w:tr>"
                "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>"
                "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>"
                "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>"
                "</w:tbl>"
            ),
        )
        root_tc = tbl.tr_lst[0].tc_lst[0]
        last_tc = tbl.tr_lst[3].tc_lst[0]

        # -- root sees the full span: 4 rows --
        assert root_tc.top == 0
        assert root_tc.bottom == 4
        # -- last continuation points back up to row 0 as its top --
        assert last_tc.top == 0
        assert last_tc.bottom == 4

    def it_handles_vMerge_chain_to_last_row(self):
        """vMerge chain that ends at the final row (no row below).

        Regression guard: ``bottom`` walks forward while there is another
        continuation; when the current row is the last, it should return
        ``_tr_idx + 1`` without trying to access a nonexistent row below.
        """
        tbl = cast(
            CT_Tbl,
            parse_xml(
                '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:tblGrid><w:gridCol/></w:tblGrid>"
                "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p/></w:tc></w:tr>"
                "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>"
                "</w:tbl>"
            ),
        )
        root_tc = tbl.tr_lst[0].tc_lst[0]
        last_tc = tbl.tr_lst[1].tc_lst[0]
        # -- root's bottom is just past the last continuation row --
        assert root_tc.bottom == 2
        # -- last continuation has no row below; bottom is its own row +1 --
        assert last_tc.bottom == 2

    def it_resolves_bottom_across_gridBefore_rows(self):
        """Regression for upstream#1458: ``cell._tc.bottom`` must not crash
        when the row directly below has ``w:gridBefore`` or gridSpan cells
        that shift the grid column of the continuation tc.
        """
        tbl = cast(
            CT_Tbl,
            parse_xml(
                '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>"
                # -- row 0 starts a vertical merge in the second grid column --
                "<w:tr>"
                "<w:tc><w:p/></w:tc>"
                "<w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p/></w:tc>"
                "</w:tr>"
                # -- row 1 starts with gridBefore=1; its single tc sits at --
                # -- grid offset 1 and continues the vMerge. --
                "<w:tr>"
                "<w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>"
                "<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>"
                "</w:tr>"
                "</w:tbl>"
            ),
        )
        top_tc = tbl.tr_lst[0].tc_lst[1]
        # -- bottom is one past the last continuation row --
        assert top_tc.bottom == 2

    def it_grows_iteratively_for_large_merges(self):
        """Regression for upstream#1208: ``_grow_to`` must not recurse per
        row, because very tall merges exceed Python's recursion limit.
        """
        import sys

        # -- a merge height larger than the default recursion limit --
        row_count = sys.getrecursionlimit() + 50
        rows_xml = "".join(
            '<w:tr><w:tc><w:p/></w:tc></w:tr>' for _ in range(row_count)
        )
        tbl = cast(
            CT_Tbl,
            parse_xml(
                '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:tblGrid><w:gridCol/></w:tblGrid>"
                + rows_xml
                + "</w:tbl>"
            ),
        )
        top_tc = tbl.tr_lst[0].tc_lst[0]

        # -- merging all rows into a single span must not raise RecursionError --
        top_tc._grow_to(1, row_count)

        # -- and the root tc's bottom now spans the full table --
        assert top_tc.bottom == row_count

    def it_merges_cells_in_a_nested_table_without_crossing_tables(self):
        """Regression for upstream#169: merging cells in a table that is
        itself nested inside another table must not leak grid-col lookups
        into the outer table's rows.
        """
        xml_str = (
            '<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>"
            "<w:tr>"
            "<w:tc><w:p/></w:tc>"
            "<w:tc>"
            # -- nested table inside the second outer cell --
            "<w:tbl>"
            "<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>"
            "<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>"
            "<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>"
            "</w:tbl>"
            "<w:p/>"
            "</w:tc>"
            "</w:tr>"
            "<w:tr>"
            "<w:tc><w:p/></w:tc>"
            "<w:tc><w:p/></w:tc>"
            "</w:tr>"
            "</w:tbl>"
        )
        outer_tbl = cast(CT_Tbl, parse_xml(xml_str))
        # -- locate the inner tbl and merge its top-left to bottom-right --
        inner_tbl = cast(CT_Tbl, outer_tbl.xpath(".//w:tc//w:tbl")[0])
        inner_top_left = inner_tbl.tr_lst[0].tc_lst[0]
        inner_bottom_right = inner_tbl.tr_lst[1].tc_lst[1]

        merged = inner_top_left.merge(inner_bottom_right)

        # -- merge returned the inner top-left, not anything from outer tbl --
        assert merged is inner_tbl.tr_lst[0].tc_lst[0]
        # -- outer table rows are untouched --
        assert len(outer_tbl.tr_lst) == 2
        assert len(outer_tbl.tr_lst[0].tc_lst) == 2
        assert len(outer_tbl.tr_lst[1].tc_lst) == 2
        # -- inner table has a 2x2 merged span on the (now-only) top cell --
        assert merged.grid_span == 2
        assert merged.bottom == 2

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def grid_span_(self, request: FixtureRequest):
        return property_mock(request, CT_Tc, "grid_span")

    @pytest.fixture
    def _grow_to_(self, request: FixtureRequest):
        return method_mock(request, CT_Tc, "_grow_to")

    @pytest.fixture
    def _move_content_to_(self, request: FixtureRequest):
        return method_mock(request, CT_Tc, "_move_content_to")

    @pytest.fixture
    def _span_dimensions_(self, request: FixtureRequest):
        return method_mock(request, CT_Tc, "_span_dimensions")

    @pytest.fixture
    def _span_to_width_(self, request: FixtureRequest):
        return method_mock(request, CT_Tc, "_span_to_width", autospec=False)

    def _snippet_tbl(self, idx: int) -> CT_Tbl:
        """A <w:tbl> element for snippet at `idx` in 'tbl-cells' snippet file."""
        return cast(CT_Tbl, parse_xml(snippet_seq("tbl-cells")[idx]))

    @pytest.fixture
    def _swallow_next_tc_(self, request: FixtureRequest):
        return method_mock(request, CT_Tc, "_swallow_next_tc")

    @pytest.fixture
    def _tbl_(self, request: FixtureRequest):
        return property_mock(request, CT_Tc, "_tbl")

    @pytest.fixture
    def top_tc_(self, request: FixtureRequest):
        return instance_mock(request, CT_Tc)

    @pytest.fixture
    def tr_(self, request: FixtureRequest):
        return instance_mock(request, CT_Row)
