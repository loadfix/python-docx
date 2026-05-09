# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.content_controls` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.content_controls import (
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
from docx.oxml.ns import qn

import pytest

from ..unitutil.cxml import element


class DescribeCT_Sdt:
    """Unit-test suite for `docx.oxml.content_controls.CT_Sdt`."""

    def it_reads_its_tag_val_from_sdtPr(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:tag{w:val=MyTag}"),
        )
        assert sdt.tag_val == "MyTag"

    def it_returns_None_for_tag_val_when_not_present(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        assert sdt.tag_val is None

    def it_can_set_tag_val(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.tag_val = "MyTag"
        assert sdt.tag_val == "MyTag"
        assert sdt.sdtPr is not None
        tag_elm = sdt.sdtPr.find(qn("w:tag"))
        assert tag_elm is not None
        assert tag_elm.get(qn("w:val")) == "MyTag"

    def it_reads_its_alias_val_from_sdtPr(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:alias{w:val=MyTitle}"),
        )
        assert sdt.alias_val == "MyTitle"

    def it_returns_None_for_alias_val_when_not_present(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        assert sdt.alias_val is None

    def it_can_set_alias_val(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.alias_val = "Hello"
        assert sdt.alias_val == "Hello"

    def it_can_remove_tag_val_by_assigning_None(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:tag{w:val=x}"),
        )
        sdt.tag_val = None
        assert sdt.tag_val is None

    def it_reads_its_id(self):
        sdt = cast(
            CT_Sdt,
            element("w:sdt/w:sdtPr/w:id{w:val=42}"),
        )
        assert sdt.sdt_id == 42

    def it_can_set_its_id(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.sdt_id = 99
        assert sdt.sdt_id == 99

    def it_detects_no_type_marker_as_rich_text_default(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        assert sdt.type_marker_tag() is None

    def it_detects_w_text_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:text"))
        assert sdt.type_marker_tag() == "w:text"

    def it_detects_w_date_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:date"))
        assert sdt.type_marker_tag() == "w:date"

    def it_reads_checked_value_when_present(self):
        xml = (
            '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            '<w:sdtPr><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>'
            "</w:sdt>"
        )
        from docx.oxml.parser import parse_xml

        sdt = cast(CT_Sdt, parse_xml(xml))
        assert sdt.checked is True

    def it_returns_None_for_checked_when_no_checkbox(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr"))
        assert sdt.checked is None

    def it_can_set_checked(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.checked = True
        assert sdt.checked is True
        sdt.checked = False
        assert sdt.checked is False

    def it_can_set_a_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.set_type_marker("w:text")
        assert sdt.type_marker_tag() == "w:text"

    def it_replaces_existing_type_marker_on_set(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:text"))
        sdt.set_type_marker("w:date")
        assert sdt.type_marker_tag() == "w:date"
        assert sdt.sdtPr is not None
        assert sdt.sdtPr.find(qn("w:text")) is None


class DescribeCT_SdtContent:
    """Unit-test suite for `docx.oxml.content_controls.CT_SdtContent`."""

    def it_concatenates_paragraph_text(self):
        sdtContent = cast(
            CT_SdtContent,
            element('w:sdtContent/(w:p/w:r/w:t"Hello")'),
        )
        assert sdtContent.text == "Hello"

    def it_concatenates_run_text(self):
        sdtContent = cast(
            CT_SdtContent,
            element('w:sdtContent/(w:r/w:t"Hi")'),
        )
        assert sdtContent.text == "Hi"

    def it_returns_empty_string_when_no_children(self):
        sdtContent = cast(CT_SdtContent, element("w:sdtContent"))
        assert sdtContent.text == ""


class DescribeCT_SdtListItem:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtListItem``."""

    def it_reads_displayText_and_value_attributes(self):
        item = cast(
            CT_SdtListItem,
            element("w:listItem{w:displayText=Red,w:value=R}"),
        )
        assert item.displayText == "Red"
        assert item.value == "R"

    def it_returns_None_when_attributes_are_absent(self):
        item = cast(CT_SdtListItem, element("w:listItem"))
        assert item.displayText is None
        assert item.value is None

    def it_can_set_displayText_and_value(self):
        item = cast(CT_SdtListItem, element("w:listItem"))
        item.displayText = "Blue"
        item.value = "B"
        assert item.get(qn("w:displayText")) == "Blue"
        assert item.get(qn("w:value")) == "B"


class DescribeCT_SdtDateMappingType:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtDateMappingType``."""

    def it_reads_its_val(self):
        sm = cast(
            CT_SdtDateMappingType,
            element("w:storeMappedDataAs{w:val=dateTime}"),
        )
        assert sm.val == "dateTime"

    def it_returns_None_when_val_is_absent(self):
        sm = cast(CT_SdtDateMappingType, element("w:storeMappedDataAs"))
        assert sm.val is None

    def it_can_set_its_val(self):
        sm = cast(CT_SdtDateMappingType, element("w:storeMappedDataAs"))
        sm.val = "date"
        assert sm.get(qn("w:val")) == "date"


class DescribeCT_SdtDate:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtDate``."""

    def it_reads_its_fullDate_attribute(self):
        date = cast(
            CT_SdtDate,
            element("w:date{w:fullDate=2026-05-09T00:00:00Z}"),
        )
        assert date.fullDate == "2026-05-09T00:00:00Z"

    def it_returns_None_for_missing_fullDate(self):
        date = cast(CT_SdtDate, element("w:date"))
        assert date.fullDate is None

    def it_can_add_child_elements_in_schema_order(self):
        date = cast(CT_SdtDate, element("w:date"))
        date.get_or_add_dateFormat()
        date.get_or_add_lid()
        date.get_or_add_storeMappedDataAs()
        date.get_or_add_calendar()
        tags = [child.tag for child in date]
        assert tags == [
            qn("w:dateFormat"),
            qn("w:lid"),
            qn("w:storeMappedDataAs"),
            qn("w:calendar"),
        ]

    def it_exposes_storeMappedDataAs_child_as_CT_SdtDateMappingType(self):
        date = cast(
            CT_SdtDate,
            element("w:date/w:storeMappedDataAs{w:val=text}"),
        )
        sm = date.storeMappedDataAs
        assert sm is not None
        assert isinstance(sm, CT_SdtDateMappingType)
        assert sm.val == "text"


class DescribeCT_SdtComboBox:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtComboBox``."""

    def it_reads_its_lastValue_attribute(self):
        combo = cast(
            CT_SdtComboBox,
            element("w:comboBox{w:lastValue=Custom}"),
        )
        assert combo.lastValue == "Custom"

    def it_returns_an_empty_listItem_lst_when_none_present(self):
        combo = cast(CT_SdtComboBox, element("w:comboBox"))
        assert combo.listItem_lst == []

    def it_can_add_a_listItem(self):
        combo = cast(CT_SdtComboBox, element("w:comboBox"))
        item = combo.add_listItem()
        item.displayText = "Red"
        item.value = "R"
        assert len(combo.listItem_lst) == 1
        assert combo.listItem_lst[0].displayText == "Red"

    def it_iterates_multiple_listItem_children(self):
        combo = cast(
            CT_SdtComboBox,
            element(
                "w:comboBox/("
                "w:listItem{w:displayText=Red,w:value=R},"
                "w:listItem{w:displayText=Blue,w:value=B}"
                ")"
            ),
        )
        assert [(i.displayText, i.value) for i in combo.listItem_lst] == [
            ("Red", "R"),
            ("Blue", "B"),
        ]


class DescribeCT_SdtDocPart:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtDocPart``."""

    def it_can_add_docPartGallery_and_category_in_order(self):
        dp = cast(CT_SdtDocPart, element("w:docPartObj"))
        dp.get_or_add_docPartGallery()
        dp.get_or_add_docPartCategory()
        dp.get_or_add_docPartUnique()
        tags = [child.tag for child in dp]
        assert tags == [
            qn("w:docPartGallery"),
            qn("w:docPartCategory"),
            qn("w:docPartUnique"),
        ]

    def it_returns_None_for_absent_children(self):
        dp = cast(CT_SdtDocPart, element("w:docPartObj"))
        assert dp.docPartGallery is None
        assert dp.docPartCategory is None
        assert dp.docPartUnique is None


class DescribeCT_SdtDropDownList:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtDropDownList``."""

    def it_reads_its_lastValue_attribute(self):
        dd = cast(
            CT_SdtDropDownList,
            element("w:dropDownList{w:lastValue=Two}"),
        )
        assert dd.lastValue == "Two"

    def it_can_add_a_listItem(self):
        dd = cast(CT_SdtDropDownList, element("w:dropDownList"))
        item = dd.add_listItem()
        item.displayText = "One"
        item.value = "1"
        assert len(dd.listItem_lst) == 1

    def it_treats_listItem_children_as_CT_SdtListItem(self):
        dd = cast(
            CT_SdtDropDownList,
            element("w:dropDownList/w:listItem{w:displayText=A,w:value=a}"),
        )
        items = dd.listItem_lst
        assert len(items) == 1
        assert isinstance(items[0], CT_SdtListItem)
        assert items[0].displayText == "A"


class DescribeCT_SdtText:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtText``."""

    def it_reads_multiLine_as_True_when_set_to_1(self):
        txt = cast(CT_SdtText, element("w:text{w:multiLine=1}"))
        assert txt.multiLine is True

    def it_reads_multiLine_as_False_when_set_to_0(self):
        txt = cast(CT_SdtText, element("w:text{w:multiLine=0}"))
        assert txt.multiLine is False

    def it_returns_None_when_multiLine_is_absent(self):
        txt = cast(CT_SdtText, element("w:text"))
        assert txt.multiLine is None


class DescribeCT_SdtEndPr:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtEndPr``."""

    def it_exposes_an_empty_rPr_lst_when_none_present(self):
        endPr = cast(CT_SdtEndPr, element("w:sdtEndPr"))
        assert endPr.rPr_lst == []

    def it_can_add_an_rPr_child(self):
        endPr = cast(CT_SdtEndPr, element("w:sdtEndPr"))
        endPr.add_rPr()
        assert len(endPr.rPr_lst) == 1


def _sdt_content_as(cls, cxml: str):
    """Parse `cxml` and re-parse under a class lookup that binds ``w:sdtContent``
    to `cls`.

    The SDT content-container types (``CT_SdtContentBlock`` et al.) all share
    the ``w:sdtContent`` tag with the generic :class:`CT_SdtContent`, which
    is the class registered for that tag in ``docx.oxml.__init__``.  These
    tests want to exercise the typed-only accessors on the specific
    container class; a dedicated class-lookup isolates that.
    """
    from lxml import etree

    parsed = element(cxml)
    xml = etree.tostring(parsed)

    lookup = etree.ElementNamespaceClassLookup()
    ns = lookup.get_namespace("http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    ns["sdtContent"] = cls

    parser = etree.XMLParser()
    parser.set_element_class_lookup(lookup)
    return etree.fromstring(xml, parser)


class DescribeCT_SdtContentBlock:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtContentBlock``."""

    def it_exposes_its_paragraph_children(self):
        block = cast(
            CT_SdtContentBlock,
            _sdt_content_as(CT_SdtContentBlock, "w:sdtContent/(w:p,w:p)"),
        )
        assert len(block.p_lst) == 2


class DescribeCT_SdtContentCell:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtContentCell``."""

    def it_exposes_its_tc_children(self):
        cell = cast(
            CT_SdtContentCell,
            _sdt_content_as(CT_SdtContentCell, "w:sdtContent/(w:tc,w:tc,w:tc)"),
        )
        assert len(cell.tc_lst) == 3


class DescribeCT_SdtContentRow:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtContentRow``."""

    def it_exposes_its_tr_children(self):
        row = cast(
            CT_SdtContentRow,
            _sdt_content_as(CT_SdtContentRow, "w:sdtContent/(w:tr,w:tr)"),
        )
        assert len(row.tr_lst) == 2


class DescribeCT_SdtContentRun:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtContentRun``."""

    def it_exposes_its_run_children(self):
        run_content = cast(
            CT_SdtContentRun,
            _sdt_content_as(CT_SdtContentRun, "w:sdtContent/(w:r,w:r)"),
        )
        assert len(run_content.r_lst) == 2


class DescribeCT_SdtContentRunRuby:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtContentRunRuby``."""

    def it_exposes_its_run_children(self):
        ruby_content = cast(
            CT_SdtContentRunRuby,
            _sdt_content_as(CT_SdtContentRunRuby, "w:sdtContent/w:r"),
        )
        assert len(ruby_content.r_lst) == 1


class DescribeCT_SdtRepeatedSection:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtRepeatedSection``."""

    def it_is_registered_for_the_w15_tag(self):
        rs = element("w15:repeatingSection")
        assert isinstance(rs, CT_SdtRepeatedSection)

    def it_reads_its_sectionTitle_attribute(self):
        rs = cast(
            CT_SdtRepeatedSection,
            element("w15:repeatingSection{w15:sectionTitle=Rows}"),
        )
        assert rs.sectionTitle == "Rows"

    def it_reads_doNotAllowInsertDeleteSection_as_True(self):
        rs = cast(
            CT_SdtRepeatedSection,
            element("w15:repeatingSection{w15:doNotAllowInsertDeleteSection=1}"),
        )
        assert rs.doNotAllowInsertDeleteSection is True

    def it_returns_None_for_absent_attributes(self):
        rs = cast(CT_SdtRepeatedSection, element("w15:repeatingSection"))
        assert rs.sectionTitle is None
        assert rs.doNotAllowInsertDeleteSection is None


class DescribeCT_Lock:
    """Unit-test suite for ``docx.oxml.content_controls.CT_Lock``."""

    def it_is_registered_for_the_w_lock_tag(self):
        lock = element("w:lock")
        assert isinstance(lock, CT_Lock)

    def it_reads_its_val_attribute(self):
        lock = cast(CT_Lock, element("w:lock{w:val=sdtContentLocked}"))
        assert lock.val == "sdtContentLocked"

    def it_returns_None_when_val_is_absent(self):
        lock = cast(CT_Lock, element("w:lock"))
        assert lock.val is None


class DescribeCT_SdtPr_lock_val:
    """Unit-test suite for ``CT_SdtPr.lock_val``."""

    def it_reads_the_lock_val(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr/w:lock{w:val=contentLocked}"))
        assert sdtPr.lock_val == "contentLocked"

    def it_returns_None_when_no_lock_child_present(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        assert sdtPr.lock_val is None

    def it_can_set_the_lock_val_creating_a_lock_child(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        sdtPr.lock_val = "sdtLocked"
        assert sdtPr.lock_val == "sdtLocked"
        assert sdtPr.find(qn("w:lock")) is not None

    def it_can_remove_the_lock_by_assigning_None(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr/w:lock{w:val=unlocked}"))
        sdtPr.lock_val = None
        assert sdtPr.lock_val is None
        assert sdtPr.find(qn("w:lock")) is None

    def it_rejects_an_unknown_lock_val(self):
        sdtPr = cast(CT_SdtPr, element("w:sdtPr"))
        with pytest.raises(ValueError):
            sdtPr.lock_val = "bogus"

    def it_round_trips_via_CT_Sdt_lock_val(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.lock_val = "sdtContentLocked"
        assert sdt.lock_val == "sdtContentLocked"
        sdt.lock_val = None
        assert sdt.lock_val is None


class DescribeCT_Sdt_type_markers_extended:
    """Unit-test suite for extended type-marker support on ``CT_Sdt``."""

    def it_detects_w_docPartObj_as_a_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:docPartObj"))
        assert sdt.type_marker_tag() == "w:docPartObj"

    def it_detects_w_docPartList_as_a_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w:docPartList"))
        assert sdt.type_marker_tag() == "w:docPartList"

    def it_detects_w15_repeatingSection_as_a_type_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt/w:sdtPr/w15:repeatingSection"))
        assert sdt.type_marker_tag() == "w15:repeatingSection"

    def it_can_set_a_w15_repeatingSection_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.set_type_marker("w15:repeatingSection")
        assert sdt.type_marker_tag() == "w15:repeatingSection"

    def it_can_set_a_w_docPartObj_marker(self):
        sdt = cast(CT_Sdt, element("w:sdt"))
        sdt.set_type_marker("w:docPartObj")
        assert sdt.type_marker_tag() == "w:docPartObj"


class DescribeCT_SdtRepeatedSectionItem:
    """Unit-test suite for ``docx.oxml.content_controls.CT_SdtRepeatedSectionItem``."""

    def it_is_registered_for_the_w15_tag(self):
        rsi = element("w15:repeatingSectionItem")
        assert isinstance(rsi, CT_SdtRepeatedSectionItem)

    def it_round_trips_inside_a_sdtPr(self):
        from docx.oxml.parser import parse_xml

        xml = (
            '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
            "<w:sdtPr><w15:repeatingSectionItem/></w:sdtPr>"
            "</w:sdt>"
        )
        sdt = cast(CT_Sdt, parse_xml(xml))
        marker = sdt.sdtPr.find(qn("w15:repeatingSectionItem"))
        assert marker is not None
        assert isinstance(marker, CT_SdtRepeatedSectionItem)
