# pyright: reportPrivateUsage=false

"""Unit test suite for the `docx.permissions` module and related integrations."""

from __future__ import annotations

from typing import cast

from docx.oxml.document import CT_Body, CT_Document
from docx.oxml.ns import qn
from docx.oxml.permissions import CT_PermStart
from docx.permissions import PermissionRange
from docx.text.paragraph import Paragraph

from .unitutil.cxml import element


class DescribePermissionRange:
    """Unit-test suite for `docx.permissions.PermissionRange`."""

    def it_knows_its_id(self):
        body = cast(CT_Body, element("w:body"))
        permStart = cast(CT_PermStart, element("w:permStart{w:id=7}"))
        pr = PermissionRange(permStart, body)
        assert pr.id == 7

    def it_knows_its_edit_group(self):
        body = cast(CT_Body, element("w:body"))
        permStart = cast(
            CT_PermStart, element("w:permStart{w:id=7,w:edGrp=everyone}")
        )
        pr = PermissionRange(permStart, body)
        assert pr.edit_group == "everyone"

    def it_knows_its_user(self):
        body = cast(CT_Body, element("w:body"))
        permStart = cast(CT_PermStart, element("w:permStart{w:id=7,w:ed=alice}"))
        pr = PermissionRange(permStart, body)
        assert pr.user == "alice"

    def it_reports_None_for_absent_optional_attributes(self):
        body = cast(CT_Body, element("w:body"))
        permStart = cast(CT_PermStart, element("w:permStart{w:id=0}"))
        pr = PermissionRange(permStart, body)
        assert pr.edit_group is None
        assert pr.user is None
        assert pr.displaced_by_custom_xml is None

    def it_knows_its_displaced_by_custom_xml(self):
        body = cast(CT_Body, element("w:body"))
        permStart = cast(
            CT_PermStart,
            element("w:permStart{w:id=0,w:displacedByCustomXml=next}"),
        )
        pr = PermissionRange(permStart, body)
        assert pr.displaced_by_custom_xml == "next"

    def it_can_delete_itself(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:permStart{w:id=0,w:edGrp=everyone}"
                ",w:r/w:t\"hello\""
                ",w:permEnd{w:id=0})"
            ),
        )
        assert len(body.xpath(".//w:permStart")) == 1
        assert len(body.xpath(".//w:permEnd")) == 1

        permStart = cast(CT_PermStart, body.xpath(".//w:permStart")[0])
        pr = PermissionRange(permStart, body)
        pr.delete()

        assert len(body.xpath(".//w:permStart")) == 0
        assert len(body.xpath(".//w:permEnd")) == 0

    def it_can_delete_a_cross_paragraph_range(self):
        body = cast(
            CT_Body,
            element(
                "w:body/(w:p/(w:permStart{w:id=1,w:edGrp=everyone},w:r/w:t\"hi\")"
                ",w:p/(w:r/w:t\"there\",w:permEnd{w:id=1}))"
            ),
        )
        permStart = cast(CT_PermStart, body.xpath(".//w:permStart")[0])
        pr = PermissionRange(permStart, body)

        pr.delete()

        assert len(body.xpath(".//w:permStart")) == 0
        assert len(body.xpath(".//w:permEnd")) == 0


class DescribeParagraph_permission_ranges:
    """Unit-test suite for `Paragraph.permission_ranges` and `.add_permission_range()`."""

    def it_returns_empty_for_a_paragraph_with_no_ranges(self):
        body = cast(CT_Body, element('w:body/w:p/w:r/w:t"hello"'))
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        assert para.permission_ranges == []

    def it_yields_a_range_proxy_for_each_permStart(self):
        body = cast(
            CT_Body,
            element(
                "w:body/w:p/(w:permStart{w:id=1,w:edGrp=everyone}"
                ",w:r/w:t\"hi\""
                ",w:permEnd{w:id=1})"
            ),
        )
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        ranges = para.permission_ranges

        assert len(ranges) == 1
        pr = ranges[0]
        assert isinstance(pr, PermissionRange)
        assert pr.id == 1
        assert pr.edit_group == "everyone"
        assert pr.user is None

    def it_can_add_a_permission_range_wrapping_whole_paragraph(self):
        body = cast(CT_Body, element('w:body/w:p/w:r/w:t"hello"'))
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        pr = para.add_permission_range("range1", edit_group="everyone")

        assert isinstance(pr, PermissionRange)
        assert pr.edit_group == "everyone"
        assert pr.user is None
        assert pr.id == 0
        # -- permStart is first child (no pPr), permEnd is last --
        children = list(p_elm)
        assert children[0].tag == qn("w:permStart")
        assert children[-1].tag == qn("w:permEnd")

    def it_places_permStart_after_pPr_when_present(self):
        body = cast(CT_Body, element('w:body/w:p/(w:pPr,w:r/w:t"hello")'))
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        para.add_permission_range(edit_group="everyone")

        children = list(p_elm)
        assert children[0].tag == qn("w:pPr")
        assert children[1].tag == qn("w:permStart")
        assert children[-1].tag == qn("w:permEnd")

    def it_allocates_unique_ids(self):
        body = cast(CT_Body, element('w:body/w:p/w:r/w:t"hello"'))
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        pr1 = para.add_permission_range(edit_group="everyone")
        pr2 = para.add_permission_range(user="alice")

        assert pr1.id == 0
        assert pr2.id == 1
        assert pr2.user == "alice"

    def it_can_delete_an_added_range(self):
        body = cast(CT_Body, element('w:body/w:p/w:r/w:t"hello"'))
        p_elm = body.p_lst[0]
        para = Paragraph(p_elm, None)  # type: ignore[arg-type]

        pr = para.add_permission_range(edit_group="everyone")
        assert len(para.permission_ranges) == 1

        pr.delete()

        assert len(para.permission_ranges) == 0
        assert len(body.xpath(".//w:permEnd")) == 0


class DescribeDocument_permission_ranges:
    """Unit-test suite for `Document.permission_ranges`."""

    def it_provides_document_level_iteration_across_paragraphs(self):
        from docx.document import Document

        doc_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                "(w:p/(w:permStart{w:id=0,w:edGrp=everyone},w:permEnd{w:id=0})"
                ",w:p/(w:permStart{w:id=1,w:ed=alice},w:permEnd{w:id=1}))"
            ),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        ranges = doc.permission_ranges

        assert len(ranges) == 2
        assert ranges[0].id == 0
        assert ranges[0].edit_group == "everyone"
        assert ranges[1].id == 1
        assert ranges[1].user == "alice"

    def it_returns_empty_list_when_no_ranges(self):
        from docx.document import Document

        doc_elm = cast(
            CT_Document,
            element('w:document/w:body/w:p/w:r/w:t"hi"'),
        )
        doc = Document(doc_elm, None)  # type: ignore[arg-type]

        assert doc.permission_ranges == []
