# pyright: reportPrivateUsage=false
# pyright: reportUnknownMemberType=false

"""Unit test suite for the docx.document module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.comments import Comment, Comments
from docx.custom_properties import CustomProperties
from docx.document import Document, _Body
from docx.enum.section import WD_SECTION
from docx.font_table import FontTable
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.coreprops import CoreProperties
from docx.oxml.document import CT_Body, CT_Document
from docx.parts.document import DocumentPart
from docx.section import Section, Sections
from docx.settings import Settings
from docx.shape import InlineShape, InlineShapes
from docx.shared import Length
from docx.styles.styles import Styles
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .unitutil.cxml import element, xml
from .unitutil.mock import (
    FixtureRequest,
    Mock,
    class_mock,
    instance_mock,
    method_mock,
    property_mock,
)


class DescribeDocument:
    """Unit-test suite for `docx.document.Document`."""

    def it_can_add_a_comment(
        self,
        document_part_: Mock,
        comments_prop_: Mock,
        comments_: Mock,
        comment_: Mock,
        run_mark_comment_range_: Mock,
    ):
        comment_.comment_id = 42
        comments_.add_comment.return_value = comment_
        comments_prop_.return_value = comments_
        document = Document(cast(CT_Document, element("w:document/w:body/w:p/w:r")), document_part_)
        run = document.paragraphs[0].runs[0]

        comment = document.add_comment(run, "Comment text.")

        comments_.add_comment.assert_called_once_with("Comment text.", "", "")
        run_mark_comment_range_.assert_called_once_with(run, run, 42)
        assert comment is comment_

    @pytest.mark.parametrize(
        ("level", "style"), [(0, "Title"), (1, "Heading 1"), (2, "Heading 2"), (9, "Heading 9")]
    )
    def it_can_add_a_heading(
        self, level: int, style: str, document: Document, add_paragraph_: Mock, paragraph_: Mock
    ):
        add_paragraph_.return_value = paragraph_

        paragraph = document.add_heading("Spam vs. Bacon", level)

        add_paragraph_.assert_called_once_with(document, "Spam vs. Bacon", style)
        assert paragraph is paragraph_

    def it_raises_on_heading_level_out_of_range(self, document: Document):
        with pytest.raises(ValueError, match="level must be in range 0-9, got -1"):
            document.add_heading(level=-1)
        with pytest.raises(ValueError, match="level must be in range 0-9, got 10"):
            document.add_heading(level=10)

    def it_can_add_a_page_break(
        self, document: Document, add_paragraph_: Mock, paragraph_: Mock
    ):
        add_paragraph_.return_value = paragraph_
        paragraph_.add_page_break.return_value = paragraph_

        paragraph = document.add_page_break()

        add_paragraph_.assert_called_once_with(document)
        paragraph_.add_page_break.assert_called_once_with()
        assert paragraph is paragraph_

    @pytest.mark.parametrize(
        ("text", "style"), [("", None), ("", "Heading 1"), ("foo\rbar", "Body Text")]
    )
    def it_can_add_a_paragraph(
        self,
        text: str,
        style: str | None,
        document: Document,
        body_: Mock,
        body_prop_: Mock,
        paragraph_: Mock,
    ):
        body_prop_.return_value = body_
        body_.add_paragraph.return_value = paragraph_

        paragraph = document.add_paragraph(text, style)

        body_.add_paragraph.assert_called_once_with(text, style)
        assert paragraph is paragraph_

    def it_can_add_a_picture(
        self, document: Document, add_paragraph_: Mock, run_: Mock, picture_: Mock
    ):
        path, width, height = "foobar.png", 100, 200
        add_paragraph_.return_value.add_run.return_value = run_
        run_.add_picture.return_value = picture_

        picture = document.add_picture(path, width, height)

        run_.add_picture.assert_called_once_with(path, width, height)
        assert picture is picture_

    @pytest.mark.parametrize(
        ("sentinel_cxml", "start_type", "new_sentinel_cxml"),
        [
            ("w:sectPr", WD_SECTION.EVEN_PAGE, "w:sectPr/w:type{w:val=evenPage}"),
            (
                "w:sectPr/w:type{w:val=evenPage}",
                WD_SECTION.ODD_PAGE,
                "w:sectPr/w:type{w:val=oddPage}",
            ),
            ("w:sectPr/w:type{w:val=oddPage}", WD_SECTION.NEW_PAGE, "w:sectPr"),
        ],
    )
    def it_can_add_a_section(
        self,
        sentinel_cxml: str,
        start_type: WD_SECTION,
        new_sentinel_cxml: str,
        Section_: Mock,
        section_: Mock,
        document_part_: Mock,
    ):
        Section_.return_value = section_
        document = Document(
            cast(CT_Document, element("w:document/w:body/(w:p,%s)" % sentinel_cxml)),
            document_part_,
        )

        section = document.add_section(start_type)

        assert document.element.xml == xml(
            "w:document/w:body/(w:p,w:p/w:pPr/%s,%s)" % (sentinel_cxml, new_sentinel_cxml)
        )
        sectPr = document.element.xpath("w:body/w:sectPr")[0]
        Section_.assert_called_once_with(sectPr, document_part_)
        assert section is section_

    def it_can_add_a_table(
        self,
        document: Document,
        _block_width_prop_: Mock,
        body_prop_: Mock,
        body_: Mock,
        table_: Mock,
    ):
        rows, cols, style = 4, 2, "Light Shading Accent 1"
        body_prop_.return_value = body_
        body_.add_table.return_value = table_
        _block_width_prop_.return_value = width = 42

        table = document.add_table(rows, cols, style)

        body_.add_table.assert_called_once_with(rows, cols, width)
        assert table == table_
        assert table.style == style

    def it_can_save_the_document_to_a_file(self, document_part_: Mock):
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        document.save("foobar.docx")

        document_part_.save.assert_called_once_with("foobar.docx")

    def it_knows_when_the_document_has_macros(self, document_part_: Mock):
        document = Document(cast(CT_Document, element("w:document")), document_part_)
        document_part_.part_related_by.return_value = Mock()

        assert document.has_macros is True
        document_part_.part_related_by.assert_called_once_with(RT.VBA_PROJECT)

    def it_knows_when_the_document_has_no_macros(self, document_part_: Mock):
        document = Document(cast(CT_Document, element("w:document")), document_part_)
        document_part_.part_related_by.side_effect = KeyError

        assert document.has_macros is False

    def it_provides_access_to_the_comments(self, document_part_: Mock, comments_: Mock):
        document_part_.comments = comments_
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.comments is comments_

    def it_provides_access_to_its_core_properties(
        self, document_part_: Mock, core_properties_: Mock
    ):
        document_part_.core_properties = core_properties_
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        core_properties = document.core_properties

        assert core_properties is core_properties_

    def it_provides_access_to_its_custom_properties(
        self, document_part_: Mock, custom_properties_: Mock
    ):
        document_part_.custom_properties = custom_properties_
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        custom_properties = document.custom_properties

        assert custom_properties is custom_properties_

    def it_provides_access_to_its_font_table(
        self, document_part_: Mock, font_table_: Mock
    ):
        document_part_.font_table = font_table_
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.font_table is font_table_

    def and_font_table_is_None_when_the_document_has_no_font_table_part(
        self, document_part_: Mock
    ):
        document_part_.font_table = None
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.font_table is None

    def it_provides_access_to_its_inline_shapes(self, document_part_: Mock, inline_shapes_: Mock):
        document_part_.inline_shapes = inline_shapes_
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.inline_shapes is inline_shapes_

    def it_can_iterate_the_inner_content_of_the_document(
        self, body_prop_: Mock, body_: Mock, document_part_: Mock
    ):
        body_prop_.return_value = body_
        body_.iter_inner_content.return_value = iter((1, 2, 3))
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert list(document.iter_inner_content()) == [1, 2, 3]

    def it_provides_access_to_its_paragraphs(
        self, document: Document, body_prop_: Mock, body_: Mock, paragraphs_: Mock
    ):
        body_prop_.return_value = body_
        body_.paragraphs = paragraphs_
        paragraphs = document.paragraphs
        assert paragraphs is paragraphs_

    def it_provides_access_to_its_sections(
        self, document_part_: Mock, Sections_: Mock, sections_: Mock
    ):
        document_elm = cast(CT_Document, element("w:document"))
        Sections_.return_value = sections_
        document = Document(document_elm, document_part_)

        sections = document.sections

        Sections_.assert_called_once_with(document_elm, document_part_)
        assert sections is sections_

    def it_provides_access_to_its_settings(self, document_part_: Mock, settings_: Mock):
        document_part_.settings = settings_
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.settings is settings_

    def it_delegates_footnote_properties_to_settings(
        self, document_part_: Mock, settings_: Mock
    ):
        document_part_.settings = settings_
        settings_.footnote_properties = "fp-sentinel"
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.footnote_properties == "fp-sentinel"

    def it_delegates_add_footnote_properties_to_settings(
        self, document_part_: Mock, settings_: Mock
    ):
        document_part_.settings = settings_
        settings_.add_footnote_properties.return_value = "added-fp"
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        result = document.add_footnote_properties()

        settings_.add_footnote_properties.assert_called_once_with()
        assert result == "added-fp"

    def it_delegates_endnote_properties_to_settings(
        self, document_part_: Mock, settings_: Mock
    ):
        document_part_.settings = settings_
        settings_.endnote_properties = "ep-sentinel"
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.endnote_properties == "ep-sentinel"

    def it_delegates_add_endnote_properties_to_settings(
        self, document_part_: Mock, settings_: Mock
    ):
        document_part_.settings = settings_
        settings_.add_endnote_properties.return_value = "added-ep"
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        result = document.add_endnote_properties()

        settings_.add_endnote_properties.assert_called_once_with()
        assert result == "added-ep"

    def it_provides_access_to_its_styles(self, document_part_: Mock, styles_: Mock):
        document_part_.styles = styles_
        document = Document(cast(CT_Document, element("w:document")), document_part_)

        assert document.styles is styles_

    def it_provides_access_to_its_tables(
        self, document: Document, body_prop_: Mock, body_: Mock, tables_: Mock
    ):
        body_prop_.return_value = body_
        body_.tables = tables_

        assert document.tables is tables_

    def it_provides_access_to_the_document_part(self, document_part_: Mock):
        document = Document(cast(CT_Document, element("w:document")), document_part_)
        assert document.part is document_part_

    def it_provides_access_to_the_document_body(
        self, _Body_: Mock, body_: Mock, document_part_: Mock
    ):
        _Body_.return_value = body_
        document_elm = cast(CT_Document, element("w:document/w:body"))
        body_elm = document_elm[0]
        document = Document(document_elm, document_part_)

        body = document._body

        _Body_.assert_called_once_with(body_elm, document)
        assert body is body_

    def it_can_accept_all_tracked_changes(self, document_part_: Mock):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/("
                'w:p/(w:r/w:t"keep ",w:ins{w:id=1,w:author=A}/w:r/w:t"added"),'
                'w:p/(w:del{w:id=2,w:author=B}/w:r/w:delText"removed",w:r/w:t" end")'
                ")"
            ),
        )
        document = Document(document_elm, document_part_)

        count = document.accept_all_changes()

        assert count == 2
        assert document_elm.xpath(".//w:ins") == []
        assert document_elm.xpath(".//w:del") == []
        # Insertion flattened, deletion gone
        paragraphs = document_elm.xpath(".//w:p")
        assert "".join(t.text for t in paragraphs[0].xpath(".//w:t")) == "keep added"
        assert "".join(t.text for t in paragraphs[1].xpath(".//w:t")) == " end"

    def it_can_reject_all_tracked_changes(self, document_part_: Mock):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/("
                'w:p/(w:r/w:t"keep ",w:ins{w:id=1,w:author=A}/w:r/w:t"added"),'
                'w:p/(w:del{w:id=2,w:author=B}/w:r/w:delText"removed",w:r/w:t" end")'
                ")"
            ),
        )
        document = Document(document_elm, document_part_)

        count = document.reject_all_changes()

        assert count == 2
        assert document_elm.xpath(".//w:ins") == []
        assert document_elm.xpath(".//w:del") == []
        assert document_elm.xpath(".//w:delText") == []
        # Insertion removed, deletion restored
        paragraphs = document_elm.xpath(".//w:p")
        assert "".join(t.text for t in paragraphs[0].xpath(".//w:t")) == "keep "
        assert "".join(t.text for t in paragraphs[1].xpath(".//w:t")) == "removed end"

    def it_resolves_formatting_changes_on_accept(self, document_part_: Mock):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/w:pPr/("
                "w:jc{w:val=center},"
                "w:pPrChange{w:id=5,w:author=C}/w:pPr/w:jc{w:val=left}"
                ")"
            ),
        )
        document = Document(document_elm, document_part_)

        count = document.accept_all_changes()

        assert count == 1
        # accept: pPrChange removed, current formatting (center) preserved
        assert document_elm.xpath(".//w:pPrChange") == []
        jc = document_elm.xpath(".//w:pPr/w:jc")
        assert len(jc) == 1
        assert jc[0].get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
        ) == "center"

    def it_resolves_formatting_changes_on_reject(self, document_part_: Mock):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/w:p/w:pPr/("
                "w:jc{w:val=center},"
                "w:pPrChange{w:id=5,w:author=C}/w:pPr/w:jc{w:val=left}"
                ")"
            ),
        )
        document = Document(document_elm, document_part_)

        count = document.reject_all_changes()

        assert count == 1
        # reject: pPrChange gone, prior formatting (left) restored
        assert document_elm.xpath(".//w:pPrChange") == []
        jc = document_elm.xpath(".//w:pPr/w:jc")
        assert len(jc) == 1
        assert jc[0].get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
        ) == "left"

    def it_renders_revision_marks_text_joined_by_blank_lines(self, document_part_: Mock):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/("
                'w:p/(w:r/w:t"first ",w:ins{w:id=1,w:author=A}/w:r/w:t"added"),'
                'w:p/(w:del{w:id=2,w:author=B}/w:r/w:delText"removed",w:r/w:t" end"),'
                'w:p/w:r/w:t"plain"'
                ")"
            ),
        )
        document = Document(document_elm, document_part_)

        rendered = document.revision_marks_text()

        assert rendered == "first [+added+]\n\n[-removed-] end\n\nplain"

    def it_supports_custom_markers_in_document_revision_marks(
        self, document_part_: Mock
    ):
        document_elm = cast(
            CT_Document,
            element(
                "w:document/w:body/"
                'w:p/(w:ins{w:id=1,w:author=A}/w:r/w:t"x",'
                'w:del{w:id=2,w:author=B}/w:r/w:delText"y")'
            ),
        )
        document = Document(document_elm, document_part_)

        rendered = document.revision_marks_text(
            open_ins="<+", close_ins="+>", open_del="<-", close_del="->"
        )

        assert rendered == "<+x+><-y->"

    def it_determines_block_width_to_help(
        self, document: Document, sections_prop_: Mock, section_: Mock
    ):
        sections_prop_.return_value = [None, section_]
        section_.page_width = 6000
        section_.left_margin = 1500
        section_.right_margin = 1000

        width = document._block_width

        assert isinstance(width, Length)
        assert width == 3500

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def add_paragraph_(self, request: FixtureRequest):
        return method_mock(request, Document, "add_paragraph")

    @pytest.fixture
    def _Body_(self, request: FixtureRequest):
        return class_mock(request, "docx.document._Body")

    @pytest.fixture
    def body_(self, request: FixtureRequest):
        return instance_mock(request, _Body)

    @pytest.fixture
    def _block_width_prop_(self, request: FixtureRequest):
        return property_mock(request, Document, "_block_width")

    @pytest.fixture
    def body_prop_(self, request: FixtureRequest):
        return property_mock(request, Document, "_body")

    @pytest.fixture
    def comment_(self, request: FixtureRequest):
        return instance_mock(request, Comment)

    @pytest.fixture
    def comments_(self, request: FixtureRequest):
        return instance_mock(request, Comments)

    @pytest.fixture
    def comments_prop_(self, request: FixtureRequest):
        return property_mock(request, Document, "comments")

    @pytest.fixture
    def core_properties_(self, request: FixtureRequest):
        return instance_mock(request, CoreProperties)

    @pytest.fixture
    def custom_properties_(self, request: FixtureRequest):
        return instance_mock(request, CustomProperties)

    @pytest.fixture
    def document(self, document_part_: Mock) -> Document:
        document_elm = cast(CT_Document, element("w:document"))
        return Document(document_elm, document_part_)

    @pytest.fixture
    def document_part_(self, request: FixtureRequest):
        return instance_mock(request, DocumentPart)

    @pytest.fixture
    def font_table_(self, request: FixtureRequest):
        return instance_mock(request, FontTable)

    @pytest.fixture
    def inline_shapes_(self, request: FixtureRequest):
        return instance_mock(request, InlineShapes)

    @pytest.fixture
    def paragraph_(self, request: FixtureRequest):
        return instance_mock(request, Paragraph)

    @pytest.fixture
    def paragraphs_(self, request: FixtureRequest):
        return instance_mock(request, list)

    @pytest.fixture
    def picture_(self, request: FixtureRequest):
        return instance_mock(request, InlineShape)

    @pytest.fixture
    def run_(self, request: FixtureRequest):
        return instance_mock(request, Run)

    @pytest.fixture
    def run_mark_comment_range_(self, request: FixtureRequest):
        return method_mock(request, Run, "mark_comment_range")

    @pytest.fixture
    def Section_(self, request: FixtureRequest):
        return class_mock(request, "docx.document.Section")

    @pytest.fixture
    def section_(self, request: FixtureRequest):
        return instance_mock(request, Section)

    @pytest.fixture
    def Sections_(self, request: FixtureRequest):
        return class_mock(request, "docx.document.Sections")

    @pytest.fixture
    def sections_(self, request: FixtureRequest):
        return instance_mock(request, Sections)

    @pytest.fixture
    def sections_prop_(self, request: FixtureRequest):
        return property_mock(request, Document, "sections")

    @pytest.fixture
    def settings_(self, request: FixtureRequest):
        return instance_mock(request, Settings)

    @pytest.fixture
    def styles_(self, request: FixtureRequest):
        return instance_mock(request, Styles)

    @pytest.fixture
    def table_(self, request: FixtureRequest):
        return instance_mock(request, Table)

    @pytest.fixture
    def tables_(self, request: FixtureRequest):
        return instance_mock(request, list)


class Describe_Body:
    """Unit-test suite for `docx.document._Body`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_cxml"),
        [
            ("w:body", "w:body"),
            ("w:body/w:p", "w:body"),
            ("w:body/w:sectPr", "w:body/w:sectPr"),
            ("w:body/(w:p, w:sectPr)", "w:body/w:sectPr"),
        ],
    )
    def it_can_clear_itself_of_all_content_it_holds(
        self, cxml: str, expected_cxml: str, document_: Mock
    ):
        body = _Body(cast(CT_Body, element(cxml)), document_)

        _body = body.clear_content()

        assert body._body.xml == xml(expected_cxml)
        assert _body is body

    def it_can_add_a_block_level_content_control(self, document_: Mock):
        from docx.content_controls import ContentControl, ContentControlType

        body = _Body(cast(CT_Body, element("w:body")), document_)

        cc = body.add_content_control(
            ContentControlType.RICH_TEXT, tag="T", title="Title"
        )

        assert isinstance(cc, ContentControl)
        assert cc.tag == "T"
        assert cc.title == "Title"
        # -- the sdt was appended to the body --
        assert len(body._body.xpath("./w:sdt")) == 1

    def it_inserts_block_level_sdt_before_trailing_sectPr(self, document_: Mock):
        from docx.content_controls import ContentControlType

        body = _Body(cast(CT_Body, element("w:body/(w:p,w:sectPr)")), document_)

        body.add_content_control(ContentControlType.RICH_TEXT, tag="T")

        # -- verify order: w:p, w:sdt, w:sectPr --
        children = [c.tag.rsplit("}", 1)[-1] for c in body._body]
        assert children == ["p", "sdt", "sectPr"]

    def it_lists_block_level_content_controls(self, document_: Mock):
        from docx.content_controls import ContentControlType

        body = _Body(cast(CT_Body, element("w:body")), document_)
        body.add_content_control(ContentControlType.RICH_TEXT, tag="A")
        body.add_content_control(ContentControlType.RICH_TEXT, tag="B")

        assert [cc.tag for cc in body.content_controls] == ["A", "B"]

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def document_(self, request: FixtureRequest):
        return instance_mock(request, Document)
