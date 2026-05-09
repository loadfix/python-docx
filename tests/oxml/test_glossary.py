# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.glossary` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.glossary import (
    CT_DocPart,
    CT_DocPartBehaviors,
    CT_DocPartBody,
    CT_DocPartCategory,
    CT_DocPartPr,
    CT_DocPartTypes,
    CT_DocParts,
    CT_GlossaryDocument,
)

from ..unitutil.cxml import element


class DescribeCT_GlossaryDocument:
    """Unit-test suite for `docx.oxml.glossary.CT_GlossaryDocument`."""

    def it_exposes_its_docParts_child(self):
        glossary = cast(CT_GlossaryDocument, element("w:glossaryDocument/w:docParts"))
        assert glossary.docParts is not None
        assert isinstance(glossary.docParts, CT_DocParts)

    def it_returns_None_for_an_absent_docParts(self):
        glossary = cast(CT_GlossaryDocument, element("w:glossaryDocument"))
        assert glossary.docParts is None

    def it_yields_an_empty_docPart_lst_when_docParts_is_absent(self):
        glossary = cast(CT_GlossaryDocument, element("w:glossaryDocument"))
        assert glossary.docPart_lst == []

    def it_yields_an_empty_docPart_lst_when_docParts_is_empty(self):
        glossary = cast(CT_GlossaryDocument, element("w:glossaryDocument/w:docParts"))
        assert glossary.docPart_lst == []

    def it_yields_each_docPart_in_order(self):
        glossary = cast(
            CT_GlossaryDocument,
            element("w:glossaryDocument/w:docParts/(w:docPart,w:docPart,w:docPart)"),
        )
        assert len(glossary.docPart_lst) == 3
        assert all(isinstance(dp, CT_DocPart) for dp in glossary.docPart_lst)


class DescribeCT_DocParts:
    """Unit-test suite for `docx.oxml.glossary.CT_DocParts`."""

    def it_exposes_its_docPart_children_in_order(self):
        docParts = cast(CT_DocParts, element("w:docParts/(w:docPart,w:docPart)"))
        assert len(docParts.docPart_lst) == 2


class DescribeCT_DocPart:
    """Unit-test suite for `docx.oxml.glossary.CT_DocPart`."""

    def it_exposes_its_docPartPr_child(self):
        doc_part = cast(CT_DocPart, element("w:docPart/w:docPartPr"))
        assert doc_part.docPartPr is not None
        assert isinstance(doc_part.docPartPr, CT_DocPartPr)

    def it_returns_None_for_absent_docPartPr(self):
        doc_part = cast(CT_DocPart, element("w:docPart"))
        assert doc_part.docPartPr is None

    def it_exposes_its_docPartBody_child(self):
        doc_part = cast(CT_DocPart, element("w:docPart/w:docPartBody"))
        assert doc_part.docPartBody is not None
        assert isinstance(doc_part.docPartBody, CT_DocPartBody)

    def it_returns_None_for_absent_docPartBody(self):
        doc_part = cast(CT_DocPart, element("w:docPart"))
        assert doc_part.docPartBody is None


class DescribeCT_DocPartPr:
    """Unit-test suite for `docx.oxml.glossary.CT_DocPartPr`."""

    def it_exposes_the_name_w_val(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr/w:name{w:val=MyBlock}"))
        assert pr.name_val == "MyBlock"

    def it_returns_None_when_w_name_is_absent(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr"))
        assert pr.name_val is None

    def it_returns_None_when_w_name_has_no_val_attribute(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr/w:name"))
        assert pr.name_val is None

    def it_exposes_the_description_w_val(self):
        pr = cast(
            CT_DocPartPr, element("w:docPartPr/w:description{w:val=a description}")
        )
        assert pr.description_val == "a description"

    def it_returns_None_when_description_is_absent(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr"))
        assert pr.description_val is None

    def it_exposes_the_guid_w_val(self):
        pr = cast(
            CT_DocPartPr,
            element("w:docPartPr/w:guid{w:val=abc-123-def}"),
        )
        assert pr.guid_val == "abc-123-def"

    def it_returns_None_when_guid_is_absent(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr"))
        assert pr.guid_val is None

    def it_exposes_its_category_child(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr/w:category"))
        assert pr.category is not None
        assert isinstance(pr.category, CT_DocPartCategory)

    def it_returns_None_when_category_is_absent(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr"))
        assert pr.category is None


class DescribeCT_DocPartCategory:
    """Unit-test suite for `docx.oxml.glossary.CT_DocPartCategory`."""

    def it_exposes_the_name_w_val(self):
        cat = cast(CT_DocPartCategory, element("w:category/w:name{w:val=General}"))
        assert cat.name_val == "General"

    def it_returns_None_when_name_is_absent(self):
        cat = cast(CT_DocPartCategory, element("w:category"))
        assert cat.name_val is None

    def it_exposes_the_gallery_w_val(self):
        cat = cast(
            CT_DocPartCategory, element("w:category/w:gallery{w:val=quickParts}")
        )
        assert cat.gallery is not None
        assert cat.gallery.val == "quickParts"

    def it_returns_None_when_gallery_is_absent(self):
        cat = cast(CT_DocPartCategory, element("w:category"))
        assert cat.gallery is None


class DescribeCT_DocPartBody:
    """Unit-test suite for `docx.oxml.glossary.CT_DocPartBody`."""

    def it_exposes_its_paragraphs(self):
        body = cast(CT_DocPartBody, element("w:docPartBody/(w:p,w:p)"))
        assert len(body.p_lst) == 2

    def it_exposes_its_tables(self):
        body = cast(CT_DocPartBody, element("w:docPartBody/w:tbl"))
        assert len(body.tbl_lst) == 1

    def it_orders_inner_content_elements_in_document_order(self):
        body = cast(CT_DocPartBody, element("w:docPartBody/(w:p,w:tbl,w:p)"))
        tags = [el.tag.rsplit("}", 1)[-1] for el in body.inner_content_elements]
        assert tags == ["p", "tbl", "p"]


class DescribeCT_DocPartTypes:
    """Unit-test suite for `docx.oxml.glossary.CT_DocPartTypes`."""

    def it_reads_each_w_type_val_attribute(self):
        types = cast(
            CT_DocPartTypes,
            element("w:types/(w:type{w:val=autoTxt},w:type{w:val=toolbar})"),
        )
        assert types.values == ["autoTxt", "toolbar"]

    def it_returns_an_empty_list_when_no_children(self):
        types = cast(CT_DocPartTypes, element("w:types"))
        assert types.values == []

    def it_can_append_a_w_type_child(self):
        types = cast(CT_DocPartTypes, element("w:types"))
        types.add_type("autoTxt")
        assert types.values == ["autoTxt"]


class DescribeCT_DocPartBehaviors:
    """Unit-test suite for `docx.oxml.glossary.CT_DocPartBehaviors`."""

    def it_reads_each_w_behavior_val_attribute(self):
        behaviors = cast(
            CT_DocPartBehaviors,
            element("w:behaviors/(w:behavior{w:val=content},w:behavior{w:val=p})"),
        )
        assert behaviors.values == ["content", "p"]

    def it_returns_an_empty_list_when_no_children(self):
        behaviors = cast(CT_DocPartBehaviors, element("w:behaviors"))
        assert behaviors.values == []

    def it_can_append_a_w_behavior_child(self):
        behaviors = cast(CT_DocPartBehaviors, element("w:behaviors"))
        behaviors.add_behavior("content")
        assert behaviors.values == ["content"]


class DescribeCT_DocPartPrWriteHelpers:
    """Unit-test suite for mutating helpers on `CT_DocPartPr`."""

    def it_creates_a_w_name_child_with_the_given_val(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr"))
        pr.set_name("Foo")
        assert pr.name_val == "Foo"

    def it_updates_an_existing_w_name_val(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr/w:name{w:val=Old}"))
        pr.set_name("New")
        assert pr.name_val == "New"

    def it_creates_a_w_description_child(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr"))
        pr.set_description("A block")
        assert pr.description_val == "A block"

    def it_creates_a_w_guid_child(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr"))
        pr.set_guid("{abc}")
        assert pr.guid_val == "{abc}"

    def it_inserts_children_in_schema_order(self):
        pr = cast(CT_DocPartPr, element("w:docPartPr"))
        # -- deliberately out-of-order calls --
        pr.set_guid("{x}")
        pr.set_name("N")
        pr.set_description("D")
        tags = [c.tag.rsplit("}", 1)[-1] for c in pr]
        assert tags == ["name", "description", "guid"]


class DescribeCT_DocPartCategoryWriteHelpers:
    """Unit-test suite for mutating helpers on `CT_DocPartCategory`."""

    def it_sets_the_category_name(self):
        cat = cast(CT_DocPartCategory, element("w:category"))
        cat.set_name("General")
        assert cat.name_val == "General"

    def it_sets_the_gallery_val(self):
        cat = cast(CT_DocPartCategory, element("w:category"))
        cat.set_gallery("quickParts")
        assert cat.gallery is not None and cat.gallery.val == "quickParts"

    def it_preserves_schema_order_name_before_gallery(self):
        cat = cast(CT_DocPartCategory, element("w:category"))
        # -- set gallery first, then name — name must come before gallery --
        cat.set_gallery("quickParts")
        cat.set_name("General")
        tags = [c.tag.rsplit("}", 1)[-1] for c in cat]
        assert tags == ["name", "gallery"]
