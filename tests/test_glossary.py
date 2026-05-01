"""Unit-test suite for the `docx.glossary` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.text import WD_BUILDING_BLOCK_GALLERY
from docx.glossary import BuildingBlock, BuildingBlockCategory, Glossary
from docx.oxml.glossary import CT_DocPart, CT_GlossaryDocument

from .unitutil.cxml import element


# -- a compact building block used in a few tests below --
_SAMPLE_BLOCK = (
    "w:docPart/("
    "w:docPartPr/("
    "w:name{w:val=MyBlock},"
    "w:category/(w:name{w:val=General},w:gallery{w:val=quickParts}),"
    "w:description{w:val=sample description},"
    "w:guid{w:val=abc-123-def}"
    "),"
    "w:docPartBody/(w:p,w:p,w:tbl)"
    ")"
)

_SAMPLE_GLOSSARY = (
    "w:glossaryDocument/w:docParts/("
    "w:docPart/(w:docPartPr/w:name{w:val=First}),"
    "w:docPart/(w:docPartPr/w:name{w:val=Second}),"
    "w:docPart/w:docPartPr/w:name{w:val=Third}"
    ")"
)


class DescribeGlossary:
    """Unit-test suite for `docx.glossary.Glossary`."""

    def it_exposes_its_building_blocks(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_SAMPLE_GLOSSARY)))
        blocks = glossary.building_blocks
        assert len(blocks) == 3
        assert all(isinstance(b, BuildingBlock) for b in blocks)
        assert [b.name for b in blocks] == ["First", "Second", "Third"]

    def it_is_iterable_over_building_blocks(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_SAMPLE_GLOSSARY)))
        assert [b.name for b in glossary] == ["First", "Second", "Third"]

    def it_supports_len(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_SAMPLE_GLOSSARY)))
        assert len(glossary) == 3

    def it_returns_zero_len_for_an_empty_docParts(self):
        glossary = Glossary(
            cast(CT_GlossaryDocument, element("w:glossaryDocument/w:docParts"))
        )
        assert len(glossary) == 0
        assert list(glossary) == []

    def it_returns_zero_len_when_docParts_is_absent(self):
        glossary = Glossary(
            cast(CT_GlossaryDocument, element("w:glossaryDocument"))
        )
        assert len(glossary) == 0

    def it_can_look_up_a_building_block_by_name(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_SAMPLE_GLOSSARY)))
        block = glossary["Second"]
        assert isinstance(block, BuildingBlock)
        assert block.name == "Second"

    def it_raises_KeyError_for_an_unknown_name(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_SAMPLE_GLOSSARY)))
        with pytest.raises(KeyError):
            _ = glossary["NoSuchBlock"]


class DescribeBuildingBlock:
    """Unit-test suite for `docx.glossary.BuildingBlock`."""

    def it_exposes_its_name(self):
        block = BuildingBlock(cast(CT_DocPart, element(_SAMPLE_BLOCK)))
        assert block.name == "MyBlock"

    def it_returns_None_for_name_when_docPartPr_is_absent(self):
        block = BuildingBlock(cast(CT_DocPart, element("w:docPart")))
        assert block.name is None

    def it_returns_None_for_name_when_w_name_is_absent(self):
        block = BuildingBlock(
            cast(CT_DocPart, element("w:docPart/w:docPartPr"))
        )
        assert block.name is None

    def it_exposes_its_category(self):
        block = BuildingBlock(cast(CT_DocPart, element(_SAMPLE_BLOCK)))
        cat = block.category
        assert isinstance(cat, BuildingBlockCategory)
        assert cat.category_name == "General"
        assert cat.gallery == "quickParts"

    def and_category_returns_a_proxy_with_None_slots_when_absent(self):
        block = BuildingBlock(cast(CT_DocPart, element("w:docPart")))
        cat = block.category
        assert isinstance(cat, BuildingBlockCategory)
        assert cat.category_name is None
        assert cat.gallery is None

    def it_exposes_its_description(self):
        block = BuildingBlock(cast(CT_DocPart, element(_SAMPLE_BLOCK)))
        assert block.description == "sample description"

    def it_returns_None_for_description_when_absent(self):
        block = BuildingBlock(cast(CT_DocPart, element("w:docPart")))
        assert block.description is None

    def it_exposes_its_guid(self):
        block = BuildingBlock(cast(CT_DocPart, element(_SAMPLE_BLOCK)))
        assert block.guid == "abc-123-def"

    def it_returns_None_for_guid_when_absent(self):
        block = BuildingBlock(cast(CT_DocPart, element("w:docPart")))
        assert block.guid is None

    def it_exposes_its_paragraphs(self):
        block = BuildingBlock(cast(CT_DocPart, element(_SAMPLE_BLOCK)))
        paragraphs = block.paragraphs
        assert len(paragraphs) == 2

    def it_returns_empty_paragraphs_when_docPartBody_is_absent(self):
        block = BuildingBlock(cast(CT_DocPart, element("w:docPart")))
        assert block.paragraphs == []

    def it_exposes_its_tables(self):
        block = BuildingBlock(cast(CT_DocPart, element(_SAMPLE_BLOCK)))
        tables = block.tables
        assert len(tables) == 1

    def it_returns_empty_tables_when_docPartBody_is_absent(self):
        block = BuildingBlock(cast(CT_DocPart, element("w:docPart")))
        assert block.tables == []


class DescribeBuildingBlockCategory:
    """Unit-test suite for `docx.glossary.BuildingBlockCategory`."""

    def it_exposes_its_name_and_gallery(self):
        cat_elm = element(
            "w:category/(w:name{w:val=General},w:gallery{w:val=quickParts})"
        )
        cat = BuildingBlockCategory(cat_elm)  # type: ignore[arg-type]
        assert cat.category_name == "General"
        assert cat.gallery == "quickParts"

    def it_returns_None_when_name_is_absent(self):
        cat_elm = element("w:category/w:gallery{w:val=quickParts}")
        cat = BuildingBlockCategory(cat_elm)  # type: ignore[arg-type]
        assert cat.category_name is None
        assert cat.gallery == "quickParts"

    def it_returns_None_when_gallery_is_absent(self):
        cat_elm = element("w:category/w:name{w:val=General}")
        cat = BuildingBlockCategory(cat_elm)  # type: ignore[arg-type]
        assert cat.category_name == "General"
        assert cat.gallery is None

    def it_returns_None_for_every_slot_when_category_element_is_None(self):
        cat = BuildingBlockCategory(None)
        assert cat.category_name is None
        assert cat.gallery is None

    def it_round_trips_a_known_gallery_through_the_enum(self):
        cat_elm = element("w:category/w:gallery{w:val=quickParts}")
        cat = BuildingBlockCategory(cat_elm)  # type: ignore[arg-type]
        assert cat.gallery_value is WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS

    @pytest.mark.parametrize(
        ("xml_val", "expected"),
        [
            ("coverPg", WD_BUILDING_BLOCK_GALLERY.COVER_PAGES),
            ("hdrs", WD_BUILDING_BLOCK_GALLERY.HEADERS),
            ("ftrs", WD_BUILDING_BLOCK_GALLERY.FOOTERS),
            ("txtBox", WD_BUILDING_BLOCK_GALLERY.TEXT_BOXES),
            ("custom1", WD_BUILDING_BLOCK_GALLERY.CUSTOM_1),
        ],
    )
    def it_maps_common_galleries_to_enum_members(self, xml_val, expected):
        cat_elm = element(f"w:category/w:gallery{{w:val={xml_val}}}")
        cat = BuildingBlockCategory(cat_elm)  # type: ignore[arg-type]
        assert cat.gallery_value is expected

    def it_returns_None_for_gallery_value_when_unknown(self):
        cat_elm = element("w:category/w:gallery{w:val=notARealGallery}")
        cat = BuildingBlockCategory(cat_elm)  # type: ignore[arg-type]
        assert cat.gallery_value is None
        # -- but raw `gallery` still round-trips the literal
        assert cat.gallery == "notARealGallery"

    def it_returns_None_for_gallery_value_when_gallery_is_absent(self):
        cat_elm = element("w:category/w:name{w:val=General}")
        cat = BuildingBlockCategory(cat_elm)  # type: ignore[arg-type]
        assert cat.gallery_value is None

    def it_compares_equal_by_gallery_and_name(self):
        cat_elm_a = element(
            "w:category/(w:name{w:val=General},w:gallery{w:val=quickParts})"
        )
        cat_elm_b = element(
            "w:category/(w:name{w:val=General},w:gallery{w:val=quickParts})"
        )
        a = BuildingBlockCategory(cat_elm_a)  # type: ignore[arg-type]
        b = BuildingBlockCategory(cat_elm_b)  # type: ignore[arg-type]
        assert a == b
        assert hash(a) == hash(b)

    def it_compares_unequal_when_slots_differ(self):
        a_elm = element(
            "w:category/(w:name{w:val=General},w:gallery{w:val=quickParts})"
        )
        b_elm = element(
            "w:category/(w:name{w:val=Other},w:gallery{w:val=quickParts})"
        )
        a = BuildingBlockCategory(a_elm)  # type: ignore[arg-type]
        b = BuildingBlockCategory(b_elm)  # type: ignore[arg-type]
        assert a != b


# -- a mixed-category glossary used for the filter tests below --
_MIXED_GLOSSARY = (
    "w:glossaryDocument/w:docParts/("
    "w:docPart/w:docPartPr/("
    "w:name{w:val=Alpha},"
    "w:category/(w:name{w:val=General},w:gallery{w:val=quickParts})"
    "),"
    "w:docPart/w:docPartPr/("
    "w:name{w:val=Beta},"
    "w:category/(w:name{w:val=General},w:gallery{w:val=quickParts})"
    "),"
    "w:docPart/w:docPartPr/("
    "w:name{w:val=Gamma},"
    "w:category/(w:name{w:val=Built-In},w:gallery{w:val=coverPg})"
    "),"
    "w:docPart/w:docPartPr/("
    "w:name{w:val=Delta},"
    "w:category/(w:name{w:val=Built-In},w:gallery{w:val=hdrs})"
    "),"
    "w:docPart/w:docPartPr/w:name{w:val=Epsilon}"
    ")"
)


class DescribeGlossaryFiltering:
    """Unit-test suite for filter/aggregation methods on `Glossary`."""

    def it_filters_by_gallery_enum(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_MIXED_GLOSSARY)))
        blocks = glossary.by_category(
            gallery=WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS
        )
        assert [b.name for b in blocks] == ["Alpha", "Beta"]

    def it_filters_by_gallery_xml_string(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_MIXED_GLOSSARY)))
        blocks = glossary.by_category(gallery="coverPg")
        assert [b.name for b in blocks] == ["Gamma"]

    def it_filters_by_category_name(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_MIXED_GLOSSARY)))
        blocks = glossary.by_category(category_name="Built-In")
        assert [b.name for b in blocks] == ["Gamma", "Delta"]

    def it_intersects_gallery_and_category_name_filters(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_MIXED_GLOSSARY)))
        blocks = glossary.by_category(
            gallery=WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS,
            category_name="General",
        )
        assert [b.name for b in blocks] == ["Alpha", "Beta"]

    def and_an_empty_intersection_returns_an_empty_list(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_MIXED_GLOSSARY)))
        blocks = glossary.by_category(
            gallery=WD_BUILDING_BLOCK_GALLERY.QUICK_PARTS,
            category_name="Built-In",
        )
        assert blocks == []

    def it_returns_all_blocks_when_called_with_no_args(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_MIXED_GLOSSARY)))
        assert len(glossary.by_category()) == 5

    def it_dedupes_categories(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_MIXED_GLOSSARY)))
        cats = glossary.categories
        # -- Alpha + Beta share (General, quickParts); Gamma + Delta have
        # -- distinct categories; Epsilon has no category and is dropped.
        assert len(cats) == 3
        keys = [(c.gallery, c.category_name) for c in cats]
        assert keys == [
            ("quickParts", "General"),
            ("coverPg", "Built-In"),
            ("hdrs", "Built-In"),
        ]

    def it_dedupes_galleries(self):
        glossary = Glossary(cast(CT_GlossaryDocument, element(_MIXED_GLOSSARY)))
        assert glossary.galleries == ["quickParts", "coverPg", "hdrs"]

    def it_returns_empty_aggregates_for_an_empty_glossary(self):
        glossary = Glossary(
            cast(CT_GlossaryDocument, element("w:glossaryDocument"))
        )
        assert glossary.categories == []
        assert glossary.galleries == []
        assert glossary.by_category(category_name="General") == []
