"""Unit-test suite for the `docx.glossary` module."""

from __future__ import annotations

from typing import cast

import pytest

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
