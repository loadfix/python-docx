"""Unit-test suite for `Styles.import_from`, `import_style`, `import_builtin`
and `Styles.document_default_font` — upstream#1375, #1083, #508, #701, #197,
#486, #383.
"""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.styles import CT_Styles
from docx.styles.styles import Styles
from docx.text.font import Font

from ..unitutil.cxml import element


class DescribeStyles_ImportFrom:
    """Unit-test suite for `Styles.import_from` — cross-document style import."""

    def it_copies_a_named_style_into_the_target(self):
        source = _styles(
            "w:styles/("
            "w:style{w:type=paragraph,w:styleId=Fancy}/w:name{w:val=Fancy},"
            "w:style{w:type=paragraph,w:styleId=Plain}/w:name{w:val=Plain}"
            ")"
        )
        target = _styles("w:styles")

        imported = target.import_from(source, names=["Fancy"])

        assert len(imported) == 1
        assert imported[0].style_id == "Fancy"
        assert "Fancy" in target

    def it_skips_styles_already_present_in_the_target(self):
        source = _styles(
            "w:styles/w:style{w:type=paragraph,w:styleId=Fancy}/w:name{w:val=Fancy}"
        )
        target = _styles(
            "w:styles/w:style{w:type=paragraph,w:styleId=Fancy}/w:name{w:val=Fancy}"
        )

        imported = target.import_from(source)

        assert imported == []
        # -- still exactly one Fancy in the target --
        assert len([s for s in target if s.style_id == "Fancy"]) == 1

    def it_imports_basedOn_link_and_next_dependencies(self):
        source = _styles(
            "w:styles/("
            "w:style{w:type=paragraph,w:styleId=Body}/w:name{w:val=Body},"
            "w:style{w:type=character,w:styleId=BodyChar}/w:name{w:val=BodyChar},"
            "w:style{w:type=paragraph,w:styleId=NextStyle}/w:name{w:val=NextStyle},"
            "w:style{w:type=paragraph,w:styleId=Fancy}/("
            "w:name{w:val=Fancy},"
            "w:basedOn{w:val=Body},"
            "w:next{w:val=NextStyle},"
            "w:link{w:val=BodyChar}"
            ")"
            ")"
        )
        target = _styles("w:styles")

        target.import_from(source, names=["Fancy"])

        ids = {s.style_id for s in target}
        assert {"Fancy", "Body", "BodyChar", "NextStyle"} <= ids

    def it_accepts_objects_with_a_styles_attribute(self):
        source = _styles(
            "w:styles/w:style{w:type=paragraph,w:styleId=Fancy}/w:name{w:val=Fancy}"
        )

        class _FakeDoc:
            styles = source

        target = _styles("w:styles")
        target.import_from(_FakeDoc())
        assert "Fancy" in target


class DescribeStyles_ImportStyle:
    """Unit-test suite for `Styles.import_style` — single-style deep copy."""

    def it_returns_an_existing_style_untouched(self):
        target = _styles(
            "w:styles/w:style{w:type=paragraph,w:styleId=Fancy}/w:name{w:val=Fancy}"
        )
        source = _styles(
            "w:styles/w:style{w:type=paragraph,w:styleId=Fancy}/w:name{w:val=Fancy}"
        )
        source_elm = source._element.get_by_id("Fancy")

        result = target.import_style(source_elm)

        assert result.style_id == "Fancy"
        assert len([s for s in target if s.style_id == "Fancy"]) == 1

    def it_imports_the_style_when_not_present(self):
        source = _styles(
            "w:styles/w:style{w:type=paragraph,w:styleId=Fancy}/w:name{w:val=Fancy}"
        )
        target = _styles("w:styles")

        target.import_style(source._element.get_by_id("Fancy"))

        assert "Fancy" in target


class DescribeStyles_ImportBuiltin:
    """Unit-test suite for `Styles.import_builtin` — upstream#486."""

    def it_materialises_List_Bullet_from_the_bundled_defaults(self):
        target = _styles("w:styles")

        style = target.import_builtin("List Bullet")

        assert style.style_id == "ListBullet"
        assert "List Bullet" in target

    def it_raises_KeyError_for_unknown_names(self):
        target = _styles("w:styles")

        with pytest.raises(KeyError):
            target.import_builtin("NoSuchBuiltinStyle")


class DescribeStyles_DocumentDefaultFont:
    """Unit-test suite for `Styles.document_default_font` — upstream#383."""

    def it_returns_a_Font_over_the_docDefaults_rPr(self):
        target = _styles("w:styles")

        font = target.document_default_font

        assert isinstance(font, Font)
        # -- Writing through the Font proxy persists on the underlying XML --
        font.bold = True
        docDefaults = target._element.docDefaults
        assert docDefaults is not None
        rPrDefault = docDefaults.rPrDefault
        assert rPrDefault is not None
        # -- rPr auto-created on write --
        assert rPrDefault.rPr is not None

    def it_returns_a_live_view_so_repeated_access_sees_writes(self):
        target = _styles("w:styles")

        target.document_default_font.italic = True
        # -- Re-read through a fresh proxy and verify the value round-trips --
        assert target.document_default_font.italic is True


def _styles(cxml: str) -> Styles:
    return Styles(cast(CT_Styles, element(cxml)))
