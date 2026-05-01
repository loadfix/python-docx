# pyright: reportPrivateUsage=false

"""Test suite for the docx.text.run module."""

from __future__ import annotations

from typing import Any, List, cast

import pytest

from docx import types as t
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK, WD_UNDERLINE
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R, CT_Sym
from docx.parts.document import DocumentPart
from docx.shape import InlineShape
from docx.text.font import Font
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.text.symbol import Symbol

from ..unitutil.cxml import element, xml
from ..unitutil.mock import FixtureRequest, Mock, class_mock, instance_mock, property_mock


class DescribeRun:
    """Unit-test suite for `docx.text.run.Run`."""

    @pytest.mark.parametrize(
        ("r_cxml", "bool_prop_name", "expected_value"),
        [
            ("w:r/w:rPr", "bold", None),
            ("w:r/w:rPr/w:b", "bold", True),
            ("w:r/w:rPr/w:b{w:val=on}", "bold", True),
            ("w:r/w:rPr/w:b{w:val=off}", "bold", False),
            ("w:r/w:rPr/w:b{w:val=1}", "bold", True),
            ("w:r/w:rPr/w:i{w:val=0}", "italic", False),
        ],
    )
    def it_knows_its_bool_prop_states(
        self, r_cxml: str, bool_prop_name: str, expected_value: bool | None, paragraph_: Mock
    ):
        run = Run(cast(CT_R, element(r_cxml)), paragraph_)
        assert getattr(run, bool_prop_name) == expected_value

    @pytest.mark.parametrize(
        ("initial_r_cxml", "bool_prop_name", "value", "expected_cxml"),
        [
            # -- nothing to True, False, and None ---------------------------
            ("w:r", "bold", True, "w:r/w:rPr/w:b"),
            ("w:r", "bold", False, "w:r/w:rPr/w:b{w:val=0}"),
            ("w:r", "italic", None, "w:r/w:rPr"),
            # -- default to True, False, and None ---------------------------
            ("w:r/w:rPr/w:b", "bold", True, "w:r/w:rPr/w:b"),
            ("w:r/w:rPr/w:b", "bold", False, "w:r/w:rPr/w:b{w:val=0}"),
            ("w:r/w:rPr/w:i", "italic", None, "w:r/w:rPr"),
            # -- True to True, False, and None ------------------------------
            ("w:r/w:rPr/w:b{w:val=on}", "bold", True, "w:r/w:rPr/w:b"),
            ("w:r/w:rPr/w:b{w:val=1}", "bold", False, "w:r/w:rPr/w:b{w:val=0}"),
            ("w:r/w:rPr/w:b{w:val=1}", "bold", None, "w:r/w:rPr"),
            # -- False to True, False, and None -----------------------------
            ("w:r/w:rPr/w:i{w:val=false}", "italic", True, "w:r/w:rPr/w:i"),
            ("w:r/w:rPr/w:i{w:val=0}", "italic", False, "w:r/w:rPr/w:i{w:val=0}"),
            ("w:r/w:rPr/w:i{w:val=off}", "italic", None, "w:r/w:rPr"),
        ],
    )
    def it_can_change_its_bool_prop_settings(
        self,
        initial_r_cxml: str,
        bool_prop_name: str,
        value: bool | None,
        expected_cxml: str,
        paragraph_: Mock,
    ):
        run = Run(cast(CT_R, element(initial_r_cxml)), paragraph_)

        setattr(run, bool_prop_name, value)

        assert run._r.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("p_cxml", "r_idx", "expected_cxml"),
        [
            # --- run is removed from paragraph ---
            ("w:p/(w:r,w:r)", 0, "w:p/w:r"),
            # --- only run removed ---
            ("w:p/w:r", 0, "w:p"),
            # --- run with text removed, sibling kept ---
            ('w:p/(w:r/w:t"delete",w:r/w:t"keep")', 0, 'w:p/w:r/w:t"keep"'),
            # --- middle run removed ---
            ("w:p/(w:r,w:r,w:r)", 1, "w:p/(w:r,w:r)"),
        ],
    )
    def it_can_delete_itself(
        self,
        p_cxml: str,
        r_idx: int,
        expected_cxml: str,
        paragraph_: Mock,
    ):
        p = element(p_cxml)
        r = p.r_lst[r_idx]
        run = Run(cast(CT_R, r), paragraph_)

        run.delete()

        assert p.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", False),
            ('w:r/w:t"foobar"', False),
            ('w:r/(w:t"abc", w:lastRenderedPageBreak, w:t"def")', True),
            ("w:r/(w:lastRenderedPageBreak, w:lastRenderedPageBreak)", True),
        ],
    )
    def it_knows_whether_it_contains_a_page_break(
        self, r_cxml: str, expected_value: bool, paragraph_: Mock
    ):
        run = Run(cast(CT_R, element(r_cxml)), paragraph_)
        assert run.contains_page_break == expected_value

    @pytest.mark.parametrize(
        ("r_cxml", "expected"),
        [
            # -- no content produces an empty iterator --
            ("w:r", []),
            # -- contiguous text content is condensed into a single str --
            ('w:r/(w:t"foo",w:cr,w:t"bar")', ["str"]),
            # -- page-breaks are a form of inner-content --
            (
                'w:r/(w:t"abc",w:br,w:lastRenderedPageBreak,w:noBreakHyphen,w:t"def")',
                ["str", "RenderedPageBreak", "str"],
            ),
            # -- as are drawings --
            (
                'w:r/(w:t"abc", w:lastRenderedPageBreak, w:drawing)',
                ["str", "RenderedPageBreak", "Drawing"],
            ),
        ],
    )
    def it_can_iterate_its_inner_content_items(
        self, r_cxml: str, expected: List[str], fake_parent: t.ProvidesStoryPart
    ):
        r = cast(CT_R, element(r_cxml))
        run = Run(r, fake_parent)

        inner_content = run.iter_inner_content()

        actual = [type(item).__name__ for item in inner_content]
        assert actual == expected, f"expected: {expected}, got: {actual}"

    def it_exposes_its_formatting_change_when_rPrChange_present(self, paragraph_: Mock):
        r = cast(
            CT_R,
            element(
                "w:r/w:rPr/(w:b,w:rPrChange{w:id=1,w:author=Alice}/w:rPr/w:i)"
            ),
        )
        run = Run(r, paragraph_)

        fc = run.formatting_change

        assert fc is not None
        assert fc.author == "Alice"
        # old formatting was italic (w:i), current is bold (w:b)
        assert fc.old_properties is not None
        assert fc.old_properties.xpath("./w:i")

    def it_returns_None_for_formatting_change_when_no_rPr(self, paragraph_: Mock):
        run = Run(cast(CT_R, element("w:r")), paragraph_)
        assert run.formatting_change is None

    def it_returns_None_for_formatting_change_when_no_rPrChange(self, paragraph_: Mock):
        run = Run(cast(CT_R, element("w:r/w:rPr/w:b")), paragraph_)
        assert run.formatting_change is None

    def it_can_mark_a_comment_reference_range(self, paragraph_: Mock):
        p = cast(CT_P, element('w:p/w:r/w:t"referenced text"'))
        run = last_run = Run(p.r_lst[0], paragraph_)

        run.mark_comment_range(last_run, comment_id=42)

        assert p.xml == xml(
            'w:p/(w:commentRangeStart{w:id=42},w:r/w:t"referenced text"'
            ",w:commentRangeEnd{w:id=42}"
            ",w:r/(w:rPr/w:rStyle{w:val=CommentReference},w:commentReference{w:id=42}))"
        )

    def it_knows_its_character_style(
        self, part_prop_: Mock, document_part_: Mock, paragraph_: Mock
    ):
        style_ = document_part_.get_style.return_value
        part_prop_.return_value = document_part_
        style_id = "Barfoo"
        run = Run(cast(CT_R, element(f"w:r/w:rPr/w:rStyle{{w:val={style_id}}}")), paragraph_)

        style = run.style

        document_part_.get_style.assert_called_once_with(style_id, WD_STYLE_TYPE.CHARACTER)
        assert style is style_

    @pytest.mark.parametrize(
        ("r_cxml", "value", "style_id", "expected_cxml"),
        [
            ("w:r", "Foo Font", "FooFont", "w:r/w:rPr/w:rStyle{w:val=FooFont}"),
            ("w:r/w:rPr", "Foo Font", "FooFont", "w:r/w:rPr/w:rStyle{w:val=FooFont}"),
            (
                "w:r/w:rPr/w:rStyle{w:val=FooFont}",
                "Bar Font",
                "BarFont",
                "w:r/w:rPr/w:rStyle{w:val=BarFont}",
            ),
            ("w:r/w:rPr/w:rStyle{w:val=FooFont}", None, None, "w:r/w:rPr"),
            ("w:r", None, None, "w:r/w:rPr"),
        ],
    )
    def it_can_change_its_character_style(
        self,
        r_cxml: str,
        value: str | None,
        style_id: str | None,
        expected_cxml: str,
        part_prop_: Mock,
        paragraph_: Mock,
    ):
        part_ = part_prop_.return_value
        part_.get_style_id.return_value = style_id
        run = Run(cast(CT_R, element(r_cxml)), paragraph_)

        run.style = value

        part_.get_style_id.assert_called_once_with(value, WD_STYLE_TYPE.CHARACTER)
        assert run._r.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("r_cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r/w:rPr/w:u", None),
            ("w:r/w:rPr/w:u{w:val=single}", True),
            ("w:r/w:rPr/w:u{w:val=none}", False),
            ("w:r/w:rPr/w:u{w:val=double}", WD_UNDERLINE.DOUBLE),
            ("w:r/w:rPr/w:u{w:val=wave}", WD_UNDERLINE.WAVY),
        ],
    )
    def it_knows_its_underline_type(
        self, r_cxml: str, expected_value: bool | WD_UNDERLINE | None, paragraph_: Mock
    ):
        run = Run(cast(CT_R, element(r_cxml)), paragraph_)
        assert run.underline is expected_value

    @pytest.mark.parametrize(
        ("initial_r_cxml", "new_underline", "expected_cxml"),
        [
            ("w:r", True, "w:r/w:rPr/w:u{w:val=single}"),
            ("w:r", False, "w:r/w:rPr/w:u{w:val=none}"),
            ("w:r", None, "w:r/w:rPr"),
            ("w:r", WD_UNDERLINE.SINGLE, "w:r/w:rPr/w:u{w:val=single}"),
            ("w:r", WD_UNDERLINE.THICK, "w:r/w:rPr/w:u{w:val=thick}"),
            ("w:r/w:rPr/w:u{w:val=single}", True, "w:r/w:rPr/w:u{w:val=single}"),
            ("w:r/w:rPr/w:u{w:val=single}", False, "w:r/w:rPr/w:u{w:val=none}"),
            ("w:r/w:rPr/w:u{w:val=single}", None, "w:r/w:rPr"),
            (
                "w:r/w:rPr/w:u{w:val=single}",
                WD_UNDERLINE.SINGLE,
                "w:r/w:rPr/w:u{w:val=single}",
            ),
            (
                "w:r/w:rPr/w:u{w:val=single}",
                WD_UNDERLINE.DOTTED,
                "w:r/w:rPr/w:u{w:val=dotted}",
            ),
        ],
    )
    def it_can_change_its_underline_type(
        self,
        initial_r_cxml: str,
        new_underline: bool | WD_UNDERLINE | None,
        expected_cxml: str,
        paragraph_: Mock,
    ):
        run = Run(cast(CT_R, element(initial_r_cxml)), paragraph_)

        run.underline = new_underline

        assert run._r.xml == xml(expected_cxml)

    @pytest.mark.parametrize("invalid_value", ["foobar", 42, "single"])
    def it_raises_on_assign_invalid_underline_value(self, invalid_value: Any, paragraph_: Mock):
        run = Run(cast(CT_R, element("w:r/w:rPr")), paragraph_)
        with pytest.raises(ValueError, match=" is not a valid WD_UNDERLINE"):
            run.underline = invalid_value

    def it_provides_access_to_its_font(self, Font_: Mock, font_: Mock, paragraph_: Mock):
        Font_.return_value = font_
        run = Run(cast(CT_R, element("w:r")), paragraph_)

        font = run.font

        Font_.assert_called_once_with(run._element)
        assert font is font_

    @pytest.mark.parametrize(
        ("r_cxml", "new_text", "expected_cxml"),
        [
            ("w:r", "foo", 'w:r/w:t"foo"'),
            ('w:r/w:t"foo"', "bar", 'w:r/(w:t"foo", w:t"bar")'),
            ("w:r", "fo ", 'w:r/w:t{xml:space=preserve}"fo "'),
            ("w:r", "f o", 'w:r/w:t"f o"'),
        ],
    )
    def it_can_add_text(
        self, r_cxml: str, new_text: str, expected_cxml: str, Text_: Mock, paragraph_: Mock
    ):
        run = Run(cast(CT_R, element(r_cxml)), paragraph_)

        text = run.add_text(new_text)

        assert run._r.xml == xml(expected_cxml)
        assert text is Text_.return_value

    @pytest.mark.parametrize(
        ("break_type", "expected_cxml"),
        [
            (WD_BREAK.LINE, "w:r/w:br"),
            (WD_BREAK.PAGE, "w:r/w:br{w:type=page}"),
            (WD_BREAK.COLUMN, "w:r/w:br{w:type=column}"),
            (WD_BREAK.LINE_CLEAR_LEFT, "w:r/w:br{w:clear=left}"),
            (WD_BREAK.LINE_CLEAR_RIGHT, "w:r/w:br{w:clear=right}"),
            (WD_BREAK.LINE_CLEAR_ALL, "w:r/w:br{w:clear=all}"),
        ],
    )
    def it_can_add_a_break(self, break_type: WD_BREAK, expected_cxml: str, paragraph_: Mock):
        run = Run(cast(CT_R, element("w:r")), paragraph_)

        run.add_break(break_type)

        assert run._r.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("r_cxml", "expected_cxml"), [('w:r/w:t"foo"', 'w:r/(w:t"foo", w:tab)')]
    )
    def it_can_add_a_tab(self, r_cxml: str, expected_cxml: str, paragraph_: Mock):
        run = Run(cast(CT_R, element(r_cxml)), paragraph_)

        run.add_tab()

        assert run._r.xml == xml(expected_cxml)

    def it_can_add_a_picture(
        self,
        part_prop_: Mock,
        document_part_: Mock,
        InlineShape_: Mock,
        picture_: Mock,
        paragraph_: Mock,
    ):
        part_prop_.return_value = document_part_
        run = Run(cast(CT_R, element("w:r/wp:x")), paragraph_)
        image = "foobar.png"
        width, height, inline = 1111, 2222, element("wp:inline{id=42}")
        document_part_.new_pic_inline.return_value = inline
        InlineShape_.return_value = picture_

        picture = run.add_picture(image, width, height)

        document_part_.new_pic_inline.assert_called_once_with(image, width, height)
        assert run._r.xml == xml("w:r/(wp:x,w:drawing/wp:inline{id=42})")
        InlineShape_.assert_called_once_with(inline)
        assert picture is picture_

    @pytest.mark.parametrize(
        ("initial_r_cxml", "expected_cxml"),
        [
            ("w:r", "w:r"),
            ('w:r/w:t"foo"', "w:r"),
            ("w:r/w:br", "w:r"),
            ("w:r/w:rPr", "w:r/w:rPr"),
            ('w:r/(w:rPr, w:t"foo")', "w:r/w:rPr"),
            (
                'w:r/(w:rPr/(w:b, w:i), w:t"foo", w:cr, w:t"bar")',
                "w:r/w:rPr/(w:b, w:i)",
            ),
        ],
    )
    def it_can_remove_its_content_but_keep_formatting(
        self, initial_r_cxml: str, expected_cxml: str, paragraph_: Mock
    ):
        run = Run(cast(CT_R, element(initial_r_cxml)), paragraph_)

        cleared_run = run.clear()

        assert run._r.xml == xml(expected_cxml)
        assert cleared_run is run

    @pytest.mark.parametrize(
        ("p_cxml", "offset", "expected_left", "expected_right"),
        [
            ('w:p/w:r/w:t"foobar"', 3, "foo", "bar"),
            ('w:p/w:r/(w:rPr/w:b,w:t"foobar")', 3, "foo", "bar"),
            ('w:p/w:r/w:t"foobar"', 0, "", "foobar"),
            ('w:p/w:r/w:t"foobar"', 6, "foobar", ""),
        ],
    )
    def it_can_split_at_a_character_position(
        self,
        p_cxml: str,
        offset: int,
        expected_left: str,
        expected_right: str,
        fake_parent: t.ProvidesStoryPart,
    ):
        p = cast(CT_P, element(p_cxml))
        run = Run(p.r_lst[0], fake_parent)

        left, right = run.split(offset)

        assert left is run
        assert isinstance(right, Run)
        assert left.text == expected_left
        assert right.text == expected_right

    def it_preserves_formatting_on_split(self, fake_parent: t.ProvidesStoryPart):
        p = cast(CT_P, element('w:p/w:r/(w:rPr/(w:b,w:i),w:t"foobar")'))
        run = Run(p.r_lst[0], fake_parent)

        left, right = run.split(3)

        assert left.bold is True
        assert right.bold is True
        assert left.italic is True
        assert right.italic is True

    @pytest.mark.parametrize(
        ("r_cxml", "expected_text"),
        [
            ("w:r", ""),
            ('w:r/w:t"foobar"', "foobar"),
            ('w:r/(w:t"abc", w:tab, w:t"def", w:cr)', "abc\tdef\n"),
            ('w:r/(w:br{w:type=page}, w:t"abc", w:t"def", w:tab)', "abcdef\t"),
        ],
    )
    def it_knows_the_text_it_contains(self, r_cxml: str, expected_text: str, paragraph_: Mock):
        run = Run(cast(CT_R, element(r_cxml)), paragraph_)
        assert run.text == expected_text

    @pytest.mark.parametrize(
        ("new_text", "expected_cxml"),
        [
            ("abc  def", 'w:r/w:t"abc  def"'),
            ("abc\tdef", 'w:r/(w:t"abc", w:tab, w:t"def")'),
            ("abc\ndef", 'w:r/(w:t"abc", w:br,  w:t"def")'),
            ("abc\rdef", 'w:r/(w:t"abc", w:br,  w:t"def")'),
        ],
    )
    def it_can_replace_the_text_it_contains(
        self, new_text: str, expected_cxml: str, paragraph_: Mock
    ):
        run = Run(cast(CT_R, element('w:r/w:t"should get deleted"')), paragraph_)

        run.text = new_text

        assert run._r.xml == xml(expected_cxml)

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def document_part_(self, request: FixtureRequest):
        return instance_mock(request, DocumentPart)

    @pytest.fixture
    def Font_(self, request: FixtureRequest):
        return class_mock(request, "docx.text.run.Font")

    @pytest.fixture
    def font_(self, request: FixtureRequest):
        return instance_mock(request, Font)

    @pytest.fixture
    def InlineShape_(self, request: FixtureRequest):
        return class_mock(request, "docx.text.run.InlineShape")

    @pytest.fixture
    def paragraph_(self, request: FixtureRequest):
        return instance_mock(request, Paragraph)

    @pytest.fixture
    def part_prop_(self, request: FixtureRequest):
        return property_mock(request, Run, "part")

    @pytest.fixture
    def picture_(self, request: FixtureRequest):
        return instance_mock(request, InlineShape)

    @pytest.fixture
    def Text_(self, request: FixtureRequest):
        return class_mock(request, "docx.text.run._Text")


class DescribeRun_Symbol:
    """Unit-test suite for `Run.add_symbol` and `Run.symbols`."""

    @pytest.mark.parametrize(
        ("char_code", "font", "expected_char_hex", "expected_cxml"),
        [
            # -- integer code ----------------------------------------------
            (0xF0E0, "Wingdings", "F0E0", "w:r/w:sym{w:font=Wingdings,w:char=F0E0}"),
            # -- integer code padded to 4 hex chars ------------------------
            (0x41, "Symbol", "0041", "w:r/w:sym{w:font=Symbol,w:char=0041}"),
            # -- 4-char hex string -----------------------------------------
            ("F0E0", "Wingdings", "F0E0", "w:r/w:sym{w:font=Wingdings,w:char=F0E0}"),
            # -- lowercase hex normalised to uppercase ---------------------
            ("f0e0", "Wingdings", "F0E0", "w:r/w:sym{w:font=Wingdings,w:char=F0E0}"),
            # -- 0x-prefixed hex string ------------------------------------
            ("0xF0E0", "Wingdings", "F0E0", "w:r/w:sym{w:font=Wingdings,w:char=F0E0}"),
        ],
    )
    def it_can_add_a_symbol(
        self,
        char_code: int | str,
        font: str,
        expected_char_hex: str,
        expected_cxml: str,
        paragraph_: Mock,
    ):
        run = Run(cast(CT_R, element("w:r")), paragraph_)

        symbol = run.add_symbol(char_code, font)

        assert run._r.xml == xml(expected_cxml)
        assert isinstance(symbol, Symbol)
        assert symbol.font == font
        assert symbol.char_hex == expected_char_hex

    def it_iterates_symbols_in_document_order(self, paragraph_: Mock):
        r_cxml = (
            "w:r/(w:sym{w:font=Wingdings,w:char=F0E0},"
            'w:t"x",'
            "w:sym{w:font=Symbol,w:char=0041})"
        )
        run = Run(cast(CT_R, element(r_cxml)), paragraph_)

        symbols = list(run.symbols)

        assert len(symbols) == 2
        assert [s.char_hex for s in symbols] == ["F0E0", "0041"]
        assert [s.font for s in symbols] == ["Wingdings", "Symbol"]

    def it_yields_no_symbols_when_none_present(self, paragraph_: Mock):
        run = Run(cast(CT_R, element('w:r/w:t"no symbols here"')), paragraph_)
        assert list(run.symbols) == []

    def its_symbols_report_int_char_code(self, paragraph_: Mock):
        run = Run(
            cast(CT_R, element("w:r/w:sym{w:font=Wingdings,w:char=F0E0}")),
            paragraph_,
        )

        (symbol,) = list(run.symbols)

        assert symbol.char_code == 0xF0E0
        assert symbol.char_hex == "F0E0"
        assert symbol.font == "Wingdings"

    def its_symbols_can_be_deleted(self, paragraph_: Mock):
        r = cast(
            CT_R,
            element(
                "w:r/(w:sym{w:font=Wingdings,w:char=F0E0},"
                'w:t"x",'
                "w:sym{w:font=Symbol,w:char=0041})"
            ),
        )
        run = Run(r, paragraph_)
        first_symbol = next(iter(run.symbols))

        first_symbol.delete()

        remaining = list(run.symbols)
        assert len(remaining) == 1
        assert remaining[0].char_hex == "0041"
        # -- `w:t` is preserved --
        assert run.text == "x"

    def it_preserves_symbols_when_paragraph_text_is_read(self, paragraph_: Mock):
        """`Run.text` skips symbol elements but leaves them in the XML."""
        r = cast(
            CT_R,
            element(
                'w:r/(w:t"before",'
                "w:sym{w:font=Wingdings,w:char=F0E0},"
                'w:t"after")'
            ),
        )
        run = Run(r, paragraph_)

        # -- symbol is not part of the text string (non-textual) --
        assert run.text == "beforeafter"
        # -- but the w:sym element is still present in the XML --
        assert len(list(run.symbols)) == 1

    @pytest.fixture
    def paragraph_(self, request: FixtureRequest):
        return instance_mock(request, Paragraph)


class DescribeSymbol:
    """Unit-test suite for `docx.text.symbol.Symbol`."""

    def it_reports_its_font_char_hex_and_char_code(self):
        sym = cast(
            CT_Sym,
            element("w:sym{w:font=Wingdings,w:char=F0E0}"),
        )

        symbol = Symbol(sym)

        assert symbol.font == "Wingdings"
        assert symbol.char_hex == "F0E0"
        assert symbol.char_code == 0xF0E0

    def its_char_hex_is_uppercase_and_padded(self):
        # -- even if the XML stored the value in lowercase or without padding,
        # -- the computed char_hex is normalised --
        sym = cast(CT_Sym, element("w:sym{w:font=Symbol,w:char=41}"))

        symbol = Symbol(sym)

        assert symbol.char_code == 0x41
        assert symbol.char_hex == "0041"

    def it_can_delete_itself(self):
        r = cast(
            CT_R,
            element(
                "w:r/(w:sym{w:font=Wingdings,w:char=F0E0},"
                "w:sym{w:font=Symbol,w:char=0041})"
            ),
        )
        first_sym = r.sym_lst[0]
        symbol = Symbol(first_sym)

        symbol.delete()

        assert len(r.sym_lst) == 1
        assert r.sym_lst[0].char == "0041"

    def delete_is_a_noop_when_element_is_orphan(self):
        sym = cast(CT_Sym, element("w:sym{w:font=Wingdings,w:char=F0E0}"))
        # -- `element()` returns an orphan (no parent) --
        symbol = Symbol(sym)

        # -- should not raise --
        symbol.delete()


class DescribeRun_Rsid:
    """Unit-test suite for the RSID accessor on `Run`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:r", None),
            ("w:r{w:rsidR=00FA1B42}", "00FA1B42"),
            ("w:r{w:rsidR=001234AB,w:rsidRPr=005678CD}", "001234AB"),
        ],
    )
    def it_reads_rsid_from_rsidR_attribute(
        self, cxml: str, expected_value: str | None, paragraph_
    ):
        run = Run(cast(CT_R, element(cxml)), paragraph_)
        assert run.rsid == expected_value

    @pytest.fixture
    def paragraph_(self, request: FixtureRequest):
        return instance_mock(request, Paragraph)


class DescribeRun_StableId:
    """Unit-test suite for the `stable_id` accessor on `Run`."""

    def it_returns_a_16_character_hex_string(self, paragraph_):
        r = cast(CT_R, element("w:r/w:t\"hello\""))
        run = Run(r, paragraph_)

        result = run.stable_id

        assert isinstance(result, str)
        assert len(result) == 16
        assert all(c in "0123456789abcdef" for c in result)

    def it_returns_the_same_id_on_repeated_access(self, paragraph_):
        r = cast(CT_R, element("w:r/w:t\"hello\""))
        run = Run(r, paragraph_)
        assert run.stable_id == run.stable_id

    def it_returns_different_ids_for_different_text(self, paragraph_):
        r1 = cast(CT_R, element("w:r/w:t\"alpha\""))
        r2 = cast(CT_R, element("w:r/w:t\"beta\""))
        assert Run(r1, paragraph_).stable_id != Run(r2, paragraph_).stable_id

    def it_returns_different_ids_for_sibling_runs_with_same_text(self, paragraph_):
        p = element("w:p/(w:r/w:t\"same\",w:r/w:t\"same\")")
        run_a = Run(cast(CT_R, p[0]), paragraph_)
        run_b = Run(cast(CT_R, p[1]), paragraph_)
        assert run_a.stable_id != run_b.stable_id

    def it_incorporates_rsidR_when_present(self, paragraph_):
        r_no_rsid = cast(CT_R, element("w:r/w:t\"hello\""))
        r_rsid = cast(CT_R, element("w:r{w:rsidR=00FA1B42}/w:t\"hello\""))
        assert (
            Run(r_no_rsid, paragraph_).stable_id
            != Run(r_rsid, paragraph_).stable_id
        )

    def it_works_when_run_lacks_rsidR(self, paragraph_):
        r = cast(CT_R, element("w:r/w:t\"hello\""))
        run = Run(r, paragraph_)
        assert run.rsid is None
        assert len(run.stable_id) == 16

    @pytest.fixture
    def paragraph_(self, request: FixtureRequest):
        return instance_mock(request, Paragraph)
