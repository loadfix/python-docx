"""Tests for the structured ``DocxError`` taxonomy.

Covers a representative slice of the new error hierarchy: that each
raise-site populates ``code`` / ``suggestion`` / ``location`` /
``operation``, that the structured subclasses still satisfy the legacy
``except KeyError:`` / ``except ValueError:`` / ``except IndexError:``
contract, and that the ``_did_you_mean`` fuzzy-match helper finds the
expected close matches.
"""

from __future__ import annotations

import pytest

from docx.bookmarks import Bookmarks
from docx.dml.color import ColorFormat
from docx.enum.style import WD_STYLE_TYPE
from docx.exceptions import (
    BookmarkNotFoundError,
    BuiltinStyleNotFoundError,
    DocxError,
    FontEmbedEmptyError,
    FontFamilyInvalidError,
    FontNotFoundError,
    InvalidBrightnessError,
    InvalidColorError,
    LatentStyleNotFoundError,
    NotAWordTemplateError,
    OutOfRangeError,
    PythonDocxError,
    StyleDuplicateError,
    StyleNotFoundError,
    StyleTypeMismatchError,
    ThemeTokenInvalidError,
    ValueOutOfRangeError,
    __all_codes__,
    closest_names,
    is_docx_error,
)
from docx.font_table import FontTable
from docx.oxml.parser import parse_xml
from docx.shared import RGBColor
from docx.styles.styles import Styles
from docx.theme import ThemeColors


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _styles_with(*names: str) -> Styles:
    """Return a |Styles| proxy populated with `names`."""
    inner = "".join(
        f'<w:style w:type="paragraph" w:styleId="{n.replace(" ", "")}">'
        f'<w:name w:val="{n}"/></w:style>'
        for n in names
    )
    xml = f'<w:styles xmlns:w="{_W_NS}">{inner}</w:styles>'
    return Styles(parse_xml(xml))


def _bookmarks_with(*names: str) -> Bookmarks:
    """Return a |Bookmarks| proxy whose body contains `names` as bookmarks."""
    inner = "".join(
        f'<w:bookmarkStart w:id="{i}" w:name="{n}"/>'
        f'<w:bookmarkEnd w:id="{i}"/>'
        for i, n in enumerate(names)
    )
    xml = f'<w:body xmlns:w="{_W_NS}"><w:p>{inner}</w:p></w:body>'
    return Bookmarks(parse_xml(xml))


def _font_table_with(*names: str) -> FontTable:
    """Return a |FontTable| with one ``w:font`` per name (no part attached)."""
    inner = "".join(f'<w:font w:name="{n}"/>' for n in names)
    xml = f'<w:fonts xmlns:w="{_W_NS}">{inner}</w:fonts>'
    return FontTable(parse_xml(xml), part=None)  # type: ignore[arg-type]


class DescribeDocxError:
    def it_carries_structured_fields(self):
        err = DocxError(
            "boom",
            code="X",
            suggestion="hint",
            location="here",
            operation="op",
        )
        assert err.code == "X"
        assert err.message == "boom"
        assert err.suggestion == "hint"
        assert err.location == "here"
        assert err.operation == "op"
        assert str(err).startswith("[X] boom")

    def it_serialises_to_dict(self):
        err = DocxError("m", code="C", suggestion="s", location="l", operation="o")
        assert err.to_dict() == {
            "code": "C",
            "message": "m",
            "suggestion": "s",
            "location": "l",
            "operation": "o",
        }

    def it_uses_default_code_when_unspecified(self):
        err = StyleNotFoundError("nope")
        assert err.code == "STYLE_NOT_FOUND"

    def it_inherits_from_pythondocx_error(self):
        err = StyleNotFoundError("nope")
        assert isinstance(err, PythonDocxError)
        assert is_docx_error(err)


class DescribeBackwardsCompat:
    """Multi-inheritance preserves legacy ``except`` blocks."""

    def it_makes_StyleNotFoundError_a_KeyError(self):
        styles = _styles_with("Heading 1", "Body Text")
        with pytest.raises(KeyError):
            _ = styles["Nonexistent"]

    def it_makes_BookmarkNotFoundError_a_KeyError(self):
        bookmarks = _bookmarks_with("intro", "conclusion")
        with pytest.raises(KeyError):
            _ = bookmarks["middle"]

    def it_makes_FontNotFoundError_a_KeyError(self):
        ft = _font_table_with("Arial", "Calibri")
        with pytest.raises(KeyError):
            _ = ft["NoSuchFont"]

    def it_makes_OutOfRangeError_an_IndexError(self):
        # -- pop() needs a Sections proxy; the simpler validation is to --
        # -- assert OutOfRangeError extends IndexError. --
        assert issubclass(OutOfRangeError, IndexError)
        assert issubclass(OutOfRangeError, DocxError)

    def it_makes_value_errors_keep_ValueError_inheritance(self):
        for cls in (
            StyleDuplicateError,
            StyleTypeMismatchError,
            FontFamilyInvalidError,
            FontEmbedEmptyError,
            InvalidColorError,
            InvalidBrightnessError,
            ValueOutOfRangeError,
            NotAWordTemplateError,
        ):
            assert issubclass(cls, ValueError)
            assert issubclass(cls, DocxError)


class DescribeStyleNotFoundError:
    def it_populates_code_and_operation(self):
        styles = _styles_with("Heading 1", "Body Text", "Normal")
        with pytest.raises(StyleNotFoundError) as exc_info:
            _ = styles["Body Texxt"]  # extra letter — typo
        err = exc_info.value
        assert err.code == "STYLE_NOT_FOUND"
        assert err.operation == "Styles.__getitem__"
        assert err.location is not None and "Body Texxt" in err.location

    def it_suggests_a_close_match(self):
        styles = _styles_with("Heading 1", "Body Text", "Normal")
        with pytest.raises(StyleNotFoundError) as exc_info:
            _ = styles["Body Texxt"]
        # -- difflib finds "Body Text" as the closest match --
        assert exc_info.value.suggestion is not None
        assert "Body Text" in exc_info.value.suggestion


class DescribeStyleDuplicateError:
    def it_fires_when_adding_a_duplicate(self):
        styles = _styles_with("Heading 1")
        with pytest.raises(StyleDuplicateError) as exc_info:
            styles.add_style("Heading 1", WD_STYLE_TYPE.PARAGRAPH)
        err = exc_info.value
        assert err.code == "STYLE_DUPLICATE"
        assert err.operation == "Styles.add_style"
        # -- The suggestion guides the caller toward the right fix. --
        assert err.suggestion is not None
        assert "unique" in err.suggestion.lower()


class DescribeBuiltinStyleNotFoundError:
    def it_fires_for_unknown_builtin(self):
        styles = _styles_with()
        with pytest.raises(BuiltinStyleNotFoundError) as exc_info:
            styles.import_builtin("ThisIsNotABuiltin")
        err = exc_info.value
        assert err.code == "BUILTIN_STYLE_NOT_FOUND"
        assert err.operation == "Styles.import_builtin"


class DescribeLatentStyleNotFoundError:
    def it_fires_for_unknown_latent(self):
        # -- A LatentStyles element with one named exception. --
        latent_xml = parse_xml(
            f'<w:latentStyles xmlns:w="{_W_NS}">'
            f'<w:lsdException w:name="Caption"/></w:latentStyles>'
        )
        from docx.styles.latent import LatentStyles
        latents = LatentStyles(latent_xml)
        with pytest.raises(LatentStyleNotFoundError) as exc_info:
            _ = latents["Capshun"]  # typo
        err = exc_info.value
        assert err.code == "LATENT_STYLE_NOT_FOUND"
        assert err.operation == "LatentStyles.__getitem__"


class DescribeBookmarkNotFoundError:
    def it_populates_suggestion_with_close_match(self):
        bookmarks = _bookmarks_with("introduction", "conclusion")
        with pytest.raises(BookmarkNotFoundError) as exc_info:
            _ = bookmarks["intoduction"]  # missing 'r'
        err = exc_info.value
        assert err.code == "BOOKMARK_NOT_FOUND"
        assert err.operation == "Bookmarks.__getitem__"
        assert err.suggestion is not None
        assert "introduction" in err.suggestion

    def it_uses_remove_operation_for_missing_remove(self):
        bookmarks = _bookmarks_with("intro")
        with pytest.raises(BookmarkNotFoundError) as exc_info:
            bookmarks.remove("does_not_exist")
        assert exc_info.value.operation == "Bookmarks.remove"


class DescribeFontNotFoundError:
    def it_suggests_a_known_name(self):
        ft = _font_table_with("Arial", "Calibri", "Times New Roman")
        with pytest.raises(FontNotFoundError) as exc_info:
            _ = ft["Ariel"]  # typo for "Arial"
        err = exc_info.value
        assert err.code == "FONT_NOT_FOUND"
        assert err.suggestion is not None and "Arial" in err.suggestion


class DescribeFontFamilyInvalidError:
    def it_fires_for_unknown_family_token(self):
        ft = _font_table_with()
        with pytest.raises(FontFamilyInvalidError) as exc_info:
            ft.add_embedded_font("/dev/null", family="reglar")  # type: ignore[arg-type]
        err = exc_info.value
        assert err.code == "FONT_FAMILY_INVALID"
        # -- 'regular' is the closest match --
        assert err.suggestion is not None
        assert "regular" in err.suggestion


class DescribeFontEmbedEmptyError:
    def it_fires_when_no_variants_supplied(self):
        ft = _font_table_with()
        with pytest.raises(FontEmbedEmptyError) as exc_info:
            ft.embed_font("Helvetica")
        err = exc_info.value
        assert err.code == "FONT_EMBED_EMPTY"
        assert err.operation == "FontTable.embed_font"


class DescribeThemeTokenInvalidError:
    def it_fires_for_unknown_token(self):
        tc = ThemeColors(None)
        with pytest.raises(ThemeTokenInvalidError) as exc_info:
            _ = tc["accent99"]
        err = exc_info.value
        assert err.code == "THEME_TOKEN_INVALID"
        assert err.suggestion is not None
        # -- "accent1".."accent6" are close to "accent99" --
        assert "accent" in err.suggestion


class DescribeInvalidColorError:
    def it_fires_for_out_of_range_components(self):
        with pytest.raises(InvalidColorError) as exc_info:
            RGBColor(300, 0, 0)
        err = exc_info.value
        assert err.code == "INVALID_COLOR"
        assert err.operation == "RGBColor.__new__"

    def it_fires_for_bad_hex_string(self):
        with pytest.raises(InvalidColorError) as exc_info:
            RGBColor.from_string("ABCD")  # 4 chars — neither 3 nor 6
        err = exc_info.value
        assert err.code == "INVALID_COLOR"
        assert err.operation == "RGBColor.from_string"


class DescribeInvalidBrightnessError:
    def it_fires_for_out_of_range_value(self):
        # -- a w:r element with a w:rPr/w:color/@w:themeColor so the second --
        # -- guard (no theme color) does not trigger first --
        r_xml = (
            '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:rPr><w:color w:val="auto" w:themeColor="accent1"/></w:rPr>'
            '</w:r>'
        )
        r = parse_xml(r_xml)
        cf = ColorFormat(r)
        with pytest.raises(InvalidBrightnessError) as exc_info:
            cf.brightness = 2.0
        assert exc_info.value.code == "INVALID_BRIGHTNESS"

    def it_fires_when_no_theme_color_is_set(self):
        r = parse_xml(f'<w:r xmlns:w="{_W_NS}"/>')
        cf = ColorFormat(r)
        with pytest.raises(InvalidBrightnessError) as exc_info:
            cf.brightness = 0.5
        assert exc_info.value.code == "BRIGHTNESS_NO_THEME"
        assert exc_info.value.suggestion is not None
        assert "theme_color" in exc_info.value.suggestion


class DescribeStyleTypeMismatch:
    def it_emits_a_structured_mismatch_error(self):
        styles = _styles_with("Heading 1")
        # -- Look up "Heading 1" (paragraph-typed in our fixture) and ask for --
        # -- its id as a *character* style — triggers the type-mismatch guard. --
        with pytest.raises(StyleTypeMismatchError) as exc_info:
            styles.get_style_id("Heading 1", WD_STYLE_TYPE.CHARACTER)
        err = exc_info.value
        assert err.code == "STYLE_TYPE_MISMATCH"
        assert err.operation == "Styles._get_style_id_from_style"


class DescribeClosestNamesHelper:
    def it_returns_close_matches_in_order(self):
        result = closest_names("acent1", ["accent1", "accent2", "lt1", "dk2"])
        assert result[0] == "accent1"

    def it_returns_empty_when_nothing_close(self):
        result = closest_names("xyz", ["aaa", "bbb"])
        assert result == []


class DescribeErrorCatalog:
    def it_lists_every_emitted_code(self):
        # -- The catalog is the public-API gesture; ensure every code we --
        # -- exercise in this test module is enumerated. --
        emitted = {
            "STYLE_NOT_FOUND",
            "STYLE_DUPLICATE",
            "STYLE_TYPE_MISMATCH",
            "LATENT_STYLE_NOT_FOUND",
            "BUILTIN_STYLE_NOT_FOUND",
            "BOOKMARK_NOT_FOUND",
            "FONT_NOT_FOUND",
            "FONT_FAMILY_INVALID",
            "FONT_EMBED_EMPTY",
            "THEME_TOKEN_INVALID",
            "INVALID_COLOR",
            "INVALID_BRIGHTNESS",
            "BRIGHTNESS_NO_THEME",
        }
        catalog = set(__all_codes__)
        missing = emitted - catalog
        assert not missing, f"emitted codes not in catalog: {sorted(missing)}"
