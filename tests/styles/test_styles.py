"""Unit test suite for the docx.styles.styles module."""

import pytest

from docx.enum.style import WD_STYLE, WD_STYLE_TYPE
from docx.oxml.styles import CT_Style, CT_Styles
from docx.styles.latent import LatentStyles
from docx.styles.style import BaseStyle
from docx.styles.styles import Styles, _builtin_style_ui_name

from ..unitutil.cxml import element
from ..unitutil.mock import call, class_mock, function_mock, instance_mock, method_mock


class DescribeStyles:
    def it_supports_the_in_operator_on_style_name(self, in_fixture):
        styles, name, expected_value = in_fixture
        assert (name in styles) is expected_value

    def it_knows_its_length(self, len_fixture):
        styles, expected_value = len_fixture
        assert len(styles) == expected_value

    def it_can_iterate_over_its_styles(self, iter_fixture):
        styles, expected_count, style_, StyleFactory_, expected_calls = iter_fixture
        count = 0
        for style in styles:
            assert style is style_
            count += 1
        assert count == expected_count
        assert StyleFactory_.call_args_list == expected_calls

    @pytest.mark.filterwarnings("ignore::UserWarning")
    def it_can_get_a_style_by_id(self, getitem_id_fixture):
        styles, key, expected_element = getitem_id_fixture
        style = styles[key]
        assert style._element is expected_element

    def it_can_get_a_style_by_name(self, getitem_name_fixture):
        styles, key, expected_element = getitem_name_fixture
        style = styles[key]
        assert style._element is expected_element

    def it_can_get_a_style_by_WD_BUILTIN_STYLE_enum_member(self):
        # -- upstream#1439: `styles[WD_STYLE.BODY_TEXT]` used to raise KeyError --
        styles = Styles(
            element(
                "w:styles/w:style{w:type=paragraph}/w:name{w:val=Body Text}"
            )
        )
        expected_element = styles._element[0]

        style = styles[WD_STYLE.BODY_TEXT]

        assert style._element is expected_element

    def it_can_get_a_heading_style_by_WD_BUILTIN_STYLE_enum_member(self):
        styles = Styles(
            element(
                "w:styles/w:style{w:type=paragraph}/w:name{w:val=heading 1}"
            )
        )
        expected_element = styles._element[0]

        style = styles[WD_STYLE.HEADING_1]

        assert style._element is expected_element

    def it_raises_KeyError_for_missing_WD_BUILTIN_STYLE_member(self):
        styles = Styles(element("w:styles"))
        with pytest.raises(KeyError):
            styles[WD_STYLE.BODY_TEXT]

    def it_can_get_INDEX_HEADING_by_enum(self):
        # -- upstream#542: ``styles[WD_STYLE.INDEX_HEADING]`` used to raise
        # -- ``KeyError: "no style with name 'INDEX_HEADING (-34)'"`` because
        # -- ``str(enum)`` was fed straight into the name lookup.
        styles = Styles(
            element(
                "w:styles/w:style{w:type=paragraph}/w:name{w:val=Index Heading}"
            )
        )
        expected_element = styles._element[0]

        style = styles[WD_STYLE.INDEX_HEADING]

        assert style._element is expected_element

    def it_can_get_TOC_HEADING_by_enum(self):
        # -- upstream#542: ``WD_STYLE.TOC_HEADING`` was entirely missing. --
        styles = Styles(
            element(
                "w:styles/w:style{w:type=paragraph}/w:name{w:val=TOC Heading}"
            )
        )
        expected_element = styles._element[0]

        style = styles[WD_STYLE.TOC_HEADING]

        assert style._element is expected_element

    def it_raises_on_style_not_found(self, get_raises_fixture):
        styles, key = get_raises_fixture
        with pytest.raises(KeyError):
            styles[key]

    def it_can_add_a_new_style(self, add_fixture):
        styles, name, style_type, builtin = add_fixture[:4]
        name_, StyleFactory_, style_elm_, style_ = add_fixture[4:]

        style = styles.add_style(name, style_type, builtin)

        styles._element.add_style_of_type.assert_called_once_with(name_, style_type, builtin)
        StyleFactory_.assert_called_once_with(style_elm_)
        assert style is style_

    def it_raises_when_style_name_already_used(self, add_raises_fixture):
        styles, name = add_raises_fixture
        with pytest.raises(ValueError, match="document already contains style 'Hea"):
            styles.add_style(name, None)

    def it_can_get_the_default_style_for_a_type(self, default_fixture):
        styles, style_type, StyleFactory_ = default_fixture[:3]
        StyleFactory_calls, style_ = default_fixture[3:]

        style = styles.default(style_type)

        assert StyleFactory_.call_args_list == StyleFactory_calls
        assert style is style_

    def it_can_get_a_style_of_type_by_id(self, _get_by_id_, style_):
        style_id, style_type = 42, 7
        _get_by_id_.return_value = style_
        styles = Styles(None)

        style = styles.get_by_id(style_id, style_type)

        _get_by_id_.assert_called_once_with(styles, style_id, style_type)
        assert style is style_

    def but_it_returns_the_default_style_for_style_id_None(self, default_, style_):
        style_type = 17
        default_.return_value = style_
        styles = Styles(None)

        style = styles.get_by_id(None, style_type)

        default_.assert_called_once_with(styles, style_type)
        assert style is style_

    def it_can_get_a_style_id_from_a_style(self, _get_style_id_from_style_):
        style = BaseStyle(None)
        style_type = 22
        _get_style_id_from_style_.return_value = "StyleId"
        styles = Styles(None)

        style_id = styles.get_style_id(style, style_type)

        _get_style_id_from_style_.assert_called_once_with(styles, style, style_type)
        assert style_id == "StyleId"

    def and_it_can_get_a_style_id_from_a_style_name(self, _get_style_id_from_name_):
        style_type = 22
        _get_style_id_from_name_.return_value = "StyleId"
        styles = Styles(None)

        style_id = styles.get_style_id("Style Name", style_type)

        _get_style_id_from_name_.assert_called_once_with(styles, "Style Name", style_type)
        assert style_id == "StyleId"

    def but_it_returns_None_for_a_style_or_name_of_None(self):
        styles = Styles(None)

        style_id = styles.get_style_id(None, style_type=22)

        assert style_id is None

    def it_gets_a_style_by_id_to_help(self, _get_by_id_fixture):
        styles, style_id, style_type, default_calls = _get_by_id_fixture[:4]
        StyleFactory_, StyleFactory_calls, style_ = _get_by_id_fixture[4:]

        style = styles._get_by_id(style_id, style_type)

        assert styles.default.call_args_list == default_calls
        assert StyleFactory_.call_args_list == StyleFactory_calls
        assert style is style_

    def it_gets_a_style_id_from_a_name_to_help(self, _getitem_, _get_style_id_from_style_, style_):
        style_name, style_type, style_id_ = "Foo Bar", 1, "FooBar"
        _getitem_.return_value = style_
        _get_style_id_from_style_.return_value = style_id_
        styles = Styles(None)

        style_id = styles._get_style_id_from_name(style_name, style_type)

        styles.__getitem__.assert_called_once_with(styles, style_name)
        _get_style_id_from_style_.assert_called_once_with(styles, style_, style_type)
        assert style_id is style_id_

    def it_gets_a_style_id_from_a_style_to_help(self, id_style_fixture):
        styles, style_, style_type, style_id_ = id_style_fixture

        style_id = styles._get_style_id_from_style(style_, style_type)

        styles.default.assert_called_once_with(styles, style_type)
        assert style_id is style_id_

    def it_raises_on_style_type_mismatch(self, id_style_raises_fixture):
        styles, style_, style_type = id_style_raises_fixture
        with pytest.raises(ValueError, match="assigned style is type 1, need type 2"):
            styles._get_style_id_from_style(style_, style_type)

    def it_provides_access_to_the_latent_styles(self, latent_styles_fixture):
        styles, LatentStyles_, latent_styles_ = latent_styles_fixture
        latent_styles = styles.latent_styles
        LatentStyles_.assert_called_once_with(styles._element.latentStyles)
        assert latent_styles is latent_styles_

    # fixture --------------------------------------------------------

    @pytest.fixture(
        params=[
            ("Foo Bar", "Foo Bar", WD_STYLE_TYPE.CHARACTER, False),
            ("Heading 1", "heading 1", WD_STYLE_TYPE.PARAGRAPH, True),
        ]
    )
    def add_fixture(self, request, styles_elm_, _getitem_, style_elm_, StyleFactory_, style_):
        name, name_, style_type, builtin = request.param
        styles = Styles(styles_elm_)
        _getitem_.return_value = None
        styles_elm_.add_style_of_type.return_value = style_elm_
        StyleFactory_.return_value = style_
        return (
            styles,
            name,
            style_type,
            builtin,
            name_,
            StyleFactory_,
            style_elm_,
            style_,
        )

    @pytest.fixture
    def add_raises_fixture(self, _getitem_):
        styles = Styles(element("w:styles/w:style/w:name{w:val=heading 1}"))
        name = "Heading 1"
        return styles, name

    @pytest.fixture(
        params=[
            ("w:styles", False, WD_STYLE_TYPE.CHARACTER),
            (
                "w:styles/w:style{w:type=paragraph,w:default=1}",
                True,
                WD_STYLE_TYPE.PARAGRAPH,
            ),
            (
                "w:styles/(w:style{w:type=table,w:default=1},w:style{w:type=table,w:default=1})",
                True,
                WD_STYLE_TYPE.TABLE,
            ),
        ]
    )
    def default_fixture(self, request, StyleFactory_, style_):
        styles_cxml, is_defined, style_type = request.param
        styles_elm = element(styles_cxml)
        styles = Styles(styles_elm)
        StyleFactory_calls = [call(styles_elm[-1])] if is_defined else []
        StyleFactory_.return_value = style_
        expected_value = style_ if is_defined else None
        return (styles, style_type, StyleFactory_, StyleFactory_calls, expected_value)

    @pytest.fixture(
        params=[
            (
                "w:styles/w:style{w:type=paragraph,w:styleId=Foo}",
                "Foo",
                WD_STYLE_TYPE.PARAGRAPH,
            ),
            (
                "w:styles/w:style{w:type=paragraph,w:styleId=Foo}",
                "Bar",
                WD_STYLE_TYPE.PARAGRAPH,
            ),
            (
                "w:styles/w:style{w:type=table,w:styleId=Bar}",
                "Bar",
                WD_STYLE_TYPE.PARAGRAPH,
            ),
        ]
    )
    def _get_by_id_fixture(self, request, default_, StyleFactory_, style_):
        styles_cxml, style_id, style_type = request.param
        styles_elm = element(styles_cxml)
        style_elm = styles_elm[0]
        styles = Styles(styles_elm)
        default_calls = [] if style_id == "Foo" else [call(styles, style_type)]
        StyleFactory_calls = [call(style_elm)] if style_id == "Foo" else []
        default_.return_value = StyleFactory_.return_value = style_
        return (
            styles,
            style_id,
            style_type,
            default_calls,
            StyleFactory_,
            StyleFactory_calls,
            style_,
        )

    @pytest.fixture(
        params=[
            ("w:styles/(w:style{%s,w:styleId=Foobar},w:style,w:style)", 0),
            ("w:styles/(w:style,w:style{%s,w:styleId=Foobar},w:style)", 1),
            ("w:styles/(w:style,w:style,w:style{%s,w:styleId=Foobar})", 2),
        ]
    )
    def getitem_id_fixture(self, request):
        styles_cxml_tmpl, style_idx = request.param
        styles_cxml = styles_cxml_tmpl % "w:type=paragraph"
        styles = Styles(element(styles_cxml))
        expected_element = styles._element[style_idx]
        return styles, "Foobar", expected_element

    @pytest.fixture(
        params=[
            ("w:styles/(w:style%s/w:name{w:val=foo},w:style)", "foo", 0),
            ("w:styles/(w:style,w:style%s/w:name{w:val=foo})", "foo", 1),
            ("w:styles/w:style%s/w:name{w:val=heading 1}", "Heading 1", 0),
        ]
    )
    def getitem_name_fixture(self, request):
        styles_cxml_tmpl, key, style_idx = request.param
        styles_cxml = styles_cxml_tmpl % "{w:type=character}"
        styles = Styles(element(styles_cxml))
        expected_element = styles._element[style_idx]
        return styles, key, expected_element

    @pytest.fixture(
        params=[
            ("w:styles/(w:style,w:style/w:name{w:val=foo},w:style)"),
            ("w:styles/(w:style{w:styleId=foo},w:style,w:style)"),
        ]
    )
    def get_raises_fixture(self, request):
        styles_cxml = request.param
        styles = Styles(element(styles_cxml))
        return styles, "bar"

    @pytest.fixture(params=[True, False])
    def id_style_fixture(self, request, default_, style_):
        style_is_default = request.param
        styles = Styles(None)
        style_id, style_type = "FooBar", 1
        default_.return_value = style_ if style_is_default else None
        style_.style_id, style_.type = style_id, style_type
        expected_value = None if style_is_default else style_id
        return styles, style_, style_type, expected_value

    @pytest.fixture
    def id_style_raises_fixture(self, style_):
        styles = Styles(None)
        style_.type = 1
        style_type = 2
        return styles, style_, style_type

    @pytest.fixture(
        params=[
            ("w:styles/w:style/w:name{w:val=heading 1}", "Heading 1", True),
            ("w:styles/w:style/w:name{w:val=Foo Bar}", "Foo Bar", True),
            ("w:styles/w:style/w:name{w:val=heading 1}", "Foobar", False),
            ("w:styles", "Foobar", False),
        ]
    )
    def in_fixture(self, request):
        styles_cxml, name, expected_value = request.param
        styles = Styles(element(styles_cxml))
        return styles, name, expected_value

    @pytest.fixture(
        params=[
            ("w:styles", 0),
            ("w:styles/w:style", 1),
            ("w:styles/(w:style,w:style)", 2),
            ("w:styles/(w:style,w:style,w:style)", 3),
        ]
    )
    def iter_fixture(self, request, StyleFactory_, style_):
        styles_cxml, expected_count = request.param
        styles_elm = element(styles_cxml)
        styles = Styles(styles_elm)
        expected_calls = [call(style_elm) for style_elm in styles_elm]
        StyleFactory_.return_value = style_
        return styles, expected_count, style_, StyleFactory_, expected_calls

    @pytest.fixture
    def latent_styles_fixture(self, LatentStyles_, latent_styles_):
        styles = Styles(element("w:styles/w:latentStyles"))
        return styles, LatentStyles_, latent_styles_

    @pytest.fixture(
        params=[
            ("w:styles", 0),
            ("w:styles/w:style", 1),
            ("w:styles/(w:style,w:style)", 2),
            ("w:styles/(w:style,w:style,w:style)", 3),
        ]
    )
    def len_fixture(self, request):
        styles_cxml, expected_value = request.param
        styles = Styles(element(styles_cxml))
        return styles, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def default_(self, request):
        return method_mock(request, Styles, "default")

    @pytest.fixture
    def _get_by_id_(self, request):
        return method_mock(request, Styles, "_get_by_id")

    @pytest.fixture
    def _getitem_(self, request):
        return method_mock(request, Styles, "__getitem__")

    @pytest.fixture
    def _get_style_id_from_name_(self, request):
        return method_mock(request, Styles, "_get_style_id_from_name")

    @pytest.fixture
    def _get_style_id_from_style_(self, request):
        return method_mock(request, Styles, "_get_style_id_from_style")

    @pytest.fixture
    def LatentStyles_(self, request, latent_styles_):
        return class_mock(request, "docx.styles.styles.LatentStyles", return_value=latent_styles_)

    @pytest.fixture
    def latent_styles_(self, request):
        return instance_mock(request, LatentStyles)

    @pytest.fixture
    def style_(self, request):
        return instance_mock(request, BaseStyle)

    @pytest.fixture
    def StyleFactory_(self, request):
        return function_mock(request, "docx.styles.styles.StyleFactory")

    @pytest.fixture
    def style_elm_(self, request):
        return instance_mock(request, CT_Style)

    @pytest.fixture
    def styles_elm_(self, request):
        return instance_mock(request, CT_Styles)


class DescribeStyles_CaseInsensitiveLookup:
    """Phase A-v2 #6: Styles.__getitem__ case-insensitive fallback.

    See upstream #494, #420, PR#239. LibreOffice saves built-in style names
    in lower-case on disk; the UI name lookup must still resolve them.
    """

    def it_looks_up_a_style_by_exact_ui_name(self):
        styles = Styles(
            element(
                "w:styles/w:style{w:type=paragraph}/w:name{w:val=Heading 1}"
            )
        )
        expected_element = styles._element[0]

        # -- exact match: on-disk name "Heading 1" matches "Heading 1" --
        style = styles["Heading 1"]
        assert style._element is expected_element

    def it_finds_a_libreoffice_cased_style_by_its_ui_name(self):
        # -- upstream#494: LibreOffice saves "Heading 1" as "Heading 1"
        # -- with lower-case variants like "heading 1"; the fallback should
        # -- also match oddly-cased on-disk names. --
        styles = Styles(
            element(
                "w:styles/w:style{w:type=paragraph}/w:name{w:val=HEADING 1}"
            )
        )
        expected_element = styles._element[0]

        style = styles["Heading 1"]
        assert style._element is expected_element

    def it_case_insensitively_matches_custom_names(self):
        styles = Styles(
            element(
                "w:styles/w:style{w:type=paragraph}/w:name{w:val=My Style}"
            )
        )
        expected_element = styles._element[0]

        style = styles["MY STYLE"]
        assert style._element is expected_element

    def it_supports_in_operator_case_insensitively(self):
        styles = Styles(
            element(
                "w:styles/w:style{w:type=paragraph}/w:name{w:val=HEADING 1}"
            )
        )
        assert "Heading 1" in styles

    def it_can_lookup_various_WD_BUILTIN_STYLE_members(self):
        # -- upstream#420: full WD_BUILTIN_STYLE → UI-name translation. --
        styles = Styles(
            element(
                "w:styles/"
                "(w:style{w:type=paragraph}/w:name{w:val=Caption},"
                "w:style{w:type=paragraph}/w:name{w:val=Body Text 2},"
                "w:style{w:type=character}/w:name{w:val=Emphasis})"
            )
        )

        assert styles[WD_STYLE.CAPTION]._element is styles._element[0]
        assert styles[WD_STYLE.BODY_TEXT_2]._element is styles._element[1]
        assert styles[WD_STYLE.EMPHASIS]._element is styles._element[2]


class DescribeBuiltinStyleUiName:
    """Unit-test suite for `_builtin_style_ui_name` helper."""

    def it_returns_Body_Text_for_BODY_TEXT(self):
        assert _builtin_style_ui_name(WD_STYLE.BODY_TEXT) == "Body Text"

    def it_returns_Heading_1_for_HEADING_1(self):
        assert _builtin_style_ui_name(WD_STYLE.HEADING_1) == "Heading 1"

    def it_returns_Normal_for_NORMAL(self):
        assert _builtin_style_ui_name(WD_STYLE.NORMAL) == "Normal"

    def it_returns_Index_Heading_for_INDEX_HEADING(self):
        # -- upstream#542 --
        assert _builtin_style_ui_name(WD_STYLE.INDEX_HEADING) == "Index Heading"

    def it_returns_TOC_Heading_for_TOC_HEADING(self):
        # -- upstream#542: TOC_HEADING was previously missing from the enum --
        assert _builtin_style_ui_name(WD_STYLE.TOC_HEADING) == "TOC Heading"


class DescribeBabelFish:
    """Unit-test suite for `docx.styles.BabelFish`."""

    def it_translates_WD_BUILTIN_STYLE_members_to_UI_names(self):
        from docx.styles import BabelFish

        assert BabelFish.enum2ui(WD_STYLE.BODY_TEXT) == "Body Text"
        assert BabelFish.enum2ui(WD_STYLE.HEADING_1) == "Heading 1"
        assert BabelFish.enum2ui(WD_STYLE.INDEX_HEADING) == "Index Heading"
        assert BabelFish.enum2ui(WD_STYLE.TOC_HEADING) == "TOC Heading"
