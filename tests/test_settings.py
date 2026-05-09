# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.settings module."""

from __future__ import annotations

import warnings

import pytest

from docx.endnotes import EndnoteProperties
from docx.enum.text import (
    WD_ENDNOTE_POSITION,
    WD_FOOTNOTE_POSITION,
    WD_NUMBER_FORMAT,
    WD_PROTECTION,
    WD_VIEW,
)
from docx.footnotes import FootnoteProperties
from docx.settings import (
    CompatFlags,
    CompatSettings,
    DocumentProtection,
    RsidList,
    Settings,
    WriteProtection,
)
from docx.shared import Twips

from .unitutil.cxml import element, xml


class DescribeSettings:
    """Unit-test suite for the `docx.settings.Settings` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", False),
            ("w:settings/w:evenAndOddHeaders", True),
            ("w:settings/w:evenAndOddHeaders{w:val=0}", False),
            ("w:settings/w:evenAndOddHeaders{w:val=1}", True),
            ("w:settings/w:evenAndOddHeaders{w:val=true}", True),
        ],
    )
    def it_knows_when_the_document_has_distinct_odd_and_even_headers(
        self, cxml: str, expected_value: bool
    ):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", DeprecationWarning)
            assert Settings(element(cxml)).odd_and_even_pages_header_footer is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", True, "w:settings/w:evenAndOddHeaders"),
            ("w:settings/w:evenAndOddHeaders", False, "w:settings"),
            ("w:settings/w:evenAndOddHeaders{w:val=1}", True, "w:settings/w:evenAndOddHeaders"),
            ("w:settings/w:evenAndOddHeaders{w:val=off}", False, "w:settings"),
        ],
    )
    def it_can_change_whether_the_document_has_distinct_odd_and_even_headers(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        settings = Settings(element(cxml))

        with warnings.catch_warnings():
            warnings.simplefilter("ignore", DeprecationWarning)
            settings.odd_and_even_pages_header_footer = new_value

        assert settings._settings.xml == xml(expected_cxml)

    def it_emits_deprecation_warning_for_odd_and_even_pages_header_footer(self):
        settings = Settings(element("w:settings/w:evenAndOddHeaders"))
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            settings.odd_and_even_pages_header_footer
            assert len(w) == 1
            assert issubclass(w[0].category, DeprecationWarning)
            assert "even_and_odd_headers" in str(w[0].message)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", False),
            ("w:settings/w:evenAndOddHeaders", True),
            ("w:settings/w:evenAndOddHeaders{w:val=0}", False),
        ],
    )
    def it_provides_even_and_odd_headers(self, cxml: str, expected_value: bool):
        assert Settings(element(cxml)).even_and_odd_headers is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", True, "w:settings/w:evenAndOddHeaders"),
            ("w:settings/w:evenAndOddHeaders", False, "w:settings"),
        ],
    )
    def it_can_change_even_and_odd_headers(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        settings = Settings(element(cxml))
        settings.even_and_odd_headers = new_value
        assert settings._settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", None),
            ("w:settings/w:zoom{w:percent=100}", 100),
            ("w:settings/w:zoom{w:percent=75}", 75),
        ],
    )
    def it_can_get_the_zoom_percent(self, cxml: str, expected_value: int | None):
        assert Settings(element(cxml)).zoom_percent == expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", 100, "w:settings/w:zoom{w:percent=100}"),
            ("w:settings/w:zoom{w:percent=75}", 150, "w:settings/w:zoom{w:percent=150}"),
            ("w:settings/w:zoom{w:percent=100}", None, "w:settings"),
        ],
    )
    def it_can_set_the_zoom_percent(
        self, cxml: str, new_value: int | None, expected_cxml: str
    ):
        settings = Settings(element(cxml))
        settings.zoom_percent = new_value
        assert settings._settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", False),
            ("w:settings/w:trackRevisions", True),
            ("w:settings/w:trackRevisions{w:val=0}", False),
        ],
    )
    def it_can_get_track_revisions(self, cxml: str, expected_value: bool):
        assert Settings(element(cxml)).track_revisions is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", True, "w:settings/w:trackRevisions"),
            ("w:settings/w:trackRevisions", False, "w:settings"),
        ],
    )
    def it_can_set_track_revisions(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        settings = Settings(element(cxml))
        settings.track_revisions = new_value
        assert settings._settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", False),
            ("w:settings/w:updateFields", True),
            ("w:settings/w:updateFields{w:val=0}", False),
            ("w:settings/w:updateFields{w:val=true}", True),
        ],
    )
    def it_can_get_update_fields_on_open(self, cxml: str, expected_value: bool):
        assert Settings(element(cxml)).update_fields_on_open is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", True, "w:settings/w:updateFields"),
            ("w:settings/w:updateFields", False, "w:settings"),
            ("w:settings/w:trackRevisions", True, "w:settings/(w:trackRevisions,w:updateFields)"),
        ],
    )
    def it_can_set_update_fields_on_open(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        settings = Settings(element(cxml))
        settings.update_fields_on_open = new_value
        assert settings._settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", None),
            ("w:settings/w:defaultTabStop{w:val=720}", Twips(720)),
        ],
    )
    def it_can_get_the_default_tab_stop(self, cxml: str, expected_value):
        assert Settings(element(cxml)).default_tab_stop == expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", Twips(720), "w:settings/w:defaultTabStop{w:val=720}"),
            ("w:settings/w:defaultTabStop{w:val=720}", None, "w:settings"),
        ],
    )
    def it_can_set_the_default_tab_stop(
        self, cxml: str, new_value, expected_cxml: str
    ):
        settings = Settings(element(cxml))
        settings.default_tab_stop = new_value
        assert settings._settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_type", "expected_enabled"),
        [
            ("w:settings", None, False),
            (
                "w:settings/w:documentProtection{w:edit=readOnly,w:enforcement=1}",
                "readOnly",
                True,
            ),
            (
                "w:settings/w:documentProtection{w:edit=comments,w:enforcement=0}",
                "comments",
                False,
            ),
        ],
    )
    def it_can_get_document_protection(
        self,
        cxml: str,
        expected_type: str | None,
        expected_enabled: bool,
    ):
        protection = Settings(element(cxml)).document_protection
        assert protection.type == expected_type
        assert protection.enabled is expected_enabled

    def it_can_get_the_compatibility_mode(self):
        settings = Settings(element("w:settings"))
        assert settings.compatibility_mode is None

    def it_can_set_the_compatibility_mode(self):
        settings = Settings(element("w:settings"))
        settings.compatibility_mode = 15
        assert settings.compatibility_mode == 15

    def it_can_remove_the_compatibility_mode(self):
        settings = Settings(element("w:settings"))
        settings.compatibility_mode = 15
        settings.compatibility_mode = None
        assert settings.compatibility_mode is None

    def it_returns_None_when_no_footnote_properties_present(self):
        settings = Settings(element("w:settings"))
        assert settings.footnote_properties is None

    def it_returns_a_FootnoteProperties_when_footnotePr_is_present(self):
        settings = Settings(element("w:settings/w:footnotePr"))
        props = settings.footnote_properties
        assert isinstance(props, FootnoteProperties)

    def it_can_add_footnote_properties(self):
        settings = Settings(element("w:settings"))

        props = settings.add_footnote_properties()

        assert isinstance(props, FootnoteProperties)
        assert settings._settings.xml == xml("w:settings/w:footnotePr")

    def add_footnote_properties_returns_existing_element_when_present(self):
        settings = Settings(element("w:settings/w:footnotePr/w:numFmt{w:val=chicago}"))
        props = settings.add_footnote_properties()
        assert props.number_format == WD_NUMBER_FORMAT.CHICAGO

    def it_can_remove_footnote_properties(self):
        settings = Settings(element("w:settings/w:footnotePr/w:pos{w:val=pageBottom}"))
        settings.remove_footnote_properties()
        assert settings._settings.xml == xml("w:settings")

    def it_round_trips_footnote_properties_through_settings(self):
        settings = Settings(element("w:settings"))
        props = settings.add_footnote_properties()

        props.number_format = WD_NUMBER_FORMAT.LOWER_ROMAN
        props.start_number = 1
        props.position = WD_FOOTNOTE_POSITION.BENEATH_TEXT

        assert settings.footnote_properties is not None
        assert settings.footnote_properties.number_format == WD_NUMBER_FORMAT.LOWER_ROMAN
        assert settings.footnote_properties.start_number == 1
        assert (
            settings.footnote_properties.position == WD_FOOTNOTE_POSITION.BENEATH_TEXT
        )

    def it_returns_None_when_no_endnote_properties_present(self):
        settings = Settings(element("w:settings"))
        assert settings.endnote_properties is None

    def it_returns_an_EndnoteProperties_when_endnotePr_is_present(self):
        settings = Settings(element("w:settings/w:endnotePr"))
        props = settings.endnote_properties
        assert isinstance(props, EndnoteProperties)

    def it_can_add_endnote_properties(self):
        settings = Settings(element("w:settings"))

        props = settings.add_endnote_properties()

        assert isinstance(props, EndnoteProperties)
        assert settings._settings.xml == xml("w:settings/w:endnotePr")

    def it_can_remove_endnote_properties(self):
        settings = Settings(element("w:settings/w:endnotePr/w:pos{w:val=docEnd}"))
        settings.remove_endnote_properties()
        assert settings._settings.xml == xml("w:settings")

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", None),
            ("w:settings/w:view{w:val=normal}", WD_VIEW.NORMAL),
            ("w:settings/w:view{w:val=outline}", WD_VIEW.OUTLINE),
            ("w:settings/w:view{w:val=print}", WD_VIEW.PRINT),
            ("w:settings/w:view{w:val=web}", WD_VIEW.WEB),
            ("w:settings/w:view{w:val=reading}", WD_VIEW.READING),
            ("w:settings/w:view{w:val=masterPages}", WD_VIEW.MASTER_PAGES),
            ("w:settings/w:view{w:val=none}", WD_VIEW.NONE),
        ],
    )
    def it_can_get_the_view(self, cxml: str, expected_value: WD_VIEW | None):
        assert Settings(element(cxml)).view == expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", WD_VIEW.PRINT, "w:settings/w:view{w:val=print}"),
            ("w:settings", WD_VIEW.OUTLINE, "w:settings/w:view{w:val=outline}"),
            (
                "w:settings/w:view{w:val=normal}",
                WD_VIEW.WEB,
                "w:settings/w:view{w:val=web}",
            ),
            ("w:settings/w:view{w:val=print}", None, "w:settings"),
        ],
    )
    def it_can_set_the_view(
        self, cxml: str, new_value: WD_VIEW | None, expected_cxml: str
    ):
        settings = Settings(element(cxml))
        settings.view = new_value
        assert settings._settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        "member",
        [
            WD_VIEW.NONE,
            WD_VIEW.PRINT,
            WD_VIEW.OUTLINE,
            WD_VIEW.MASTER_PAGES,
            WD_VIEW.NORMAL,
            WD_VIEW.WEB,
            WD_VIEW.READING,
        ],
    )
    def it_round_trips_every_view_mode(self, member: WD_VIEW):
        settings = Settings(element("w:settings"))
        settings.view = member
        assert settings.view == member

    def it_returns_None_for_view_when_no_view_child_present(self):
        assert Settings(element("w:settings")).view is None

    def it_removes_the_view_child_when_view_set_to_None(self):
        settings = Settings(element("w:settings/w:view{w:val=print}"))
        settings.view = None
        assert settings.view is None
        assert settings._settings.xml == xml("w:settings")

    def it_round_trips_endnote_properties_through_settings(self):
        settings = Settings(element("w:settings"))
        props = settings.add_endnote_properties()

        props.number_format = WD_NUMBER_FORMAT.UPPER_LETTER
        props.position = WD_ENDNOTE_POSITION.END_OF_SECTION

        assert settings.endnote_properties is not None
        assert settings.endnote_properties.number_format == WD_NUMBER_FORMAT.UPPER_LETTER
        assert settings.endnote_properties.position == WD_ENDNOTE_POSITION.END_OF_SECTION


class DescribeSettings_Rsids:
    """Unit-test suite for RSID access on `docx.settings.Settings`."""

    # -- rsid_root ----------------------------------------------------------

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", None),
            ("w:settings/w:rsids", None),
            ("w:settings/w:rsids/w:rsidRoot", None),
            ("w:settings/w:rsids/w:rsidRoot{w:val=00FA1B42}", "00FA1B42"),
            (
                "w:settings/w:rsids/("
                "w:rsidRoot{w:val=00ABCDEF},"
                "w:rsid{w:val=001234AB})",
                "00ABCDEF",
            ),
        ],
    )
    def it_can_get_the_rsid_root(self, cxml: str, expected_value: str | None):
        assert Settings(element(cxml)).rsid_root == expected_value

    # -- rsids --------------------------------------------------------------

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", []),
            ("w:settings/w:rsids", []),
            ("w:settings/w:rsids/w:rsidRoot{w:val=00FA1B42}", []),
            (
                "w:settings/w:rsids/w:rsid{w:val=001234AB}",
                ["001234AB"],
            ),
            (
                "w:settings/w:rsids/("
                "w:rsidRoot{w:val=00FA1B42},"
                "w:rsid{w:val=001234AB},"
                "w:rsid{w:val=00567890})",
                ["001234AB", "00567890"],
            ),
        ],
    )
    def it_can_get_the_rsids_in_document_order(
        self, cxml: str, expected_value: list[str]
    ):
        assert Settings(element(cxml)).rsids == expected_value


class DescribeRsidList:
    """Unit-test suite for the `docx.settings.RsidList` proxy."""

    def it_is_a_list_subclass_for_backward_compat(self):
        rsids = Settings(element("w:settings")).rsids
        assert isinstance(rsids, RsidList)
        assert isinstance(rsids, list)
        assert rsids == []

    def it_exposes_root_as_None_when_absent(self):
        rsids = Settings(element("w:settings/w:rsids")).rsids
        assert rsids.root is None

    def it_exposes_root_from_rsidRoot(self):
        rsids = Settings(
            element("w:settings/w:rsids/w:rsidRoot{w:val=00CAFE00}")
        ).rsids
        assert rsids.root == "00CAFE00"

    def it_returns_empty_ids_set_when_no_rsids(self):
        rsids = Settings(element("w:settings")).rsids
        assert rsids.ids == set()

    def it_returns_ids_as_set_of_root_and_children(self):
        rsids = Settings(
            element(
                "w:settings/w:rsids/("
                "w:rsidRoot{w:val=00CAFE00},"
                "w:rsid{w:val=001234AB},"
                "w:rsid{w:val=00567890})"
            )
        ).rsids
        assert rsids.ids == {"00CAFE00", "001234AB", "00567890"}

    def it_ids_is_a_set_and_supports_constant_time_membership(self):
        rsids = Settings(
            element("w:settings/w:rsids/w:rsid{w:val=001234AB}")
        ).rsids
        ids = rsids.ids
        assert isinstance(ids, set)
        assert "001234AB" in ids
        assert "deadbeef" not in ids

    def it_can_add_a_new_session_rsid(self):
        settings = Settings(element("w:settings"))
        token = settings.rsids.new_session()

        assert isinstance(token, str)
        assert len(token) == 8
        # -- materialised ``w:rsids`` with root + rsid on first call
        fresh = settings.rsids
        assert fresh.root == token
        assert token in fresh.ids
        assert token in fresh

    def it_new_session_mints_unique_tokens(self):
        settings = Settings(element("w:settings"))
        a = settings.rsids.new_session()
        b = settings.rsids.new_session()
        assert a != b
        # -- root is fixed to the first-ever session, second call only appends
        assert settings.rsids.root == a
        assert {a, b}.issubset(settings.rsids.ids)

    def it_can_add_a_caller_supplied_rsid(self):
        settings = Settings(element("w:settings"))
        settings.rsids.add("00ABCDEF")
        assert "00ABCDEF" in settings.rsids.ids
        # -- idempotent: calling again does not duplicate
        settings.rsids.add("00ABCDEF")
        assert settings.rsids.count("00ABCDEF") == 1


class DescribeCompatSettings:
    """Unit-test suite for `docx.settings.CompatSettings`."""

    def it_is_exposed_on_Settings(self):
        settings = Settings(element("w:settings"))
        assert isinstance(settings.compat_settings, CompatSettings)

    def it_is_empty_when_no_compat_child_present(self):
        settings = Settings(element("w:settings"))
        assert len(settings.compat_settings) == 0
        assert list(settings.compat_settings) == []
        assert "x" not in settings.compat_settings

    def it_reports_len_and_iter(self):
        settings = Settings(
            element(
                "w:settings/w:compat/("
                "w:compatSetting{w:name=a,w:uri=http://x,w:val=1},"
                "w:compatSetting{w:name=b,w:uri=http://x,w:val=2})"
            )
        )
        assert len(settings.compat_settings) == 2
        assert list(settings.compat_settings) == ["a", "b"]

    def it_supports_contains(self):
        settings = Settings(
            element(
                "w:settings/w:compat/w:compatSetting"
                "{w:name=compatibilityMode,w:uri=http://x,w:val=15}"
            )
        )
        assert "compatibilityMode" in settings.compat_settings
        assert "other" not in settings.compat_settings

    def it_returns_the_value_via_getitem(self):
        settings = Settings(
            element(
                "w:settings/w:compat/w:compatSetting"
                "{w:name=compatibilityMode,w:uri=http://x,w:val=15}"
            )
        )
        assert settings.compat_settings["compatibilityMode"] == "15"

    def it_raises_KeyError_for_missing_name(self):
        settings = Settings(element("w:settings"))
        with pytest.raises(KeyError):
            settings.compat_settings["missing"]

    def it_supports_get_with_default(self):
        settings = Settings(
            element(
                "w:settings/w:compat/w:compatSetting"
                "{w:name=foo,w:uri=http://x,w:val=1}"
            )
        )
        assert settings.compat_settings.get("foo") == "1"
        assert settings.compat_settings.get("bar") is None
        assert settings.compat_settings.get("bar", "fallback") == "fallback"

    def it_creates_w_compat_on_first_set(self):
        settings = Settings(element("w:settings"))
        settings.compat_settings["compatibilityMode"] = "15"
        assert settings.compat_settings["compatibilityMode"] == "15"
        assert "compatibilityMode" in settings.compat_settings

    def it_updates_an_existing_setting_in_place(self):
        settings = Settings(
            element(
                "w:settings/w:compat/w:compatSetting"
                "{w:name=compatibilityMode,w:uri=http://x,w:val=14}"
            )
        )
        settings.compat_settings["compatibilityMode"] = "15"
        assert settings.compat_settings["compatibilityMode"] == "15"
        assert len(settings.compat_settings) == 1

    def it_can_remove_via_delitem(self):
        settings = Settings(
            element(
                "w:settings/w:compat/w:compatSetting"
                "{w:name=foo,w:uri=http://x,w:val=1}"
            )
        )
        del settings.compat_settings["foo"]
        assert "foo" not in settings.compat_settings
        # -- w:compat child is pruned once empty --
        assert settings._settings.xml == xml("w:settings")

    def it_can_remove_via_remove_method(self):
        settings = Settings(
            element(
                "w:settings/w:compat/("
                "w:compatSetting{w:name=a,w:uri=http://x,w:val=1},"
                "w:compatSetting{w:name=b,w:uri=http://x,w:val=2})"
            )
        )
        settings.compat_settings.remove("a")
        assert list(settings.compat_settings) == ["b"]

    def it_raises_KeyError_when_removing_missing_name(self):
        settings = Settings(element("w:settings"))
        with pytest.raises(KeyError):
            settings.compat_settings.remove("missing")

    def it_coerces_set_values_to_str(self):
        settings = Settings(element("w:settings"))
        settings.compat_settings["compatibilityMode"] = 15  # type: ignore[assignment]
        assert settings.compat_settings["compatibilityMode"] == "15"


class DescribeCompatFlags:
    """Unit-test suite for `docx.settings.CompatFlags`."""

    def it_is_exposed_on_Settings(self):
        settings = Settings(element("w:settings"))
        assert isinstance(settings.compat_flags, CompatFlags)

    def it_returns_False_for_missing_flag_without_raising(self):
        settings = Settings(element("w:settings"))
        assert settings.compat_flags["growAutofit"] is False

    def it_returns_True_when_flag_element_is_present(self):
        settings = Settings(element("w:settings/w:compat/w:growAutofit"))
        assert settings.compat_flags["growAutofit"] is True

    def it_creates_the_flag_element_when_set_True(self):
        settings = Settings(element("w:settings"))
        settings.compat_flags["growAutofit"] = True
        assert settings.compat_flags["growAutofit"] is True
        assert settings._settings.xml == xml(
            "w:settings/w:compat/w:growAutofit"
        )

    def it_is_idempotent_when_setting_True_twice(self):
        settings = Settings(element("w:settings"))
        settings.compat_flags["growAutofit"] = True
        settings.compat_flags["growAutofit"] = True
        assert len(settings.compat_flags) == 1

    def it_removes_the_element_when_set_False(self):
        settings = Settings(element("w:settings/w:compat/w:growAutofit"))
        settings.compat_flags["growAutofit"] = False
        assert settings.compat_flags["growAutofit"] is False
        # -- empty w:compat is pruned --
        assert settings._settings.xml == xml("w:settings")

    def it_tolerates_set_False_for_missing_flag(self):
        settings = Settings(element("w:settings"))
        settings.compat_flags["growAutofit"] = False
        assert settings._settings.xml == xml("w:settings")

    def it_supports_contains(self):
        settings = Settings(element("w:settings/w:compat/w:growAutofit"))
        assert "growAutofit" in settings.compat_flags
        assert "useFELayout" not in settings.compat_flags

    def it_iterates_present_flag_names(self):
        settings = Settings(
            element(
                "w:settings/w:compat/("
                "w:growAutofit,w:useFELayout,"
                "w:compatSetting{w:name=n,w:uri=http://x,w:val=1})"
            )
        )
        assert list(settings.compat_flags) == ["growAutofit", "useFELayout"]

    def it_reports_len_as_number_of_present_flags(self):
        settings = Settings(
            element(
                "w:settings/w:compat/("
                "w:growAutofit,w:useFELayout,"
                "w:compatSetting{w:name=n,w:uri=http://x,w:val=1})"
            )
        )
        assert len(settings.compat_flags) == 2

    def it_can_clear_all_flags_and_prunes_empty_compat(self):
        settings = Settings(
            element(
                "w:settings/w:compat/(w:growAutofit,w:useFELayout)"
            )
        )
        settings.compat_flags.clear()
        assert len(settings.compat_flags) == 0
        assert settings._settings.xml == xml("w:settings")

    def it_preserves_compat_settings_when_clearing_flags(self):
        settings = Settings(
            element(
                "w:settings/w:compat/("
                "w:growAutofit,"
                "w:compatSetting{w:name=n,w:uri=http://x,w:val=1})"
            )
        )
        settings.compat_flags.clear()
        assert len(settings.compat_flags) == 0
        assert "n" in settings.compat_settings

    def it_accepts_unknown_flag_names(self):
        settings = Settings(element("w:settings"))
        # -- any local name gets the w: prefix applied --
        settings.compat_flags["someCustomFlag"] = True
        assert settings.compat_flags["someCustomFlag"] is True
        assert settings._settings.xml == xml(
            "w:settings/w:compat/w:someCustomFlag"
        )

    def it_can_delete_a_present_flag_via_delitem(self):
        settings = Settings(element("w:settings/w:compat/w:growAutofit"))
        del settings.compat_flags["growAutofit"]
        assert "growAutofit" not in settings.compat_flags
        assert settings._settings.xml == xml("w:settings")

    def it_raises_KeyError_when_deleting_absent_flag(self):
        settings = Settings(element("w:settings"))
        with pytest.raises(KeyError):
            del settings.compat_flags["growAutofit"]

    def it_exposes_a_list_of_known_flag_names(self):
        names = CompatFlags.names()
        assert isinstance(names, tuple)
        assert "growAutofit" in names
        assert "doNotBreakWrappedTables" in names
        assert "cachedColBalance" in names
        # -- a reasonable coverage threshold --
        assert len(names) >= 50


class DescribeDocumentProtection:
    """Unit-test suite for `docx.settings.DocumentProtection`."""

    # -- backward-compat: legacy read-only API still works -----------------

    def it_exposes_legacy_type_attribute(self):
        settings = Settings(
            element("w:settings/w:documentProtection{w:edit=readOnly,w:enforcement=1}")
        )
        assert settings.document_protection.type == "readOnly"

    def it_exposes_legacy_enabled_attribute(self):
        settings = Settings(
            element("w:settings/w:documentProtection{w:edit=readOnly,w:enforcement=1}")
        )
        assert settings.document_protection.enabled is True

    # -- mode / enforce ----------------------------------------------------

    @pytest.mark.parametrize(
        ("cxml", "expected_mode"),
        [
            ("w:settings", None),
            ("w:settings/w:documentProtection", None),
            (
                "w:settings/w:documentProtection{w:edit=readOnly}",
                WD_PROTECTION.READ_ONLY,
            ),
            (
                "w:settings/w:documentProtection{w:edit=comments}",
                WD_PROTECTION.COMMENTS,
            ),
            (
                "w:settings/w:documentProtection{w:edit=trackedChanges}",
                WD_PROTECTION.TRACKED_CHANGES,
            ),
            (
                "w:settings/w:documentProtection{w:edit=forms}",
                WD_PROTECTION.FORMS,
            ),
        ],
    )
    def it_can_get_mode(self, cxml: str, expected_mode: WD_PROTECTION | None):
        settings = Settings(element(cxml))
        assert settings.document_protection.mode == expected_mode

    @pytest.mark.parametrize(
        "member",
        [
            WD_PROTECTION.READ_ONLY,
            WD_PROTECTION.COMMENTS,
            WD_PROTECTION.TRACKED_CHANGES,
            WD_PROTECTION.FORMS,
        ],
    )
    def it_round_trips_every_mode(self, member: WD_PROTECTION):
        settings = Settings(element("w:settings"))
        settings.document_protection.mode = member
        assert settings.document_protection.mode == member

    def it_creates_the_documentProtection_element_on_first_mode_set(self):
        settings = Settings(element("w:settings"))
        settings.document_protection.mode = WD_PROTECTION.COMMENTS
        assert settings._settings.xml == xml(
            "w:settings/w:documentProtection{w:edit=comments}"
        )

    def it_clears_mode_when_assigned_None(self):
        settings = Settings(
            element("w:settings/w:documentProtection{w:edit=readOnly}")
        )
        settings.document_protection.mode = None
        assert settings.document_protection.mode is None

    @pytest.mark.parametrize(
        ("cxml", "expected_enforce"),
        [
            ("w:settings", False),
            ("w:settings/w:documentProtection", False),
            ("w:settings/w:documentProtection{w:enforcement=0}", False),
            ("w:settings/w:documentProtection{w:enforcement=1}", True),
        ],
    )
    def it_can_get_enforce(self, cxml: str, expected_enforce: bool):
        assert Settings(element(cxml)).document_protection.enforce is expected_enforce

    def it_can_set_enforce(self):
        settings = Settings(element("w:settings"))
        settings.document_protection.enforce = True
        assert settings.document_protection.enforce is True
        settings.document_protection.enforce = False
        assert settings.document_protection.enforce is False

    # -- formatting_locked -------------------------------------------------

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("w:settings", False),
            ("w:settings/w:documentProtection", False),
            ("w:settings/w:documentProtection{w:formatting=1}", True),
            ("w:settings/w:documentProtection{w:formatting=0}", False),
        ],
    )
    def it_can_get_formatting_locked(self, cxml: str, expected: bool):
        assert (
            Settings(element(cxml)).document_protection.formatting_locked is expected
        )

    def it_can_set_formatting_locked(self):
        settings = Settings(element("w:settings"))
        settings.document_protection.formatting_locked = True
        assert settings.document_protection.formatting_locked is True
        settings.document_protection.formatting_locked = False
        assert settings.document_protection.formatting_locked is False

    # -- password_hash / password_salt -------------------------------------

    def it_round_trips_password_hash_and_salt(self):
        settings = Settings(element("w:settings"))
        settings.document_protection.password_hash = "deadbeef=="
        settings.document_protection.password_salt = "cafebabe+/"
        assert settings.document_protection.password_hash == "deadbeef=="
        assert settings.document_protection.password_salt == "cafebabe+/"

    def it_returns_None_for_password_fields_when_absent(self):
        settings = Settings(element("w:settings"))
        assert settings.document_protection.password_hash is None
        assert settings.document_protection.password_salt is None

    def it_can_clear_password_hash_by_assigning_None(self):
        settings = Settings(element("w:settings"))
        settings.document_protection.password_hash = "abc"
        settings.document_protection.password_hash = None
        assert settings.document_protection.password_hash is None

    # -- algorithm metadata ------------------------------------------------

    def it_round_trips_algorithm_metadata(self):
        settings = Settings(element("w:settings"))
        protection = settings.document_protection
        protection.crypto_provider_type = "rsaAES"
        protection.crypto_algorithm_class = "hash"
        protection.crypto_algorithm_type = "typeAny"
        protection.crypto_algorithm_sid = 4
        protection.spin_count = 100000

        assert protection.crypto_provider_type == "rsaAES"
        assert protection.crypto_algorithm_class == "hash"
        assert protection.crypto_algorithm_type == "typeAny"
        assert protection.crypto_algorithm_sid == 4
        assert protection.spin_count == 100000

    def it_returns_None_for_algorithm_metadata_when_absent(self):
        settings = Settings(element("w:settings"))
        protection = settings.document_protection
        assert protection.crypto_provider_type is None
        assert protection.crypto_algorithm_class is None
        assert protection.crypto_algorithm_type is None
        assert protection.crypto_algorithm_sid is None
        assert protection.spin_count is None

    # -- high-level enable / disable --------------------------------------

    def it_returns_a_DocumentProtection_from_document_protection(self):
        settings = Settings(element("w:settings"))
        assert isinstance(settings.document_protection, DocumentProtection)

    def it_can_enable_protection_without_a_password(self):
        settings = Settings(element("w:settings"))
        settings.enable_protection(mode=WD_PROTECTION.COMMENTS)
        protection = settings.document_protection
        assert protection.mode == WD_PROTECTION.COMMENTS
        assert protection.enforce is True
        assert protection.password_hash is None
        assert protection.password_salt is None
        assert protection.crypto_provider_type is None
        assert protection.spin_count is None

    def it_can_enable_protection_with_a_password(self):
        settings = Settings(element("w:settings"))
        settings.enable_protection(mode=WD_PROTECTION.READ_ONLY, password="secret")
        protection = settings.document_protection
        assert protection.mode == WD_PROTECTION.READ_ONLY
        assert protection.enforce is True
        assert protection.password_hash is not None
        assert protection.password_salt is not None
        # -- base64 strings decode cleanly to their expected byte lengths --
        import base64 as _b64

        assert len(_b64.b64decode(protection.password_salt)) == 16
        assert len(_b64.b64decode(protection.password_hash)) == 20  # SHA-1
        assert protection.crypto_provider_type == "rsaAES"
        assert protection.crypto_algorithm_class == "hash"
        assert protection.crypto_algorithm_type == "typeAny"
        assert protection.crypto_algorithm_sid == 4
        assert protection.spin_count == 100000

    def it_produces_a_different_hash_per_call(self):
        settings = Settings(element("w:settings"))
        settings.enable_protection(mode=WD_PROTECTION.READ_ONLY, password="secret")
        hash1 = settings.document_protection.password_hash
        salt1 = settings.document_protection.password_salt
        settings.enable_protection(mode=WD_PROTECTION.READ_ONLY, password="secret")
        hash2 = settings.document_protection.password_hash
        salt2 = settings.document_protection.password_salt
        # -- fresh random salt means the hashes differ too --
        assert salt1 != salt2
        assert hash1 != hash2

    def it_can_enable_protection_without_enforcement(self):
        settings = Settings(element("w:settings"))
        settings.enable_protection(mode=WD_PROTECTION.FORMS, enforce=False)
        protection = settings.document_protection
        assert protection.mode == WD_PROTECTION.FORMS
        assert protection.enforce is False

    def it_defaults_to_READ_ONLY_when_no_mode_given(self):
        settings = Settings(element("w:settings"))
        settings.enable_protection()
        assert settings.document_protection.mode == WD_PROTECTION.READ_ONLY
        assert settings.document_protection.enforce is True

    def it_can_disable_protection(self):
        settings = Settings(
            element("w:settings/w:documentProtection{w:edit=readOnly,w:enforcement=1}")
        )
        settings.disable_protection()
        assert settings._settings.xml == xml("w:settings")

    def it_tolerates_disable_when_already_absent(self):
        settings = Settings(element("w:settings"))
        settings.disable_protection()
        assert settings._settings.xml == xml("w:settings")

    def it_clears_stale_password_fields_when_enabling_without_password(self):
        settings = Settings(element("w:settings"))
        settings.enable_protection(mode=WD_PROTECTION.READ_ONLY, password="secret")
        # -- now re-enable without a password: hash/salt/crypto-meta should be cleared --
        settings.enable_protection(mode=WD_PROTECTION.READ_ONLY)
        protection = settings.document_protection
        assert protection.password_hash is None
        assert protection.password_salt is None
        assert protection.crypto_provider_type is None
        assert protection.spin_count is None


class DescribeWriteProtection:
    """Unit-test suite for `docx.settings.WriteProtection`."""

    def it_returns_a_WriteProtection_from_write_protection(self):
        settings = Settings(element("w:settings"))
        assert isinstance(settings.write_protection, WriteProtection)

    def it_reads_present_as_True_when_element_exists(self):
        settings = Settings(element("w:settings/w:writeProtection"))
        assert settings.write_protection.present is True

    def it_reads_present_as_False_when_element_absent(self):
        settings = Settings(element("w:settings"))
        assert settings.write_protection.present is False

    # -- recommended_read_only ---------------------------------------------

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("w:settings", False),
            ("w:settings/w:writeProtection", False),
            ("w:settings/w:writeProtection{w:recommended=1}", True),
            ("w:settings/w:writeProtection{w:recommended=0}", False),
        ],
    )
    def it_can_get_recommended_read_only(self, cxml: str, expected: bool):
        assert (
            Settings(element(cxml)).write_protection.recommended_read_only is expected
        )

    def it_can_set_recommended_read_only(self):
        settings = Settings(element("w:settings"))
        settings.write_protection.recommended_read_only = True
        assert settings.write_protection.recommended_read_only is True
        settings.write_protection.recommended_read_only = False
        assert settings.write_protection.recommended_read_only is False

    def it_creates_the_writeProtection_element_on_first_write(self):
        settings = Settings(element("w:settings"))
        settings.write_protection.recommended_read_only = True
        assert settings._settings.xml == xml(
            "w:settings/w:writeProtection{w:recommended=1}"
        )

    def it_exposes_enforcement_alias_for_recommended_read_only(self):
        settings = Settings(element("w:settings"))
        settings.write_protection.enforcement = True
        assert settings.write_protection.enforcement is True
        assert settings.write_protection.recommended_read_only is True

    # -- password round-trip ----------------------------------------------

    def it_round_trips_password_hash_and_salt(self):
        settings = Settings(element("w:settings"))
        settings.write_protection.password_hash = "deadbeef=="
        settings.write_protection.password_salt = "cafebabe+/"
        assert settings.write_protection.password_hash == "deadbeef=="
        assert settings.write_protection.password_salt == "cafebabe+/"

    def it_returns_None_for_password_fields_when_absent(self):
        settings = Settings(element("w:settings"))
        wp = settings.write_protection
        assert wp.password_hash is None
        assert wp.password_salt is None
        assert wp.crypto_provider_type is None
        assert wp.crypto_algorithm_class is None
        assert wp.crypto_algorithm_type is None
        assert wp.crypto_algorithm_sid is None
        assert wp.spin_count is None

    def it_round_trips_algorithm_metadata(self):
        settings = Settings(element("w:settings"))
        wp = settings.write_protection
        wp.crypto_provider_type = "rsaAES"
        wp.crypto_algorithm_class = "hash"
        wp.crypto_algorithm_type = "typeAny"
        wp.crypto_algorithm_sid = 4
        wp.spin_count = 100000
        assert wp.crypto_provider_type == "rsaAES"
        assert wp.crypto_algorithm_class == "hash"
        assert wp.crypto_algorithm_type == "typeAny"
        assert wp.crypto_algorithm_sid == 4
        assert wp.spin_count == 100000

    # -- high-level enable / disable --------------------------------------

    def it_can_enable_write_protection_without_a_password(self):
        settings = Settings(element("w:settings"))
        settings.enable_write_protection(recommended=True)
        wp = settings.write_protection
        assert wp.recommended_read_only is True
        assert wp.password_hash is None
        assert wp.password_salt is None
        assert wp.crypto_provider_type is None
        assert wp.spin_count is None

    def it_can_enable_write_protection_with_a_password(self):
        settings = Settings(element("w:settings"))
        settings.enable_write_protection(recommended=True, password="s3cret")
        wp = settings.write_protection
        assert wp.recommended_read_only is True
        assert wp.password_hash is not None
        assert wp.password_salt is not None
        import base64 as _b64

        assert len(_b64.b64decode(wp.password_salt)) == 16
        assert len(_b64.b64decode(wp.password_hash)) == 20  # SHA-1
        assert wp.crypto_provider_type == "rsaAES"
        assert wp.crypto_algorithm_class == "hash"
        assert wp.crypto_algorithm_type == "typeAny"
        assert wp.crypto_algorithm_sid == 4
        assert wp.spin_count == 100000

    def it_produces_a_different_hash_per_call(self):
        settings = Settings(element("w:settings"))
        settings.enable_write_protection(password="s3cret")
        hash1 = settings.write_protection.password_hash
        salt1 = settings.write_protection.password_salt
        settings.enable_write_protection(password="s3cret")
        hash2 = settings.write_protection.password_hash
        salt2 = settings.write_protection.password_salt
        assert salt1 != salt2
        assert hash1 != hash2

    def it_can_disable_write_protection(self):
        settings = Settings(
            element("w:settings/w:writeProtection{w:recommended=1}")
        )
        settings.disable_write_protection()
        assert settings._settings.xml == xml("w:settings")

    def it_tolerates_disable_when_already_absent(self):
        settings = Settings(element("w:settings"))
        settings.disable_write_protection()
        assert settings._settings.xml == xml("w:settings")

    def it_clears_stale_password_fields_when_enabling_without_password(self):
        settings = Settings(element("w:settings"))
        settings.enable_write_protection(recommended=True, password="s3cret")
        settings.enable_write_protection(recommended=True)
        wp = settings.write_protection
        assert wp.password_hash is None
        assert wp.password_salt is None
        assert wp.crypto_provider_type is None
        assert wp.spin_count is None

    def it_preserves_document_protection_when_removing_write_protection(self):
        cxml = (
            "w:settings/("
            "w:writeProtection{w:recommended=1},"
            "w:documentProtection{w:edit=readOnly,w:enforcement=1})"
        )
        settings = Settings(element(cxml))
        settings.disable_write_protection()
        # -- documentProtection still present --
        assert settings.document_protection.mode == WD_PROTECTION.READ_ONLY


class DescribeSettings_themeFontLanguage:
    """Unit-test suite for :attr:`Settings.theme_font_language`."""

    def it_returns_a_tuple_of_Nones_when_element_absent(self):
        settings = Settings(element("w:settings"))

        assert settings.theme_font_language == (None, None, None)

    def it_reads_the_three_language_attrs(self):
        settings = Settings(
            element(
                "w:settings/w:themeFontLang{w:val=en-US,w:eastAsia=zh-CN,w:bidi=ar-SA}"
            )
        )

        assert settings.theme_font_language == ("en-US", "zh-CN", "ar-SA")

    def it_sets_val_when_assigning_a_plain_string(self):
        settings = Settings(element("w:settings"))

        settings.theme_font_language = "en-GB"

        assert settings.theme_font_language == ("en-GB", None, None)

    def it_sets_all_three_when_assigning_a_tuple(self):
        settings = Settings(element("w:settings"))

        settings.theme_font_language = ("en-US", "zh-CN", "ar-SA")

        assert settings.theme_font_language == ("en-US", "zh-CN", "ar-SA")

    def it_removes_the_element_when_assigned_None(self):
        settings = Settings(element("w:settings/w:themeFontLang{w:val=en-US}"))

        settings.theme_font_language = None

        assert settings.theme_font_language == (None, None, None)


class DescribeSettings_spellAndGrammar:
    """Unit-test suite for hide_spelling_errors / hide_grammatical_errors."""

    def it_defaults_to_False_when_absent(self):
        settings = Settings(element("w:settings"))

        assert settings.hide_spelling_errors is False
        assert settings.hide_grammatical_errors is False

    def it_returns_True_when_element_present(self):
        settings = Settings(
            element("w:settings/(w:hideSpellingErrors,w:hideGrammaticalErrors)")
        )

        assert settings.hide_spelling_errors is True
        assert settings.hide_grammatical_errors is True

    def it_adds_and_removes_hide_spelling_errors(self):
        settings = Settings(element("w:settings"))

        settings.hide_spelling_errors = True
        assert settings.hide_spelling_errors is True

        settings.hide_spelling_errors = False
        assert settings.hide_spelling_errors is False
        # -- element should be removed, not merely set to 0 --
        assert settings._settings.hideSpellingErrors is None

    def it_adds_and_removes_hide_grammatical_errors(self):
        settings = Settings(element("w:settings"))

        settings.hide_grammatical_errors = True
        assert settings.hide_grammatical_errors is True

        settings.hide_grammatical_errors = False
        assert settings.hide_grammatical_errors is False
        assert settings._settings.hideGrammaticalErrors is None


class DescribeSettings_autoHyphenation:
    """Unit-test suite for auto-hyphenation related settings."""

    def it_defaults_to_False_when_absent(self):
        settings = Settings(element("w:settings"))

        assert settings.auto_hyphenation is False
        assert settings.do_not_hyphenate_caps is False
        assert settings.consecutive_hyphen_limit is None
        assert settings.hyphenation_zone is None

    def it_reads_auto_hyphenation_presence(self):
        settings = Settings(element("w:settings/w:autoHyphenation"))

        assert settings.auto_hyphenation is True

    def it_can_enable_and_disable_auto_hyphenation(self):
        settings = Settings(element("w:settings"))

        settings.auto_hyphenation = True
        assert settings.auto_hyphenation is True

        settings.auto_hyphenation = False
        assert settings.auto_hyphenation is False
        assert settings._settings.autoHyphenation is None

    def it_can_set_consecutive_hyphen_limit(self):
        settings = Settings(element("w:settings"))

        settings.consecutive_hyphen_limit = 3

        assert settings.consecutive_hyphen_limit == 3

    def it_removes_consecutive_hyphen_limit_when_zero_or_None(self):
        settings = Settings(element("w:settings/w:consecutiveHyphenLimit{w:val=5}"))

        settings.consecutive_hyphen_limit = None

        assert settings.consecutive_hyphen_limit is None

    def it_can_roundtrip_hyphenation_zone(self):
        settings = Settings(element("w:settings"))

        settings.hyphenation_zone = Twips(720)

        assert settings.hyphenation_zone == Twips(720)


class DescribeSettings_docVars:
    """Unit-test suite for :attr:`Settings.doc_vars`."""

    def it_is_empty_when_no_container_present(self):
        settings = Settings(element("w:settings"))

        assert len(settings.doc_vars) == 0
        assert list(settings.doc_vars) == []

    def it_can_add_and_read_back_a_variable(self):
        settings = Settings(element("w:settings"))

        settings.doc_vars["project"] = "Alpha"

        assert settings.doc_vars["project"] == "Alpha"
        assert len(settings.doc_vars) == 1

    def it_reads_existing_doc_vars(self):
        settings = Settings(
            element("w:settings/w:docVars/w:docVar{w:name=color,w:val=red}")
        )

        assert settings.doc_vars["color"] == "red"

    def it_overwrites_existing_value_on_setitem(self):
        settings = Settings(element("w:settings"))
        settings.doc_vars["x"] = "1"

        settings.doc_vars["x"] = "2"

        assert settings.doc_vars["x"] == "2"
        assert len(settings.doc_vars) == 1

    def it_raises_KeyError_on_missing_name(self):
        settings = Settings(element("w:settings"))

        with pytest.raises(KeyError):
            _ = settings.doc_vars["missing"]

    def it_removes_variable_and_prunes_empty_container(self):
        settings = Settings(element("w:settings"))
        settings.doc_vars["x"] = "1"

        del settings.doc_vars["x"]

        assert len(settings.doc_vars) == 0
        # -- the w:docVars container should be pruned on empty --
        assert settings._settings.docVars is None

    def it_supports_contains_get_and_items(self):
        settings = Settings(element("w:settings"))
        settings.doc_vars["a"] = "1"
        settings.doc_vars["b"] = "2"

        assert "a" in settings.doc_vars
        assert settings.doc_vars.get("missing") is None
        assert settings.doc_vars.get("missing", "default") == "default"
        assert settings.doc_vars.items() == [("a", "1"), ("b", "2")]


class DescribeSettings_docId:
    """Unit-test suite for Word-extension doc-identifier accessors."""

    def it_is_None_when_neither_docId_present(self):
        settings = Settings(element("w:settings"))

        assert settings.doc_id is None

    def it_reads_and_writes_round_trip(self):
        settings = Settings(element("w:settings"))

        settings.doc_id = "AAAAAAAA-1111-1111-1111-111111111111"

        assert settings.doc_id == "{AAAAAAAA-1111-1111-1111-111111111111}"

    def it_accepts_preformatted_braced_guid_without_double_wrapping(self):
        settings = Settings(element("w:settings"))

        settings.doc_id = "{12345678-1234-1234-1234-123456789012}"

        assert settings.doc_id == "{12345678-1234-1234-1234-123456789012}"

    def it_removes_both_w14_and_w15_docIds_on_None(self):
        settings = Settings(element("w:settings"))
        settings.doc_id = "11111111-1111-1111-1111-111111111111"

        settings.doc_id = None

        assert settings.doc_id is None
        assert settings._settings.w14_docId is None
        assert settings._settings.w15_docId is None


class DescribeSettings_chartTrackingRefBased:
    """Unit-test suite for ``w15:chartTrackingRefBased`` accessor."""

    def it_is_False_when_element_absent(self):
        settings = Settings(element("w:settings"))

        assert settings.chart_tracking_ref_based is False

    def it_is_True_when_element_present(self):
        settings = Settings(element("w:settings/w15:chartTrackingRefBased"))

        assert settings.chart_tracking_ref_based is True

    def it_adds_the_element_when_enabled(self):
        settings = Settings(element("w:settings"))

        settings.chart_tracking_ref_based = True

        assert settings.chart_tracking_ref_based is True

    def it_removes_the_element_when_disabled(self):
        settings = Settings(element("w:settings/w15:chartTrackingRefBased"))

        settings.chart_tracking_ref_based = False

        assert settings.chart_tracking_ref_based is False
