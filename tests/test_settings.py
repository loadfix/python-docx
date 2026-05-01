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
    WD_VIEW,
)
from docx.footnotes import FootnoteProperties
from docx.settings import CompatFlags, CompatSettings, Settings
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
