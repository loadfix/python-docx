# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.settings module."""

from __future__ import annotations

import warnings

import pytest

from docx.settings import Settings
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
