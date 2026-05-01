# pyright: reportPrivateUsage=false

"""Unit-test suite for the `docx.web_settings` module."""

from __future__ import annotations

import pytest

from docx.web_settings import WebSettings

from .unitutil.cxml import element, xml


class DescribeWebSettings:
    """Unit-test suite for `docx.web_settings.WebSettings`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", None),
            ("w:webSettings/w:encoding{w:val=utf-8}", "utf-8"),
            ("w:webSettings/w:encoding{w:val=windows-1252}", "windows-1252"),
            ("w:webSettings/w:encoding", None),
        ],
    )
    def it_provides_access_to_the_encoding(
        self, cxml: str, expected_value: str | None
    ):
        assert WebSettings(element(cxml)).encoding == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", False),
            ("w:webSettings/w:optimizeForBrowser", True),
            ("w:webSettings/w:optimizeForBrowser{w:val=0}", False),
        ],
    )
    def it_provides_access_to_optimize_for_browser(
        self, cxml: str, expected_value: bool
    ):
        assert WebSettings(element(cxml)).optimize_for_browser is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:webSettings", True, "w:webSettings/w:optimizeForBrowser"),
            ("w:webSettings/w:optimizeForBrowser", False, "w:webSettings"),
            ("w:webSettings/w:optimizeForBrowser", None, "w:webSettings"),
        ],
    )
    def it_can_change_optimize_for_browser(
        self, cxml: str, new_value: bool | None, expected_cxml: str
    ):
        web_settings = WebSettings(element(cxml))
        web_settings.optimize_for_browser = new_value
        assert web_settings._web_settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", False),
            ("w:webSettings/w:allowPNG", True),
            ("w:webSettings/w:allowPNG{w:val=0}", False),
        ],
    )
    def it_provides_access_to_allow_png(self, cxml: str, expected_value: bool):
        assert WebSettings(element(cxml)).allow_png is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:webSettings", True, "w:webSettings/w:allowPNG"),
            ("w:webSettings/w:allowPNG", False, "w:webSettings"),
        ],
    )
    def it_can_change_allow_png(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        web_settings = WebSettings(element(cxml))
        web_settings.allow_png = new_value
        assert web_settings._web_settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", False),
            ("w:webSettings/w:doNotSaveAsSingleFile", True),
            ("w:webSettings/w:doNotSaveAsSingleFile{w:val=0}", False),
        ],
    )
    def it_provides_access_to_do_not_save_as_single_file(
        self, cxml: str, expected_value: bool
    ):
        assert (
            WebSettings(element(cxml)).do_not_save_as_single_file is expected_value
        )

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:webSettings", True, "w:webSettings/w:doNotSaveAsSingleFile"),
            ("w:webSettings/w:doNotSaveAsSingleFile", False, "w:webSettings"),
        ],
    )
    def it_can_change_do_not_save_as_single_file(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        web_settings = WebSettings(element(cxml))
        web_settings.do_not_save_as_single_file = new_value
        assert web_settings._web_settings.xml == xml(expected_cxml)
