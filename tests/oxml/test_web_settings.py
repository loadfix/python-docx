# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.web_settings` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.web_settings import CT_WebSettings

from ..unitutil.cxml import element, xml


class DescribeCT_WebSettings:
    """Unit-test suite for `docx.oxml.web_settings.CT_WebSettings`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", None),
            ("w:webSettings/w:encoding", None),
            ("w:webSettings/w:encoding{w:val=utf-8}", "utf-8"),
            ("w:webSettings/w:encoding{w:val=windows-1252}", "windows-1252"),
        ],
    )
    def it_can_get_the_encoding_val(self, cxml: str, expected_value: str | None):
        web_settings = cast(CT_WebSettings, element(cxml))
        assert web_settings.encoding_val == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", False),
            ("w:webSettings/w:optimizeForBrowser", True),
            ("w:webSettings/w:optimizeForBrowser{w:val=0}", False),
            ("w:webSettings/w:optimizeForBrowser{w:val=1}", True),
            ("w:webSettings/w:optimizeForBrowser{w:val=true}", True),
        ],
    )
    def it_can_get_the_optimizeForBrowser_val(self, cxml: str, expected_value: bool):
        web_settings = cast(CT_WebSettings, element(cxml))
        assert web_settings.optimizeForBrowser_val is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:webSettings", True, "w:webSettings/w:optimizeForBrowser"),
            ("w:webSettings/w:optimizeForBrowser", False, "w:webSettings"),
            ("w:webSettings/w:optimizeForBrowser{w:val=0}", True,
             "w:webSettings/w:optimizeForBrowser"),
            ("w:webSettings/w:optimizeForBrowser", None, "w:webSettings"),
        ],
    )
    def it_can_set_the_optimizeForBrowser_val(
        self, cxml: str, new_value: bool | None, expected_cxml: str
    ):
        web_settings = cast(CT_WebSettings, element(cxml))
        web_settings.optimizeForBrowser_val = new_value
        assert web_settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", False),
            ("w:webSettings/w:allowPNG", True),
            ("w:webSettings/w:allowPNG{w:val=0}", False),
            ("w:webSettings/w:allowPNG{w:val=true}", True),
        ],
    )
    def it_can_get_the_allowPNG_val(self, cxml: str, expected_value: bool):
        web_settings = cast(CT_WebSettings, element(cxml))
        assert web_settings.allowPNG_val is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:webSettings", True, "w:webSettings/w:allowPNG"),
            ("w:webSettings/w:allowPNG", False, "w:webSettings"),
            ("w:webSettings/w:allowPNG{w:val=0}", True, "w:webSettings/w:allowPNG"),
        ],
    )
    def it_can_set_the_allowPNG_val(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        web_settings = cast(CT_WebSettings, element(cxml))
        web_settings.allowPNG_val = new_value
        assert web_settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", False),
            ("w:webSettings/w:doNotSaveAsSingleFile", True),
            ("w:webSettings/w:doNotSaveAsSingleFile{w:val=0}", False),
        ],
    )
    def it_can_get_the_doNotSaveAsSingleFile_val(
        self, cxml: str, expected_value: bool
    ):
        web_settings = cast(CT_WebSettings, element(cxml))
        assert web_settings.doNotSaveAsSingleFile_val is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:webSettings", True, "w:webSettings/w:doNotSaveAsSingleFile"),
            ("w:webSettings/w:doNotSaveAsSingleFile", False, "w:webSettings"),
        ],
    )
    def it_can_set_the_doNotSaveAsSingleFile_val(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        web_settings = cast(CT_WebSettings, element(cxml))
        web_settings.doNotSaveAsSingleFile_val = new_value
        assert web_settings.xml == xml(expected_cxml)

    def it_preserves_child_order_when_adding_children(self):
        """Setters must insert children in the schema-prescribed order."""
        web_settings = cast(CT_WebSettings, element("w:webSettings"))
        # -- add in reverse-schema order to exercise the successors lists --
        web_settings.doNotSaveAsSingleFile_val = True
        web_settings.allowPNG_val = True
        web_settings.optimizeForBrowser_val = True
        expected = xml(
            "w:webSettings/("
            "w:optimizeForBrowser,"
            "w:allowPNG,"
            "w:doNotSaveAsSingleFile"
            ")"
        )
        assert web_settings.xml == expected

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:webSettings", False),
            ("w:webSettings/w:relyOnVML", True),
            ("w:webSettings/w:relyOnVML{w:val=0}", False),
            ("w:webSettings/w:relyOnVML{w:val=true}", True),
        ],
    )
    def it_can_get_the_relyOnVML_val(self, cxml: str, expected_value: bool):
        web_settings = cast(CT_WebSettings, element(cxml))
        assert web_settings.relyOnVML_val is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:webSettings", True, "w:webSettings/w:relyOnVML"),
            ("w:webSettings/w:relyOnVML", False, "w:webSettings"),
            ("w:webSettings/w:relyOnVML{w:val=0}", True, "w:webSettings/w:relyOnVML"),
            ("w:webSettings/w:relyOnVML", None, "w:webSettings"),
        ],
    )
    def it_can_set_the_relyOnVML_val(
        self, cxml: str, new_value: bool | None, expected_cxml: str
    ):
        web_settings = cast(CT_WebSettings, element(cxml))
        web_settings.relyOnVML_val = new_value
        assert web_settings.xml == xml(expected_cxml)

    def it_exposes_a_frameset_child_when_present(self):
        web_settings = cast(
            CT_WebSettings,
            element("w:webSettings/w:frameset/(w:frame,w:frame)"),
        )
        frameset = web_settings.frameset
        assert frameset is not None
        assert len(frameset.frame_lst) == 2
