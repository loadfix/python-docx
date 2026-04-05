# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.settings` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.settings import CT_Settings
from docx.shared import Twips

from ..unitutil.cxml import element, xml


class DescribeCT_Settings:
    """Unit-test suite for `docx.oxml.settings.CT_Settings`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", None),
            ("w:settings/w:zoom{w:percent=100}", 100),
            ("w:settings/w:zoom{w:percent=75}", 75),
            ("w:settings/w:zoom", None),
        ],
    )
    def it_can_get_the_zoom_percent(self, cxml: str, expected_value: int | None):
        settings = cast(CT_Settings, element(cxml))
        assert settings.zoom_percent == expected_value

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
        settings = cast(CT_Settings, element(cxml))
        settings.zoom_percent = new_value
        assert settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", False),
            ("w:settings/w:trackRevisions", True),
            ("w:settings/w:trackRevisions{w:val=0}", False),
            ("w:settings/w:trackRevisions{w:val=true}", True),
        ],
    )
    def it_can_get_trackRevisions(self, cxml: str, expected_value: bool):
        settings = cast(CT_Settings, element(cxml))
        assert settings.trackRevisions_val is expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", True, "w:settings/w:trackRevisions"),
            ("w:settings/w:trackRevisions", False, "w:settings"),
            ("w:settings/w:trackRevisions{w:val=0}", True, "w:settings/w:trackRevisions"),
        ],
    )
    def it_can_set_trackRevisions(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        settings = cast(CT_Settings, element(cxml))
        settings.trackRevisions_val = new_value
        assert settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", None),
            ("w:settings/w:defaultTabStop{w:val=720}", Twips(720)),
            ("w:settings/w:defaultTabStop{w:val=360}", Twips(360)),
        ],
    )
    def it_can_get_the_defaultTabStop(self, cxml: str, expected_value):
        settings = cast(CT_Settings, element(cxml))
        assert settings.defaultTabStop_val == expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", Twips(720), "w:settings/w:defaultTabStop{w:val=720}"),
            (
                "w:settings/w:defaultTabStop{w:val=720}",
                Twips(360),
                "w:settings/w:defaultTabStop{w:val=360}",
            ),
            ("w:settings/w:defaultTabStop{w:val=720}", None, "w:settings"),
        ],
    )
    def it_can_set_the_defaultTabStop(
        self, cxml: str, new_value, expected_cxml: str
    ):
        settings = cast(CT_Settings, element(cxml))
        settings.defaultTabStop_val = new_value
        assert settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "expected_edit", "expected_enforcement"),
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
            ("w:settings/w:documentProtection{w:edit=forms}", "forms", False),
        ],
    )
    def it_can_get_document_protection(
        self,
        cxml: str,
        expected_edit: str | None,
        expected_enforcement: bool,
    ):
        settings = cast(CT_Settings, element(cxml))
        assert settings.documentProtection_edit == expected_edit
        assert settings.documentProtection_enforcement is expected_enforcement

    def it_can_get_the_compatibilityMode_when_absent(self):
        settings = cast(CT_Settings, element("w:settings"))
        assert settings.compatibilityMode is None

    def it_can_get_the_compatibilityMode_when_present(self):
        settings = cast(CT_Settings, element("w:settings/w:compat"))
        # -- no compatSetting children yet, so None --
        assert settings.compatibilityMode is None

    def it_can_set_the_compatibilityMode(self):
        settings = cast(CT_Settings, element("w:settings"))
        settings.compatibilityMode = 15
        assert settings.compatibilityMode == 15

    def it_can_change_the_compatibilityMode(self):
        settings = cast(CT_Settings, element("w:settings"))
        settings.compatibilityMode = 14
        assert settings.compatibilityMode == 14
        settings.compatibilityMode = 15
        assert settings.compatibilityMode == 15

    def it_can_remove_the_compatibilityMode(self):
        settings = cast(CT_Settings, element("w:settings"))
        settings.compatibilityMode = 15
        assert settings.compatibilityMode == 15
        settings.compatibilityMode = None
        assert settings.compatibilityMode is None

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", False),
            ("w:settings/w:evenAndOddHeaders", True),
            ("w:settings/w:evenAndOddHeaders{w:val=0}", False),
            ("w:settings/w:evenAndOddHeaders{w:val=1}", True),
        ],
    )
    def it_can_get_evenAndOddHeaders(self, cxml: str, expected_value: bool):
        settings = cast(CT_Settings, element(cxml))
        assert settings.evenAndOddHeaders_val is expected_value
