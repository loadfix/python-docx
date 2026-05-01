# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.settings` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.settings import CT_DocProtect, CT_Rsids, CT_Settings
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
            ("w:settings/w:documentProtection", None, False),
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

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", "readOnly", "w:settings/w:documentProtection{w:edit=readOnly}"),
            (
                "w:settings/w:documentProtection{w:edit=readOnly}",
                "comments",
                "w:settings/w:documentProtection{w:edit=comments}",
            ),
            (
                "w:settings/w:documentProtection{w:edit=readOnly}",
                None,
                "w:settings/w:documentProtection",
            ),
        ],
    )
    def it_can_set_documentProtection_edit(
        self, cxml: str, new_value: str | None, expected_cxml: str
    ):
        settings = cast(CT_Settings, element(cxml))
        settings.documentProtection_edit = new_value
        assert settings.xml == xml(expected_cxml)

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", True, "w:settings/w:documentProtection{w:enforcement=1}"),
            (
                "w:settings/w:documentProtection{w:enforcement=1}",
                False,
                "w:settings/w:documentProtection",
            ),
        ],
    )
    def it_can_set_documentProtection_enforcement(
        self, cxml: str, new_value: bool, expected_cxml: str
    ):
        settings = cast(CT_Settings, element(cxml))
        settings.documentProtection_enforcement = new_value
        assert settings.xml == xml(expected_cxml)

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
        assert settings.xml == xml("w:settings")

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

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:settings", None),
            ("w:settings/w:view", None),
            ("w:settings/w:view{w:val=normal}", "normal"),
            ("w:settings/w:view{w:val=outline}", "outline"),
            ("w:settings/w:view{w:val=print}", "print"),
            ("w:settings/w:view{w:val=web}", "web"),
            ("w:settings/w:view{w:val=reading}", "reading"),
            ("w:settings/w:view{w:val=masterPages}", "masterPages"),
            ("w:settings/w:view{w:val=none}", "none"),
        ],
    )
    def it_can_get_the_view_val(self, cxml: str, expected_value: str | None):
        settings = cast(CT_Settings, element(cxml))
        assert settings.view_val == expected_value

    @pytest.mark.parametrize(
        ("cxml", "new_value", "expected_cxml"),
        [
            ("w:settings", "print", "w:settings/w:view{w:val=print}"),
            (
                "w:settings/w:view{w:val=print}",
                "outline",
                "w:settings/w:view{w:val=outline}",
            ),
            ("w:settings/w:view{w:val=print}", None, "w:settings"),
            ("w:settings/w:zoom{w:percent=100}", "web",
             "w:settings/(w:view{w:val=web},w:zoom{w:percent=100})"),
        ],
    )
    def it_can_set_the_view_val(
        self, cxml: str, new_value: str | None, expected_cxml: str
    ):
        settings = cast(CT_Settings, element(cxml))
        settings.view_val = new_value
        assert settings.xml == xml(expected_cxml)


class DescribeCT_Rsids:
    """Unit-test suite for `docx.oxml.settings.CT_Rsids`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:rsids", None),
            ("w:rsids/w:rsidRoot{w:val=00FA1B42}", "00FA1B42"),
            ("w:rsids/w:rsidRoot", None),
            (
                "w:rsids/(w:rsidRoot{w:val=00ABCDEF},w:rsid{w:val=001234AB})",
                "00ABCDEF",
            ),
        ],
    )
    def it_reads_rsidRoot_val(self, cxml: str, expected_value: str | None):
        rsids = cast(CT_Rsids, element(cxml))
        assert rsids.rsidRoot_val == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("w:rsids", []),
            ("w:rsids/w:rsidRoot{w:val=00FA1B42}", []),
            ("w:rsids/w:rsid{w:val=001234AB}", ["001234AB"]),
            (
                "w:rsids/("
                "w:rsidRoot{w:val=00FA1B42},"
                "w:rsid{w:val=001234AB},"
                "w:rsid{w:val=00567890},"
                "w:rsid{w:val=00ABCDEF})",
                ["001234AB", "00567890", "00ABCDEF"],
            ),
        ],
    )
    def it_reads_rsid_vals_in_document_order(
        self, cxml: str, expected_value: list[str]
    ):
        rsids = cast(CT_Rsids, element(cxml))
        assert rsids.rsid_vals == expected_value


class DescribeCT_Settings_Rsids:
    """Unit-test suite for RSID access via `docx.oxml.settings.CT_Settings`."""

    def it_returns_None_for_rsids_when_not_present(self):
        settings = cast(CT_Settings, element("w:settings"))
        assert settings.rsids is None

    def it_returns_the_rsids_element_when_present(self):
        settings = cast(
            CT_Settings,
            element("w:settings/w:rsids/w:rsidRoot{w:val=00FA1B42}"),
        )
        rsids = settings.rsids
        assert rsids is not None
        assert rsids.rsidRoot_val == "00FA1B42"


class DescribeCT_Compat:
    """Unit-test suite for `docx.oxml.settings.CT_Compat`."""

    # -- compatSetting dict-style helpers -----------------------------------

    def it_returns_None_for_unknown_compat_setting_name(self):
        compat = cast(
            CT_Settings, element("w:settings/w:compat")
        ).compat
        assert compat is not None
        assert compat.get_compat_setting("notThere") is None

    def it_can_get_a_compat_setting_by_name(self):
        settings = cast(
            CT_Settings,
            element(
                "w:settings/w:compat/w:compatSetting"
                "{w:name=compatibilityMode,w:uri=http://x,w:val=15}"
            ),
        )
        assert settings.compat is not None
        assert settings.compat.get_compat_setting("compatibilityMode") == "15"

    def it_can_add_a_new_compat_setting(self):
        settings = cast(CT_Settings, element("w:settings/w:compat"))
        assert settings.compat is not None
        settings.compat.set_compat_setting("foo", "1", uri="http://bar")
        assert settings.xml == xml(
            "w:settings/w:compat/w:compatSetting"
            "{w:name=foo,w:uri=http://bar,w:val=1}"
        )

    def it_can_update_an_existing_compat_setting_in_place(self):
        settings = cast(
            CT_Settings,
            element(
                "w:settings/w:compat/w:compatSetting"
                "{w:name=foo,w:uri=http://keep,w:val=old}"
            ),
        )
        assert settings.compat is not None
        settings.compat.set_compat_setting("foo", "new", uri="http://ignored")
        # -- URI is left unchanged when the setting already exists --
        assert settings.xml == xml(
            "w:settings/w:compat/w:compatSetting"
            "{w:name=foo,w:uri=http://keep,w:val=new}"
        )

    def it_can_remove_a_compat_setting(self):
        settings = cast(
            CT_Settings,
            element(
                "w:settings/w:compat/w:compatSetting"
                "{w:name=foo,w:uri=http://x,w:val=1}"
            ),
        )
        assert settings.compat is not None
        assert settings.compat.remove_compat_setting("foo") is True
        assert settings.xml == xml("w:settings/w:compat")

    def it_returns_False_when_removing_a_missing_compat_setting(self):
        settings = cast(CT_Settings, element("w:settings/w:compat"))
        assert settings.compat is not None
        assert settings.compat.remove_compat_setting("absent") is False

    def it_iterates_compat_setting_names_in_document_order(self):
        settings = cast(
            CT_Settings,
            element(
                "w:settings/w:compat/("
                "w:compatSetting{w:name=a,w:uri=http://x,w:val=1},"
                "w:compatSetting{w:name=b,w:uri=http://x,w:val=2},"
                "w:compatSetting{w:name=c,w:uri=http://x,w:val=3})"
            ),
        )
        assert settings.compat is not None
        assert list(settings.compat.iter_compat_setting_names()) == ["a", "b", "c"]

    # -- direct flag helpers ------------------------------------------------

    def it_reports_has_flag_for_present_child(self):
        settings = cast(
            CT_Settings, element("w:settings/w:compat/w:growAutofit")
        )
        assert settings.compat is not None
        assert settings.compat.has_flag("growAutofit") is True
        assert settings.compat.has_flag("useFELayout") is False

    def it_can_add_a_flag(self):
        settings = cast(CT_Settings, element("w:settings/w:compat"))
        assert settings.compat is not None
        settings.compat.set_flag("growAutofit", True)
        assert settings.xml == xml("w:settings/w:compat/w:growAutofit")

    def it_does_not_duplicate_existing_flag(self):
        settings = cast(
            CT_Settings, element("w:settings/w:compat/w:growAutofit")
        )
        assert settings.compat is not None
        settings.compat.set_flag("growAutofit", True)
        assert settings.xml == xml("w:settings/w:compat/w:growAutofit")

    def it_can_remove_a_flag(self):
        settings = cast(
            CT_Settings, element("w:settings/w:compat/w:growAutofit")
        )
        assert settings.compat is not None
        settings.compat.set_flag("growAutofit", False)
        assert settings.xml == xml("w:settings/w:compat")

    def it_iterates_flag_names_skipping_compatSetting(self):
        settings = cast(
            CT_Settings,
            element(
                "w:settings/w:compat/("
                "w:growAutofit,"
                "w:compatSetting{w:name=n,w:uri=http://x,w:val=1},"
                "w:useFELayout)"
            ),
        )
        assert settings.compat is not None
        assert list(settings.compat.iter_flag_names()) == [
            "growAutofit",
            "useFELayout",
        ]

    def it_can_clear_all_flags_but_preserve_compat_settings(self):
        settings = cast(
            CT_Settings,
            element(
                "w:settings/w:compat/("
                "w:growAutofit,"
                "w:compatSetting{w:name=n,w:uri=http://x,w:val=1},"
                "w:useFELayout)"
            ),
        )
        assert settings.compat is not None
        settings.compat.clear_flags()
        assert settings.xml == xml(
            "w:settings/w:compat/w:compatSetting"
            "{w:name=n,w:uri=http://x,w:val=1}"
        )


class DescribeCT_DocProtect:
    """Unit-test suite for `docx.oxml.settings.CT_DocProtect`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_edit", "expected_enforcement", "expected_formatting"),
        [
            ("w:documentProtection", None, False, False),
            (
                "w:documentProtection{w:edit=readOnly,w:enforcement=1}",
                "readOnly",
                True,
                False,
            ),
            (
                "w:documentProtection{w:edit=comments,w:formatting=1}",
                "comments",
                False,
                True,
            ),
        ],
    )
    def it_can_get_core_attributes(
        self,
        cxml: str,
        expected_edit: str | None,
        expected_enforcement: bool,
        expected_formatting: bool,
    ):
        dp = cast(CT_DocProtect, element(cxml))
        assert dp.edit == expected_edit
        assert dp.enforcement is expected_enforcement
        assert dp.formatting is expected_formatting

    @pytest.mark.parametrize(
        ("attr_name", "xml_name", "value"),
        [
            ("hash", "w:hash", "abc123"),
            ("salt", "w:salt", "def456"),
            ("cryptProviderType", "w:cryptProviderType", "rsaAES"),
            ("cryptAlgorithmClass", "w:cryptAlgorithmClass", "hash"),
            ("cryptAlgorithmType", "w:cryptAlgorithmType", "typeAny"),
        ],
    )
    def it_round_trips_each_string_attribute(
        self, attr_name: str, xml_name: str, value: str
    ):
        dp = cast(
            CT_DocProtect,
            element("w:documentProtection{%s=%s}" % (xml_name, value)),
        )
        assert getattr(dp, attr_name) == value

    @pytest.mark.parametrize(
        ("attr_name", "xml_name", "value"),
        [
            ("cryptAlgorithmSid", "w:cryptAlgorithmSid", 4),
            ("cryptSpinCount", "w:cryptSpinCount", 100000),
        ],
    )
    def it_round_trips_each_int_attribute(
        self, attr_name: str, xml_name: str, value: int
    ):
        dp = cast(
            CT_DocProtect,
            element("w:documentProtection{%s=%d}" % (xml_name, value)),
        )
        assert getattr(dp, attr_name) == value

    def it_can_set_hash_and_salt(self):
        dp = cast(CT_DocProtect, element("w:documentProtection"))
        dp.hash = "deadbeef=="
        dp.salt = "cafebabe+/"
        # -- hashes/salts are base64 strings in real use, which can include
        # -- characters the cxml parser rejects; exercise the setter path
        # -- using raw lxml access, not via cxml element syntax.
        assert dp.hash == "deadbeef=="
        assert dp.salt == "cafebabe+/"

    def it_can_set_algorithm_metadata(self):
        dp = cast(CT_DocProtect, element("w:documentProtection"))
        dp.cryptProviderType = "rsaAES"
        dp.cryptAlgorithmClass = "hash"
        dp.cryptAlgorithmType = "typeAny"
        dp.cryptAlgorithmSid = 4
        dp.cryptSpinCount = 100000
        assert dp.cryptProviderType == "rsaAES"
        assert dp.cryptAlgorithmClass == "hash"
        assert dp.cryptAlgorithmType == "typeAny"
        assert dp.cryptAlgorithmSid == 4
        assert dp.cryptSpinCount == 100000
