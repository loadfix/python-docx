# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.permissions` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.permissions import CT_PermEnd, CT_PermStart

from ..unitutil.cxml import element


class DescribeCT_PermStart:
    """Unit-test suite for `docx.oxml.permissions.CT_PermStart`."""

    def it_knows_its_id(self):
        permStart = cast(CT_PermStart, element("w:permStart{w:id=3}"))
        assert permStart.id == 3

    def it_knows_its_edit_group(self):
        permStart = cast(
            CT_PermStart, element("w:permStart{w:id=3,w:edGrp=everyone}")
        )
        assert permStart.edit_group == "everyone"

    def it_knows_its_user(self):
        permStart = cast(
            CT_PermStart, element("w:permStart{w:id=3,w:ed=alice}")
        )
        assert permStart.user == "alice"

    def it_reports_None_for_absent_optional_attributes(self):
        permStart = cast(CT_PermStart, element("w:permStart{w:id=3}"))
        assert permStart.edit_group is None
        assert permStart.user is None
        assert permStart.displaced_by_custom_xml is None
        assert permStart.col_first is None
        assert permStart.col_last is None

    def it_knows_its_displaced_by_custom_xml(self):
        permStart = cast(
            CT_PermStart,
            element("w:permStart{w:id=3,w:displacedByCustomXml=next}"),
        )
        assert permStart.displaced_by_custom_xml == "next"


class DescribeCT_PermEnd:
    """Unit-test suite for `docx.oxml.permissions.CT_PermEnd`."""

    def it_knows_its_id(self):
        permEnd = cast(CT_PermEnd, element("w:permEnd{w:id=3}"))
        assert permEnd.id == 3
