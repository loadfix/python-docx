# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.shared` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.ns import qn
from docx.oxml.shared import CT_DecimalNumber, CT_OnOff, CT_String

from ..unitutil.cxml import element


class DescribeCT_DecimalNumber:
    """Unit-test suite for `docx.oxml.shared.CT_DecimalNumber`."""

    def it_knows_its_val(self):
        dn = cast(CT_DecimalNumber, element("w:numId{w:val=42}"))
        assert dn.val == 42

    def it_can_construct_a_new_element(self):
        dn = CT_DecimalNumber.new("w:abstractNumId", 7)
        assert dn.tag == qn("w:abstractNumId")
        assert dn.get(qn("w:val")) == "7"


class DescribeCT_OnOff:
    """Unit-test suite for `docx.oxml.shared.CT_OnOff`."""

    def it_defaults_to_True_when_val_attribute_is_absent(self):
        onoff = cast(CT_OnOff, element("w:b"))
        assert onoff.val is True

    def it_reads_an_explicit_false_val(self):
        onoff = cast(CT_OnOff, element("w:b{w:val=false}"))
        assert onoff.val is False

    def it_reads_an_explicit_true_val(self):
        onoff = cast(CT_OnOff, element("w:b{w:val=true}"))
        assert onoff.val is True


class DescribeCT_String:
    """Unit-test suite for `docx.oxml.shared.CT_String`."""

    def it_knows_its_val(self):
        s = cast(CT_String, element("w:pStyle{w:val=Heading1}"))
        assert s.val == "Heading1"

    def it_can_construct_a_new_element(self):
        s = CT_String.new("w:pStyle", "Normal")
        assert s.tag == qn("w:pStyle")
        assert s.val == "Normal"
