# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.math` module."""

from __future__ import annotations

from typing import cast

from docx.oxml.math import CT_MathR, CT_MathT, CT_OMath, CT_OMathPara

from ..unitutil.cxml import element


class DescribeCT_OMath:
    """Unit-test suite for `docx.oxml.math.CT_OMath`."""

    def it_is_registered_for_m_oMath(self):
        el = element('m:oMath/m:r/m:t"x"')
        assert isinstance(el, CT_OMath)

    def it_concatenates_m_t_text(self):
        el = cast(CT_OMath, element('m:oMath/(m:r/m:t"a",m:r/m:t"bc")'))
        assert el.text == "abc"

    def it_returns_empty_text_when_no_m_t_children(self):
        el = cast(CT_OMath, element("m:oMath"))
        assert el.text == ""


class DescribeCT_OMathPara:
    """Unit-test suite for `docx.oxml.math.CT_OMathPara`."""

    def it_is_registered_for_m_oMathPara(self):
        el = element("m:oMathPara/m:oMath")
        assert isinstance(el, CT_OMathPara)

    def it_concatenates_descendant_m_t_text(self):
        el = cast(
            CT_OMathPara,
            element('m:oMathPara/m:oMath/(m:r/m:t"a",m:r/m:t"b")'),
        )
        assert el.text == "ab"


class DescribeCT_MathR:
    def it_is_registered_for_m_r(self):
        el = element('m:oMath/m:r/m:t"x"')
        child = el[0]
        assert isinstance(child, CT_MathR)


class DescribeCT_MathT:
    def it_is_registered_for_m_t(self):
        el = element('m:oMath/m:r/m:t"x"')
        t = el[0][0]
        assert isinstance(t, CT_MathT)
