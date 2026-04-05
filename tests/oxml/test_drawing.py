# pyright: reportPrivateUsage=false

"""Unit test suite for the docx.oxml.drawing module."""

from __future__ import annotations

from typing import cast

from docx.oxml.drawing import CT_Drawing, CT_TxbxContent

from ..unitutil.cxml import element


class DescribeCT_Drawing:
    """Unit test suite for `docx.oxml.drawing.CT_Drawing` objects."""

    def it_provides_access_to_txbxContent_descendants(self):
        drawing = cast(
            CT_Drawing,
            element(
                "w:drawing/wp:inline/a:graphic/a:graphicData"
                "/wps:wsp/wps:txbx/w:txbxContent/w:p"
            ),
        )

        txbx_contents = drawing.txbxContent_lst

        assert len(txbx_contents) == 1
        assert isinstance(txbx_contents[0], CT_TxbxContent)

    def it_returns_empty_list_when_no_txbxContent(self):
        drawing = cast(
            CT_Drawing,
            element("w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic"),
        )

        assert drawing.txbxContent_lst == []


class DescribeCT_TxbxContent:
    """Unit test suite for `docx.oxml.drawing.CT_TxbxContent` objects."""

    def it_provides_access_to_its_paragraph_children(self):
        txbxContent = cast(
            CT_TxbxContent,
            element("w:txbxContent/(w:p,w:p)"),
        )

        assert len(txbxContent.p_lst) == 2

    def it_can_get_concatenated_text(self):
        txbxContent = cast(
            CT_TxbxContent,
            element('w:txbxContent/(w:p/w:r/w:t"Hello",w:p/w:r/w:t"World")'),
        )

        assert txbxContent.text == "Hello\nWorld"

    def it_returns_empty_string_when_no_text(self):
        txbxContent = cast(
            CT_TxbxContent,
            element("w:txbxContent/w:p"),
        )

        assert txbxContent.text == ""
