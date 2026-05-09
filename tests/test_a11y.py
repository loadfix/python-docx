# pyright: reportPrivateUsage=false

"""Unit tests for accessibility helpers on InlineShape, Table, and Document.

Covers the round-trip behaviour of :attr:`InlineShape.alt_text`,
:attr:`InlineShape.a11y_role`, :attr:`Table.accessibility_title` /
:attr:`Table.accessibility_summary`, and :meth:`Document.check_alt_text`.
"""

from __future__ import annotations

from typing import cast

import pytest

from docx.document import Document
from docx.oxml.shape import CT_Inline
from docx.oxml.table import CT_Tbl
from docx.shape import InlineShape
from docx.table import Table

from .unitutil.cxml import element
from .unitutil.mock import FixtureRequest, Mock, instance_mock


class DescribeInlineShape_AltText:
    """Round-trip behaviour for ``InlineShape.alt_text``."""

    def it_returns_None_when_descr_is_absent(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))

        assert InlineShape(inline).alt_text is None

    def it_round_trips_a_simple_value(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        shape = InlineShape(inline)

        shape.alt_text = "a dog sleeping"

        assert shape.alt_text == "a dog sleeping"
        assert shape._inline.docPr.descr == "a dog sleeping"

    def it_populates_title_as_a_fallback_when_title_is_absent(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        shape = InlineShape(inline)

        shape.alt_text = "a dog sleeping"

        assert shape._inline.docPr.title == "a dog sleeping"

    def it_does_not_overwrite_an_existing_title(self):
        inline = cast(
            CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1,title=Fido}")
        )
        shape = InlineShape(inline)

        shape.alt_text = "a dog sleeping"

        # -- @title is preserved when it already had a value --
        assert shape._inline.docPr.title == "Fido"
        assert shape._inline.docPr.descr == "a dog sleeping"

    def it_does_not_touch_title_when_alt_text_is_cleared(self):
        inline = cast(
            CT_Inline,
            element("wp:inline/wp:docPr{id=1,name=P1,descr=x,title=Fido}"),
        )
        shape = InlineShape(inline)

        shape.alt_text = None

        assert shape._inline.docPr.title == "Fido"
        assert shape._inline.docPr.descr is None

    def it_strips_a_role_prefix_on_read(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        inline.docPr.descr = "[figure] Q3 chart"

        assert InlineShape(inline).alt_text == "Q3 chart"

    def it_preserves_the_role_prefix_when_updating_alt_text(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        inline.docPr.descr = "[logo] Acme"
        shape = InlineShape(inline)

        shape.alt_text = "Acme Corp"

        assert shape._inline.docPr.descr == "[logo] Acme Corp"
        assert shape.alt_text == "Acme Corp"
        assert shape.a11y_role == "logo"


class DescribeInlineShape_A11yRole:
    """Round-trip behaviour for ``InlineShape.a11y_role``."""

    def it_returns_None_when_no_prefix_is_present(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        inline.docPr.descr = "plain description"

        assert InlineShape(inline).a11y_role is None

    @pytest.mark.parametrize("role", ["figure", "decorative", "logo"])
    def it_round_trips_each_valid_role(self, role: str):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        shape = InlineShape(inline)

        shape.a11y_role = role

        assert shape.a11y_role == role

    def it_combines_role_and_alt_text_into_a_single_descr(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        shape = InlineShape(inline)

        shape.alt_text = "Q3 chart"
        shape.a11y_role = "figure"

        assert shape._inline.docPr.descr == "[figure] Q3 chart"

    def it_removes_the_prefix_when_set_to_None(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        inline.docPr.descr = "[logo] Acme"
        shape = InlineShape(inline)

        shape.a11y_role = None

        assert shape.a11y_role is None
        assert shape._inline.docPr.descr == "Acme"
        assert shape.alt_text == "Acme"

    def it_rejects_an_unknown_role(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        shape = InlineShape(inline)

        with pytest.raises(ValueError, match="a11y_role must be one of"):
            shape.a11y_role = "banner"

    def it_can_mark_a_shape_as_decorative_without_alt_text(self):
        inline = cast(CT_Inline, element("wp:inline/wp:docPr{id=1,name=P1}"))
        shape = InlineShape(inline)

        shape.a11y_role = "decorative"

        assert shape._inline.docPr.descr == "[decorative]"
        assert shape.a11y_role == "decorative"
        assert shape.alt_text == ""


class DescribeTable_AccessibilityTitleAndSummary:
    """Round-trip behaviour for ``Table.accessibility_title`` / ``_summary``."""

    def it_reads_the_tblCaption(self, document_: Mock):
        tbl = cast(CT_Tbl, element("w:tbl/w:tblPr/w:tblCaption{w:val=Revenue}"))
        table = Table(tbl, document_)

        assert table.accessibility_title == "Revenue"

    def it_sets_the_tblCaption(self, document_: Mock):
        table = Table(cast(CT_Tbl, element("w:tbl/w:tblPr")), document_)

        table.accessibility_title = "Revenue by quarter"

        assert table.accessibility_title == "Revenue by quarter"
        # -- setter is aliased onto alt_text; same underlying attribute --
        assert table.alt_text == "Revenue by quarter"

    def it_clears_the_tblCaption_when_set_to_None(self, document_: Mock):
        tbl = cast(CT_Tbl, element("w:tbl/w:tblPr/w:tblCaption{w:val=x}"))
        table = Table(tbl, document_)

        table.accessibility_title = None

        assert table.accessibility_title is None

    def it_reads_the_tblDescription(self, document_: Mock):
        tbl = cast(
            CT_Tbl,
            element("w:tbl/w:tblPr/w:tblDescription{w:val=Monthly totals}"),
        )
        table = Table(tbl, document_)

        assert table.accessibility_summary == "Monthly totals"

    def it_sets_the_tblDescription(self, document_: Mock):
        table = Table(cast(CT_Tbl, element("w:tbl/w:tblPr")), document_)

        table.accessibility_summary = "Table of monthly totals for FY26"

        assert table.accessibility_summary == "Table of monthly totals for FY26"
        assert table.alt_description == "Table of monthly totals for FY26"

    def it_clears_the_tblDescription_when_set_to_None(self, document_: Mock):
        tbl = cast(
            CT_Tbl, element("w:tbl/w:tblPr/w:tblDescription{w:val=x}")
        )
        table = Table(tbl, document_)

        table.accessibility_summary = None

        assert table.accessibility_summary is None

    # -- fixtures ----------------------------------------------------

    @pytest.fixture
    def document_(self, request: FixtureRequest):
        return instance_mock(request, Document)


class DescribeDocument_CheckAltText:
    """``Document.check_alt_text`` flags images/tables missing a11y data."""

    def it_returns_an_empty_list_when_no_issues_are_present(self):
        shape_ok = _stub_shape(alt_text="A cat on a mat", role=None)
        table_ok = _stub_table(summary="Summary of Q3 revenue")
        document = _stub_document([shape_ok], [table_ok])

        assert document.check_alt_text() == []

    def it_flags_a_shape_with_no_alt_text(self):
        shape_missing = _stub_shape(alt_text=None, role=None)
        document = _stub_document([shape_missing], [])

        assert document.check_alt_text() == [(shape_missing, "missing_alt_text")]

    def it_flags_a_shape_with_whitespace_only_alt_text(self):
        shape_blank = _stub_shape(alt_text="   ", role=None)
        document = _stub_document([shape_blank], [])

        assert document.check_alt_text() == [(shape_blank, "missing_alt_text")]

    def it_does_not_flag_a_decorative_shape_without_alt_text(self):
        decorative = _stub_shape(alt_text=None, role="decorative")
        document = _stub_document([decorative], [])

        assert document.check_alt_text() == []

    def it_flags_a_table_with_no_summary(self):
        table_no_summary = _stub_table(summary=None)
        document = _stub_document([], [table_no_summary])

        assert document.check_alt_text() == [(table_no_summary, "missing_summary")]

    def it_flags_a_table_with_whitespace_only_summary(self):
        table_blank = _stub_table(summary="  \t  ")
        document = _stub_document([], [table_blank])

        assert document.check_alt_text() == [(table_blank, "missing_summary")]

    def it_reports_issues_for_shapes_before_tables(self):
        bad_shape = _stub_shape(alt_text=None, role=None)
        bad_table = _stub_table(summary=None)
        document = _stub_document([bad_shape], [bad_table])

        issues = document.check_alt_text()

        assert issues == [
            (bad_shape, "missing_alt_text"),
            (bad_table, "missing_summary"),
        ]


# ====================================================================
# test helpers
# ====================================================================


def _stub_shape(alt_text: str | None, role: str | None):
    """Return a Mock that looks like an :class:`InlineShape` for a11y purposes."""
    shape = Mock(name="InlineShape")
    shape.alt_text = alt_text
    shape.a11y_role = role
    return shape


def _stub_table(summary: str | None):
    """Return a Mock that looks like a :class:`Table` for a11y purposes."""
    table = Mock(name="Table")
    table.accessibility_summary = summary
    return table


def _stub_document(shapes, tables):
    """Return a real Document subclass instance wired to the stubbed collections.

    We can't construct a real :class:`Document` here without an underlying
    package, so we install a minimal subclass that overrides the two
    collections :meth:`Document.check_alt_text` iterates over.
    """
    cls = type(
        "_StubDocument",
        (Document,),
        {
            "inline_shapes": property(lambda self: shapes),
            "tables": property(lambda self: tables),
        },
    )
    # -- bypass Document.__init__; we only need check_alt_text --
    return cls.__new__(cls)
