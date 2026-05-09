# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.mail_merge` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.mail_merge import (
    CT_Base64Binary,
    CT_DataSourceObject,
    CT_MailMergeDataType,
    CT_MailMergeDest,
    CT_MailMergeDocType,
    CT_MailMergeOdsoFMDFieldType,
    CT_MailMergeSourceType,
    CT_Odso,
    CT_OdsoFieldMapData,
    CT_OdsoRecipientData,
    CT_RecipientData,
    CT_TargetScreenSz,
)

from ..unitutil.cxml import element


# ---------------------------------------------------------------------------
# Val-wrapper CTs
# ---------------------------------------------------------------------------


class DescribeCT_MailMergeDocType:
    """Unit-test suite for `docx.oxml.mail_merge.CT_MailMergeDocType`."""

    @pytest.mark.parametrize(
        "val",
        ["catalog", "envelopes", "mailingLabels", "formLetters", "email", "fax"],
    )
    def it_reads_its_val_attr(self, val: str):
        mdt = cast(
            CT_MailMergeDocType, element(f"w:mainDocumentType{{w:val={val}}}")
        )
        assert mdt.val == val

    def it_registers_on_w_mainDocumentType(self):
        mdt = element("w:mainDocumentType{w:val=email}")
        assert isinstance(mdt, CT_MailMergeDocType)


class DescribeCT_MailMergeDataType:
    """Unit-test suite for `docx.oxml.mail_merge.CT_MailMergeDataType`."""

    @pytest.mark.parametrize(
        "val",
        ["textFile", "database", "spreadsheet", "query", "odbc", "native"],
    )
    def it_reads_its_val_attr(self, val: str):
        dt = cast(CT_MailMergeDataType, element(f"w:dataType{{w:val={val}}}"))
        assert dt.val == val

    def it_registers_on_w_dataType(self):
        dt = element("w:dataType{w:val=odbc}")
        assert isinstance(dt, CT_MailMergeDataType)


class DescribeCT_MailMergeDest:
    """Unit-test suite for `docx.oxml.mail_merge.CT_MailMergeDest`."""

    @pytest.mark.parametrize(
        "val", ["newDocument", "printer", "email", "fax"]
    )
    def it_reads_its_val_attr(self, val: str):
        dest = cast(CT_MailMergeDest, element(f"w:destination{{w:val={val}}}"))
        assert dest.val == val

    def it_registers_on_w_destination(self):
        dest = element("w:destination{w:val=email}")
        assert isinstance(dest, CT_MailMergeDest)


class DescribeCT_MailMergeSourceType:
    """Unit-test suite for `docx.oxml.mail_merge.CT_MailMergeSourceType`.

    ``w:type`` is a polymorphic QName owned by ``CT_SectType`` at the parser
    registry; this CT class is exported for programmatic construction and
    type-hinting of the ``w:odso/w:type`` slot. Verify the descriptor layer.
    """

    def it_exposes_a_required_val_attribute_descriptor(self):
        assert "val" in CT_MailMergeSourceType.__dict__ or any(
            "val" in base.__dict__ for base in CT_MailMergeSourceType.__mro__
        )


class DescribeCT_MailMergeOdsoFMDFieldType:
    """Unit-test suite for `docx.oxml.mail_merge.CT_MailMergeOdsoFMDFieldType`.

    Same polymorphic-QName caveat as :class:`CT_MailMergeSourceType` — the
    class is exported for the ``w:fieldMapData/w:type`` slot but not
    registered against ``w:type`` (which the parser already owns for
    ``CT_SectType``). Verify the descriptor layer.
    """

    def it_exposes_a_required_val_attribute_descriptor(self):
        assert "val" in CT_MailMergeOdsoFMDFieldType.__dict__ or any(
            "val" in base.__dict__
            for base in CT_MailMergeOdsoFMDFieldType.__mro__
        )


# ---------------------------------------------------------------------------
# CT_Base64Binary (w:uniqueTag)
# ---------------------------------------------------------------------------


class DescribeCT_Base64Binary:
    """Unit-test suite for `docx.oxml.mail_merge.CT_Base64Binary`."""

    def it_reads_its_val_attr(self):
        ut = cast(CT_Base64Binary, element("w:uniqueTag{w:val=abc123}"))
        assert ut.val == "abc123"

    def it_registers_on_w_uniqueTag(self):
        ut = element("w:uniqueTag{w:val=x}")
        assert isinstance(ut, CT_Base64Binary)


# ---------------------------------------------------------------------------
# CT_DataSourceObject (w:dataSource / w:headerSource / w:src)
# ---------------------------------------------------------------------------


class DescribeCT_DataSourceObject:
    """Unit-test suite for `docx.oxml.mail_merge.CT_DataSourceObject`."""

    def it_reads_its_rId(self):
        ds = cast(CT_DataSourceObject, element("w:dataSource{r:id=rId42}"))
        assert ds.rId == "rId42"

    def it_is_None_rId_when_not_present(self):
        ds = cast(CT_DataSourceObject, element("w:dataSource"))
        assert ds.rId is None

    def it_can_set_its_rId(self):
        ds = cast(CT_DataSourceObject, element("w:dataSource"))
        ds.rId = "rId7"
        assert ds.rId == "rId7"

    def it_registers_on_w_dataSource(self):
        assert isinstance(element("w:dataSource"), CT_DataSourceObject)

    def it_registers_on_w_headerSource(self):
        assert isinstance(element("w:headerSource"), CT_DataSourceObject)

    def it_registers_on_w_src(self):
        assert isinstance(element("w:src"), CT_DataSourceObject)


# ---------------------------------------------------------------------------
# CT_OdsoFieldMapData (w:fieldMapData)
# ---------------------------------------------------------------------------


class DescribeCT_OdsoFieldMapData:
    """Unit-test suite for `docx.oxml.mail_merge.CT_OdsoFieldMapData`."""

    def it_registers_on_w_fieldMapData(self):
        fmd = element("w:fieldMapData")
        assert isinstance(fmd, CT_OdsoFieldMapData)

    def it_parses_its_children_in_order(self):
        fmd = cast(
            CT_OdsoFieldMapData,
            element(
                "w:fieldMapData/(w:type{w:val=dbColumn},w:name{w:val=First},"
                "w:mappedName{w:val=FirstName},w:column{w:val=0})"
            ),
        )
        # All four descriptor-typed children should be reachable.
        assert fmd.type is not None
        assert fmd.name is not None
        assert fmd.mappedName is not None
        assert fmd.column is not None

    def it_round_trips_empty_fieldMapData(self):
        fmd = cast(CT_OdsoFieldMapData, element("w:fieldMapData"))
        assert fmd.type is None
        assert fmd.name is None
        assert fmd.mappedName is None
        assert fmd.column is None
        assert fmd.lid is None
        assert fmd.dynamicAddress is None


# ---------------------------------------------------------------------------
# CT_Odso (w:odso)
# ---------------------------------------------------------------------------


class DescribeCT_Odso:
    """Unit-test suite for `docx.oxml.mail_merge.CT_Odso`."""

    def it_registers_on_w_odso(self):
        odso = element("w:odso")
        assert isinstance(odso, CT_Odso)

    def it_reads_its_udl_table_colDelim_fHdr(self):
        odso = cast(
            CT_Odso,
            element(
                "w:odso/(w:udl{w:val=Provider},w:table{w:val=Cust},"
                "w:colDelim{w:val=44},w:fHdr)"
            ),
        )
        assert odso.udl is not None
        assert odso.table is not None
        assert odso.colDelim is not None
        assert odso.fHdr is not None

    def it_reads_its_src_as_a_DataSourceObject(self):
        odso = cast(
            CT_Odso, element("w:odso/w:src{r:id=rId9}")
        )
        assert isinstance(odso.src, CT_DataSourceObject)
        assert odso.src is not None
        assert odso.src.rId == "rId9"

    def it_reads_unbounded_fieldMapData_children(self):
        odso = cast(
            CT_Odso,
            element(
                "w:odso/(w:fieldMapData/w:name{w:val=A},"
                "w:fieldMapData/w:name{w:val=B},"
                "w:fieldMapData/w:name{w:val=C})"
            ),
        )
        assert len(odso.fieldMapData_lst) == 3

    def it_reads_unbounded_recipientData_refs(self):
        odso = cast(
            CT_Odso,
            element(
                "w:odso/(w:recipientData{r:id=rA},w:recipientData{r:id=rB})"
            ),
        )
        assert len(odso.recipientData_lst) == 2

    def it_has_empty_lists_when_absent(self):
        odso = cast(CT_Odso, element("w:odso"))
        assert odso.fieldMapData_lst == []
        assert odso.recipientData_lst == []


# ---------------------------------------------------------------------------
# CT_RecipientData (w:recipientData — the rich row variant)
# ---------------------------------------------------------------------------


class DescribeCT_RecipientData:
    """Unit-test suite for `docx.oxml.mail_merge.CT_RecipientData`."""

    def it_registers_on_w_recipientData(self):
        rd = element("w:recipientData")
        assert isinstance(rd, CT_RecipientData)

    def it_reads_its_active_column_uniqueTag_children(self):
        rd = cast(
            CT_RecipientData,
            element(
                "w:recipientData/(w:active,w:column{w:val=3},"
                "w:uniqueTag{w:val=Zm9v})"
            ),
        )
        assert rd.active is not None
        assert rd.column is not None
        assert rd.uniqueTag is not None
        assert cast(CT_Base64Binary, rd.uniqueTag).val == "Zm9v"

    def it_round_trips_empty_recipientData(self):
        rd = cast(CT_RecipientData, element("w:recipientData"))
        assert rd.active is None
        assert rd.column is None
        assert rd.uniqueTag is None


# ---------------------------------------------------------------------------
# CT_OdsoRecipientData (w:recipients)
# ---------------------------------------------------------------------------


class DescribeCT_OdsoRecipientData:
    """Unit-test suite for `docx.oxml.mail_merge.CT_OdsoRecipientData`."""

    def it_registers_on_w_recipients(self):
        rcp = element("w:recipients")
        assert isinstance(rcp, CT_OdsoRecipientData)

    def it_holds_an_unbounded_list_of_recipientData(self):
        rcp = cast(
            CT_OdsoRecipientData,
            element(
                "w:recipients/(w:recipientData/w:column{w:val=0},"
                "w:recipientData/w:column{w:val=1})"
            ),
        )
        assert len(rcp.recipientData_lst) == 2


# ---------------------------------------------------------------------------
# CT_TargetScreenSz (w:targetScreenSz)
# ---------------------------------------------------------------------------


class DescribeCT_TargetScreenSz:
    """Unit-test suite for `docx.oxml.mail_merge.CT_TargetScreenSz`."""

    @pytest.mark.parametrize(
        "val",
        [
            "544x376",
            "640x480",
            "720x512",
            "800x600",
            "1024x768",
            "1152x882",
            "1152x900",
            "1280x1024",
            "1600x1200",
            "1800x1440",
            "1920x1200",
        ],
    )
    def it_reads_its_val_attr(self, val: str):
        tss = cast(
            CT_TargetScreenSz, element(f"w:targetScreenSz{{w:val={val}}}")
        )
        assert tss.val == val

    def it_registers_on_w_targetScreenSz(self):
        tss = element("w:targetScreenSz{w:val=1024x768}")
        assert isinstance(tss, CT_TargetScreenSz)


# ---------------------------------------------------------------------------
# Integration: CT_MailMerge wires dataSource / headerSource / odso
# ---------------------------------------------------------------------------


class DescribeCT_MailMerge_OdsoIntegration:
    """Verify `CT_MailMerge` exposes the new descriptor slots for ODSO."""

    def it_exposes_dataSource_descriptor(self):
        from docx.oxml.settings import CT_MailMerge

        mm = cast(
            CT_MailMerge, element("w:mailMerge/w:dataSource{r:id=rId3}")
        )
        assert mm.dataSource is not None
        assert cast(CT_DataSourceObject, mm.dataSource).rId == "rId3"

    def it_exposes_headerSource_descriptor(self):
        from docx.oxml.settings import CT_MailMerge

        mm = cast(
            CT_MailMerge, element("w:mailMerge/w:headerSource{r:id=rId4}")
        )
        assert mm.headerSource is not None
        assert cast(CT_DataSourceObject, mm.headerSource).rId == "rId4"

    def it_exposes_odso_descriptor(self):
        from docx.oxml.settings import CT_MailMerge

        mm = cast(
            CT_MailMerge,
            element(
                "w:mailMerge/w:odso/w:udl{w:val=Provider}"
            ),
        )
        assert mm.odso is not None
        assert isinstance(mm.odso, CT_Odso)
