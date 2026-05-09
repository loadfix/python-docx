"""Unit-test suite for `docx.settings.MailMerge` and related oxml classes."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.text import (
    WD_MAIL_MERGE_DATA_TYPE,
    WD_MAIL_MERGE_DESTINATION,
    WD_MAIL_MERGE_DOCUMENT_TYPE,
    WD_MAIL_MERGE_TYPE,
    WD_ODSO_TYPE,
)
from docx.oxml.mail_merge import CT_Odso
from docx.oxml.settings import CT_MailMerge, CT_Settings
from docx.settings import MailMerge, OdsoSettings, Settings

from .unitutil.cxml import element


class DescribeMailMerge:
    """Unit-test suite for `docx.settings.MailMerge` proxy."""

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("w:mailMerge", None),
            ("w:mailMerge/w:mainDocumentType{w:val=formLetters}", WD_MAIL_MERGE_TYPE.FORM_LETTERS),
            ("w:mailMerge/w:mainDocumentType{w:val=email}", WD_MAIL_MERGE_TYPE.EMAIL),
            ("w:mailMerge/w:mainDocumentType{w:val=catalog}", WD_MAIL_MERGE_TYPE.CATALOG),
            ("w:mailMerge/w:mainDocumentType{w:val=envelopes}", WD_MAIL_MERGE_TYPE.ENVELOPES),
            ("w:mailMerge/w:mainDocumentType{w:val=mailingLabels}", WD_MAIL_MERGE_TYPE.MAILING_LABELS),
            ("w:mailMerge/w:mainDocumentType{w:val=fax}", WD_MAIL_MERGE_TYPE.FAX),
        ],
    )
    def it_reads_the_main_document_type(self, cxml, expected):
        mm = cast(CT_MailMerge, element(cxml))
        assert MailMerge(mm).main_document_type == expected

    def it_writes_the_main_document_type(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        MailMerge(mm).main_document_type = WD_MAIL_MERGE_TYPE.EMAIL
        assert mm.xpath("./w:mainDocumentType")[0].val == "email"

    def it_clears_the_main_document_type(self):
        mm = cast(CT_MailMerge, element("w:mailMerge/w:mainDocumentType{w:val=email}"))
        MailMerge(mm).main_document_type = None
        assert mm.xpath("./w:mainDocumentType") == []

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("w:mailMerge", None),
            ("w:mailMerge/w:destination{w:val=newDocument}", WD_MAIL_MERGE_DESTINATION.NEW_DOCUMENT),
            ("w:mailMerge/w:destination{w:val=printer}", WD_MAIL_MERGE_DESTINATION.PRINTER),
            ("w:mailMerge/w:destination{w:val=email}", WD_MAIL_MERGE_DESTINATION.EMAIL),
            ("w:mailMerge/w:destination{w:val=fax}", WD_MAIL_MERGE_DESTINATION.FAX),
        ],
    )
    def it_reads_the_destination(self, cxml, expected):
        mm = cast(CT_MailMerge, element(cxml))
        assert MailMerge(mm).destination == expected

    def it_writes_the_destination(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        MailMerge(mm).destination = WD_MAIL_MERGE_DESTINATION.EMAIL
        assert mm.xpath("./w:destination")[0].val == "email"

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("w:mailMerge", None),
            ("w:mailMerge/w:dataType{w:val=spreadsheet}", WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET),
            ("w:mailMerge/w:dataType{w:val=odbc}", WD_MAIL_MERGE_DATA_TYPE.ODBC),
        ],
    )
    def it_reads_the_data_type(self, cxml, expected):
        mm = cast(CT_MailMerge, element(cxml))
        assert MailMerge(mm).data_type == expected

    def it_returns_None_for_unknown_data_type_xml(self):
        mm = cast(CT_MailMerge, element("w:mailMerge/w:dataType{w:val=zzz}"))
        assert MailMerge(mm).data_type is None

    def it_round_trips_connect_string(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        proxy = MailMerge(mm)
        proxy.connect_string = "DSN=MyData;UID=user"
        assert proxy.connect_string == "DSN=MyData;UID=user"
        proxy.connect_string = None
        assert proxy.connect_string is None
        assert mm.xpath("./w:connectString") == []

    def it_round_trips_query(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        proxy = MailMerge(mm)
        proxy.query = "SELECT FirstName, LastName FROM Customers"
        assert proxy.query == "SELECT FirstName, LastName FROM Customers"

    def it_round_trips_mail_subject(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        proxy = MailMerge(mm)
        proxy.mail_subject = "Hello"
        assert proxy.mail_subject == "Hello"

    def it_round_trips_address_field_name(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        proxy = MailMerge(mm)
        proxy.address_field_name = "Email_Address"
        assert proxy.address_field_name == "Email_Address"

    def it_round_trips_active_record_as_int(self):
        mm = cast(CT_MailMerge, element("w:mailMerge/w:activeRecord{w:val=7}"))
        proxy = MailMerge(mm)
        assert proxy.active_record == 7
        proxy.active_record = 12
        assert proxy.active_record == 12
        proxy.active_record = None
        assert proxy.active_record is None

    def it_round_trips_check_errors_as_int(self):
        mm = cast(CT_MailMerge, element("w:mailMerge/w:checkErrors{w:val=2}"))
        assert MailMerge(mm).check_errors == 2

    def it_returns_None_for_bad_active_record_value(self):
        mm = cast(CT_MailMerge, element("w:mailMerge/w:activeRecord{w:val=abc}"))
        assert MailMerge(mm).active_record is None

    @pytest.mark.parametrize(
        ("flag", "cxml"),
        [
            ("link_to_query", "w:mailMerge/w:linkToQuery"),
            ("do_not_suppress_blank_lines", "w:mailMerge/w:doNotSuppressBlankLines"),
            ("mail_as_attachment", "w:mailMerge/w:mailAsAttachment"),
            ("view_merged_data", "w:mailMerge/w:viewMergedData"),
        ],
    )
    def it_reads_bool_flags_as_True_when_present(self, flag, cxml):
        mm = cast(CT_MailMerge, element(cxml))
        assert getattr(MailMerge(mm), flag) is True

    @pytest.mark.parametrize(
        ("flag",),
        [
            ("link_to_query",),
            ("do_not_suppress_blank_lines",),
            ("mail_as_attachment",),
            ("view_merged_data",),
        ],
    )
    def it_reads_bool_flags_as_False_when_absent(self, flag):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        assert getattr(MailMerge(mm), flag) is False

    def it_sets_a_bool_flag(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        proxy = MailMerge(mm)
        proxy.mail_as_attachment = True
        assert proxy.mail_as_attachment is True
        proxy.mail_as_attachment = False
        assert proxy.mail_as_attachment is False
        assert mm.xpath("./w:mailAsAttachment") == []


class DescribeSettings_mail_merge:
    """Integration of MailMerge with `Settings`."""

    def it_returns_None_when_no_mailMerge_element(self):
        settings = cast(CT_Settings, element("w:settings"))
        assert Settings(settings).mail_merge is None

    def it_returns_the_proxy_when_element_is_present(self):
        settings = cast(
            CT_Settings,
            element("w:settings/w:mailMerge/w:mainDocumentType{w:val=formLetters}"),
        )
        proxy = Settings(settings).mail_merge
        assert proxy is not None
        assert proxy.main_document_type == WD_MAIL_MERGE_TYPE.FORM_LETTERS

    def it_can_enable_mail_merge_with_defaults(self):
        settings = cast(CT_Settings, element("w:settings"))
        proxy = Settings(settings).enable_mail_merge()
        assert proxy.main_document_type == WD_MAIL_MERGE_TYPE.FORM_LETTERS
        assert settings.xpath("./w:mailMerge") != []

    def it_can_enable_mail_merge_with_args(self):
        settings = cast(CT_Settings, element("w:settings"))
        proxy = Settings(settings).enable_mail_merge(
            main_document_type=WD_MAIL_MERGE_TYPE.EMAIL,
            destination=WD_MAIL_MERGE_DESTINATION.EMAIL,
            data_type=WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET,
            connect_string="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=...",
            query="SELECT * FROM [Sheet1$]",
            mail_subject="Hi",
            address_field_name="Email",
        )
        assert proxy.main_document_type == WD_MAIL_MERGE_TYPE.EMAIL
        assert proxy.destination == WD_MAIL_MERGE_DESTINATION.EMAIL
        assert proxy.data_type == WD_MAIL_MERGE_DATA_TYPE.SPREADSHEET
        assert proxy.connect_string == (
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=..."
        )
        assert proxy.query == "SELECT * FROM [Sheet1$]"
        assert proxy.mail_subject == "Hi"
        assert proxy.address_field_name == "Email"

    def it_can_disable_mail_merge(self):
        settings = cast(
            CT_Settings,
            element("w:settings/w:mailMerge/w:mainDocumentType{w:val=formLetters}"),
        )
        Settings(settings).disable_mail_merge()
        assert settings.xpath("./w:mailMerge") == []

    def it_is_idempotent_to_disable_when_absent(self):
        settings = cast(CT_Settings, element("w:settings"))
        Settings(settings).disable_mail_merge()  # no-op, no raise
        assert settings.xpath("./w:mailMerge") == []


class DescribeMailMerge_data_source:
    """`MailMerge.data_source` — rId reference to the merge data-source part."""

    def it_reads_None_when_no_dataSource_child(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        assert MailMerge(mm).data_source is None

    def it_reads_the_rId(self):
        mm = cast(CT_MailMerge, element("w:mailMerge/w:dataSource{r:id=rId5}"))
        assert MailMerge(mm).data_source == "rId5"

    def it_round_trips_the_rId(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        proxy = MailMerge(mm)
        proxy.data_source = "rId9"
        assert proxy.data_source == "rId9"
        assert mm.xpath("./w:dataSource")[0].rId == "rId9"

    def it_can_clear_the_rId(self):
        mm = cast(CT_MailMerge, element("w:mailMerge/w:dataSource{r:id=rId5}"))
        MailMerge(mm).data_source = None
        assert mm.xpath("./w:dataSource") == []

    def it_round_trips_header_source(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        proxy = MailMerge(mm)
        proxy.header_source = "rId7"
        assert proxy.header_source == "rId7"
        proxy.header_source = None
        assert proxy.header_source is None


class DescribeOdsoSettings:
    """`OdsoSettings` proxy — dedicated unit tests for the ODSO sub-block."""

    # -- udl / table --------------------------------------------------------

    def it_round_trips_udl(self):
        odso = cast(CT_Odso, element("w:odso"))
        proxy = OdsoSettings(odso)
        proxy.udl = "my.udl"
        assert proxy.udl == "my.udl"
        assert odso.xpath("./w:udl")[0].get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
        ) == "my.udl"
        proxy.udl = None
        assert proxy.udl is None
        assert odso.xpath("./w:udl") == []

    def it_round_trips_table(self):
        odso = cast(CT_Odso, element("w:odso/w:table{w:val=Sheet1}"))
        proxy = OdsoSettings(odso)
        assert proxy.table == "Sheet1"
        proxy.table = "Customers"
        assert proxy.table == "Customers"

    # -- src (relationship reference) --------------------------------------

    def it_reads_None_src_when_absent(self):
        odso = cast(CT_Odso, element("w:odso"))
        assert OdsoSettings(odso).src is None

    def it_round_trips_src_rId(self):
        odso = cast(CT_Odso, element("w:odso"))
        proxy = OdsoSettings(odso)
        proxy.src = "rId12"
        assert proxy.src == "rId12"
        assert odso.xpath("./w:src")[0].rId == "rId12"
        proxy.src = None
        assert odso.xpath("./w:src") == []

    # -- column_delimiter ---------------------------------------------------

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("w:odso", None),
            ("w:odso/w:colDelim{w:val=44}", 44),
            ("w:odso/w:colDelim{w:val=9}", 9),
        ],
    )
    def it_reads_column_delimiter(self, cxml, expected):
        odso = cast(CT_Odso, element(cxml))
        assert OdsoSettings(odso).column_delimiter == expected

    def it_writes_column_delimiter(self):
        odso = cast(CT_Odso, element("w:odso"))
        proxy = OdsoSettings(odso)
        proxy.column_delimiter = 124  # pipe
        assert proxy.column_delimiter == 124

    def it_returns_None_for_bad_column_delimiter(self):
        odso = cast(CT_Odso, element("w:odso/w:colDelim{w:val=abc}"))
        assert OdsoSettings(odso).column_delimiter is None

    # -- type (WD_ODSO_TYPE) ------------------------------------------------

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("database", WD_ODSO_TYPE.DATABASE),
            ("addressBook", WD_ODSO_TYPE.ADDRESS_BOOK),
            ("text", WD_ODSO_TYPE.TEXT),
            ("native", WD_ODSO_TYPE.NATIVE),
            ("legacy", WD_ODSO_TYPE.LEGACY),
        ],
    )
    def it_reads_the_type(self, raw, expected):
        odso = cast(CT_Odso, element(f"w:odso/w:type{{w:val={raw}}}"))
        assert OdsoSettings(odso).type == expected

    def it_returns_None_type_when_absent(self):
        odso = cast(CT_Odso, element("w:odso"))
        assert OdsoSettings(odso).type is None

    def it_returns_None_for_unknown_type_xml(self):
        odso = cast(CT_Odso, element("w:odso/w:type{w:val=unknown}"))
        assert OdsoSettings(odso).type is None

    def it_writes_the_type(self):
        odso = cast(CT_Odso, element("w:odso"))
        proxy = OdsoSettings(odso)
        proxy.type = WD_ODSO_TYPE.ADDRESS_BOOK
        assert proxy.type == WD_ODSO_TYPE.ADDRESS_BOOK

    def it_clears_the_type(self):
        odso = cast(CT_Odso, element("w:odso/w:type{w:val=database}"))
        OdsoSettings(odso).type = None
        assert odso.xpath("./w:type") == []

    # -- first_row_has_column_names (fHdr) ---------------------------------

    def it_reads_False_when_fHdr_absent(self):
        odso = cast(CT_Odso, element("w:odso"))
        assert OdsoSettings(odso).first_row_has_column_names is False

    def it_reads_True_when_fHdr_present_no_val(self):
        odso = cast(CT_Odso, element("w:odso/w:fHdr"))
        assert OdsoSettings(odso).first_row_has_column_names is True

    def it_round_trips_fHdr(self):
        odso = cast(CT_Odso, element("w:odso"))
        proxy = OdsoSettings(odso)
        proxy.first_row_has_column_names = True
        assert proxy.first_row_has_column_names is True
        proxy.first_row_has_column_names = False
        assert proxy.first_row_has_column_names is False
        assert odso.xpath("./w:fHdr") == []

    # -- field_mapping ------------------------------------------------------

    def it_reads_empty_mapping_when_no_fieldMapData(self):
        odso = cast(CT_Odso, element("w:odso"))
        assert OdsoSettings(odso).field_mapping == {}

    def it_reads_a_populated_field_mapping(self):
        odso = cast(
            CT_Odso,
            element(
                "w:odso/("
                "w:fieldMapData/(w:name{w:val=FirstName},w:mappedName{w:val=First_Name}),"
                "w:fieldMapData/(w:name{w:val=LastName},w:mappedName{w:val=Last_Name})"
                ")"
            ),
        )
        assert OdsoSettings(odso).field_mapping == {
            "FirstName": "First_Name",
            "LastName": "Last_Name",
        }

    def it_writes_a_field_mapping_replacing_any_existing(self):
        odso = cast(
            CT_Odso,
            element(
                "w:odso/w:fieldMapData/(w:name{w:val=Old},w:mappedName{w:val=Legacy})"
            ),
        )
        proxy = OdsoSettings(odso)
        proxy.field_mapping = {"FirstName": "FN", "LastName": "LN"}
        # -- The old mapping is gone; two new ones are present --
        assert proxy.field_mapping == {"FirstName": "FN", "LastName": "LN"}
        assert len(odso.xpath("./w:fieldMapData")) == 2

    def it_skips_records_missing_name_or_mappedName(self):
        odso = cast(
            CT_Odso,
            element(
                "w:odso/("
                "w:fieldMapData/w:name{w:val=NoMap},"
                "w:fieldMapData/(w:name{w:val=Good},w:mappedName{w:val=Column1})"
                ")"
            ),
        )
        assert OdsoSettings(odso).field_mapping == {"Good": "Column1"}

    def it_clears_field_mapping_by_assigning_None(self):
        odso = cast(
            CT_Odso,
            element(
                "w:odso/w:fieldMapData/(w:name{w:val=FN},w:mappedName{w:val=First})"
            ),
        )
        OdsoSettings(odso).field_mapping = None
        assert odso.xpath("./w:fieldMapData") == []


class DescribeMailMerge_odso:
    """Integration of `MailMerge` with `OdsoSettings`."""

    def it_returns_None_when_no_odso_child(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        assert MailMerge(mm).odso is None

    def it_returns_the_proxy_when_odso_is_present(self):
        mm = cast(
            CT_MailMerge,
            element("w:mailMerge/w:odso/w:udl{w:val=foo.udl}"),
        )
        proxy = MailMerge(mm).odso
        assert proxy is not None
        assert proxy.udl == "foo.udl"

    def it_can_add_an_odso_block(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        odso = MailMerge(mm).add_odso()
        odso.type = WD_ODSO_TYPE.DATABASE
        assert MailMerge(mm).odso is not None
        assert MailMerge(mm).odso.type == WD_ODSO_TYPE.DATABASE

    def it_can_remove_the_odso_block(self):
        mm = cast(
            CT_MailMerge,
            element("w:mailMerge/w:odso/w:udl{w:val=foo.udl}"),
        )
        MailMerge(mm).remove_odso()
        assert mm.xpath("./w:odso") == []


class DescribeWD_MAIL_MERGE_DOCUMENT_TYPE:
    """The long-form alias for `WD_MAIL_MERGE_TYPE`."""

    def it_is_an_alias_for_WD_MAIL_MERGE_TYPE(self):
        assert WD_MAIL_MERGE_DOCUMENT_TYPE is WD_MAIL_MERGE_TYPE


class DescribeMailMerge_full_roundtrip:
    """End-to-end: create, populate, and round-trip a full mail-merge block."""

    def it_round_trips_a_full_config_through_xml(self):
        mm = cast(CT_MailMerge, element("w:mailMerge"))
        proxy = MailMerge(mm)
        proxy.main_document_type = WD_MAIL_MERGE_DOCUMENT_TYPE.EMAIL
        proxy.destination = WD_MAIL_MERGE_DESTINATION.EMAIL
        proxy.data_type = WD_MAIL_MERGE_DATA_TYPE.ODBC
        proxy.connect_string = "DSN=Customers"
        proxy.query = "SELECT * FROM main"
        proxy.data_source = "rId99"

        odso = proxy.add_odso()
        odso.udl = "customers.udl"
        odso.table = "Customers"
        odso.src = "rId42"
        odso.column_delimiter = 44
        odso.type = WD_ODSO_TYPE.DATABASE
        odso.first_row_has_column_names = True
        odso.field_mapping = {"First": "FirstName", "Last": "LastName"}

        # Serialize → parse → re-open — the document round-trips byte-wise.
        from lxml import etree as _etree

        from docx.oxml.parser import parse_xml

        serialised = _etree.tostring(mm)
        reloaded = cast(CT_MailMerge, parse_xml(serialised))
        reproxy = MailMerge(reloaded)
        assert reproxy.main_document_type == WD_MAIL_MERGE_DOCUMENT_TYPE.EMAIL
        assert reproxy.destination == WD_MAIL_MERGE_DESTINATION.EMAIL
        assert reproxy.data_type == WD_MAIL_MERGE_DATA_TYPE.ODBC
        assert reproxy.connect_string == "DSN=Customers"
        assert reproxy.query == "SELECT * FROM main"
        assert reproxy.data_source == "rId99"

        reodso = reproxy.odso
        assert reodso is not None
        assert reodso.udl == "customers.udl"
        assert reodso.table == "Customers"
        assert reodso.src == "rId42"
        assert reodso.column_delimiter == 44
        assert reodso.type == WD_ODSO_TYPE.DATABASE
        assert reodso.first_row_has_column_names is True
        assert reodso.field_mapping == {"First": "FirstName", "Last": "LastName"}
