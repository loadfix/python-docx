"""Unit-test suite for `docx.settings.MailMerge` and related oxml classes."""

from __future__ import annotations

from typing import cast

import pytest

from docx.enum.text import (
    WD_MAIL_MERGE_DATA_TYPE,
    WD_MAIL_MERGE_DESTINATION,
    WD_MAIL_MERGE_TYPE,
)
from docx.oxml.settings import CT_MailMerge, CT_Settings
from docx.settings import MailMerge, Settings

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
