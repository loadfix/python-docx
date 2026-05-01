"""Test suite for the docx.api module."""

import io
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx.api import Document as DocumentFactoryFn
from docx.document import Document as DocumentCls
from docx.exceptions import EncryptedDocumentError
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.exceptions import PackageNotFoundError

from .unitutil.mock import FixtureRequest, Mock, class_mock, function_mock, instance_mock

_OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _make_malformed_docx_bytes() -> bytes:
    """Return a zip-packaged .docx whose `word/document.xml` is truncated mid-tag.

    The rest of the package is valid so recovery mode has something to graft
    the degraded document part onto.
    """
    tpl_path = Path(__file__).parent.parent / "src" / "docx" / "templates" / "default.docx"
    with open(tpl_path, "rb") as f:
        blob = f.read()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data[: len(data) // 2]  # -- truncate mid-element --
                zout.writestr(item, data)
    return out.getvalue()


def _make_valid_docx_bytes() -> bytes:
    tpl_path = Path(__file__).parent.parent / "src" / "docx" / "templates" / "default.docx"
    with open(tpl_path, "rb") as f:
        return f.read()


def _make_empty_document_xml_docx_bytes() -> bytes:
    """Return a valid .docx whose `word/document.xml` is an empty byte string.

    Empty content is unrecoverable even with ``recover=True`` — forces the
    stub-element fallback in ``XmlPart.load``.
    """
    tpl_path = Path(__file__).parent.parent / "src" / "docx" / "templates" / "default.docx"
    with open(tpl_path, "rb") as f:
        blob = f.read()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = b""
                zout.writestr(item, data)
    return out.getvalue()


class DescribeDocument:
    """Unit-test suite for `docx.api.Document` factory function."""

    def it_opens_a_docx_file(self, Package_: Mock, document_: Mock):
        document_part = Package_.open.return_value.main_document_part
        document_part.document = document_
        document_part.content_type = CT.WML_DOCUMENT_MAIN

        document = DocumentFactoryFn("foobar.docx")

        Package_.open.assert_called_once_with("foobar.docx", recover=False)
        assert document is document_

    def it_opens_the_default_docx_if_none_specified(
        self, _default_docx_path_: Mock, Package_: Mock, document_: Mock
    ):
        _default_docx_path_.return_value = "default-document.docx"
        document_part = Package_.open.return_value.main_document_part
        document_part.document = document_
        document_part.content_type = CT.WML_DOCUMENT_MAIN

        document = DocumentFactoryFn()

        Package_.open.assert_called_once_with("default-document.docx", recover=False)
        assert document is document_

    def it_opens_a_docm_file(self, Package_: Mock, document_: Mock):
        document_part = Package_.open.return_value.main_document_part
        document_part.document = document_
        document_part.content_type = CT.WML_DOCUMENT_MACRO

        document = DocumentFactoryFn("foobar.docm")

        Package_.open.assert_called_once_with("foobar.docm", recover=False)
        assert document is document_

    def it_raises_on_not_a_Word_file(self, Package_: Mock):
        Package_.open.return_value.main_document_part.content_type = "BOGUS"

        with pytest.raises(ValueError, match="file 'foobar.xlsx' is not a Word file,"):
            DocumentFactoryFn("foobar.xlsx")

    def it_raises_EncryptedDocumentError_on_password_protected_path(self, tmp_path):
        encrypted_path = tmp_path / "encrypted.docx"
        encrypted_path.write_bytes(_OLE_SIGNATURE + b"\x00" * 512)

        with pytest.raises(EncryptedDocumentError, match="msoffcrypto-tool"):
            DocumentFactoryFn(str(encrypted_path))

    def it_raises_EncryptedDocumentError_on_password_protected_stream(self):
        stream = io.BytesIO(_OLE_SIGNATURE + b"\x00" * 512)

        with pytest.raises(EncryptedDocumentError, match="password-protected"):
            DocumentFactoryFn(stream)

    def it_raises_on_malformed_document_xml_by_default(self):
        stream = io.BytesIO(_make_malformed_docx_bytes())

        with pytest.raises(etree.XMLSyntaxError):
            DocumentFactoryFn(stream)

    def it_opens_malformed_document_in_recover_mode(self):
        stream = io.BytesIO(_make_malformed_docx_bytes())

        document = DocumentFactoryFn(stream, recover=True)

        assert isinstance(document, DocumentCls)
        assert len(document.recovery_warnings) > 0
        assert all(isinstance(w, str) for w in document.recovery_warnings)

    def it_reports_no_warnings_for_valid_document_in_recover_mode(self):
        stream = io.BytesIO(_make_valid_docx_bytes())

        document = DocumentFactoryFn(stream, recover=True)

        assert document.recovery_warnings == []

    def it_recovery_mode_still_raises_for_invalid_zip(self, tmp_path):
        not_a_zip = tmp_path / "bogus.docx"
        not_a_zip.write_bytes(b"this is not a zip file")

        with pytest.raises(PackageNotFoundError):
            DocumentFactoryFn(str(not_a_zip), recover=True)

    def it_recovery_mode_still_raises_for_encrypted_docx(self, tmp_path):
        encrypted_path = tmp_path / "encrypted.docx"
        encrypted_path.write_bytes(_OLE_SIGNATURE + b"\x00" * 512)

        with pytest.raises(EncryptedDocumentError):
            DocumentFactoryFn(str(encrypted_path), recover=True)

    def it_defaults_recover_to_False_for_valid_document(self):
        stream = io.BytesIO(_make_valid_docx_bytes())

        document = DocumentFactoryFn(stream)

        assert document.recovery_warnings == []

    def it_falls_back_to_stub_when_document_xml_is_empty(self):
        stream = io.BytesIO(_make_empty_document_xml_docx_bytes())

        document = DocumentFactoryFn(stream, recover=True)

        assert isinstance(document, DocumentCls)
        assert document.paragraphs == []
        assert len(document.recovery_warnings) >= 1

    def it_passes_recover_True_through_to_Package_open(
        self, Package_: Mock, document_: Mock
    ):
        document_part = Package_.open.return_value.main_document_part
        document_part.document = document_
        document_part.content_type = CT.WML_DOCUMENT_MAIN

        DocumentFactoryFn("foobar.docx", recover=True)

        Package_.open.assert_called_once_with("foobar.docx", recover=True)

    # -- fixtures --------------------------------------------------------------------------------

    @pytest.fixture
    def _default_docx_path_(self, request: FixtureRequest):
        return function_mock(request, "docx.api._default_docx_path")

    @pytest.fixture
    def document_(self, request: FixtureRequest):
        return instance_mock(request, DocumentCls)

    @pytest.fixture
    def Package_(self, request: FixtureRequest):
        return class_mock(request, "docx.api.Package")
