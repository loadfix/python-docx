"""Unit tests for Flat-OPC (``<pkg:package>``) read and write.

Closes upstream#892.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx.api import Document as DocumentFactoryFn
from docx.opc.flat_opc import (
    PKG_NS,
    PKG_PACKAGE,
    PKG_PART,
    expand_flat_opc_to_zip_stream,
    looks_like_flat_opc,
    write_flat_opc,
)


def _default_template_bytes() -> bytes:
    tpl = Path(__file__).parent.parent / "src" / "docx" / "templates" / "default.docx"
    with open(tpl, "rb") as f:
        return f.read()


class DescribeLooksLikeFlatOpc:
    def it_returns_True_for_a_pkg_package_stream(self):
        xml = (
            '<?xml version="1.0"?><pkg:package xmlns:pkg="%s"/>'
            % PKG_NS
        ).encode("utf-8")
        stream = io.BytesIO(xml)

        assert looks_like_flat_opc(stream) is True
        # -- stream position restored for downstream reader --
        assert stream.tell() == 0

    def it_returns_False_for_a_zip_stream(self):
        stream = io.BytesIO(_default_template_bytes())

        assert looks_like_flat_opc(stream) is False

    def it_returns_False_for_a_missing_path(self):
        assert looks_like_flat_opc("/nope/does/not/exist") is False

    def it_returns_False_for_junk_bytes(self):
        stream = io.BytesIO(b"not xml at all")

        assert looks_like_flat_opc(stream) is False

    def it_returns_False_for_non_streamable_input(self):
        # -- mock-ish object without read/seek/tell should be rejected cleanly --
        assert looks_like_flat_opc(object()) is False


class DescribeFlatOpcWriteRead:
    def it_round_trips_a_document_through_flat_opc(self, tmp_path):
        flat_path = tmp_path / "doc.xml"
        zip_blob = _default_template_bytes()

        write_flat_opc(str(flat_path), zip_blob)

        # -- output must be a valid `<pkg:package>` XML document --
        data = flat_path.read_bytes()
        assert data.startswith(b"<?xml")
        root = etree.fromstring(data)
        assert root.tag == PKG_PACKAGE
        parts = root.findall(PKG_PART)
        assert len(parts) >= 1
        # -- each part carries a pkg:name attribute --
        for part in parts:
            assert part.get("{%s}name" % PKG_NS)

    def it_expands_flat_opc_back_to_a_zip_stream(self, tmp_path):
        flat_path = tmp_path / "doc.xml"
        zip_blob = _default_template_bytes()
        write_flat_opc(str(flat_path), zip_blob)

        expanded = expand_flat_opc_to_zip_stream(str(flat_path))

        with zipfile.ZipFile(expanded, "r") as zf:
            names = zf.namelist()
            assert "[Content_Types].xml" in names
            assert "word/document.xml" in names

    def it_opens_a_flat_opc_package_via_Document(self, tmp_path):
        flat_path = tmp_path / "doc.xml"
        zip_blob = _default_template_bytes()
        write_flat_opc(str(flat_path), zip_blob)

        document = DocumentFactoryFn(str(flat_path))

        assert document is not None

    def it_saves_as_flat_opc_when_flag_is_set(self, tmp_path):
        document = DocumentFactoryFn()
        flat_path = tmp_path / "out.xml"

        document.save(str(flat_path), flat_opc=True)

        data = flat_path.read_bytes()
        assert data.startswith(b"<?xml")
        root = etree.fromstring(data)
        assert root.tag == PKG_PACKAGE

    def it_rejects_non_flat_opc_xml(self, tmp_path):
        bad = tmp_path / "bad.xml"
        bad.write_bytes(b'<?xml version="1.0"?><not_a_package/>')

        with pytest.raises(ValueError, match="pkg:package"):
            expand_flat_opc_to_zip_stream(str(bad))
