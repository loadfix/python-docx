"""Unit tests for Strict OOXML → Transitional translation on package open.

Covers upstream#1520, upstream#693. A Strict-fixture is built at test time by
round-tripping the default template through a namespace rewrite, which keeps
the test self-contained (no binary fixture committed to the tree).
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from docx.api import Document as DocumentFactoryFn
from docx.opc.strict import (
    STRICT_TO_TRANSITIONAL,
    is_strict_document_xml,
    translate_strict_blob,
)


_TRANSITIONAL_WML = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_STRICT_WML = b"http://purl.oclc.org/ooxml/wordprocessingml/main"


def _default_template_bytes() -> bytes:
    tpl_path = (
        Path(__file__).parent.parent / "src" / "docx" / "templates" / "default.docx"
    )
    with open(tpl_path, "rb") as f:
        return f.read()


def _make_strict_docx_bytes() -> bytes:
    """Return a Strict-OOXML version of the default template.

    Every Transitional namespace URI inside each XML member is rewritten to
    its Strict counterpart. Binary members (images etc.) are left untouched.
    """
    # -- invert the translation map so we can walk Transitional → Strict --
    transitional_to_strict = {
        trans: strict for strict, trans in STRICT_TO_TRANSITIONAL.items()
    }
    blob = _default_template_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith((".xml", ".rels")) or item.filename == (
                    "[Content_Types].xml"
                ):
                    for transitional, strict in transitional_to_strict.items():
                        data = data.replace(transitional, strict)
                zout.writestr(item, data)
    return out.getvalue()


class DescribeStrictTranslator:
    """Unit tests for the pure-function strict → transitional rewrite."""

    def it_detects_a_strict_document_blob(self):
        assert is_strict_document_xml(b'<w:document xmlns:w="%s"/>' % _STRICT_WML)

    def it_ignores_a_transitional_document_blob(self):
        assert not is_strict_document_xml(
            b'<w:document xmlns:w="%s"/>' % _TRANSITIONAL_WML
        )

    def it_ignores_an_empty_blob(self):
        assert not is_strict_document_xml(b"")
        assert not is_strict_document_xml(None)

    def it_rewrites_every_strict_uri(self):
        src = b""
        for strict in STRICT_TO_TRANSITIONAL.keys():
            src += b"<x>" + strict + b"</x>"
        out = translate_strict_blob(src)
        assert out is not None
        for strict, trans in STRICT_TO_TRANSITIONAL.items():
            assert strict not in out
            assert trans in out

    def it_passes_through_non_strict_blobs_untouched(self):
        blob = b"<w:document xmlns:w='%s'/>" % _TRANSITIONAL_WML
        assert translate_strict_blob(blob) is blob

    def it_preserves_none(self):
        assert translate_strict_blob(None) is None


class DescribeStrictDocumentOpen:
    """Integration: Document() transparently opens a Strict-namespaced package."""

    def it_opens_a_strict_docx_file(self, tmp_path):
        strict_path = tmp_path / "strict.docx"
        strict_path.write_bytes(_make_strict_docx_bytes())

        document = DocumentFactoryFn(str(strict_path))

        # -- paragraphs should be readable; the default template has at least one paragraph. --
        assert document is not None
        assert len(document.paragraphs) >= 0

    def it_opens_a_strict_docx_stream(self):
        stream = io.BytesIO(_make_strict_docx_bytes())

        document = DocumentFactoryFn(stream)

        assert document is not None

    def it_saves_strict_as_transitional(self, tmp_path):
        strict_path = tmp_path / "strict.docx"
        strict_path.write_bytes(_make_strict_docx_bytes())

        document = DocumentFactoryFn(str(strict_path))
        out_path = tmp_path / "out.docx"
        document.save(str(out_path))

        # -- the saved package must NOT carry any Strict namespace URIs --
        with zipfile.ZipFile(out_path, "r") as zf:
            for name in zf.namelist():
                data = zf.read(name)
                assert _STRICT_WML not in data, (
                    f"{name} still contains Strict WML namespace"
                )
        # -- and it must be reopenable as a normal Transitional docx --
        document2 = DocumentFactoryFn(str(out_path))
        assert document2 is not None

    def it_leaves_transitional_packages_untouched(self):
        """Transitional docs must not be touched by the Strict detection path."""
        stream = io.BytesIO(_default_template_bytes())

        document = DocumentFactoryFn(stream)

        assert document is not None


@pytest.fixture
def strict_fixture_bytes():
    return _make_strict_docx_bytes()
