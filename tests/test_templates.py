"""Unit tests for ``.dotx`` / ``.dotm`` template loading and ``from_template``.

Closes upstream#1532, upstream#363, upstream-PR#1537, upstream-PR#522,
upstream-PR#523.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from docx import from_template
from docx.api import Document as DocumentFactoryFn
from docx.opc.constants import CONTENT_TYPE as CT


def _default_template_bytes() -> bytes:
    tpl_path = (
        Path(__file__).parent.parent / "src" / "docx" / "templates" / "default.docx"
    )
    with open(tpl_path, "rb") as f:
        return f.read()


def _make_template_bytes(source_ct: str, target_ct: str) -> bytes:
    """Return a new package whose main-document content type is `target_ct`.

    The caller passes `source_ct` = the content type currently in the template
    (``WML_DOCUMENT_MAIN`` for the bundled default) and `target_ct` = the
    content type we want to swap in. Every override in ``[Content_Types].xml``
    is inspected and the matching entry rewritten.
    """
    blob = _default_template_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zin:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.replace(
                        source_ct.encode("ascii"), target_ct.encode("ascii")
                    )
                zout.writestr(item, data)
    return out.getvalue()


class DescribeDotxSupport:
    """Open a ``.dotx`` template directly via :func:`Document`."""

    def it_opens_a_dotx_template(self, tmp_path):
        dotx = tmp_path / "tpl.dotx"
        dotx.write_bytes(_make_template_bytes(CT.WML_DOCUMENT_MAIN, CT.WML_TEMPLATE_MAIN))

        document = DocumentFactoryFn(str(dotx))

        assert document is not None

    def it_opens_a_dotm_template(self, tmp_path):
        dotm = tmp_path / "tpl.dotm"
        dotm.write_bytes(
            _make_template_bytes(CT.WML_DOCUMENT_MAIN, CT.WML_TEMPLATE_MACRO)
        )

        document = DocumentFactoryFn(str(dotm))

        assert document is not None

    def it_still_rejects_non_Word_content_types(self, tmp_path):
        bogus = tmp_path / "tpl.xlsx"
        bogus.write_bytes(
            _make_template_bytes(CT.WML_DOCUMENT_MAIN, CT.SML_SHEET_MAIN)
        )

        with pytest.raises(ValueError, match="not a Word file"):
            DocumentFactoryFn(str(bogus))


class DescribeFromTemplate:
    """``from_template`` switches template content-type to the document variant."""

    def it_swaps_dotx_to_docx_content_type(self, tmp_path):
        dotx_bytes = _make_template_bytes(CT.WML_DOCUMENT_MAIN, CT.WML_TEMPLATE_MAIN)
        stream = io.BytesIO(dotx_bytes)

        document = from_template(stream)

        assert document.part.content_type == CT.WML_DOCUMENT_MAIN

    def it_swaps_dotm_to_docm_content_type(self, tmp_path):
        dotm_bytes = _make_template_bytes(CT.WML_DOCUMENT_MAIN, CT.WML_TEMPLATE_MACRO)
        stream = io.BytesIO(dotm_bytes)

        document = from_template(stream)

        assert document.part.content_type == CT.WML_DOCUMENT_MACRO

    def it_rejects_non_template_packages(self):
        # -- a plain .docx is not a template --
        stream = io.BytesIO(_default_template_bytes())

        with pytest.raises(ValueError, match="not a Word template"):
            from_template(stream)

    def it_is_exposed_as_Document_from_template(self):
        from docx.api import Document as DocumentFn

        assert hasattr(DocumentFn, "from_template")
        assert DocumentFn.from_template is from_template

    def it_produces_a_document_that_saves_as_docx(self, tmp_path):
        dotx_bytes = _make_template_bytes(CT.WML_DOCUMENT_MAIN, CT.WML_TEMPLATE_MAIN)
        stream = io.BytesIO(dotx_bytes)

        document = from_template(stream)
        out_path = tmp_path / "derived.docx"
        document.save(str(out_path))

        with zipfile.ZipFile(out_path, "r") as zf:
            content_types = zf.read("[Content_Types].xml").decode("utf-8")
            assert CT.WML_DOCUMENT_MAIN in content_types
            assert CT.WML_TEMPLATE_MAIN not in content_types
