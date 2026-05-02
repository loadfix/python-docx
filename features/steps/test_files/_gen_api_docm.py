"""Generate ``api-demo.docm`` fixture for macro-enabled document scenarios.

``.docm`` documents are OOXML packages whose main document part uses the
``application/vnd.ms-word.document.macroEnabled.main+xml`` content type and
carry a relationship of type ``vbaProject`` pointing at a binary VBA project
part. python-docx supports *reading* these packages and exposes
:attr:`~docx.document.Document.has_macros` so callers can detect them;
authoring VBA is out of scope.

This generator starts from the default python-docx template, rewrites
``[Content_Types].xml`` to declare the macro-enabled content type, injects a
``vbaProject`` relationship, and emits a placeholder ``word/vbaProject.bin``
stub. The VBA blob is **not** a valid compiled VBA project — it exists only
so the relationship target resolves and ``has_macros`` returns |True|. Word
itself would refuse to open this fixture; use it only for python-docx
detection tests.

Run ``python features/steps/test_files/_gen_api_docm.py`` to regenerate.
"""

from __future__ import annotations

import os
import tempfile
import zipfile

from docx import Document

HERE = os.path.abspath(os.path.dirname(__file__))
SOURCE = os.path.join(
    HERE, "..", "..", "..", "src", "docx", "templates", "default.docx"
)
OUT_PATH = os.path.join(HERE, "api-demo.docm")

_MACRO_MAIN = (
    "application/vnd.ms-word.document.macroEnabled.main+xml"
)
_VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
_VBA_REL_TYPE = (
    "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
)


def _patch_content_types(xml: bytes) -> bytes:
    xml = xml.replace(
        (
            b"application/vnd.openxmlformats-officedocument"
            b".wordprocessingml.document.main+xml"
        ),
        _MACRO_MAIN.encode("utf-8"),
    )
    extra = (
        f'<Default Extension="bin" ContentType="{_VBA_CONTENT_TYPE}"/>'
    ).encode("utf-8")
    end = xml.rfind(b"</Types>")
    if end == -1:
        raise ValueError("[Content_Types].xml missing </Types>")
    return xml[:end] + extra + xml[end:]


def _patch_rels(xml: bytes) -> bytes:
    rel = (
        f'<Relationship Id="rIdVba" Type="{_VBA_REL_TYPE}"'
        f' Target="vbaProject.bin"/>'
    ).encode("utf-8")
    end = xml.rfind(b"</Relationships>")
    if end == -1:
        raise ValueError("document.xml.rels missing </Relationships>")
    return xml[:end] + rel + xml[end:]


def build() -> str:
    source = os.path.normpath(SOURCE)
    if not os.path.isfile(source):
        raise FileNotFoundError(source)

    with zipfile.ZipFile(source, "r") as zi, zipfile.ZipFile(
        OUT_PATH, "w", zipfile.ZIP_DEFLATED
    ) as zo:
        for info in zi.infolist():
            data = zi.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = _patch_content_types(data)
            elif info.filename == "word/_rels/document.xml.rels":
                data = _patch_rels(data)
            zo.writestr(info, data)
        # -- placeholder VBA blob: enough to satisfy relationship resolution --
        zo.writestr("word/vbaProject.bin", b"\x00" * 32)
    return OUT_PATH


def validate(path: str) -> None:
    document = Document(path)
    assert document.has_macros is True, "has_macros should be True for .docm"


if __name__ == "__main__":
    out = build()
    validate(out)
    print(f"wrote {out}")
