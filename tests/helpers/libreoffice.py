"""Layer 5: LibreOffice headless validation helpers.

Validates .docx files by converting them with LibreOffice in headless mode.
If LibreOffice is not installed, tests using these helpers should be skipped.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile

import pytest

def libreoffice_available() -> bool:
    """Return True if LibreOffice is available on the system PATH."""
    return shutil.which("libreoffice") is not None


# -- pytest marker for tests that require LibreOffice --
requires_libreoffice = pytest.mark.skipif(
    not libreoffice_available(),
    reason="LibreOffice not installed",
)


def validate_with_libreoffice(docx_path: str) -> tuple[bool, str]:
    """Validate a .docx file by converting it to PDF with LibreOffice headless.

    Returns (success, message) where success is True if the conversion completed
    without error.
    """
    if not libreoffice_available():
        return False, "LibreOffice not installed"

    with tempfile.TemporaryDirectory() as tmpdir:
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                tmpdir,
                docx_path,
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )

        if result.returncode != 0:
            return False, f"LibreOffice conversion failed: {result.stderr}"

        # -- check that a PDF was produced --
        basename = os.path.splitext(os.path.basename(docx_path))[0]
        pdf_path = os.path.join(tmpdir, f"{basename}.pdf")
        if not os.path.isfile(pdf_path):
            return False, "LibreOffice did not produce a PDF output file"

        pdf_size = os.path.getsize(pdf_path)
        if pdf_size == 0:
            return False, "LibreOffice produced an empty PDF"

    return True, "OK"
