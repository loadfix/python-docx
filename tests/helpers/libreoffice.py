"""LibreOffice headless validation for .docx files.

Converts .docx files to PDF using LibreOffice in headless mode. If the conversion
fails, it indicates the file is malformed or contains unsupported content.

This validation layer is optional and requires LibreOffice to be installed. Tests
using this helper should be marked with `@pytest.mark.libreoffice`.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile


class LibreOfficeNotAvailable(RuntimeError):
    """Raised when LibreOffice is not installed or not on PATH."""


class LibreOfficeConversionError(RuntimeError):
    """Raised when LibreOffice fails to convert a .docx file."""


def is_libreoffice_available() -> bool:
    """Return True if LibreOffice is available on the system PATH."""
    return shutil.which("libreoffice") is not None


def validate_with_libreoffice(
    docx_path: str, timeout: int = 60, outdir: str | None = None
) -> tuple[str, str]:
    """Validate a .docx file by converting it to PDF with LibreOffice headless.

    Returns a (pdf_path, outdir) tuple on success. The caller is responsible for
    cleaning up `outdir` (e.g. via `shutil.rmtree(outdir)`).

    Raises LibreOfficeConversionError if the conversion fails.
    Raises LibreOfficeNotAvailable if LibreOffice is not installed.

    Args:
        docx_path: Path to the .docx file to validate.
        timeout: Maximum seconds to wait for conversion (default 60).
        outdir: Optional output directory. A temporary directory is created if None.
    """
    if not is_libreoffice_available():
        raise LibreOfficeNotAvailable(
            "LibreOffice is not installed. Install with: "
            "sudo apt-get install libreoffice-writer"
        )

    created_outdir = outdir is None
    if outdir is None:
        outdir = tempfile.mkdtemp(prefix="docx_lo_validate_")

    try:
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                outdir,
                docx_path,
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        if created_outdir:
            shutil.rmtree(outdir, ignore_errors=True)
        raise LibreOfficeConversionError(
            f"LibreOffice conversion timed out after {timeout}s for {docx_path}"
        )

    if result.returncode != 0:
        if created_outdir:
            shutil.rmtree(outdir, ignore_errors=True)
        raise LibreOfficeConversionError(
            f"LibreOffice conversion failed (exit code {result.returncode}):\n"
            f"stdout: {result.stdout}\n"
            f"stderr: {result.stderr}"
        )

    # Find the generated PDF
    basename = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(outdir, f"{basename}.pdf")

    if not os.path.exists(pdf_path):
        if created_outdir:
            shutil.rmtree(outdir, ignore_errors=True)
        raise LibreOfficeConversionError(
            f"LibreOffice conversion produced no output PDF for {docx_path}.\n"
            f"stdout: {result.stdout}\n"
            f"stderr: {result.stderr}"
        )

    return pdf_path, outdir
