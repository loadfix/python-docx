#!/usr/bin/env python3
"""Download OOXML (ECMA-376) XSD schemas for validation testing.

Downloads the ECMA-376 5th Edition Part 4 schemas and extracts the relevant XSD
files into the `tests/schema/` directory.

Usage:
    python -m tests.helpers.download_schema

The schemas are publicly available from ECMA International. This script downloads
them on demand so they don't need to be committed to the repository.
"""

from __future__ import annotations

import io
import os
import sys
import zipfile
from urllib.request import urlopen

# -- ECMA-376 5th Edition schemas URL --
SCHEMA_URL = (
    "https://www.ecma-international.org/wp-content/uploads/ECMA-376-Fifth-Edition-Part-4"
    "-Transitional-Migration-Features.zip"
)

# -- Alternative: Office Open XML schemas from the standard --
SCHEMA_URL_ALT = (
    "https://www.ecma-international.org/wp-content/uploads/"
    "Office-Open-XML-WML-RefSch.zip"
)

SCHEMA_DIR = os.path.join(os.path.dirname(__file__), "..", "schema")

# -- XSD files we need from the archive --
WANTED_FILES = {
    "wml.xsd",
    "shared-commonSimpleTypes.xsd",
    "shared-math.xsd",
    "shared-bibliography.xsd",
    "shared-customXmlSchemaProperties.xsd",
    "shared-relationshipReference.xsd",
    "dml-main.xsd",
    "dml-wordprocessingDrawing.xsd",
}


def download_schemas() -> None:
    """Download and extract OOXML XSD schemas into tests/schema/."""
    os.makedirs(SCHEMA_DIR, exist_ok=True)

    marker = os.path.join(SCHEMA_DIR, ".downloaded")
    if os.path.exists(marker):
        print(f"Schemas already downloaded in {SCHEMA_DIR}")
        return

    print("Downloading OOXML schemas from ECMA-376...")
    for url in [SCHEMA_URL_ALT, SCHEMA_URL]:
        try:
            resp = urlopen(url, timeout=60)  # noqa: S310
            data = resp.read()
            break
        except Exception as exc:
            print(f"  Failed to download from {url}: {exc}")
            continue
    else:
        print(
            "Could not download schemas. You can manually place XSD files in "
            f"{SCHEMA_DIR}"
        )
        return

    zf = zipfile.ZipFile(io.BytesIO(data))
    extracted = 0
    for entry in zf.namelist():
        basename = os.path.basename(entry)
        if basename in WANTED_FILES or basename.endswith(".xsd"):
            target = os.path.join(SCHEMA_DIR, basename)
            if not os.path.exists(target):
                with open(target, "wb") as f:
                    f.write(zf.read(entry))
                extracted += 1

    # -- write marker file --
    with open(marker, "w") as f:
        f.write("downloaded\n")

    print(f"Extracted {extracted} schema files to {SCHEMA_DIR}")


if __name__ == "__main__":
    download_schemas()
    sys.exit(0)
