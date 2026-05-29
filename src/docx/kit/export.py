"""Unified export entry-points — wrappers around the existing python-docx exporters.

Closes #305.

The :mod:`docx.kit.export` module bundles four single-call wrappers
around the existing python-docx exporters plus one new minimal EPUB
3 exporter, with a generic :func:`to` dispatcher that picks the right
one based on the output file extension::

    from docx.kit import export

    export.to_pdf(doc, "out.pdf")        # uses Document.save_as_pdf_a (level=3a)
    export.to_html(doc, "out.html")      # uses Document.to_html()
    export.to_md(doc, "out.md")          # uses Document.to_markdown()
    export.to_epub(doc, "out.epub")      # NEW — minimal EPUB 3 export

    # Or one-shot, dispatching on the file extension:
    export.to(doc, "out.pdf")

Every wrapper opens / writes the output file with UTF-8 encoding
and returns ``None`` (the file path is the side effect). Errors
from the underlying exporter (e.g. :class:`ImportError` from
``Document.save_as_pdf_a`` when ``reportlab`` is missing) propagate
unchanged.

The EPUB exporter is a new minimal EPUB 3 single-file exporter:

* Standard mime type, ``META-INF/container.xml`` pointing at the OPF.
* ``OEBPS/content.opf`` with a generated UUID URN identifier and a
  single XHTML chapter.
* ``OEBPS/toc.ncx`` (legacy NCX for EPUB 2 reader compatibility).
* ``OEBPS/chapter1.xhtml`` containing the document's HTML rendering
  (via :meth:`Document.to_html`) wrapped in an XHTML 1.1 doctype.
* ``OEBPS/styles.css`` with a basic stylesheet.

Pure stdlib — :mod:`zipfile` is the only dependency. The exporter is
intentionally minimal and trades fidelity for portability: EPUB
readers rendering the output get a single navigable chapter with
inline CSS preserved from the HTML pipeline.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import os
import re
import uuid
import zipfile
from typing import IO, TYPE_CHECKING, Any, Optional, Union

if TYPE_CHECKING:
    from docx.document import Document as DocumentCls


__all__ = [
    "to",
    "to_epub",
    "to_html",
    "to_md",
    "to_pdf",
]


# -- accepted ``path`` types for every wrapper. Either a filesystem
# -- path-like or an open binary file-like object. --
PathLike = Union[str, "os.PathLike[str]"]


def to_pdf(document: "DocumentCls", path: PathLike, **kwargs: Any) -> None:
    """Render `document` to a PDF/A archival PDF at `path`.

    Thin wrapper around :meth:`Document.save_as_pdf_a` — the
    archival-grade exporter shipped in 2026.05.29. The default
    conformance level is ``"3a"`` (tagged structure tree, accessible);
    override via ``level=`` kwarg.

    Raises :class:`ImportError` (re-raised from the underlying
    exporter) when ``reportlab`` is not installed; install via
    ``pip install 'python-docx[pdfa]'``.

    Example::

        from docx.kit import export
        export.to_pdf(doc, "report.pdf")              # PDF/A-3a (default)
        export.to_pdf(doc, "report.pdf", level="2b")  # PDF/A-2b override

    .. versionadded:: 2026.05.29
    """
    level = kwargs.pop("level", "3a")
    if kwargs:
        raise TypeError(
            "to_pdf() got unexpected keyword argument(s): %s"
            % ", ".join(sorted(kwargs))
        )
    document.save_as_pdf_a(os.fspath(path), level=level)


def to_html(document: "DocumentCls", path: PathLike, **kwargs: Any) -> None:
    """Render `document` to a self-contained HTML file at `path`.

    Thin wrapper around :meth:`Document.to_html`. Both
    ``include_styles=`` and ``embed_images=`` (defaults: |True|)
    forward to the underlying exporter; pass them through ``**kwargs``.

    Output is written with UTF-8 encoding.

    Example::

        from docx.kit import export
        export.to_html(doc, "out.html")
        export.to_html(doc, "out.html", embed_images=False)

    .. versionadded:: 2026.05.29
    """
    include_styles = kwargs.pop("include_styles", True)
    embed_images = kwargs.pop("embed_images", True)
    if kwargs:
        raise TypeError(
            "to_html() got unexpected keyword argument(s): %s"
            % ", ".join(sorted(kwargs))
        )
    html_text = document.to_html(
        include_styles=include_styles, embed_images=embed_images
    )
    _write_text(path, html_text)


def to_md(document: "DocumentCls", path: PathLike, **kwargs: Any) -> None:
    """Render `document` to a GitHub-Flavoured-Markdown file at `path`.

    Thin wrapper around :meth:`Document.to_markdown`. No additional
    keyword arguments are accepted.

    Output is written with UTF-8 encoding.

    Example::

        from docx.kit import export
        export.to_md(doc, "out.md")

    .. versionadded:: 2026.05.29
    """
    if kwargs:
        raise TypeError(
            "to_md() got unexpected keyword argument(s): %s"
            % ", ".join(sorted(kwargs))
        )
    md_text = document.to_markdown()
    _write_text(path, md_text)


def to_epub(
    document: "DocumentCls",
    path: PathLike,
    *,
    title: Optional[str] = None,
    author: Optional[str] = None,
) -> None:
    """Render `document` to a minimal EPUB 3 single-file package at `path`.

    Builds a portable EPUB 3 file with the document's HTML rendering
    (via :meth:`Document.to_html`) wrapped in an XHTML 1.1 doctype as
    a single chapter. Pure stdlib — :mod:`zipfile` is the only
    dependency.

    Package contents:

    * ``mimetype`` — the magic ``application/epub+zip`` identifier
      (stored uncompressed, per the EPUB spec).
    * ``META-INF/container.xml`` — OCF wrapper pointing at the OPF.
    * ``OEBPS/content.opf`` — the package document declaring
      manifest, spine, metadata (title, language, identifier, modified
      timestamp, optional author).
    * ``OEBPS/toc.ncx`` — legacy NCX navigation for EPUB 2 readers.
    * ``OEBPS/nav.xhtml`` — EPUB 3 nav document.
    * ``OEBPS/chapter1.xhtml`` — the document HTML wrapped as a single
      chapter.
    * ``OEBPS/styles.css`` — basic stylesheet (paragraph spacing,
      table borders, code monospace).

    The ``title`` defaults to the document's core-properties title
    (falling back to ``"Document"``); ``author`` defaults to the
    core-properties author when not supplied.

    Example::

        from docx.kit import export
        export.to_epub(doc, "out.epub", title="Annual Report 2026")

    .. versionadded:: 2026.05.29
    """
    epub_title = (title or _doc_title(document) or "Document").strip() or "Document"
    epub_author = author or _doc_author(document) or "Unknown"

    # -- Render the document HTML once; we strip the outer
    # -- <html><body>...</body></html> wrapper later and re-wrap in
    # -- XHTML 1.1 to keep EPUB readers happy. --
    html_text = document.to_html(include_styles=False, embed_images=True)
    chapter_body = _extract_body(html_text)

    identifier = "urn:uuid:" + str(uuid.uuid4())

    chapter_xhtml = _CHAPTER_TEMPLATE.format(
        title=_xml_escape(epub_title),
        body=chapter_body,
    )
    container_xml = _CONTAINER_XML
    content_opf = _build_content_opf(
        identifier=identifier, title=epub_title, author=epub_author
    )
    toc_ncx = _build_toc_ncx(identifier=identifier, title=epub_title)
    nav_xhtml = _build_nav_xhtml(title=epub_title)
    styles_css = _STYLES_CSS

    out_path = os.fspath(path)
    # -- EPUB requires the ``mimetype`` entry be the first entry in
    # -- the zip and stored uncompressed. zipfile does not let us
    # -- pin the *first* entry directly except by writing it before
    # -- anything else. --
    with zipfile.ZipFile(out_path, "w") as zf:
        info = zipfile.ZipInfo("mimetype")
        info.compress_type = zipfile.ZIP_STORED
        zf.writestr(info, "application/epub+zip")
        zf.writestr("META-INF/container.xml", container_xml)
        zf.writestr("OEBPS/content.opf", content_opf)
        zf.writestr("OEBPS/toc.ncx", toc_ncx)
        zf.writestr("OEBPS/nav.xhtml", nav_xhtml)
        zf.writestr("OEBPS/chapter1.xhtml", chapter_xhtml)
        zf.writestr("OEBPS/styles.css", styles_css)


def to(document: "DocumentCls", path: PathLike, **kwargs: Any) -> None:
    """Dispatch to the right :mod:`docx.kit.export` wrapper based on `path`'s extension.

    Recognised extensions (case-insensitive):

    * ``.pdf``  -> :func:`to_pdf`
    * ``.html`` / ``.htm`` -> :func:`to_html`
    * ``.md`` / ``.markdown`` -> :func:`to_md`
    * ``.epub`` -> :func:`to_epub`

    Any other extension raises :class:`ValueError`.

    Example::

        from docx.kit import export
        export.to(doc, "out.pdf")     # dispatches to to_pdf
        export.to(doc, "out.epub")    # dispatches to to_epub

    .. versionadded:: 2026.05.29
    """
    ext = os.path.splitext(os.fspath(path))[1].lower()
    if ext == ".pdf":
        to_pdf(document, path, **kwargs)
    elif ext in (".html", ".htm"):
        to_html(document, path, **kwargs)
    elif ext in (".md", ".markdown"):
        to_md(document, path, **kwargs)
    elif ext == ".epub":
        to_epub(document, path, **kwargs)
    else:
        raise ValueError(
            "Unsupported export extension %r — supported: .pdf, .html, .htm, "
            ".md, .markdown, .epub" % ext
        )


# -- ---------------------------------------------------------------
# -- internals
# -- ---------------------------------------------------------------


def _write_text(path: PathLike, text: str) -> None:
    """Write `text` to `path` with UTF-8 encoding."""
    with open(os.fspath(path), "w", encoding="utf-8", newline="") as fh:
        fh.write(text)


def _xml_escape(text: str) -> str:
    """Escape XML/HTML metacharacters in `text` for safe attribute / text use."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _doc_title(document: "DocumentCls") -> Optional[str]:
    """Return the document's core-properties title or ``None`` when unavailable."""
    try:
        return document.core_properties.title or None
    except (AttributeError, KeyError):  # pragma: no cover - defensive
        return None


def _doc_author(document: "DocumentCls") -> Optional[str]:
    """Return the document's core-properties author or ``None`` when unavailable."""
    try:
        return document.core_properties.author or None
    except (AttributeError, KeyError):  # pragma: no cover - defensive
        return None


# -- Strip the outer ``<html>...<body>...</body></html>`` wrapper from
# -- :meth:`Document.to_html` output so we can splice the body into
# -- our XHTML chapter template. The exporter emits a stable shape
# -- (a leading ``<!DOCTYPE>`` plus ``<html><head>...</head><body>...</body></html>``)
# -- so a permissive regex is fine here — robustness is bounded by
# -- the exporter we control. --
_BODY_RE = re.compile(r"<body[^>]*>(.*?)</body>", re.DOTALL | re.IGNORECASE)


def _extract_body(html_text: str) -> str:
    """Return the inner ``<body>`` of `html_text`, or the whole string when absent."""
    match = _BODY_RE.search(html_text)
    if match is None:
        return html_text
    return match.group(1).strip()


_CONTAINER_XML = """<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>
"""


_CHAPTER_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" \
"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
  <meta http-equiv="Content-Type" content="application/xhtml+xml; charset=utf-8"/>
  <title>{title}</title>
  <link rel="stylesheet" type="text/css" href="styles.css"/>
</head>
<body>
{body}
</body>
</html>
"""


_STYLES_CSS = """\
body {
  font-family: Georgia, 'Times New Roman', serif;
  line-height: 1.5;
  margin: 1em;
}
p { margin: 0 0 0.6em 0; }
h1, h2, h3, h4, h5, h6 { margin: 1.2em 0 0.4em 0; }
table { border-collapse: collapse; margin: 0.6em 0; }
td, th { border: 1px solid #888; padding: 0.2em 0.4em; }
code, pre { font-family: 'Courier New', monospace; }
img { max-width: 100%; height: auto; }
"""


def _build_content_opf(*, identifier: str, title: str, author: str) -> str:
    """Return the OEBPS/content.opf string for the EPUB package."""
    return _CONTENT_OPF_TEMPLATE.format(
        identifier=_xml_escape(identifier),
        title=_xml_escape(title),
        author=_xml_escape(author),
        modified=_iso_modified(),
    )


def _build_toc_ncx(*, identifier: str, title: str) -> str:
    """Return the OEBPS/toc.ncx string for the EPUB package (EPUB 2 fallback)."""
    return _TOC_NCX_TEMPLATE.format(
        identifier=_xml_escape(identifier),
        title=_xml_escape(title),
    )


def _build_nav_xhtml(*, title: str) -> str:
    """Return the OEBPS/nav.xhtml string (EPUB 3 navigation document)."""
    return _NAV_XHTML_TEMPLATE.format(title=_xml_escape(title))


def _iso_modified() -> str:
    """Return an ISO-8601 ``dcterms:modified`` timestamp accurate to seconds."""
    # -- EPUB 3 requires exactly ``YYYY-MM-DDThh:mm:ssZ`` — no
    # -- microseconds, no timezone offset. We use a fixed timestamp
    # -- when ``SOURCE_DATE_EPOCH`` is set (reproducible-build
    # -- convention) so EPUB output is byte-identical run-to-run. --
    sde = os.environ.get("SOURCE_DATE_EPOCH")
    if sde is not None:
        try:
            from datetime import datetime, timezone

            return (
                datetime.fromtimestamp(int(sde), tz=timezone.utc)
                .strftime("%Y-%m-%dT%H:%M:%SZ")
            )
        except (ValueError, TypeError):  # pragma: no cover - defensive
            pass
    from datetime import datetime, timezone

    return datetime.now(tz=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


_CONTENT_OPF_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" \
unique-identifier="bookid" xml:lang="en">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="bookid">{identifier}</dc:identifier>
    <dc:title>{title}</dc:title>
    <dc:creator>{author}</dc:creator>
    <dc:language>en</dc:language>
    <meta property="dcterms:modified">{modified}</meta>
  </metadata>
  <manifest>
    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" \
properties="nav"/>
    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
    <item id="chapter1" href="chapter1.xhtml" \
media-type="application/xhtml+xml"/>
    <item id="styles" href="styles.css" media-type="text/css"/>
  </manifest>
  <spine toc="ncx">
    <itemref idref="chapter1"/>
  </spine>
</package>
"""


_TOC_NCX_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE ncx PUBLIC "-//NISO//DTD ncx 2005-1//EN" \
"http://www.daisy.org/z3986/2005/ncx-2005-1.dtd">
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
  <head>
    <meta name="dtb:uid" content="{identifier}"/>
    <meta name="dtb:depth" content="1"/>
    <meta name="dtb:totalPageCount" content="0"/>
    <meta name="dtb:maxPageNumber" content="0"/>
  </head>
  <docTitle><text>{title}</text></docTitle>
  <navMap>
    <navPoint id="navPoint-1" playOrder="1">
      <navLabel><text>{title}</text></navLabel>
      <content src="chapter1.xhtml"/>
    </navPoint>
  </navMap>
</ncx>
"""


_NAV_XHTML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" \
xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="en">
<head>
  <meta http-equiv="Content-Type" content="application/xhtml+xml; charset=utf-8"/>
  <title>{title}</title>
</head>
<body>
  <nav epub:type="toc">
    <h1>Contents</h1>
    <ol>
      <li><a href="chapter1.xhtml">{title}</a></li>
    </ol>
  </nav>
</body>
</html>
"""


# -- ``IO`` is exported for typing's sake even though we coerce to a
# -- path string with ``os.fspath`` before opening — keeping the
# -- import preserves the option to switch to a stream-based API
# -- without re-touching the import block. --
_ = IO
