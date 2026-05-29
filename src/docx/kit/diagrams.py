"""Inline Mermaid / PlantUML / Graphviz-DOT diagrams in a docx body.

Closes #292.

Three public helpers (:func:`mermaid`, :func:`plantuml`, :func:`dot`)
take the diagram source as a string, render it to PNG (or SVG), and
embed the result as an inline picture in a
:class:`~docx.document.Document`. Each returns the
:class:`~docx.shape.InlineShape` for further customisation::

    from docx import Document
    from docx.kit import diagrams

    doc = Document()
    diagrams.mermaid(doc, "flowchart TD\\n A --> B", caption="Workflow")
    diagrams.plantuml(doc, "@startuml\\nA -> B\\n@enduml")
    diagrams.dot(doc, "digraph G { A -> B; }")
    doc.save("out.docx")

Render path
-----------

**The default render path is a network call** to the public
`kroki.io <https://kroki.io>`_ service: the diagram source is POSTed
to ``https://kroki.io/<language>/<format>`` and the rendered bytes
come back. The ``timeout=`` parameter (default 30 s) bounds the wait;
the underlying HTTP client (``requests`` first, then ``httpx``) is
soft-imported so non-callers pay no install cost.

For an *offline* render — air-gapped builds, deterministic CI,
vendor-policy reasons — install one of the language CLI binaries and
the kit will prefer it automatically:

* **Mermaid** — ``npm i -g @mermaid-js/mermaid-cli`` (kit shells out to
  ``mmdc``).
* **PlantUML** — install ``plantuml`` (or ``plantuml.jar`` + a JRE).
* **DOT / Graphviz** — install Graphviz (kit shells out to ``dot``).

Pass ``backend="kroki"`` to force the network path even when a binary
is present; pass ``backend="local"`` to force the binary (raising
:class:`FileNotFoundError` when missing); the default
``backend="auto"`` prefers local and falls back to kroki.

When ``caption`` is supplied, a ``"Figure N: caption"`` paragraph in
the ``Caption`` style is appended after the image via
:meth:`Document.add_caption <docx.document.Document.add_caption>` (Word
auto-numbers via the ``SEQ`` field). When ``alt_text`` is supplied, it
is set on the inline picture for screen readers.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import base64
import os
import shutil
import subprocess
import tempfile
import zlib
from typing import TYPE_CHECKING, Optional, Tuple, Union

from docx.shared import Inches

if TYPE_CHECKING:
    from docx.document import Document
    from docx.shape import InlineShape

# -- Kroki public endpoint.  POSTing the raw diagram source to
# -- ``<base>/<language>/<format>`` returns the rendered bytes.  The
# -- service supports many languages (mermaid, plantuml, graphviz,
# -- ditaa, etc.) under a uniform routing convention.
_KROKI_BASE_URL = "https://kroki.io"

# -- Default network timeout, in seconds.  30 s is generous enough for
# -- a cold-start kroki request including DNS, TLS, and a complex
# -- diagram render; callers who need a tighter or looser bound supply
# -- their own ``timeout=`` argument.
_DEFAULT_TIMEOUT = 30.0

# -- Default rendered-image width in inches.  5 inches fits inside a
# -- standard 6-inch text column on US Letter / A4 with margin to
# -- spare; callers override via ``width_in=``.
_DEFAULT_WIDTH_IN = 5.0

# -- Output image formats kroki understands.  PNG is the default
# -- because Word's renderer is most consistent with PNG; SVG is
# -- offered for callers who want vector output.
_VALID_FORMATS = ("png", "svg")
_DEFAULT_FORMAT = "png"

# -- Backend selection values for the ``backend=`` argument.
_VALID_BACKENDS = ("auto", "kroki", "local")
_DEFAULT_BACKEND = "auto"

# -- Per-language metadata.  Maps the public helper name ("mermaid",
# -- "plantuml", "dot") to the kroki language slug and the local CLI
# -- binary names tried in order.  ``dot`` resolves to the graphviz
# -- "dot" binary; kroki's slug for the same language is "graphviz".
_LANGUAGES = {
    "mermaid": {
        "kroki_slug": "mermaid",
        "binaries": ("mmdc",),
    },
    "plantuml": {
        "kroki_slug": "plantuml",
        "binaries": ("plantuml",),
    },
    "dot": {
        "kroki_slug": "graphviz",
        "binaries": ("dot",),
    },
}


# ---------------------------------------------------------------------------
# Soft-import shim for the HTTP client.
# ---------------------------------------------------------------------------


def _http_post(url, data, timeout):
    # type: (str, bytes, float) -> bytes
    """POST ``data`` to ``url`` and return the response body bytes.

    Soft-imports ``requests`` first, then ``httpx``. Raises
    :class:`ImportError` (with both candidate names) if neither is
    installed.

    The body is the raw diagram source; kroki returns the rendered
    bytes with the requested image format.
    """
    try:
        import requests  # type: ignore[import-not-found]
    except ImportError:
        requests = None  # type: ignore[assignment]
    if requests is not None:
        response = requests.post(url, data=data, timeout=timeout)
        response.raise_for_status()
        return response.content

    try:
        import httpx  # type: ignore[import-not-found]
    except ImportError:
        httpx = None  # type: ignore[assignment]
    if httpx is not None:
        response = httpx.post(url, content=data, timeout=timeout)
        response.raise_for_status()
        return response.content

    raise ImportError(
        "docx.kit.diagrams needs 'requests' or 'httpx' to call the kroki.io "
        "rendering service. Install one (e.g. `pip install requests`) or "
        "install a local diagram binary (mmdc / plantuml / dot) and pass "
        "backend='local'."
    )


# ---------------------------------------------------------------------------
# Local CLI shells.
# ---------------------------------------------------------------------------


def _find_local_binary(language):
    # type: (str) -> Optional[str]
    """Return the absolute path of the first available local binary.

    Checks each candidate in ``_LANGUAGES[language]["binaries"]`` via
    :func:`shutil.which`. Returns ``None`` when no candidate is on
    ``PATH``.
    """
    for candidate in _LANGUAGES[language]["binaries"]:
        path = shutil.which(candidate)
        if path:
            return path
    return None


def _render_local(language, source, image_format):
    # type: (str, str, str) -> bytes
    """Render ``source`` via the language-specific local CLI binary.

    Mermaid uses ``mmdc -i in -o out.<fmt>``. PlantUML uses
    ``plantuml -t<fmt> -pipe`` (stdin -> stdout). DOT uses
    ``dot -T<fmt>`` (stdin -> stdout).

    Returns the rendered image bytes. Propagates
    :class:`subprocess.CalledProcessError` on non-zero exit.
    """
    binary = _find_local_binary(language)
    if binary is None:
        raise FileNotFoundError(
            "no local binary for %s on PATH (looked for %s)"
            % (language, ", ".join(_LANGUAGES[language]["binaries"]))
        )

    # -- mermaid-cli does not read from stdin; it requires real files.
    # -- We stage source in a tempdir, run mmdc, read the output.
    if language == "mermaid":
        with tempfile.TemporaryDirectory() as tmp:
            in_path = os.path.join(tmp, "in.mmd")
            out_path = os.path.join(tmp, "out." + image_format)
            with open(in_path, "w", encoding="utf-8") as fh:
                fh.write(source)
            subprocess.run(
                [binary, "-i", in_path, "-o", out_path],
                check=True,
                capture_output=True,
            )
            with open(out_path, "rb") as fh:
                return fh.read()

    # -- plantuml + dot accept stdin -> stdout with their respective
    # -- pipe / format flags.
    if language == "plantuml":
        argv = [binary, "-t" + image_format, "-pipe"]
    else:  # -- "dot"
        argv = [binary, "-T" + image_format]

    completed = subprocess.run(
        argv,
        input=source.encode("utf-8"),
        check=True,
        capture_output=True,
    )
    return completed.stdout


# ---------------------------------------------------------------------------
# Kroki transport.
# ---------------------------------------------------------------------------


def kroki_url(language, image_format=_DEFAULT_FORMAT, base_url=_KROKI_BASE_URL):
    # type: (str, str, str) -> str
    """Return the kroki POST endpoint for ``language`` and ``image_format``.

    Surfaced as a public helper so callers can override the base URL
    (e.g. point at a self-hosted kroki container) by passing
    ``base_url=``.
    """
    if language not in _LANGUAGES:
        raise ValueError(
            "language must be one of %s; got %r"
            % (sorted(_LANGUAGES), language)
        )
    slug = _LANGUAGES[language]["kroki_slug"]
    return "%s/%s/%s" % (base_url.rstrip("/"), slug, image_format)


def kroki_get_url(
    language, source, image_format=_DEFAULT_FORMAT, base_url=_KROKI_BASE_URL
):
    # type: (str, str, str, str) -> str
    """Return the kroki *GET*-style URL with ``source`` deflate-encoded.

    Kroki accepts both POST (raw body) and GET (URL-embedded
    ``deflate + urlsafe-base64`` of the source). The GET form is handy
    for embedding shareable links in non-HTTP contexts; the kit's
    transport uses POST for reliability with large diagrams. Exposed
    publicly because callers occasionally want the URL itself
    (e.g. to render in a Word hyperlink rather than embed the bytes).
    """
    if language not in _LANGUAGES:
        raise ValueError(
            "language must be one of %s; got %r"
            % (sorted(_LANGUAGES), language)
        )
    slug = _LANGUAGES[language]["kroki_slug"]
    compressed = zlib.compress(source.encode("utf-8"), 9)
    encoded = base64.urlsafe_b64encode(compressed).decode("ascii")
    return "%s/%s/%s/%s" % (base_url.rstrip("/"), slug, image_format, encoded)


def _render_kroki(language, source, image_format, timeout, base_url):
    # type: (str, str, str, float, str) -> bytes
    """POST ``source`` to kroki and return the rendered bytes."""
    url = kroki_url(language, image_format=image_format, base_url=base_url)
    return _http_post(url, source.encode("utf-8"), timeout=timeout)


# ---------------------------------------------------------------------------
# Render dispatch (auto / kroki / local).
# ---------------------------------------------------------------------------


def _render(language, source, image_format, timeout, backend, base_url):
    # type: (str, str, str, float, str, str) -> Tuple[bytes, str]
    """Dispatch the render request to the configured backend.

    Returns ``(image_bytes, backend_used)`` where ``backend_used`` is
    one of ``"local"`` or ``"kroki"`` — the actual backend that
    produced the bytes (handy for tests and observability).

    ``backend="auto"`` prefers the local binary (no network, no
    third-party dependency) and falls back to kroki when the binary is
    not on ``PATH``.
    """
    if backend not in _VALID_BACKENDS:
        raise ValueError(
            "backend must be one of %s; got %r" % (_VALID_BACKENDS, backend)
        )
    if image_format not in _VALID_FORMATS:
        raise ValueError(
            "image_format must be one of %s; got %r"
            % (_VALID_FORMATS, image_format)
        )
    if not source or not source.strip():
        raise ValueError("source must be a non-empty diagram description")

    if backend == "local":
        return _render_local(language, source, image_format), "local"
    if backend == "kroki":
        return (
            _render_kroki(language, source, image_format, timeout, base_url),
            "kroki",
        )
    # -- auto: prefer local when available.
    if _find_local_binary(language) is not None:
        return _render_local(language, source, image_format), "local"
    return (
        _render_kroki(language, source, image_format, timeout, base_url),
        "kroki",
    )


def _embed(
    document,
    image_bytes,
    image_format,
    width_in,
    caption,
    alt_text,
):
    # type: (Document, bytes, str, float, Optional[str], Optional[str]) -> InlineShape
    """Embed ``image_bytes`` into ``document`` and return the InlineShape.

    The bytes are staged through a temporary file because
    :meth:`Document.add_picture` accepts a path or file-like; we use a
    :class:`io.BytesIO` (file-like) so the caller never sees a stray
    tempfile on disk.

    When ``caption`` is supplied, appends a ``"Figure N: caption"``
    paragraph in the ``Caption`` style after the image.

    When ``alt_text`` is supplied, sets it on the inline picture so
    screen readers and Word's accessibility checker pick it up.
    """
    import io

    stream = io.BytesIO(image_bytes)
    width = Inches(width_in) if width_in is not None else None
    inline_shape = document.add_picture(stream, width=width)

    if alt_text:
        # -- Public ``InlineShape`` accessibility helpers.  These are
        # -- exposed at the proxy level (no ``oxml`` reach-down).
        try:
            inline_shape.alt_text_title = alt_text
            inline_shape.alt_text_descr = alt_text
        except AttributeError:
            # -- Older python-docx revisions exposed only one of the
            # -- two; fall back to whichever attribute exists.
            for attr in ("alt_text_descr", "alt_text_title"):
                if hasattr(inline_shape, attr):
                    setattr(inline_shape, attr, alt_text)
                    break

    if caption is not None:
        document.add_caption(caption, label="Figure")

    # -- Silence the unused-import lint when image_format is reserved
    # -- for future format-specific embed tweaks (e.g. SVG <-> PNG).
    del image_format
    return inline_shape


# ---------------------------------------------------------------------------
# Public helpers.
# ---------------------------------------------------------------------------


def _diagram(
    language,
    document,
    source,
    caption,
    width_in,
    alt_text,
    image_format,
    timeout,
    backend,
    base_url,
):
    # type: (str, Document, str, Optional[str], Optional[float], Optional[str], str, float, str, str) -> InlineShape
    """Shared body for :func:`mermaid` / :func:`plantuml` / :func:`dot`.

    Renders ``source`` via the selected backend and embeds the
    resulting image bytes inline in ``document``. Public helpers exist
    as individual functions so callers see ``diagrams.mermaid(...)``
    rather than ``diagrams.diagram("mermaid", ...)`` — but the body is
    identical, parameterised by ``language``.
    """
    image_bytes, _backend_used = _render(
        language, source, image_format, timeout, backend, base_url
    )
    return _embed(document, image_bytes, image_format, width_in, caption, alt_text)


def mermaid(
    document,
    source,
    caption=None,
    width_in=_DEFAULT_WIDTH_IN,
    alt_text=None,
    image_format=_DEFAULT_FORMAT,
    timeout=_DEFAULT_TIMEOUT,
    backend=_DEFAULT_BACKEND,
    base_url=_KROKI_BASE_URL,
):
    # type: (Document, str, Optional[str], Optional[float], Optional[str], str, float, str, str) -> InlineShape
    """Render ``source`` as a Mermaid diagram and embed it in ``document``.

    Parameters
    ----------
    document
        The :class:`Document` to mutate.
    source
        The Mermaid diagram description (e.g. ``flowchart TD ...``).
    caption
        Optional figure caption. When supplied, a ``"Figure N: caption"``
        paragraph in the ``Caption`` style is appended after the image.
    width_in
        Rendered image width in inches. Defaults to 5.0.
    alt_text
        Optional accessibility alt-text applied to the picture.
    image_format
        ``"png"`` (default) or ``"svg"``.
    timeout
        Network timeout in seconds for the kroki call (default 30 s).
        Ignored when the local binary is used.
    backend
        ``"auto"`` (default — prefer local binary, fall back to kroki),
        ``"kroki"`` (force network), or ``"local"`` (force local binary).
    base_url
        Override for the kroki endpoint. Defaults to
        ``https://kroki.io``; supply your own to point at a self-hosted
        kroki instance.

    Returns
    -------
    InlineShape
        The newly inserted inline picture, suitable for further
        customisation.

    Notes
    -----
    The default render path is a **network call** to ``kroki.io``
    (POST). Set ``backend="local"`` and install
    ``@mermaid-js/mermaid-cli`` to render offline.
    """
    return _diagram(
        "mermaid", document, source, caption, width_in, alt_text,
        image_format, timeout, backend, base_url,
    )


def plantuml(
    document,
    source,
    caption=None,
    width_in=_DEFAULT_WIDTH_IN,
    alt_text=None,
    image_format=_DEFAULT_FORMAT,
    timeout=_DEFAULT_TIMEOUT,
    backend=_DEFAULT_BACKEND,
    base_url=_KROKI_BASE_URL,
):
    # type: (Document, str, Optional[str], Optional[float], Optional[str], str, float, str, str) -> InlineShape
    """Render ``source`` as a PlantUML diagram and embed it in ``document``.

    See :func:`mermaid` for the parameter contract — this helper has
    the same shape, only the diagram language changes. The default
    render path is a **network call** to ``kroki.io`` (PlantUML
    renderer); set ``backend="local"`` and install ``plantuml.jar`` +
    a JRE (``plantuml`` on ``PATH``) to render offline.
    """
    return _diagram(
        "plantuml", document, source, caption, width_in, alt_text,
        image_format, timeout, backend, base_url,
    )


def dot(
    document,
    source,
    caption=None,
    width_in=_DEFAULT_WIDTH_IN,
    alt_text=None,
    image_format=_DEFAULT_FORMAT,
    timeout=_DEFAULT_TIMEOUT,
    backend=_DEFAULT_BACKEND,
    base_url=_KROKI_BASE_URL,
):
    # type: (Document, str, Optional[str], Optional[float], Optional[str], str, float, str, str) -> InlineShape
    """Render ``source`` as a Graphviz / DOT diagram and embed it in ``document``.

    See :func:`mermaid` for the parameter contract. The default render
    path is a **network call** to ``kroki.io`` (Graphviz renderer);
    set ``backend="local"`` and install Graphviz (``dot`` on ``PATH``)
    to render offline.
    """
    return _diagram(
        "dot", document, source, caption, width_in, alt_text,
        image_format, timeout, backend, base_url,
    )


__all__ = [
    "mermaid",
    "plantuml",
    "dot",
    "kroki_url",
    "kroki_get_url",
]
