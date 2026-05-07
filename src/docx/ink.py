"""Read-only proxy objects for ink annotations.

Word supports ink annotations (stylus/pen strokes) authored on touch-enabled
devices. They are stored as separate ``word/ink/*.xml`` parts using the
`InkML <http://www.w3.org/2003/InkML>`_ specification and referenced from
``<w:contentPart>`` elements inside runs.

python-docx exposes these annotations read-only. Callers can enumerate the
annotations at the |Document| or |Paragraph| level, inspect their stroke
count, or retrieve the raw InkML bytes for further processing.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.ns import nsmap
from docx.oxml.parser import parse_xml

if TYPE_CHECKING:
    from docx.parts.ink import InkPart
    from docx.text.paragraph import Paragraph

    # -- Optional shared-library type. ``ooxml_ink`` may not be importable at
    # -- runtime; TYPE_CHECKING imports are only seen by type-checkers.
    from ooxml_ink.proxies import InkContent


class InkAnnotation:
    """Proxy for an ink annotation referenced from a paragraph.

    Wraps a ``<w:contentPart>`` element (via the paragraph that contains it)
    and its related ink part. Read-only.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, paragraph: Paragraph, ink_part: InkPart):
        self._paragraph = paragraph
        self._ink_part = ink_part

    @property
    def blob(self) -> bytes:
        """Raw InkML XML bytes of the referenced ink part.

        .. versionadded:: 2026.05.0
        """
        return self._ink_part.blob

    @property
    def paragraph(self) -> Paragraph:
        """The |Paragraph| that contains the ``w:contentPart`` for this annotation.

        .. versionadded:: 2026.05.0
        """
        return self._paragraph

    @property
    def partname(self) -> str:
        """The OPC partname of the related ink part, e.g. ``/word/ink/ink1.xml``.

        .. versionadded:: 2026.05.0
        """
        return str(self._ink_part.partname)

    @property
    def stroke_count(self) -> int:
        """Count of ``<inkml:trace>`` elements in the referenced ink XML.

        Returns ``0`` when the part is missing, empty, or cannot be parsed. The
        count includes traces at any depth (typical InkML wraps them in a
        ``<inkml:traceGroup>`` but they may also appear as direct children of
        the root ``<inkml:ink>`` element).

        .. versionadded:: 2026.05.0
        """
        blob = self._ink_part.blob
        if not blob:
            return 0
        try:
            root = parse_xml(blob)
        except Exception:  # pragma: no cover - defensive against malformed XML
            return 0
        inkml_uri = nsmap["inkml"]
        traces = root.xpath(
            ".//inkml:trace | self::inkml:trace",
            namespaces={"inkml": inkml_uri},
        )
        return len(traces)

    @property
    def ink_content(self) -> "InkContent | None":
        """Shared-library :class:`ooxml_ink.InkContent` facade, or |None|.

        Returns a fully-parsed :class:`~ooxml_ink.proxies.InkContent` when the
        shared ``python-ooxml-ink`` package is importable and the underlying
        ink bytes are well-formed. Returns |None| when the package is not
        installed or the payload cannot be parsed.

        Use this when you need richer structured access (traceGroups,
        annotations, sample-point text) than the bare :attr:`stroke_count`.

        .. versionadded:: 2026.05.11
        """
        blob = self._ink_part.blob
        if not blob:
            return None
        try:
            from ooxml_ink.oxml import parse_xml as ooxml_parse_xml
            from ooxml_ink.oxml.inkml import CT_Ink
            from ooxml_ink.proxies import InkContent
        except ImportError:
            return None
        try:
            ink_root = ooxml_parse_xml(blob)
        except Exception:  # noqa: BLE001  — malformed bytes degrade to None
            return None
        if not isinstance(ink_root, CT_Ink):
            return None
        return InkContent(ink_root, blob=blob)
