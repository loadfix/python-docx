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


class InkAnnotation:
    """Proxy for an ink annotation referenced from a paragraph.

    Wraps a ``<w:contentPart>`` element (via the paragraph that contains it)
    and its related ink part. Read-only.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, paragraph: Paragraph, ink_part: InkPart):
        self._paragraph = paragraph
        self._ink_part = ink_part

    @property
    def blob(self) -> bytes:
        """Raw InkML XML bytes of the referenced ink part.

        .. versionadded:: 1.3.0.dev0
        """
        return self._ink_part.blob

    @property
    def paragraph(self) -> Paragraph:
        """The |Paragraph| that contains the ``w:contentPart`` for this annotation.

        .. versionadded:: 1.3.0.dev0
        """
        return self._paragraph

    @property
    def partname(self) -> str:
        """The OPC partname of the related ink part, e.g. ``/word/ink/ink1.xml``.

        .. versionadded:: 1.3.0.dev0
        """
        return str(self._ink_part.partname)

    @property
    def stroke_count(self) -> int:
        """Count of ``<inkml:trace>`` elements in the referenced ink XML.

        Returns ``0`` when the part is missing, empty, or cannot be parsed. The
        count includes traces at any depth (typical InkML wraps them in a
        ``<inkml:traceGroup>`` but they may also appear as direct children of
        the root ``<inkml:ink>`` element).

        .. versionadded:: 1.3.0.dev0
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
