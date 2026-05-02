"""Custom element classes related to extended document properties (``app.xml``).

The extended properties part (``docProps/app.xml``) contains application-specific
metadata that the authoring application records for a document -- things like the
``Company``, ``Manager``, ``Application``, ``AppVersion``, ``TotalTime``, page
count, word count, and template name. The part is distinct from core
(Dublin-Core) properties and user-defined custom properties.

Only scalar (text-content) children of the root ``<Properties>`` element are
exposed here; structured vector children such as ``HeadingPairs`` and
``TitlesOfParts`` are left as raw XML and accessed only via the generic
getter/setter used by :class:`docx.extended_properties.ExtendedProperties`.
"""

from __future__ import annotations

from docx.oxml.ns import nsmap
from docx.oxml.xmlchemy import BaseOxmlElement

# -- namespace-prefixed tag name for children in the extended-properties ns --
_EXT_NS = nsmap["extprops"]


def _clark(local: str) -> str:
    """Return the Clark-notation tag for `local` in the extended-properties ns."""
    return "{%s}%s" % (_EXT_NS, local)


class CT_ExtendedProperties(BaseOxmlElement):
    """`<Properties>` root element of the extended-properties part (``app.xml``).

    Exposes a small number of well-known scalar children as typed Python
    attributes; also supports arbitrary ``(name, value)`` get/set via
    :meth:`get_text` / :meth:`set_text`.
    """

    # -- --- generic element text helpers ------------------------------------

    def get_text(self, local_name: str) -> str | None:
        """Return the text content of direct child ``local_name``, or |None|.

        ``local_name`` is the unqualified element name (e.g. ``"Pages"``).
        Returns |None| when no matching child exists; returns ``""`` when the
        child exists but is empty.
        """
        child = self.find(_clark(local_name))
        if child is None:
            return None
        return child.text or ""

    def set_text(self, local_name: str, value: str | int | float | bool | None) -> None:
        """Set direct child ``local_name``'s text to ``value``.

        Creates the child if it does not already exist. Passing |None| removes
        the child entirely. Boolean values are serialised using the ``true`` /
        ``false`` lexical form used by the extended-properties schema.
        """
        child = self.find(_clark(local_name))
        if value is None:
            if child is not None:
                self.remove(child)
            return
        if isinstance(value, bool):
            text = "true" if value else "false"
        else:
            text = str(value)
        if child is None:
            from lxml import etree

            child = etree.SubElement(self, _clark(local_name))
        child.text = text

    # -- --- well-known scalar property helpers ------------------------------

    @property
    def pages(self) -> int | None:
        """Value of the ``<Pages>`` child as an int, or |None| if absent/invalid."""
        text = self.get_text("Pages")
        if text is None or text == "":
            return None
        try:
            return int(text)
        except ValueError:
            return None

    @pages.setter
    def pages(self, value: int | None) -> None:
        self.set_text("Pages", value)

    @property
    def words(self) -> int | None:
        text = self.get_text("Words")
        if text is None or text == "":
            return None
        try:
            return int(text)
        except ValueError:
            return None

    @words.setter
    def words(self, value: int | None) -> None:
        self.set_text("Words", value)

    @property
    def characters(self) -> int | None:
        text = self.get_text("Characters")
        if text is None or text == "":
            return None
        try:
            return int(text)
        except ValueError:
            return None

    @characters.setter
    def characters(self, value: int | None) -> None:
        self.set_text("Characters", value)

    @property
    def characters_with_spaces(self) -> int | None:
        text = self.get_text("CharactersWithSpaces")
        if text is None or text == "":
            return None
        try:
            return int(text)
        except ValueError:
            return None

    @characters_with_spaces.setter
    def characters_with_spaces(self, value: int | None) -> None:
        self.set_text("CharactersWithSpaces", value)

    @property
    def lines(self) -> int | None:
        text = self.get_text("Lines")
        if text is None or text == "":
            return None
        try:
            return int(text)
        except ValueError:
            return None

    @lines.setter
    def lines(self, value: int | None) -> None:
        self.set_text("Lines", value)

    @property
    def paragraphs(self) -> int | None:
        text = self.get_text("Paragraphs")
        if text is None or text == "":
            return None
        try:
            return int(text)
        except ValueError:
            return None

    @paragraphs.setter
    def paragraphs(self, value: int | None) -> None:
        self.set_text("Paragraphs", value)

    @property
    def total_time(self) -> int | None:
        """``<TotalTime>`` child, in minutes."""
        text = self.get_text("TotalTime")
        if text is None or text == "":
            return None
        try:
            return int(text)
        except ValueError:
            return None

    @total_time.setter
    def total_time(self, value: int | None) -> None:
        self.set_text("TotalTime", value)
