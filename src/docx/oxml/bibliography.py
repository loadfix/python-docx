"""Custom element classes for the bibliography / citation-sources part.

The bibliography part lives at ``/customXml/item{N}.xml`` and carries a
``<b:Sources>`` root holding zero or more ``<b:Source>`` children in the
``http://schemas.openxmlformats.org/officeDocument/2006/bibliography``
namespace. Each ``<b:Source>`` describes one bibliographic entry.

This module only covers the *structural* shape of ``b:Sources`` and
``b:Source`` — the full type catalogue (books, journal articles, websites,
etc.) is left to the caller via ``source_type`` / author-name handling.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.oxml.ns import nsmap, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore

if TYPE_CHECKING:
    pass


# -- standard bibliography namespace --
_B_NS = nsmap["b"]


def _new_child(parent: "BaseOxmlElement", tag: str, text: str | None = None):
    """Append a child ``<b:{tag}>`` with ``text`` to ``parent`` and return it."""
    elm = OxmlElement(f"b:{tag}", nsdecls={"b": _B_NS})
    if text is not None:
        elm.text = text
    parent.append(elm)
    return elm


class CT_Sources(BaseOxmlElement):
    """``<b:Sources>`` — root element of a bibliography part.

    Holds zero or more ``<b:Source>`` children plus optional
    ``@SelectedStyle`` / ``@StyleName`` / ``@Version`` attributes used by
    Word to select a citation-formatting style (APA, MLA, etc.).
    """

    source_lst: "list[CT_Source]"
    add_source: "callable"

    source = ZeroOrMore("b:Source")

    @property
    def selected_style(self) -> "str | None":
        """Value of ``@SelectedStyle`` (e.g. ``"/APA.XSL"``), or |None|."""
        return self.get("SelectedStyle")

    @selected_style.setter
    def selected_style(self, value: "str | None") -> None:
        if value is None:
            if "SelectedStyle" in self.attrib:
                del self.attrib["SelectedStyle"]
            return
        self.set("SelectedStyle", value)

    @property
    def style_name(self) -> "str | None":
        """Value of ``@StyleName`` (e.g. ``"APA"``), or |None|."""
        return self.get("StyleName")

    @style_name.setter
    def style_name(self, value: "str | None") -> None:
        if value is None:
            if "StyleName" in self.attrib:
                del self.attrib["StyleName"]
            return
        self.set("StyleName", value)

    def get_source_by_tag(self, tag: str) -> "CT_Source | None":
        """Return the first ``<b:Source>`` whose ``<b:Tag>`` text equals ``tag``.

        Returns |None| when no matching source is present.
        """
        for source in self.source_lst:
            if source.tag_val == tag:
                return source
        return None

    def add_source_from_kwargs(
        self,
        tag: str,
        title: str | None = None,
        author: str | None = None,
        year: str | int | None = None,
        source_type: str = "Book",
        **extra: str,
    ) -> "CT_Source":
        """Append a ``<b:Source>`` with the supplied fields and return it.

        ``tag`` is the citation key used by ``<w:sdt>`` references and must
        be unique within the part. ``author`` is accepted as a single
        string and written out as a corporate-style ``<b:Corporate>`` name
        for simplicity; richer name lists are left to callers that need
        them. Any ``**extra`` kwargs become simple text-only children —
        e.g. ``city="London"`` becomes ``<b:City>London</b:City>``.
        """
        source = self.add_source()
        # -- element-children order per ECMA spec (friendly subset): tag,
        # -- SourceType, then the commonly-used text fields. Word tolerates
        # -- additional order so we keep it simple. --
        _new_child(source, "Tag", tag)
        _new_child(source, "SourceType", source_type)
        if title is not None:
            _new_child(source, "Title", title)
        if year is not None:
            _new_child(source, "Year", str(year))
        if author is not None:
            author_elm = _new_child(source, "Author")
            inner_author = _new_child(author_elm, "Author")
            _new_child(inner_author, "Corporate", author)
        for key, value in extra.items():
            if value is None:
                continue
            # -- capitalize the first letter to match Word's element-name
            # -- convention (Title, Year, City, Publisher, ...). --
            element_name = key[:1].upper() + key[1:]
            _new_child(source, element_name, str(value))
        return source


class CT_Source(BaseOxmlElement):
    """``<b:Source>`` — a single bibliographic entry inside ``<b:Sources>``."""

    @property
    def tag_val(self) -> "str | None":
        """Value of the child ``<b:Tag>`` element, or |None| if absent."""
        elm = self.find(qn("b:Tag"))
        if elm is None:
            return None
        return elm.text

    @property
    def source_type(self) -> "str | None":
        """Value of the child ``<b:SourceType>`` element, or |None| if absent."""
        elm = self.find(qn("b:SourceType"))
        if elm is None:
            return None
        return elm.text

    @property
    def title(self) -> "str | None":
        elm = self.find(qn("b:Title"))
        if elm is None:
            return None
        return elm.text

    @property
    def year(self) -> "str | None":
        elm = self.find(qn("b:Year"))
        if elm is None:
            return None
        return elm.text

    def field(self, name: str) -> "str | None":
        """Return the text of the ``<b:{name}>`` child element, or |None|."""
        elm = self.find(qn(f"b:{name}"))
        if elm is None:
            return None
        return elm.text

    @property
    def author(self) -> "str | None":
        """First author name (``Corporate`` or formatted ``First Last``), or |None|."""
        author_root = self.find(qn("b:Author"))
        if author_root is None:
            return None
        inner = author_root.find(qn("b:Author"))
        if inner is None:
            inner = author_root
        corp = inner.find(qn("b:Corporate"))
        if corp is not None:
            return corp.text
        name_list = inner.find(qn("b:NameList"))
        if name_list is not None:
            person = name_list.find(qn("b:Person"))
            if person is not None:
                first = person.find(qn("b:First"))
                last = person.find(qn("b:Last"))
                parts = [
                    part.text for part in (first, last) if part is not None and part.text
                ]
                return " ".join(parts) if parts else None
        return None


def new_sources_root(
    selected_style: str = "/APA.XSL", style_name: str = "APA"
) -> "CT_Sources":
    """Return a fresh ``<b:Sources>`` element with default style attributes."""
    xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<b:Sources xmlns:b="{_B_NS}" xmlns="{_B_NS}" '
        f'SelectedStyle="{selected_style}" StyleName="{style_name}"/>'
    )
    return parse_xml(xml.encode("utf-8"))
