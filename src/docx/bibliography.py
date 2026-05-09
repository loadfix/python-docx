"""Bibliography / citation-source proxy types.

A bibliography is a collection of :class:`Source` entries persisted in a
``/customXml/item{N}.xml`` part with a ``<b:Sources>`` root element (in the
``http://schemas.openxmlformats.org/officeDocument/2006/bibliography``
namespace). Each entry has a unique ``tag`` that citation SDTs in the main
document part refer to via a ``CITATION`` complex-field instruction.

The read path is exposed via :attr:`Document.bibliography`; the write path is
rooted in :meth:`Document.add_citation`.

.. versionadded:: 2026.05.7
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from docx.fields import Field
    from docx.oxml.bibliography import CT_Source, CT_Sources
    from docx.parts.bibliography import BibliographyPart


# -- ECMA-376 Part 1 §22.6.2.21 source-type catalogue. These are the values
# -- Word writes into <b:SourceType>. Unknown values are accepted on read
# -- but rejected (with a helpful message) when supplied to add_source. --
VALID_SOURCE_TYPES: "frozenset[str]" = frozenset(
    {
        "Book",
        "BookSection",
        "JournalArticle",
        "ArticleInAPeriodical",
        "ConferenceProceedings",
        "Report",
        "Misc",
        "InternetSite",
        "Film",
        "SoundRecording",
        "Performance",
        "Art",
        "DocumentFromInternetSite",
        "ElectronicSource",
        "Case",
        "Patent",
        "Interview",
    }
)


class Source:
    """Proxy for a single ``<b:Source>`` element.

    Exposes the commonly-used bibliographic fields (``tag``, ``title``,
    ``author``, ``year``, ``source_type``) as read-only properties.

    .. versionadded:: 2026.05.7
    """

    def __init__(self, element: "CT_Source"):
        self._element = element

    @property
    def element(self) -> "CT_Source":
        """The underlying ``<b:Source>`` lxml element.

        .. versionadded:: 2026.05.7
        """
        return self._element

    @property
    def tag(self) -> "str | None":
        """Unique citation key (``<b:Tag>`` child text), or |None| when absent.

        .. versionadded:: 2026.05.7
        """
        return self._element.tag_val

    @property
    def title(self) -> "str | None":
        """``<b:Title>`` text, or |None|.

        .. versionadded:: 2026.05.7
        """
        return self._element.title

    @property
    def author(self) -> "str | None":
        """Flattened author name for this source, or |None|.

        Returns the ``<b:Corporate>`` value when set, otherwise a best-effort
        ``First Last`` join for the first ``<b:Person>`` of the first
        ``<b:NameList>``.

        .. versionadded:: 2026.05.7
        """
        return self._element.author

    @property
    def year(self) -> "str | None":
        """``<b:Year>`` text, or |None|.

        .. versionadded:: 2026.05.7
        """
        return self._element.year

    @property
    def source_type(self) -> "str | None":
        """``<b:SourceType>`` text — e.g. ``"Book"``, ``"JournalArticle"``.

        .. versionadded:: 2026.05.7
        """
        return self._element.source_type

    @property
    def publisher(self) -> "str | None":
        """``<b:Publisher>`` text, or |None|.

        .. versionadded:: 2026.05.10
        """
        return self._element.field("Publisher")

    @property
    def city(self) -> "str | None":
        """``<b:City>`` text, or |None|.

        .. versionadded:: 2026.05.10
        """
        return self._element.field("City")

    def field(self, name: str) -> "str | None":
        """Return the text of the ``<b:{name}>`` child element, or |None|.

        Generic accessor for source fields not exposed as dedicated
        properties — e.g. ``source.field("Medium")``, ``source.field("URL")``.
        Names are matched exactly (Word uses PascalCase for every field).

        .. versionadded:: 2026.05.10
        """
        return self._element.field(name)

    def __repr__(self) -> str:
        return f"<Source tag={self.tag!r} title={self.title!r} year={self.year!r}>"


class Bibliography:
    """Collection proxy for the ``<b:Sources>`` element of a bibliography part.

    Iteration yields :class:`Source` proxies in document order. Per-tag
    lookup is supported via :meth:`get_by_tag`.

    .. versionadded:: 2026.05.7
    """

    def __init__(self, sources: "CT_Sources", part: "BibliographyPart | None" = None):
        self._sources = sources
        self._part = part

    @property
    def element(self) -> "CT_Sources":
        """The underlying ``<b:Sources>`` lxml element.

        .. versionadded:: 2026.05.7
        """
        return self._sources

    @property
    def part(self) -> "BibliographyPart | None":
        """The :class:`BibliographyPart` that holds this collection, or |None|.

        |None| when the :class:`Bibliography` was constructed from a bare
        ``<b:Sources>`` element (e.g. in unit tests). Otherwise the part
        provides the ``{GUID}`` store-item id used to bind citation SDTs.

        .. versionadded:: 2026.05.7
        """
        return self._part

    @property
    def sources(self) -> "list[Source]":
        """List of every :class:`Source` in this bibliography, in document order.

        .. versionadded:: 2026.05.7
        """
        return [Source(e) for e in self._sources.source_lst]

    def __iter__(self) -> "Iterator[Source]":
        return iter(self.sources)

    def __len__(self) -> int:
        return len(self._sources.source_lst)

    def get_by_tag(self, tag: str) -> "Source | None":
        """Return the :class:`Source` whose ``tag`` matches, or |None| if none do.

        .. versionadded:: 2026.05.7
        """
        found = self._sources.get_source_by_tag(tag)
        if found is None:
            return None
        return Source(found)

    @property
    def selected_style(self) -> "str | None":
        """Value of ``<b:Sources>/@SelectedStyle`` (e.g. ``"/APA.XSL"``), or |None|.

        .. versionadded:: 2026.05.7
        """
        return self._sources.selected_style

    @selected_style.setter
    def selected_style(self, value: "str | None") -> None:
        self._sources.selected_style = value

    @property
    def style_name(self) -> "str | None":
        """Value of ``<b:Sources>/@StyleName`` (e.g. ``"APA"``), or |None|.

        .. versionadded:: 2026.05.7
        """
        return self._sources.style_name

    @style_name.setter
    def style_name(self, value: "str | None") -> None:
        self._sources.style_name = value

    def add_source(
        self,
        tag: str,
        title: "str | None" = None,
        author: "str | None" = None,
        year: "str | int | None" = None,
        source_type: str = "Book",
        **extra: str,
    ) -> Source:
        """Append a new :class:`Source` with the given fields and return it.

        ``tag`` is the citation key and must be unique within this
        bibliography; reusing an existing tag will raise
        :class:`ValueError`.

        ``extra`` kwargs become text-only children under ``<b:Source>`` —
        e.g. ``city="London"`` becomes ``<b:City>London</b:City>``. Unknown
        keys are pass-through; no validation is performed against the
        ECMA-376 type catalogue.

        .. versionadded:: 2026.05.7
        """
        if self._sources.get_source_by_tag(tag) is not None:
            raise ValueError(f"bibliography already has a source with tag {tag!r}")
        if source_type not in VALID_SOURCE_TYPES:
            raise ValueError(
                f"unknown source_type {source_type!r}; expected one of "
                f"{sorted(VALID_SOURCE_TYPES)}"
            )
        elm = self._sources.add_source_from_kwargs(
            tag,
            title=title,
            author=author,
            year=year,
            source_type=source_type,
            **extra,
        )
        return Source(elm)


# -- CITATION field instruction parsing ----------------------------------

# -- The CITATION instruction looks like:
# --    CITATION smith2020 \p "45-48" \f "cf. " \s ", et al."
# -- with optional \l <lcid>, \n (suppress author), \t (suppress title),
# -- \m <tag> (combined multi-source) switches. We only surface the fields
# -- the spec asks for: source_tag, pages, prefix, suffix. --
_CITATION_INSTR_RX = re.compile(
    r"^\s*CITATION\s+(?P<tag>\S+)(?P<rest>.*)$", re.DOTALL
)


def _parse_switch_value(rest: str, switch: str) -> "str | None":
    """Extract the quoted or bare argument following ``\\<switch>`` in `rest`.

    Returns |None| if the switch is not present. Handles both
    ``\\p "45-48"`` and ``\\p 45-48`` spellings.
    """
    # -- look for \<switch> followed by whitespace --
    pattern = re.compile(
        rf"\\{re.escape(switch)}\s+(?:\"(?P<q>[^\"]*)\"|(?P<b>\S+))"
    )
    match = pattern.search(rest)
    if match is None:
        return None
    return match.group("q") if match.group("q") is not None else match.group("b")


class Citation:
    """Proxy for a single ``CITATION`` field inside the document body.

    Exposes the citation's source tag and the optional ``\\p``/``\\f``/``\\s``
    switch values (pages, prefix, suffix) as read-only properties. The
    underlying |Field| is available as :attr:`field` for callers that need
    to mutate the rendered result or dig into the raw runs.

    .. versionadded:: 2026.05.10
    """

    def __init__(self, field: "Field"):
        self._field = field

    @property
    def field(self) -> "Field":
        """The underlying |Field| proxy wrapping this CITATION.

        .. versionadded:: 2026.05.10
        """
        return self._field

    @property
    def source_tag(self) -> str:
        """The bibliography ``tag`` this citation refers to."""
        match = _CITATION_INSTR_RX.match(self._field.instruction)
        if match is None:
            return ""
        return match.group("tag")

    @property
    def pages(self) -> "str | None":
        """Value of the ``\\p`` switch (page-range override), or |None|."""
        return self._switch("p")

    @property
    def prefix(self) -> "str | None":
        """Value of the ``\\f`` switch (citation prefix), or |None|."""
        return self._switch("f")

    @property
    def suffix(self) -> "str | None":
        """Value of the ``\\s`` switch (citation suffix), or |None|."""
        return self._switch("s")

    def _switch(self, name: str) -> "str | None":
        match = _CITATION_INSTR_RX.match(self._field.instruction)
        if match is None:
            return None
        return _parse_switch_value(match.group("rest"), name)

    def __repr__(self) -> str:
        return (
            f"<Citation source_tag={self.source_tag!r} pages={self.pages!r}>"
        )


def is_citation_instruction(instruction: str) -> bool:
    """Return ``True`` if `instruction` is a ``CITATION`` field instruction.

    Leading whitespace and case-exact ``CITATION`` required, matching
    Word's field-grammar convention.
    """
    return _CITATION_INSTR_RX.match(instruction) is not None
