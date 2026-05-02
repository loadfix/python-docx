# pyright: reportImportCycles=false

"""XML parser for python-docx."""

from __future__ import annotations

from contextlib import contextmanager
from typing import TYPE_CHECKING, Iterator, cast

from lxml import etree

from docx.oxml.ns import NamespacePrefixedTag, nsmap

if TYPE_CHECKING:
    from docx.oxml.xmlchemy import BaseOxmlElement


# -- configure XML parser --
# Security: resolve_entities=False prevents XXE attacks, no_network=True prevents
# network access during parsing, huge_tree=False prevents XML bombs (billion laughs).
element_class_lookup = etree.ElementNamespaceClassLookup()
oxml_parser = etree.XMLParser(
    remove_blank_text=True,
    resolve_entities=False,
    no_network=True,
    huge_tree=False,
)
oxml_parser.set_element_class_lookup(element_class_lookup)

# -- a second parser configured with recover=True is used for recovery mode. It
# -- shares the same element-class lookup so custom classes are still produced for
# -- elements that survive recovery.
_recovery_parser = etree.XMLParser(
    remove_blank_text=True,
    resolve_entities=False,
    no_network=True,
    huge_tree=False,
    recover=True,
)
_recovery_parser.set_element_class_lookup(element_class_lookup)

# -- a third parser with `huge_tree=True` used opt-in via ``Document(..., huge_tree=True)``
# -- so very large documents (AttValue > 10 MB, > 256 nested elements, etc.) can be
# -- parsed. This relaxes libxml2's built-in XML-bomb protections and should only be
# -- enabled for trusted input. upstream#1086.
_huge_tree_parser = etree.XMLParser(
    remove_blank_text=True,
    resolve_entities=False,
    no_network=True,
    huge_tree=True,
)
_huge_tree_parser.set_element_class_lookup(element_class_lookup)

# -- matching recovery parser with huge_tree enabled, used when both
# -- `recover=True` and `huge_tree=True` are active simultaneously.
_huge_tree_recovery_parser = etree.XMLParser(
    remove_blank_text=True,
    resolve_entities=False,
    no_network=True,
    huge_tree=True,
    recover=True,
)
_huge_tree_recovery_parser.set_element_class_lookup(element_class_lookup)


class _RecoveryState:
    """Opt-in, process-wide recovery-mode state for `parse_xml()`.

    While active, `parse_xml()` falls back to the lxml recover parser when no
    explicit ``recover`` argument is passed, and collected parse warnings are
    appended to :attr:`warnings`. Not thread-safe — intended for single-threaded
    package-open flows.
    """

    def __init__(self) -> None:
        self.active: bool = False
        self.warnings: list[str] = []


_recovery_state = _RecoveryState()


class _HugeTreeState:
    """Process-wide flag enabling the ``huge_tree=True`` lxml parser.

    Controlled via :func:`huge_tree_mode` and used by :func:`parse_xml` to pick
    between the default and huge-tree parsers. Not thread-safe.
    """

    def __init__(self) -> None:
        self.active: bool = False


_huge_tree_state = _HugeTreeState()


@contextmanager
def recovery_mode() -> Iterator[list[str]]:
    """Context manager enabling recovery parsing for the duration of the block.

    Yields the shared list of warning strings collected while active. The list
    is cleared on entry and left populated on exit so callers can read it.
    """
    _recovery_state.active = True
    _recovery_state.warnings = []
    try:
        yield _recovery_state.warnings
    finally:
        _recovery_state.active = False


@contextmanager
def huge_tree_mode() -> Iterator[None]:
    """Context manager enabling the ``huge_tree=True`` lxml parser.

    While active, :func:`parse_xml` uses the huge-tree parser variant, which
    disables libxml2's built-in safety limits (notably the 10 MB AttValue cap
    and the default 256-deep nesting limit). Only enable for trusted input —
    the security guarantees of the default parser no longer apply. Nestable
    with :func:`recovery_mode`.

    .. versionadded:: 1.3.0.dev0
    """
    previous = _huge_tree_state.active
    _huge_tree_state.active = True
    try:
        yield
    finally:
        _huge_tree_state.active = previous


def parse_xml(xml: str | bytes, recover: bool | None = None) -> "BaseOxmlElement":
    """Root lxml element obtained by parsing XML character string `xml`.

    The custom parser is used, so custom element classes are produced for elements in
    `xml` that have them.

    When `recover` is True (or when the ambient :func:`recovery_mode` context is
    active and `recover` is not explicitly False), lxml's recovering parser is
    used: malformed XML yields a best-effort partial element tree instead of
    raising :class:`lxml.etree.XMLSyntaxError`. If the content is completely
    irrecoverable, lxml returns ``None`` — callers in recovery mode must be
    prepared to substitute an empty stub. Any errors encountered during
    recovery are appended as strings to the active recovery-mode warnings list.
    """
    use_recover = recover if recover is not None else _recovery_state.active
    use_huge = _huge_tree_state.active
    if not use_recover:
        parser = _huge_tree_parser if use_huge else oxml_parser
        return cast("BaseOxmlElement", etree.fromstring(xml, parser))

    recovery_parser = _huge_tree_recovery_parser if use_huge else _recovery_parser
    try:
        element = etree.fromstring(xml, recovery_parser)
    except etree.XMLSyntaxError as exc:
        # -- lxml still raises for entirely empty input even with recover=True --
        if _recovery_state.active:
            _recovery_state.warnings.append(str(exc))
        element = None
    # -- collect parse warnings when the ambient recovery context is active --
    if _recovery_state.active:
        for entry in recovery_parser.error_log:
            _recovery_state.warnings.append(str(entry))
    return cast("BaseOxmlElement", element)


def register_element_cls(tag: str, cls: type["BaseOxmlElement"]):
    """Register an lxml custom element-class to use for `tag`.

    A instance of `cls` to be constructed when the oxml parser encounters an element
    with matching `tag`. `tag` is a string of the form `nspfx:tagroot`, e.g.
    `'w:document'`.
    """
    nspfx, tagroot = tag.split(":")
    namespace = element_class_lookup.get_namespace(nsmap[nspfx])
    namespace[tagroot] = cls


def OxmlElement(
    nsptag_str: str,
    attrs: dict[str, str] | None = None,
    nsdecls: dict[str, str] | None = None,
) -> BaseOxmlElement | etree._Element:  # pyright: ignore[reportPrivateUsage]
    """Return a 'loose' lxml element having the tag specified by `nsptag_str`.

    The tag in `nsptag_str` must contain the standard namespace prefix, e.g. `a:tbl`.
    The resulting element is an instance of the custom element class for this tag name
    if one is defined. A dictionary of attribute values may be provided as `attrs`; they
    are set if present. All namespaces defined in the dict `nsdecls` are declared in the
    element using the key as the prefix and the value as the namespace name. If
    `nsdecls` is not provided, a single namespace declaration is added based on the
    prefix on `nsptag_str`.
    """
    nsptag = NamespacePrefixedTag(nsptag_str)
    if nsdecls is None:
        nsdecls = nsptag.nsmap
    return oxml_parser.makeelement(nsptag.clark_name, attrib=attrs, nsmap=nsdecls)
