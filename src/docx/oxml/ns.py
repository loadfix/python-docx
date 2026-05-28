"""Namespace-related objects.

The ``nsmap`` is assembled from the shared ``ooxml_opc.namespaces`` extension
registry (the canonical source for Microsoft 2010+ extension prefixes that
Office writes on every default-authored document) plus docx-local entries
for aliases and Dublin-Core / OPC core namespaces.

See ``ooxml_opc.namespaces`` for the catalogue of ``w14`` / ``w15`` /
``w16*`` / ``wp14`` / ``cx1..cx8`` / ``oel`` / ``aink`` / ``am3d``
prefixes that all default-authored Word documents declare on the
``<w:document>`` root. Closes the "Extension namespaces beyond w14 are
unregistered" gap from the docx audit (2026-05-08).
"""

from __future__ import annotations

from ooxml_opc.namespaces import EXTENSION_NAMESPACES


# -- Start from the legacy docx-local nsmap (what docx has always carried).
# -- The extension subset is merged in below *additively*: only the
# -- Microsoft 2010+ extension URIs that docx never registered locally
# -- (``w15``, ``w16*``, ``wp14``, ``oel``, ``aink``, ``am3d``, ``cx1..cx8``)
# -- are added. The OPC-layer prefixes (``ct:``, ``pr:``) are deliberately
# -- NOT pulled in — those stay owned by ``ooxml_opc`` and are resolved via
# -- the composite xmlchemy namespace registry at descriptor lookup time.
# -- Adding ``ct:`` here would route ``ct:Default`` element-class lookup
# -- through docx's lxml parser (which doesn't know about ``CT_Default``)
# -- instead of ooxml_opc's, breaking Content_Types.xml serialisation. --
nsmap = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
    "b": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    # -- ``cst`` / ``ep`` are the prefixes the shared ``ooxml_docprops``
    # -- package uses in its xmlchemy descriptors (``ZeroOrMore("cst:property")``
    # -- etc.). Keep the docx-native ``custprops`` / ``extprops`` prefixes as the
    # -- primary entries (they appear in fixture XML, serialisation output, and
    # -- historical element-class registrations); the ``cst`` / ``ep`` aliases
    # -- map to the same URIs so descriptors declared against the shared prefix
    # -- resolve correctly when evaluated under docx's namespace registry.
    "cst": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "custprops": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "extprops": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "dcterms": "http://purl.org/dc/terms/",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "inkml": "http://www.w3.org/2003/InkML",
    # -- ``lfxbind`` (loadfix bind tokens) is a fork-defined namespace used
    # -- to preserve the original token string of a save-time-resolved
    # -- text run (e.g. ``"Dear {customer.name}"``) so that subsequent
    # -- ``load -> bind -> save`` cycles re-resolve against the new
    # -- record instead of the previously-stamped literal. Word and
    # -- every other consumer follow the "preserve but ignore unknown
    # -- children" convention. See :mod:`docx.bind_tokens` (issue #68).
    "lfxbind": "https://loadfix.dev/docx/bind-tokens",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    # -- Microsoft-extension namespaces emitted by Word 2012+ on every
    # -- default-authored document, gated via `mc:Ignorable`. Required for
    # -- `qn()` lookups from proxy / element classes that touch docId,
    # -- chartTrackingRefBased, commentsExtensible, symEx, and related
    # -- extension children. The cross-format opc-extensions effort
    # -- consolidates these into a shared registry; the local copy here
    # -- keeps `qn()` working until that migration lands.
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "xml": "http://www.w3.org/XML/1998/namespace",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

# -- Pull in every Microsoft extension prefix Word writes on the document
# -- root that was NOT already in the docx-local nsmap. These are additive:
# -- docx.oxml.ns never owned its own copy of these URIs, so adding them
# -- closes the audit gap "qn('w15:docId') raises KeyError" without
# -- re-routing any existing prefix through a different element-class
# -- lookup. Word 2012+ writes these namespaces in every default-authored
# -- .docx, gated by ``mc:Ignorable`` on the root. --
_WORD_DOC_ROOT_EXTENSIONS = (
    "w15",
    "w16",
    "w16cex",
    "w16du",
    "w16sdtdh",
    "w16sdtfl",
    "w16se",
    "wp14",
    "wpi",
    "wne",
    "oel",
    "aink",
    "am3d",
    "cx",
    "cx1",
    "cx2",
    "cx3",
    "cx4",
    "cx5",
    "cx6",
    "cx7",
    "cx8",
)
for _pfx in _WORD_DOC_ROOT_EXTENSIONS:
    nsmap.setdefault(_pfx, EXTENSION_NAMESPACES[_pfx])
del _pfx, _WORD_DOC_ROOT_EXTENSIONS

pfxmap = {value: key for key, value in nsmap.items()}


class NamespacePrefixedTag(str):
    """Value object that knows the semantics of an XML tag having a namespace prefix."""

    def __new__(cls, nstag: str):
        return super(NamespacePrefixedTag, cls).__new__(cls, nstag)

    def __init__(self, nstag: str):
        self._pfx, self._local_part = nstag.split(":")
        self._ns_uri = nsmap[self._pfx]

    @property
    def clark_name(self) -> str:
        return "{%s}%s" % (self._ns_uri, self._local_part)

    @classmethod
    def from_clark_name(cls, clark_name: str) -> NamespacePrefixedTag:
        nsuri, local_name = clark_name[1:].split("}")
        nstag = "%s:%s" % (pfxmap[nsuri], local_name)
        return cls(nstag)

    @property
    def local_part(self) -> str:
        """The local part of this tag.

        E.g. "foobar" is returned for tag "f:foobar".
        """
        return self._local_part

    @property
    def nsmap(self) -> dict[str, str]:
        """Single-member dict mapping prefix of this tag to it's namespace name.

        Example: `{"f": "http://foo/bar"}`. This is handy for passing to xpath calls
        and other uses.
        """
        return {self._pfx: self._ns_uri}

    @property
    def nspfx(self) -> str:
        """The namespace-prefix for this tag.

        For example, "f" is returned for tag "f:foobar".
        """
        return self._pfx

    @property
    def nsuri(self) -> str:
        """The namespace URI for this tag.

        For example, "http://foo/bar" would be returned for tag "f:foobar" if the "f"
        prefix maps to "http://foo/bar" in nsmap.
        """
        return self._ns_uri


def nsdecls(*prefixes: str) -> str:
    """Namespace declaration including each namespace-prefix in `prefixes`.

    Handy for adding required namespace declarations to a tree root element.
    """
    return " ".join(['xmlns:%s="%s"' % (pfx, nsmap[pfx]) for pfx in prefixes])


def nspfxmap(*nspfxs: str) -> dict[str, str]:
    """Subset namespace-prefix mappings specified by *nspfxs*.

    Any number of namespace prefixes can be supplied, e.g. namespaces("a", "r", "p").
    """
    return {pfx: nsmap[pfx] for pfx in nspfxs}


def qn(tag: str) -> str:
    """Stands for "qualified name".

    This utility function converts a familiar namespace-prefixed tag name like "w:p"
    into a Clark-notation qualified tag name for lxml. For example, `qn("w:p")` returns
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p".
    """
    prefix, tagroot = tag.split(":")
    uri = nsmap[prefix]
    return "{%s}%s" % (uri, tagroot)
