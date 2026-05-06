"""Namespace-related objects."""

from __future__ import annotations



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
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "xml": "http://www.w3.org/XML/1998/namespace",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

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
