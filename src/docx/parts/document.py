"""|DocumentPart| and closely related objects."""

from __future__ import annotations

import secrets
from typing import IO, TYPE_CHECKING, cast

from docx.oxml.ns import qn

from docx.document import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.bibliography import BibliographyPart
from docx.parts.comments import CommentsPart
from docx.parts.custom_properties import CustomPropertiesPart
from docx.parts.endnotes import EndnotesPart
from docx.parts.extended_properties import ExtendedPropertiesPart
from docx.parts.font_table import FontTablePart
from docx.parts.footnotes import FootnotesPart
from docx.parts.glossary import GlossaryPart
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.numbering import NumberingPart
from docx.parts.settings import SettingsPart
from docx.parts.story import StoryPart
from docx.parts.styles import StylesPart
from docx.parts.theme import ThemePart
from docx.parts.web_settings import WebSettingsPart
from docx.shape import InlineShapes
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from ooxml_comments import CommentIds, CommentsExtensible

    from docx.bibliography import Bibliography
    from docx.comments import Comments
    from docx.custom_properties import CustomProperties
    from docx.custom_xml import CustomXmlPart as CustomXmlPartProxy
    from docx.endnotes import Endnotes
    from docx.enum.style import WD_STYLE_TYPE
    from docx.extended_properties import ExtendedProperties
    from docx.font_table import FontTable
    from docx.footnotes import Footnotes
    from docx.glossary import Glossary
    from docx.opc.coreprops import CoreProperties
    from docx.settings import Settings
    from docx.styles.style import BaseStyle
    from docx.theme import Theme
    from docx.web_settings import WebSettings


class DocumentPart(StoryPart):
    """Main document part of a WordprocessingML (WML) package, aka a .docx file.

    Acts as broker to other parts such as image, core properties, and style parts. It
    also acts as a convenient delegate when a mid-document object needs a service
    involving a remote ancestor. The `Parented.part` property inherited by many content
    objects provides access to this part object for that purpose.
    """

    def add_footer_part(self):
        """Return (footer_part, rId) pair for newly-created footer part."""
        footer_part = FooterPart.new(self.package)
        rId = self.relate_to(footer_part, RT.FOOTER)
        return footer_part, rId

    def add_header_part(self):
        """Return (header_part, rId) pair for newly-created header part."""
        header_part = HeaderPart.new(self.package)
        rId = self.relate_to(header_part, RT.HEADER)
        return header_part, rId

    @property
    def bibliography(self) -> "Bibliography":
        """|Bibliography| collection of citation sources for this document.

        Lazily creates a ``/customXml/item{N}.xml`` part carrying an empty
        ``<b:Sources>`` root (plus the sibling ``itemProps{N}.xml`` datastore
        part) if no bibliography is already related to the document.

        .. versionadded:: 2026.05.7
        """
        from docx.bibliography import Bibliography

        part = self._bibliography_part
        return Bibliography(part.sources_element, part)

    @property
    def _bibliography_part(self) -> BibliographyPart:
        """Existing |BibliographyPart|, or a newly created empty one.

        Searches every ``customXml`` relationship for a data part whose
        root element is ``<b:Sources>``; when none is found, builds a fresh
        :class:`BibliographyPart` and attaches a sibling properties part.

        .. versionadded:: 2026.05.7
        """
        from lxml import etree

        from docx.oxml.bibliography import CT_Sources
        from docx.oxml.parser import parse_xml

        b_sources_tag = qn("b:Sources")

        for rel in self.rels.values():
            if rel.is_external or rel.reltype != RT.CUSTOM_XML:
                continue
            try:
                target = rel.target_part
            except ValueError:
                continue
            blob = getattr(target, "blob", b"")
            if not blob:
                continue
            # -- peek at the root element name using the hardened parser
            # -- (resolve_entities=False, no_network=True) so malicious
            # -- customXml data parts cannot exfiltrate files via XXE. --
            try:
                root = parse_xml(blob)
            except etree.XMLSyntaxError:
                continue
            if root.tag != b_sources_tag:
                continue
            # -- found one; upgrade the part into a BibliographyPart if it
            # -- isn't already. The PartFactory default for CT.XML is the
            # -- plain CustomXmlPart, so most existing packages need this
            # -- in-place rebind. --
            if isinstance(target, BibliographyPart):
                return target
            # -- swap for a BibliographyPart holding the same blob/partname --
            bib_part = BibliographyPart(target.partname, target.content_type, blob)
            # -- preserve the sibling props rel(s) --
            for props_rel in list(target.rels.values()):
                if props_rel.is_external:
                    continue
                try:
                    props_target = props_rel.target_part
                except ValueError:
                    continue
                bib_part.relate_to(props_target, props_rel.reltype)
            # -- rewire package: replace old part --
            assert self.package is not None
            # -- rewire the existing relationship rather than dropping/re-adding
            # -- so we preserve the existing rId value. --
            rel._target = bib_part
            return bib_part

        # -- none found: create from scratch --
        assert self.package is not None
        bib_part = BibliographyPart.default(self.package)
        self.relate_to(bib_part, RT.CUSTOM_XML)
        bib_part.attach_itemProps(self.package)
        return bib_part

    @property
    def comments(self) -> Comments:
        """|Comments| object providing access to the comments added to this document."""
        return self._comments_part.comments

    @property
    def comments_ids(self) -> "CommentIds":
        """|CommentIds| proxy over ``word/commentsIds.xml``.

        Lazily creates the part (and its relationship from
        ``word/comments.xml``) on first access so callers can read/write
        without juggling the relationship manually. The returned proxy
        wraps the live ``<w16cid:commentsIds>`` element, so mutations
        through it persist on save.

        .. versionadded:: 2026.05.10
        """
        from ooxml_comments import CommentIds

        part = self._comments_part.comments_ids_part_or_add()
        return CommentIds(part.element)

    @property
    def comments_extensible(self) -> "CommentsExtensible":
        """|CommentsExtensible| proxy over ``word/commentsExtensible.xml``.

        Lazily creates the part (and its relationship from
        ``word/comments.xml``) on first access. See :attr:`comments_ids`
        for semantics.

        .. versionadded:: 2026.05.10
        """
        from ooxml_comments import CommentsExtensible

        part = self._comments_part.comments_extensible_part_or_add()
        return CommentsExtensible(part.element)

    @property
    def endnotes(self) -> Endnotes:
        """|Endnotes| object providing access to the endnotes in this document."""
        return self._endnotes_part.endnotes

    @property
    def _endnotes_part(self) -> EndnotesPart:
        """A |EndnotesPart| providing access to the endnotes for this document.

        Creates a default endnotes part if one is not present.
        """
        try:
            return cast(EndnotesPart, self.part_related_by(RT.ENDNOTES))
        except KeyError:
            assert self.package is not None
            endnotes_part = EndnotesPart.default(self.package)
            self.relate_to(endnotes_part, RT.ENDNOTES)
            return endnotes_part

    @property
    def font_table(self) -> FontTable | None:
        """A |FontTable| for this document, or |None| if no font-table part is related.

        Read access returns |None| when no ``fontTable`` relationship exists.
        To embed a font use :meth:`font_table_or_new` (or call
        :meth:`FontTable.add_embedded_font` through it), which will create an
        empty font-table part on demand.
        """
        font_table_part = self._font_table_part
        if font_table_part is None:
            return None
        return font_table_part.font_table

    @property
    def font_table_or_new(self) -> FontTable:
        """A |FontTable| for this document, creating an empty one if needed.

        Unlike :attr:`font_table` this always returns a live |FontTable|; a
        default (empty) ``word/fontTable.xml`` part is added to the package and
        related to the document on the first call.

        .. versionadded:: 2026.05.0
        """
        return self._font_table_part_or_new.font_table

    @property
    def _font_table_part(self) -> FontTablePart | None:
        """The |FontTablePart| related to this document, or |None| if not present.

        A default font-table part is not created on demand; use
        :attr:`_font_table_part_or_new` to force creation.
        """
        try:
            return cast(FontTablePart, self.part_related_by(RT.FONT_TABLE))
        except KeyError:
            return None

    @property
    def _font_table_part_or_new(self) -> FontTablePart:
        """Existing |FontTablePart|, or a newly created empty one, related by ``fontTable``.

        .. versionadded:: 2026.05.0
        """
        try:
            return cast(FontTablePart, self.part_related_by(RT.FONT_TABLE))
        except KeyError:
            assert self.package is not None
            part = FontTablePart.default(self.package)
            self.relate_to(part, RT.FONT_TABLE)
            return part

    @property
    def footnotes(self) -> Footnotes:
        """|Footnotes| object providing access to the footnotes in this document."""
        return self._footnotes_part.footnotes

    @property
    def _footnotes_part(self) -> FootnotesPart:
        """A |FootnotesPart| providing access to the footnotes for this document.

        Creates a default footnotes part if one is not present.
        """
        try:
            return cast(FootnotesPart, self.part_related_by(RT.FOOTNOTES))
        except KeyError:
            assert self.package is not None
            footnotes_part = FootnotesPart.default(self.package)
            self.relate_to(footnotes_part, RT.FOOTNOTES)
            return footnotes_part

    @property
    def glossary(self) -> Glossary | None:
        """A |Glossary| proxy for this document, or |None| when no glossary part is related.

        Prefer :meth:`ensure_glossary` when you want a writable glossary —
        this property preserves backward compatibility (returns |None| for
        read-only callers inspecting a document that has no glossary).
        """
        glossary_part = self._glossary_part
        if glossary_part is None:
            return None
        return glossary_part.glossary

    def ensure_glossary(self) -> Glossary:
        """A |Glossary| proxy, creating an empty glossary part if none exists.

        When the document already has a ``glossaryDocument`` relationship
        the existing part is used; otherwise a fresh, empty
        :class:`GlossaryPart` is created and related. Callers that need to
        mutate the glossary (for example by adding a building block) should
        use this method rather than the :attr:`glossary` property, which
        returns |None| for documents without a pre-existing glossary.

        .. versionadded:: 2026.05.10
        """
        glossary_part = self._glossary_part
        if glossary_part is None:
            assert self.package is not None
            glossary_part = GlossaryPart.default(self.package)
            self.relate_to(glossary_part, RT.GLOSSARY_DOCUMENT)
        return glossary_part.glossary

    def remove_glossary(self) -> bool:
        """Drop the ``glossaryDocument`` relationship, if one exists.

        Returns ``True`` when a glossary relationship was dropped, ``False``
        when there was no glossary to remove.

        .. versionadded:: 2026.05.10
        """
        for rId, rel in list(self.rels.items()):
            if rel.reltype == RT.GLOSSARY_DOCUMENT:
                self.drop_rel(rId)
                return True
        return False

    @property
    def _glossary_part(self) -> GlossaryPart | None:
        """The |GlossaryPart| related to this document, or |None| if not present.

        A default glossary part is **not** created on demand here — use
        :meth:`ensure_glossary` for write access.
        """
        try:
            return cast(GlossaryPart, self.part_related_by(RT.GLOSSARY_DOCUMENT))
        except KeyError:
            return None

    @property
    def core_properties(self) -> CoreProperties:
        """A |CoreProperties| object providing read/write access to the core properties
        of this document."""
        return self.package.core_properties

    @property
    def custom_properties(self) -> CustomProperties:
        """A |CustomProperties| collection for this document.

        Creates an empty custom properties part lazily if one is not already present.
        """
        return self._custom_properties_part.custom_properties

    @property
    def custom_xml_parts(self) -> list[CustomXmlPartProxy]:
        """List of |CustomXmlPart| proxies for each related custom XML data part.

        Empty when the document has no ``customXml`` relationships. Read-only.
        """
        from docx.custom_xml import iter_custom_xml_parts

        return iter_custom_xml_parts(self)

    @property
    def _custom_properties_part(self) -> CustomPropertiesPart:
        """Return the package-scoped |CustomPropertiesPart| for this document.

        The custom-properties part is related to the **package** (not the
        main-document part), following the convention used for core and
        extended properties. Microsoft Word rejects packages whose
        ``custom-properties`` relationship lands on
        ``word/_rels/document.xml.rels`` instead of the package-root
        ``_rels/.rels`` (issue #712), so this delegates to
        :attr:`OpcPackage._custom_properties_part` to ensure the
        relationship is written in the canonical location. A default
        (empty) part is created on demand when none is present.

        .. versionchanged:: 2026.06.0
           Delegate to the package; previously created the rel on the
           main-document part, producing a docx Word refused to open.
        """
        assert self.package is not None
        return cast(
            CustomPropertiesPart, self.package._custom_properties_part
        )  # pyright: ignore[reportPrivateUsage]

    @property
    def extended_properties(self) -> ExtendedProperties:
        """An |ExtendedProperties| proxy for the document's ``app.xml`` part.

        The extended-properties part is related to the package (not the
        main-document part), following the convention used for core properties.
        A default (empty) part is created on demand when none is present.

        .. versionadded:: 2026.05.0
        """
        return self._extended_properties_part.extended_properties

    @property
    def _extended_properties_part(self) -> ExtendedPropertiesPart:
        """Return the package-scoped |ExtendedPropertiesPart| for this document.

        .. versionadded:: 2026.05.0
        """
        assert self.package is not None
        return cast(
            ExtendedPropertiesPart, self.package._extended_properties_part
        )  # pyright: ignore[reportPrivateUsage]

    @property
    def document(self):
        """A |Document| object providing access to the content of this document."""
        return Document(self._element, self)

    def drop_header_part(self, rId: str) -> None:
        """Remove related header part identified by `rId`."""
        self.drop_rel(rId)

    def footer_part(self, rId: str):
        """Return |FooterPart| related by `rId`."""
        return self.related_parts[rId]

    def get_style(self, style_id: str | None, style_type: WD_STYLE_TYPE) -> BaseStyle:
        """Return the style in this document matching `style_id`.

        Returns the default style for `style_type` if `style_id` is |None| or does not
        match a defined style of `style_type`.
        """
        return self.styles.get_by_id(style_id, style_type)

    def get_style_id(self, style_or_name, style_type):
        """Return the style_id (|str|) of the style of `style_type` matching
        `style_or_name`.

        Returns |None| if the style resolves to the default style for `style_type` or if
        `style_or_name` is itself |None|. Raises if `style_or_name` is a style of the
        wrong type or names a style not present in the document.
        """
        return self.styles.get_style_id(style_or_name, style_type)

    def header_part(self, rId: str):
        """Return |HeaderPart| related by `rId`."""
        return self.related_parts[rId]

    @lazyproperty
    def inline_shapes(self):
        """The |InlineShapes| instance containing the inline shapes in the document."""
        return InlineShapes(self._element.body, self)

    @lazyproperty
    def numbering_part(self) -> NumberingPart:
        """A |NumberingPart| object providing access to the numbering definitions for this document.

        Creates an empty numbering part if one is not present.
        """
        try:
            return cast(NumberingPart, self.part_related_by(RT.NUMBERING))
        except KeyError:
            numbering_part = NumberingPart.new()
            self.relate_to(numbering_part, RT.NUMBERING)
            return numbering_part

    def before_marshal(self, reproducible: bool = False) -> None:
        """Stamp Word-style identifiers on paragraphs/runs just before save.

        Word emits ``w14:paraId``, ``w14:textId``, and session-wide
        ``w:rsidR`` / ``w:rsidRDefault`` on every paragraph it authors,
        plus ``w:rsidR`` on every run. These are collaboration/merge
        tracking identifiers — missing ones cause diffing tools and
        modern-comments features to behave inconsistently.

        This hook mints the identifiers lazily: existing values are
        preserved (so round-trips stay deterministic), missing ones are
        stamped with newly-generated 8-hex-digit tokens. The session's
        ``rsidRoot`` is generated once per save call.

        Also records the session's rsid in the settings part's
        ``<w:rsids>`` table so the values are reachable from
        ``w:rsids`` consumers.

        When ``reproducible`` is True the rsid-family attributes
        (``w:rsidR`` / ``w:rsidRDefault``) are **not** minted onto
        elements that don't already carry them. Existing rsid values
        are preserved (so a round-trip is faithful), but no new
        session-scoped churn markers are introduced. ``w14:paraId``
        and ``w14:textId`` are still stamped when missing because
        they are derived deterministically from each paragraph's
        content, so repeated saves remain byte-identical.

        When the document part was loaded from an on-disk package
        (``_loaded_from_package=True``) the minting steps are
        suppressed outright — Word itself does not retroactively stamp
        rsid/paraId/textId attributes on paragraphs and runs it did
        not author in the current editing session, and doing so on a
        pure read+save round-trip silently changes user data and
        breaks byte-identical fidelity. Any paragraphs python-docx
        subsequently appends via its API inherit the underlying
        element defaults (no rsid/paraId/textId), which are all
        optional per ECMA-376 and accepted by Word.

        .. versionadded:: 2026.05.2
        .. versionchanged:: 2026.05.11
            Suppress rsid/paraId/textId minting on parts loaded from a
            package — preserves byte-identical fidelity on read+save
            round-trips of Office-authored files.
        """
        # Fidelity gate: parts that were loaded from a package are
        # preserved verbatim. Minting new rsid/paraId/textId attributes
        # onto already-authored paragraphs and runs causes every
        # ``Document(path).save(out)`` to drift by several attributes
        # and force a new ``<w:rsid>`` entry into ``word/settings.xml``
        # — user-data-mutating behaviour for a nominally read-only
        # round-trip. The ``reproducible`` branch below still runs its
        # paraId/textId deterministic minting pass because those are
        # content-derived tokens that remain byte-stable across
        # identical inputs.
        loaded_from_package = getattr(self, "_loaded_from_package", False)
        if loaded_from_package and not reproducible:
            # Still want to keep settings/<w:rsids> consistent with
            # whatever rsids the loaded document already references,
            # but we do not introduce new session-scoped tokens.
            _mirror_run_formatting_to_paragraph_mark(self.element)
            self._resolve_data_source_bindings()
            self._stamp_data_source_markers()
            self._drop_unused_optional_parts(self.element)
            return

        mint = _DeterministicMinter() if reproducible else _RandomMinter()
        rsid_root = mint.rsid_root()
        rsids_seen: set[str] = set()

        paraId_tag = qn("w14:paraId")
        textId_tag = qn("w14:textId")
        rsidR_tag = qn("w:rsidR")
        rsidRDefault_tag = qn("w:rsidRDefault")

        root = self.element
        for p in root.iter(qn("w:p")):
            if not p.get(paraId_tag):
                p.set(paraId_tag, mint.paraId(p))
            if not p.get(textId_tag):
                p.set(textId_tag, mint.textId(p))
            # Under reproducible=True do not mint rsid-family attributes
            # on elements that don't already carry them. Those markers
            # are session-scoped churn that serve no purpose in a
            # content-deterministic artefact (W8-B).
            if not reproducible:
                if not p.get(rsidR_tag):
                    p.set(rsidR_tag, rsid_root)
                if not p.get(rsidRDefault_tag):
                    p.set(rsidRDefault_tag, rsid_root)
            existing_p_rsidR = p.get(rsidR_tag)
            if existing_p_rsidR:
                rsids_seen.add(existing_p_rsidR)

            for r in p.iter(qn("w:r")):
                if not reproducible and not r.get(rsidR_tag):
                    r.set(rsidR_tag, rsid_root)
                existing_r_rsidR = r.get(rsidR_tag)
                if existing_r_rsidR:
                    rsids_seen.add(existing_r_rsidR)

        # Persist the session's rsid into settings so Word accepts the
        # file without warning. Skip this in reproducible mode when no
        # rsids were seen (nothing to persist and no point adding a
        # synthetic one). The settings part may not yet exist (empty
        # template) — the property getter creates it on demand.
        try:
            settings = self.settings
            if reproducible:
                if rsids_seen:
                    settings.add_rsids(next(iter(sorted(rsids_seen))), extra=rsids_seen)
            else:
                settings.add_rsids(rsid_root, extra=rsids_seen)
        except Exception:  # pragma: no cover - defensive: don't break save
            pass

        # Mirror run formatting onto the paragraph mark (pPr/rPr) so that
        # typing after the last run in Word continues with the same
        # formatting. This is the "keep typing in bold" convention Word
        # emits by default.
        _mirror_run_formatting_to_paragraph_mark(root)

        # Resolve and inline custom-XML data-binding values for every SDT
        # whose <w:dataBinding> points at a registered data source (#80).
        # Re-bound sources are honoured automatically because the resolver
        # reads the live blob from the data part each save.
        self._resolve_data_source_bindings()
        # Persist the logical-name marker into each bound data part's
        # payload so the next ``Document(path)`` reload re-discovers the
        # source under the same name.
        self._stamp_data_source_markers()

        # Drop optional parts that the template carries but the document
        # doesn't actually use. Mirrors Word's behaviour — Word only
        # writes numbering.xml when lists are present, customXml when
        # content-control bindings are present, stylesWithEffects never
        # for new docs, and never ships a thumbnail for library-authored
        # files.
        self._drop_unused_optional_parts(root)

    def _resolve_data_source_bindings(self) -> None:
        """Save-time hook — resolve every SDT data binding against its source.

        Delegates to :func:`docx.data_sources.resolve_bindings_in_document`.
        Hardened against an environment without ``ooxml_customxml`` installed
        — the helper bails out gracefully (returns ``0``).
        """
        try:
            from docx.data_sources import resolve_bindings_in_document

            resolve_bindings_in_document(self)
        except Exception:  # pragma: no cover - defensive: never break save
            pass

    def _stamp_data_source_markers(self) -> None:
        """Stamp the logical-source name into each bound data part's payload.

        Called from :meth:`before_marshal` so the marker is written *just*
        before the package is serialised. The marker is an ignorable
        attribute on the payload root — readers without the marker (Word,
        third-party tools) ignore it; python-docx re-imports it on load via
        :meth:`_recover_data_source_markers`.
        """
        try:
            from docx.data_sources import _name_marker, stamp_name_into_payload
            from docx.opc.constants import RELATIONSHIP_TYPE as _RT
            from docx.parts.custom_xml import CustomXmlPart as _DataPart

            for rel in self.rels.values():
                if rel.is_external or rel.reltype != _RT.CUSTOM_XML:
                    continue
                try:
                    target = rel.target_part
                except ValueError:
                    continue
                if not isinstance(target, _DataPart):
                    continue
                name = _name_marker(target)
                if name is None:
                    continue
                stamp_name_into_payload(target, name)
        except Exception:  # pragma: no cover - defensive
            pass

    def _drop_unused_optional_parts(self, root) -> None:
        """Drop template-default rels whose target parts are unused.

        Mirrors Microsoft Word's "emit the minimum" behaviour for
        library-authored content: parts that the default template carried
        but the document never references are pruned so the resulting
        package contains only what's needed.

        Narrowed policy (W8-A, 2026.05.7): parts are only dropped when
        python-docx itself created them (no ``_loaded_from_package``
        flag). Parts that shipped in the source package are preserved
        verbatim — dropping them silently destroys user data from
        Word-authored files whose structure python-docx can't fully
        reason about at save time.

        Drop candidates:

        - ``RT.STYLES_WITH_EFFECTS`` — a Word 2013-compat duplicate of
          ``styles.xml``. Dropped only when python-docx authored it;
          preserved if the source package shipped it.
        - ``RT.NUMBERING`` — only needed when a paragraph carries
          ``<w:numPr>`` directly OR uses a style whose definition in
          ``styles.xml`` declares ``<w:numPr>`` (e.g. ``List Bullet``,
          ``List Number``). Also preserved if the source package shipped
          it — the numbering part frequently carries abstract numbering
          definitions referenced through indirection python-docx doesn't
          model.
        - ``RT.CUSTOM_XML`` — conservatively preserved whenever the
          source package shipped it. Previously dropped unless a
          ``<w:dataBinding>`` was present, which false-negatived on any
          customXml used for non-binding purposes (Power BI, Office Add-ins,
          Bibliographic sources, etc.).
        """
        uses_numbering = self._document_uses_numbering(root)
        uses_custom_xml = any(
            sdt.find(f".//{qn('w:dataBinding')}") is not None
            for sdt in root.iter(qn("w:sdt"))
        )

        rels_to_drop: list[str] = []
        for rId, rel in list(self.rels.items()):
            if rel.is_external:
                continue
            target = rel.target_part
            shipped = getattr(target, "_loaded_from_package", False)
            if rel.reltype == RT.STYLES_WITH_EFFECTS:
                # Keep if it shipped in the source package; only drop
                # when python-docx authored it (template default).
                if not shipped:
                    rels_to_drop.append(rId)
            elif rel.reltype == RT.NUMBERING:
                # Drop only when the document doesn't use numbering AND
                # python-docx created the numbering part itself. A shipped
                # numbering part is preserved even without a usage match —
                # it may carry abstract-num definitions referenced through
                # style-indirect or list-override chains the heuristic
                # doesn't model.
                if not uses_numbering and not shipped:
                    rels_to_drop.append(rId)
            elif rel.reltype == RT.CUSTOM_XML:
                # Preserve customXml parts that shipped in the source
                # package unconditionally. They cost nothing to keep and
                # dropping them silently destroys user data (Power BI
                # datasets, bibliography sources, content-control backing
                # data — the static heuristic misses all of these).
                # Also preserve freshly-authored bibliography parts: they
                # link to citations implicitly via w:citation markers +
                # matching <b:Tag> values rather than via w:dataBinding,
                # so the uses_custom_xml heuristic doesn't flag them.
                # Also preserve named data-source parts authored via
                # ``Document.bind_data_source`` (#80) — the document may
                # bind a source eagerly before any SDTs reference it,
                # and the bound-source registry must survive a save → reload
                # cycle.
                if (
                    not uses_custom_xml
                    and not shipped
                    and not self._rel_targets_nonempty_bibliography(rel)
                    and not self._rel_targets_named_data_source(rel)
                ):
                    rels_to_drop.append(rId)

        for rId in rels_to_drop:
            self.drop_rel(rId)

    @staticmethod
    def _rel_targets_named_data_source(rel) -> bool:
        """Return ``True`` when ``rel`` targets a ``Document.bind_data_source``
        part — either via the in-memory marker stamped at bind time or via
        the ``lfxbind:name`` attribute we persist into the payload at save
        time.
        """
        try:
            from docx.data_sources import _name_marker, recover_name_from_payload
        except ImportError:  # pragma: no cover - circular guard
            return False
        try:
            target = rel.target_part
        except ValueError:
            return False
        if _name_marker(target) is not None:
            return True
        return recover_name_from_payload(target) is not None

    @staticmethod
    def _rel_targets_nonempty_bibliography(rel) -> bool:
        """Return True when ``rel`` targets a ``<b:Sources>`` part with >=1 child."""
        from lxml import etree

        from docx.oxml.parser import parse_xml

        try:
            target = rel.target_part
        except ValueError:
            return False
        blob = getattr(target, "blob", b"")
        if not blob:
            return False
        # -- hardened parser (resolve_entities=False, no_network=True) stops
        # -- XXE / SSRF via attacker-supplied bibliography customXml parts.
        try:
            root = parse_xml(blob)
        except etree.XMLSyntaxError:
            return False
        if root.tag != qn("b:Sources"):
            return False
        return len(root) > 0

    def _document_uses_numbering(self, root) -> bool:
        """Return ``True`` if the document references numbering at all.

        A document uses numbering if any of the following holds:

        - a paragraph (or its paragraph-mark ``<w:pPr>``) carries a
          direct ``<w:numPr>`` reference, or
        - a paragraph uses a style whose definition (or any style it
          inherits from via ``w:basedOn``) carries ``<w:numPr>`` in
          ``styles.xml``. This catches built-in styles like ``List
          Bullet`` and ``List Number`` plus any user-defined style
          chains rooted in them.

        Failing the styles.xml lookup is treated as "uses numbering" —
        erring on the side of keeping the numbering part is always safe,
        dropping it risks a broken list in the output.
        """
        if root.find(f".//{qn('w:numPr')}") is not None:
            return True

        pStyle_tag = qn("w:pStyle")
        used_styles = {
            pstyle.get(qn("w:val"))
            for pPr in root.iter(qn("w:pPr"))
            for pstyle in pPr.iter(pStyle_tag)
            if pstyle.get(qn("w:val"))
        }
        if not used_styles:
            return False

        try:
            styles_part = self._styles_part
            styles_root = styles_part.element
        except Exception:  # pragma: no cover - defensive
            return True

        style_tag = qn("w:style")
        styleId_attr = qn("w:styleId")
        basedOn_tag = qn("w:basedOn")
        val_attr = qn("w:val")

        # Build a map styleId -> (has_direct_numPr, basedOnId) so we can
        # walk the w:basedOn chain for each used style.
        style_info: dict[str, tuple[bool, str | None]] = {}
        for style in styles_root.iter(style_tag):
            sid = style.get(styleId_attr)
            if not sid:
                continue
            has_numPr = style.find(f".//{qn('w:numPr')}") is not None
            basedOn = style.find(basedOn_tag)
            parent = basedOn.get(val_attr) if basedOn is not None else None
            style_info[sid] = (has_numPr, parent)

        def inherits_numPr(style_id: str) -> bool:
            visited: set[str] = set()
            current: str | None = style_id
            while current and current not in visited:
                visited.add(current)
                info = style_info.get(current)
                if info is None:
                    return False
                has_numPr, parent = info
                if has_numPr:
                    return True
                current = parent
            return False

        return any(inherits_numPr(sid) for sid in used_styles if sid)

    def save(
        self,
        path_or_stream: str | IO[bytes],
        reproducible: bool = False,
        password: str | None = None,
        strict: bool | None = None,
    ):
        """Save this document to `path_or_stream`, which can be either a path to a
        filesystem location (a string) or a file-like object.

        When `reproducible` is True the underlying zip writer emits fixed
        timestamps and sorted member names so repeated saves yield byte-identical
        output (closes upstream#1042 / upstream-PR#810).

        When `password` is provided the saved ``.docx`` is password-protected
        using ECMA-376 Agile Encryption. Encryption requires the optional
        ``python-ooxml-crypto`` dependency.

        `strict` is forwarded to :meth:`docx.opc.package.OpcPackage.save`
        for ECMA-376 conformance-class handling; see
        :meth:`docx.document.Document.save`.

        .. versionadded:: 2026.05.0
           The `reproducible` parameter.
        .. versionadded:: 2026.05.10
           The `password` parameter.
        .. versionadded:: 2026.05.11
           The `strict` parameter.
        """
        self.package.save(
            path_or_stream, reproducible=reproducible, password=password,
            strict=strict,
        )

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the settings in the settings part of
        this document."""
        return self._settings_part.settings

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in the styles part of this
        document."""
        return self._styles_part.styles

    @property
    def theme(self) -> Theme | None:
        """A |Theme| proxy for this document, or |None| when no theme part is related.

        The theme part is managed by Word; python-docx does not create one on
        demand, so this returns |None| when the document has no existing
        ``theme`` relationship.
        """
        theme_part = self._theme_part
        if theme_part is None:
            return None
        return theme_part.theme

    @property
    def _theme_part(self) -> ThemePart | None:
        """The |ThemePart| related to this document, or |None| if not present.

        A default theme part is not created on demand; the part is exposed
        read-only and only available when the source document already contains
        one.
        """
        try:
            return cast(ThemePart, self.part_related_by(RT.THEME))
        except KeyError:
            return None

    @property
    def web_settings(self) -> WebSettings | None:
        """A |WebSettings| proxy for this document, or |None| when no web-settings part exists.

        The web-settings part is managed by Word; python-docx does not create one
        on demand, so this returns |None| when the document has no existing
        ``webSettings`` relationship.
        """
        web_settings_part = self._web_settings_part
        if web_settings_part is None:
            return None
        return web_settings_part.web_settings

    @property
    def _web_settings_part(self) -> WebSettingsPart | None:
        """The |WebSettingsPart| related to this document, or |None| if not present.

        Unlike the settings part, a default web-settings part is not created on
        demand; the part is exposed read-only and only available when the source
        document already contains one.
        """
        try:
            return cast(WebSettingsPart, self.part_related_by(RT.WEB_SETTINGS))
        except KeyError:
            return None

    @property
    def _comments_part(self) -> CommentsPart:
        """A |CommentsPart| object providing access to the comments added to this document.

        Creates a default comments part if one is not present.
        """
        try:
            return cast(CommentsPart, self.part_related_by(RT.COMMENTS))
        except KeyError:
            assert self.package is not None
            comments_part = CommentsPart.default(self.package)
            self.relate_to(comments_part, RT.COMMENTS)
            return comments_part

    @property
    def _settings_part(self) -> SettingsPart:
        """A |SettingsPart| object providing access to the document-level settings for
        this document.

        Creates a default settings part if one is not present.
        """
        try:
            return cast(SettingsPart, self.part_related_by(RT.SETTINGS))
        except KeyError:
            settings_part = SettingsPart.default(self.package)
            self.relate_to(settings_part, RT.SETTINGS)
            return settings_part

    @property
    def _styles_part(self) -> StylesPart:
        """Instance of |StylesPart| for this document.

        Creates an empty styles part if one is not present.
        """
        try:
            return cast(StylesPart, self.part_related_by(RT.STYLES))
        except KeyError:
            package = self.package
            assert package is not None
            styles_part = StylesPart.default(package)
            self.relate_to(styles_part, RT.STYLES)
            return styles_part


# Tags Word mirrors from a run's <w:rPr> onto the paragraph mark's
# <w:pPr>/<w:rPr>. Toggle properties (bold/italic/underline/strike,
# caps variants) and character-shape properties (size, color, font
# name). Deliberately excludes lang, spacing, and the border/shd
# family because Word doesn't mirror those onto paragraph marks.
_MIRROR_RUN_PROP_TAGS = frozenset(
    qn(t)
    for t in (
        "w:b",
        "w:bCs",
        "w:i",
        "w:iCs",
        "w:u",
        "w:strike",
        "w:dstrike",
        "w:caps",
        "w:smallCaps",
        "w:color",
        "w:sz",
        "w:szCs",
        "w:rFonts",
        "w:vertAlign",
    )
)


def _mirror_run_formatting_to_paragraph_mark(root) -> None:
    """Copy the first run's rPr formatting onto each paragraph's pPr/rPr.

    Word emits the run formatting of (roughly) the last run in each
    paragraph onto the paragraph mark via ``<w:pPr><w:rPr>`` so that
    typing past the paragraph continues in the same formatting. We
    mirror the FIRST run's formatting because python-docx's idiomatic
    one-run-per-paragraph usage makes first == last in the common case;
    multi-run paragraphs will get the first run's formatting on the
    mark, which matches the Word "select all, format" pattern.

    Only mirrors for paragraphs that have exactly one direct <w:r>
    child whose <w:rPr> carries any of the whitelisted tags. Avoids
    over-writing an explicit pPr/rPr on the paragraph.
    """
    from copy import deepcopy

    from docx.oxml.parser import OxmlElement

    w_r = qn("w:r")
    w_pPr = qn("w:pPr")
    w_rPr = qn("w:rPr")

    for p in root.iter(qn("w:p")):
        # Find the first direct <w:r> child, ignoring hyperlinks and
        # other wrapped content where mirroring would be surprising.
        direct_runs = [child for child in p if child.tag == w_r]
        if len(direct_runs) != 1:
            continue
        source_rPr = direct_runs[0].find(w_rPr)
        if source_rPr is None:
            continue

        mirror_children = [
            child for child in source_rPr if child.tag in _MIRROR_RUN_PROP_TAGS
        ]
        if not mirror_children:
            continue

        pPr = p.find(w_pPr)
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)

        target_rPr = pPr.find(w_rPr)
        if target_rPr is None:
            target_rPr = OxmlElement("w:rPr")
            pPr.append(target_rPr)

        existing = {child.tag for child in target_rPr}
        for src in mirror_children:
            if src.tag in existing:
                continue
            target_rPr.append(deepcopy(src))


class _RandomMinter:
    """Minter that generates random 8-hex-digit tokens per call.

    Normal save mode — every call gives a fresh token, so re-saving
    an unchanged document produces different rsid/paraId values on
    each run.
    """

    def rsid_root(self) -> str:
        # Word's rsids have an "00"-prefix in practice, so match the shape.
        return "00" + secrets.token_hex(3).upper()

    def paraId(self, _paragraph) -> str:
        return secrets.token_hex(4).upper()

    def textId(self, _paragraph) -> str:
        return secrets.token_hex(4).upper()


class _DeterministicMinter:
    """Minter that derives every identifier from stable paragraph content.

    Reproducible save mode — two saves of the same document produce
    byte-identical output. Identifiers are 8-hex-digit SHA-1 prefixes
    of each paragraph's text plus a role tag.
    """

    def rsid_root(self) -> str:
        return "00000001"

    def paraId(self, paragraph) -> str:
        return self._hash8(paragraph, "paraId")

    def textId(self, paragraph) -> str:
        return self._hash8(paragraph, "textId")

    @staticmethod
    def _hash8(paragraph, role: str) -> str:
        import hashlib

        # Use the paragraph's serialised text content plus the role as
        # the stability seed. Two paragraphs with identical text will
        # share an id, which is an acceptable collision for a
        # reproducible-save corner case.
        text = "".join(paragraph.itertext())
        digest = hashlib.sha1(f"{role}:{text}".encode("utf-8")).hexdigest()
        return digest[:8].upper()
