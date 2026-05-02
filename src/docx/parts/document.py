"""|DocumentPart| and closely related objects."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, cast

from docx.document import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
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
    def comments(self) -> Comments:
        """|Comments| object providing access to the comments added to this document."""
        return self._comments_part.comments

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

        The font-table part is managed by Word — python-docx does not create one on
        demand, so this returns |None| when the document has no existing ``fontTable``
        relationship.
        """
        font_table_part = self._font_table_part
        if font_table_part is None:
            return None
        return font_table_part.font_table

    @property
    def _font_table_part(self) -> FontTablePart | None:
        """The |FontTablePart| related to this document, or |None| if not present.

        Unlike the comments or footnotes parts, a default font-table part is not
        created on demand; the part is read-only and only available when the source
        document already contains one.
        """
        try:
            return cast(FontTablePart, self.part_related_by(RT.FONT_TABLE))
        except KeyError:
            return None

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

        The glossary-document part is managed by Word; python-docx does not
        create one on demand, so this returns |None| when the document has
        no existing ``glossaryDocument`` relationship.
        """
        glossary_part = self._glossary_part
        if glossary_part is None:
            return None
        return glossary_part.glossary

    @property
    def _glossary_part(self) -> GlossaryPart | None:
        """The |GlossaryPart| related to this document, or |None| if not present.

        A default glossary part is not created on demand; the part is
        exposed read-only and only available when the source document
        already contains one.
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
        """A |CustomPropertiesPart| for this document.

        Creates a default (empty) custom properties part if one is not present.
        """
        try:
            return cast(CustomPropertiesPart, self.part_related_by(RT.CUSTOM_PROPERTIES))
        except KeyError:
            assert self.package is not None
            custom_properties_part = CustomPropertiesPart.default(self.package)
            self.relate_to(custom_properties_part, RT.CUSTOM_PROPERTIES)
            return custom_properties_part

    @property
    def extended_properties(self) -> ExtendedProperties:
        """An |ExtendedProperties| proxy for the document's ``app.xml`` part.

        The extended-properties part is related to the package (not the
        main-document part), following the convention used for core properties.
        A default (empty) part is created on demand when none is present.

        .. versionadded:: 1.3.0.dev0
        """
        return self._extended_properties_part.extended_properties

    @property
    def _extended_properties_part(self) -> ExtendedPropertiesPart:
        """Return the package-scoped |ExtendedPropertiesPart| for this document.

        .. versionadded:: 1.3.0.dev0
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

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this document to `path_or_stream`, which can be either a path to a
        filesystem location (a string) or a file-like object."""
        self.package.save(path_or_stream)

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
