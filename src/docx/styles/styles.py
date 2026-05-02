"""Styles object, container for all objects in the styles part."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import TYPE_CHECKING, Iterable
from warnings import warn

from docx.enum.style import WD_BUILTIN_STYLE, WD_STYLE_TYPE
from docx.oxml.parser import parse_xml
from docx.oxml.styles import CT_Styles
from docx.shared import ElementProxy
from docx.styles import BabelFish
from docx.styles.latent import LatentStyles
from docx.styles.style import BaseStyle, StyleFactory
from docx.text.font import Font

if TYPE_CHECKING:
    from docx.oxml.styles import CT_Style


class Styles(ElementProxy):
    """Provides access to the styles defined in a document.

    Accessed using the :attr:`.Document.styles` property. Supports ``len()``, iteration,
    and dictionary-style access by style name.
    """

    def __init__(self, styles: CT_Styles):
        super().__init__(styles)
        self._element = styles

    def __contains__(self, name):
        """Enables `in` operator on style name."""
        internal_name = BabelFish.ui2internal(name)
        if any(style.name_val == internal_name for style in self._element.style_lst):
            return True
        # -- fall back to case-insensitive match (LibreOffice saves "Heading 1"
        # -- as "heading 1" in the on-disk name; upstream #494) --
        lowered = internal_name.lower()
        return any(
            (style.name_val or "").lower() == lowered
            for style in self._element.style_lst
        )

    def __getitem__(self, key):
        """Enables dictionary-style access by UI name.

        `key` may also be a :class:`docx.enum.style.WD_BUILTIN_STYLE` member,
        in which case the member's canonical UI name is used for lookup.

        When an exact-match lookup fails, a case-insensitive name match is
        attempted so that documents saved by LibreOffice (which lower-cases
        some built-in style names — e.g. ``"heading 1"`` for ``"Heading 1"``)
        can still be looked up with their conventional UI name
        (upstream #494, #420, PR#239).

        Lookup by style id is deprecated, triggers a warning, and will be
        removed in a near-future release.
        """
        # -- translate an enum member (e.g. `WD_STYLE.BODY_TEXT`) into its UI name --
        if isinstance(key, WD_BUILTIN_STYLE):
            key = BabelFish.enum2ui(key)

        style_elm = self._element.get_by_name(BabelFish.ui2internal(key))
        if style_elm is not None:
            return StyleFactory(style_elm)

        # -- case-insensitive name fallback (covers LibreOffice-cased names
        # -- like "heading 1" and WD_BUILTIN_STYLE members whose UI name
        # -- casing differs from the on-disk `w:name/@w:val`) --
        if isinstance(key, str):
            lowered = key.lower()
            ui_lowered = BabelFish.ui2internal(key).lower()
            for style in self._element.style_lst:
                sv = (style.name_val or "").lower()
                if sv == lowered or sv == ui_lowered:
                    return StyleFactory(style)

        style_elm = self._element.get_by_id(key)
        if style_elm is not None:
            msg = "style lookup by style_id is deprecated. Use style name as key instead."
            warn(msg, UserWarning, stacklevel=2)
            return StyleFactory(style_elm)

        raise KeyError("no style with name '%s'" % key)

    def __iter__(self):
        return (StyleFactory(style) for style in self._element.style_lst)

    def __len__(self):
        return len(self._element.style_lst)

    def add_style(self, name, style_type, builtin=False):
        """Return a newly added style object of `style_type` and identified by `name`.

        A builtin style can be defined by passing True for the optional `builtin`
        argument.
        """
        if name in self:
            raise ValueError("document already contains style '%s'" % name)
        style_name = BabelFish.ui2internal(name)
        style = self._element.add_style_of_type(style_name, style_type, builtin)
        return StyleFactory(style)

    def default(self, style_type: WD_STYLE_TYPE):
        """Return the default style for `style_type` or |None| if no default is defined
        for that type (not common)."""
        style = self._element.default_for(style_type)
        if style is None:
            return None
        return StyleFactory(style)

    def get_by_id(self, style_id: str | None, style_type: WD_STYLE_TYPE):
        """Return the style of `style_type` matching `style_id`.

        Returns the default for `style_type` if `style_id` is not found or is |None|, or
        if the style having `style_id` is not of `style_type`.
        """
        if style_id is None:
            return self.default(style_type)
        return self._get_by_id(style_id, style_type)

    def get_style_id(self, style_or_name, style_type):
        """Return the id of the style corresponding to `style_or_name`, or |None| if
        `style_or_name` is |None|.

        If `style_or_name` is not a style object, the style is looked up using
        `style_or_name` as a style name, raising |ValueError| if no style with that name
        is defined. Raises |ValueError| if the target style is not of `style_type`.
        """
        if style_or_name is None:
            return None
        elif isinstance(style_or_name, BaseStyle):
            return self._get_style_id_from_style(style_or_name, style_type)
        else:
            return self._get_style_id_from_name(style_or_name, style_type)

    @property
    def document_default_font(self) -> Font:
        """A |Font| proxy for ``w:styles/w:docDefaults/w:rPrDefault/w:rPr``.

        Reading and writing properties on the returned font modifies the
        document-wide default character formatting that applies to every run
        that does not otherwise override a property via its style chain.

        The ``<w:docDefaults>``/``<w:rPrDefault>`` ancestors are created on
        demand the first time this property is accessed. Closes upstream#383.

        .. versionadded:: 2026.05.0
        """
        docDefaults = self._element.get_or_add_docDefaults()
        rPrDefault = docDefaults.get_or_add_rPrDefault()
        return Font(rPrDefault)

    def import_from(
        self,
        source: "Styles | object",
        names: Iterable[str] | None = None,
    ) -> list[BaseStyle]:
        """Deep-copy styles from `source` into this document's styles.

        `source` can be a |Styles| instance or any object exposing a ``styles``
        attribute that returns one (e.g. a |Document|). When `names` is |None|,
        every style in `source` is considered; otherwise only styles whose UI
        name matches an entry in `names` are considered (other entries are
        silently ignored). Any style already present in this document (matched
        by ``w:styleId``) is skipped.

        Styles referenced by imported styles via ``w:basedOn``, ``w:link`` or
        ``w:next`` are imported transitively so the dependency chain resolves
        after the call. Returns the list of newly-imported |BaseStyle| proxies
        in import order. Closes upstream#1375, #1083, #508, #701, #197.

        .. versionadded:: 2026.05.0
        """
        source_styles = source.styles if hasattr(source, "styles") else source
        assert isinstance(source_styles, Styles)

        if names is None:
            candidate_elms = list(source_styles._element.style_lst)
        else:
            wanted = {BabelFish.ui2internal(n) for n in names}
            candidate_elms = [
                s for s in source_styles._element.style_lst if s.name_val in wanted
            ]

        imported: list[BaseStyle] = []
        visited: set[str] = set()

        for style_elm in candidate_elms:
            imported.extend(self._import_style_with_deps(style_elm, source_styles, visited))
        return imported

    def import_style(self, style: "BaseStyle | CT_Style") -> BaseStyle:
        """Deep-copy a single style (and its transitive dependencies) into this document.

        `style` may be a |BaseStyle| proxy or a raw ``<w:style>`` element. If a
        style with the same ``w:styleId`` already exists in this document the
        existing style is returned unchanged (no overwrite). Otherwise the
        style is imported together with any styles it references via
        ``w:basedOn``, ``w:link`` and ``w:next``.

        .. versionadded:: 2026.05.0
        """
        style_elm = style._element if isinstance(style, BaseStyle) else style
        source_elm = style_elm.getparent()
        # -- source_elm is a CT_Styles --
        source_styles = Styles(source_elm)
        imported = self._import_style_with_deps(style_elm, source_styles, set())
        # -- always return a proxy for the target style (imported or pre-existing) --
        target = self._element.get_by_id(style_elm.styleId)
        if target is None:
            # -- couldn't import; shouldn't normally happen --
            return imported[-1] if imported else StyleFactory(style_elm)
        return StyleFactory(target)

    def import_builtin(self, name: str) -> BaseStyle:
        """Materialise the builtin style `name` into this document.

        Word ships with a predefined "latent" style set (e.g. ``"List Bullet"``,
        ``"FollowedHyperlink"``) whose visible definitions live in the default
        ``styles.xml`` of a freshly-created document. This method deep-copies
        the style definition from the bundled default template into this
        document so it (and any styles it depends on via ``basedOn``/``link``/
        ``next``) becomes a first-class style the caller can apply to
        paragraphs and runs. If the style is already present the existing
        style is returned unchanged. Closes upstream#486.

        Raises :class:`KeyError` if `name` is not present in the bundled
        defaults.

        .. versionadded:: 2026.05.0
        """
        default_styles = _default_template_styles()
        source_elm = default_styles._element.get_by_name(BabelFish.ui2internal(name))
        if source_elm is None:
            raise KeyError(name)
        return self.import_style(source_elm)

    @property
    def latent_styles(self):
        """A |LatentStyles| object providing access to the default behaviors for latent
        styles and the collection of |_LatentStyle| objects that define overrides of
        those defaults for a particular named latent style."""
        return LatentStyles(self._element.get_or_add_latentStyles())

    def _get_by_id(self, style_id: str | None, style_type: WD_STYLE_TYPE):
        """Return the style of `style_type` matching `style_id`.

        Returns the default for `style_type` if `style_id` is not found or if the style
        having `style_id` is not of `style_type`.
        """
        style = self._element.get_by_id(style_id) if style_id else None
        if style is None or style.type != style_type:
            return self.default(style_type)
        return StyleFactory(style)

    def _get_style_id_from_name(self, style_name: str, style_type: WD_STYLE_TYPE) -> str | None:
        """Return the id of the style of `style_type` corresponding to `style_name`.

        Returns |None| if that style is the default style for `style_type`. Raises
        |ValueError| if the named style is not found in the document or does not match
        `style_type`.
        """
        return self._get_style_id_from_style(self[style_name], style_type)

    def _get_style_id_from_style(self, style: BaseStyle, style_type: WD_STYLE_TYPE) -> str | None:
        """Id of `style`, or |None| if it is the default style of `style_type`.

        Raises |ValueError| if style is not of `style_type`.
        """
        if style.type != style_type:
            raise ValueError("assigned style is type %s, need type %s" % (style.type, style_type))
        if style == self.default(style_type):
            return None
        return style.style_id

    def _import_style_with_deps(
        self,
        style_elm: "CT_Style",
        source_styles: "Styles",
        visited: set[str],
    ) -> list[BaseStyle]:
        """Deep-copy `style_elm` and its ``basedOn``/``link``/``next`` dependencies.

        Returns the list of styles newly added to this document, in the order
        they were appended. `visited` tracks styleIds already processed across
        the current import to break cycles and avoid redundant work.
        """
        imported: list[BaseStyle] = []
        style_id = style_elm.styleId
        if not style_id or style_id in visited:
            return imported
        visited.add(style_id)

        # -- skip if already present in this document --
        if self._element.get_by_id(style_id) is not None:
            return imported

        # -- resolve dependencies first so references remain valid --
        for ref in (style_elm.basedOn_val, style_elm.link_val, style_elm.next_val):
            if not ref:
                continue
            dep_elm = source_styles._element.get_by_id(ref)
            if dep_elm is None:
                continue
            imported.extend(
                self._import_style_with_deps(dep_elm, source_styles, visited)
            )

        # -- deep-copy the style element into our styles tree --
        new_elm = deepcopy(style_elm)
        self._element.append(new_elm)
        imported.append(StyleFactory(new_elm))
        return imported


def _default_template_styles() -> Styles:
    """Return a |Styles| view over the bundled default ``styles.xml``.

    The templates ship with python-docx under ``src/docx/templates/``. The
    richer ``default-docx-template/word/styles.xml`` defines every builtin
    style Word materialises from its latent-style set (``"List Bullet"``,
    ``"Hyperlink"``, ``"FollowedHyperlink"``, etc.) and is therefore the
    preferred source. The smaller ``default-styles.xml`` is consulted as a
    fallback. The returned |Styles| instance is freshly parsed on each call
    so callers can mutate it without contaminating shared state.
    """
    base = Path(__file__).resolve().parent.parent / "templates"
    rich = base / "default-docx-template" / "word" / "styles.xml"
    template = rich if rich.is_file() else base / "default-styles.xml"
    element = parse_xml(template.read_bytes())
    return Styles(element)


def _builtin_style_ui_name(member: WD_BUILTIN_STYLE) -> str:
    """Return the canonical UI style name for a `WD_BUILTIN_STYLE` member.

    Thin shim kept for backward compatibility; delegates to `BabelFish.enum2ui`.
    """
    return BabelFish.enum2ui(member)
