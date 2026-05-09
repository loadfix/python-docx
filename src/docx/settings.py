"""Settings object, providing access to document-level settings."""

from __future__ import annotations

import base64
import hashlib
import os
import secrets
import warnings
from typing import TYPE_CHECKING, Iterator, cast

from docx.enum.text import (
    WD_MAIL_MERGE_DATA_TYPE,
    WD_MAIL_MERGE_DESTINATION,
    WD_MAIL_MERGE_TYPE,
    WD_ODSO_TYPE,
    WD_PROTECTION,
    WD_VIEW,
)
from docx.oxml.ns import qn as _qn
from docx.shared import ElementProxy


def qn_w(local: str) -> str:
    """Return the Clark-notation form of a ``w:{local}`` QName.

    Tiny local alias — keeps the OdsoSettings proxy reads and writes
    compact without littering the module with full ``qn("w:...")`` calls.
    """
    return _qn(f"w:{local}")

if TYPE_CHECKING:
    import docx.types as t
    from docx.endnotes import EndnoteProperties
    from docx.footnotes import FootnoteProperties
    from docx.oxml.mail_merge import CT_DataSourceObject, CT_Odso
    from docx.oxml.settings import (
        CT_Compat,
        CT_DocVars,
        CT_MailMerge,
        CT_Settings,
        CT_WriteProtection,
    )
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.shared import Length


# -- Curated list of well-known direct-child compat-flag element names. Used by
# -- `CompatFlags.names()` for discoverability. Ordering follows roughly the
# -- historical / schema order of common Word compatibility flags.
_KNOWN_COMPAT_FLAG_NAMES: tuple[str, ...] = (
    "useSingleBorderforContiguousCells",
    "wpJustification",
    "noTabHangInd",
    "noLeading",
    "spaceForUL",
    "balanceSingleByteDoubleByteWidth",
    "noExtraLineSpacing",
    "doNotLeaveBackslashAlone",
    "ulTrailSpace",
    "doNotExpandShiftReturn",
    "spacingInWholePoints",
    "lineWrapLikeWord6",
    "printBodyTextBeforeHeader",
    "printColBlack",
    "wpSpaceWidth",
    "showBreaksInFrames",
    "subFontBySize",
    "suppressBottomSpacing",
    "suppressTopSpacing",
    "suppressSpacingAtTopOfPage",
    "suppressTopSpacingWP",
    "suppressSpBfAfterPgBrk",
    "swapBordersFacingPages",
    "convMailMergeEsc",
    "truncateFontHeightsLikeWP6",
    "mwSmallCaps",
    "usePrinterMetrics",
    "doNotSuppressParagraphBorders",
    "wrapTrailSpaces",
    "footnoteLayoutLikeWW8",
    "shapeLayoutLikeWW8",
    "alignTablesRowByRow",
    "forgetLastTabAlignment",
    "adjustLineHeightInTable",
    "autoSpaceLikeWord95",
    "noSpaceRaiseLower",
    "doNotUseHTMLParagraphAutoSpacing",
    "layoutRawTableWidth",
    "layoutTableRowsApart",
    "useWord97LineBreakRules",
    "doNotBreakWrappedTables",
    "doNotSnapToGridInCell",
    "selectFldWithFirstOrLastChar",
    "applyBreakingRules",
    "doNotWrapTextWithPunct",
    "doNotUseEastAsianBreakRules",
    "useWord2002TableStyleRules",
    "growAutofit",
    "useFELayout",
    "useNormalStyleForList",
    "doNotUseIndentAsNumberingTabStop",
    "useAltKinsokuLineBreakRules",
    "allowSpaceOfSameStyleInTable",
    "doNotSuppressIndentation",
    "doNotAutofitConstrainedTables",
    "autofitToFirstFixedWidthCell",
    "underlineTabInNumList",
    "displayHangulFixedWidth",
    "splitPgBreakAndParaMark",
    "doNotVertAlignCellWithSp",
    "doNotBreakConstrainedForcedTable",
    "doNotVertAlignInTxbx",
    "useAnsiKerningPairs",
    "cachedColBalance",
)


class Settings(ElementProxy):
    """Provides access to document-level settings for a document.

    Accessed using the :attr:`.Document.settings` property.
    """

    def __init__(self, element: BaseOxmlElement, parent: t.ProvidesXmlPart | None = None):
        super().__init__(element, parent)
        self._settings = cast("CT_Settings", element)

    @property
    def compatibility_mode(self) -> int | None:
        """The target Word compatibility-mode version (e.g. 15 for Word 2013+).

        Read/write. None when no compatibility mode is specified.

        .. versionadded:: 2026.05.0
        """
        return self._settings.compatibilityMode

    @compatibility_mode.setter
    def compatibility_mode(self, value: int | None):
        self._settings.compatibilityMode = value

    @property
    def compat_settings(self) -> CompatSettings:
        """Dict-like access to ``w:compat/w:compatSetting`` entries.

        Keys are the ``@w:name`` strings; values are the ``@w:val`` strings. The
        returned object is a live view -- assignments and deletions are reflected in
        the underlying XML immediately and create/remove the ``w:compat`` element as
        needed.

        .. versionadded:: 2026.05.0
        """
        return CompatSettings(self._settings)

    @property
    def compat_flags(self) -> CompatFlags:
        """Dict-like access to direct-child flag elements under ``w:compat``.

        Each known Word compatibility flag (e.g. ``growAutofit``,
        ``doNotBreakWrappedTables``, ...) is represented as a direct child of
        ``w:compat`` whose mere presence turns the behaviour on. Keys are the flag
        names without the ``w:`` prefix; values are booleans. Unknown keys are also
        accepted and written/read using the ``w:`` namespace.

        .. versionadded:: 2026.05.0
        """
        return CompatFlags(self._settings)

    @property
    def default_tab_stop(self) -> Length | None:
        """The default tab-stop interval for the document as a |Length| value.

        Read/write. Assign a |Length| value (e.g. ``Twips(720)``) or |None| to remove.

        .. versionadded:: 2026.05.0
        """
        return self._settings.defaultTabStop_val

    @default_tab_stop.setter
    def default_tab_stop(self, value: int | Length | None):
        self._settings.defaultTabStop_val = value

    @property
    def document_protection(self) -> DocumentProtection:
        """Access to document protection settings.

        Provides read/write access to the ``w:documentProtection`` element and its
        attributes. Use :meth:`enable_protection` and :meth:`disable_protection` for
        common high-level operations.

        .. versionadded:: 2026.05.0
        """
        return DocumentProtection(self._settings)

    def enable_protection(
        self,
        mode: WD_PROTECTION = WD_PROTECTION.READ_ONLY,
        enforce: bool = True,
        password: str | None = None,
    ) -> DocumentProtection:
        """Enable document protection with the given `mode`.

        Equivalent to populating ``w:documentProtection`` with ``@w:edit=<mode>``
        and ``@w:enforcement=enforce``. If `password` is given, compute Word's
        password hash (SHA-1 based) with a fresh 16-byte random salt and populate
        the ``@w:hash``/``@w:salt`` and associated ``@w:crypt*`` attributes; if
        `password` is |None|, no password is set.

        Returns the :class:`DocumentProtection` proxy so callers can chain further
        attribute assignments (e.g. ``settings.enable_protection(...).formatting_locked = True``).

        .. versionadded:: 2026.05.0
        """
        protection = self.document_protection
        protection.mode = mode
        protection.enforce = bool(enforce)
        if password is None:
            # -- clear any stale hash/salt/crypto metadata --
            protection.password_hash = None
            protection.password_salt = None
            protection.crypto_provider_type = None
            protection.crypto_algorithm_class = None
            protection.crypto_algorithm_type = None
            protection.crypto_algorithm_sid = None
            protection.spin_count = None
        else:
            protection.set_password(password)
        return protection

    def disable_protection(self) -> None:
        """Remove the ``w:documentProtection`` element entirely.

        .. versionadded:: 2026.05.0
        """
        self._settings._remove_documentProtection()  # pyright: ignore[reportPrivateUsage]

    @property
    def write_protection(self) -> WriteProtection:
        """Access to the document's write-protection (password-to-modify) settings.

        Provides read/write access to the ``w:writeProtection`` element. The
        ``w:writeProtection`` element is distinct from ``w:documentProtection``
        — it governs *save* access to the current file rather than editing
        modes within the document. Use :meth:`enable_write_protection` and
        :meth:`disable_write_protection` for common high-level operations.

        .. versionadded:: 2026.05.10
        """
        return WriteProtection(self._settings)

    def enable_write_protection(
        self,
        recommended: bool = False,
        password: str | None = None,
    ) -> WriteProtection:
        """Enable write-protection on the document.

        When `recommended` is |True|, Word displays the "Read-only
        recommended" banner on open. When `password` is supplied, Word's
        SHA-1 password hash (with a fresh 16-byte random salt and 100,000
        iterations) is written into the ``@w:hash``/``@w:salt`` attributes
        along with the matching ``@w:crypt*`` metadata, and Word will
        prompt for the password before allowing save.

        Returns the :class:`WriteProtection` proxy for further tweaks.

        .. versionadded:: 2026.05.10
        """
        wp = self.write_protection
        wp.recommended_read_only = bool(recommended)
        if password is None:
            wp.password_hash = None
            wp.password_salt = None
            wp.crypto_provider_type = None
            wp.crypto_algorithm_class = None
            wp.crypto_algorithm_type = None
            wp.crypto_algorithm_sid = None
            wp.spin_count = None
        else:
            wp.set_password(password)
        return wp

    def disable_write_protection(self) -> None:
        """Remove the ``w:writeProtection`` element entirely.

        .. versionadded:: 2026.05.10
        """
        self._settings._remove_writeProtection()  # pyright: ignore[reportPrivateUsage]

    @property
    def mail_merge(self) -> MailMerge | None:
        """Access the mail-merge configuration or |None| when not configured.

        Returns a |MailMerge| proxy providing read/write access to the
        ``w:mailMerge`` element's fields (main document type, destination,
        data source, query, etc.). Returns |None| when the document has no
        mail-merge block.

        .. versionadded:: 2026.05.0
        """
        mm = self._settings.mailMerge
        if mm is None:
            return None
        return MailMerge(mm)

    def enable_mail_merge(
        self,
        main_document_type: WD_MAIL_MERGE_TYPE = WD_MAIL_MERGE_TYPE.FORM_LETTERS,
        destination: WD_MAIL_MERGE_DESTINATION | None = None,
        data_type: WD_MAIL_MERGE_DATA_TYPE | None = None,
        connect_string: str | None = None,
        query: str | None = None,
        mail_subject: str | None = None,
        address_field_name: str | None = None,
    ) -> MailMerge:
        """Create or replace the ``w:mailMerge`` element and return a proxy.

        `main_document_type` selects the merge style (form letters by default).
        Any other argument left as |None| is omitted from the XML.

        .. versionadded:: 2026.05.0
        """
        mm = self._settings.get_or_add_mailMerge()
        proxy = MailMerge(mm)
        proxy.main_document_type = main_document_type
        if destination is not None:
            proxy.destination = destination
        if data_type is not None:
            proxy.data_type = data_type
        if connect_string is not None:
            proxy.connect_string = connect_string
        if query is not None:
            proxy.query = query
        if mail_subject is not None:
            proxy.mail_subject = mail_subject
        if address_field_name is not None:
            proxy.address_field_name = address_field_name
        return proxy

    def disable_mail_merge(self) -> None:
        """Remove the ``w:mailMerge`` element entirely.

        .. versionadded:: 2026.05.0
        """
        self._settings._remove_mailMerge()  # pyright: ignore[reportPrivateUsage]

    @property
    def even_and_odd_headers(self) -> bool:
        """True if this document has distinct odd and even page headers and footers.

        Read/write.

        .. versionadded:: 2026.05.0
        """
        return self._settings.evenAndOddHeaders_val

    @even_and_odd_headers.setter
    def even_and_odd_headers(self, value: bool):
        self._settings.evenAndOddHeaders_val = value

    @property
    def odd_and_even_pages_header_footer(self) -> bool:
        """True if this document has distinct odd and even page headers and footers.

        Read/write. Deprecated: use `even_and_odd_headers` instead.
        """
        warnings.warn(
            "odd_and_even_pages_header_footer is deprecated, use even_and_odd_headers instead",
            DeprecationWarning,
            stacklevel=2,
        )
        return self.even_and_odd_headers

    @odd_and_even_pages_header_footer.setter
    def odd_and_even_pages_header_footer(self, value: bool):
        warnings.warn(
            "odd_and_even_pages_header_footer is deprecated, use even_and_odd_headers instead",
            DeprecationWarning,
            stacklevel=2,
        )
        self.even_and_odd_headers = value

    @property
    def view(self) -> WD_VIEW | None:
        """The preferred view mode for this document as a |WD_VIEW| member.

        Read/write. |None| when no ``w:view`` child is present in settings.
        Assign |None| to remove the element. Recognized OOXML values are
        ``none``, ``print``, ``outline``, ``masterPages``, ``normal``,
        ``web``, and ``reading``.

        .. versionadded:: 2026.05.0
        """
        val = self._settings.view_val
        if val is None:
            return None
        return WD_VIEW.from_xml(val)

    @view.setter
    def view(self, value: WD_VIEW | None):
        if value is None:
            self._settings.view_val = None
            return
        self._settings.view_val = WD_VIEW.to_xml(value)

    @property
    def track_revisions(self) -> bool:
        """True when revision tracking is enabled for this document.

        Read/write. Maps to ``w:trackRevisions`` — the session-level flag
        that records every edit as a tracked change. Distinct from
        ``w:trackChanges`` (a per-run marker on existing changes); this is
        the global toggle for *new* edits.

        .. versionadded:: 2026.05.0
        """
        return self._settings.trackRevisions_val

    @track_revisions.setter
    def track_revisions(self, value: bool):
        self._settings.trackRevisions_val = value

    @property
    def remove_personal_information(self) -> bool:
        """True when Word should strip author/reviewer personal info on save.

        Backed by the ``w:removePersonalInformation`` element. Read/write.
        When |True|, Word removes user names from comments, revision
        tracking, and properties when the document is saved.

        .. versionadded:: 2026.05.10
        """
        return self._settings.removePersonalInformation_val

    @remove_personal_information.setter
    def remove_personal_information(self, value: bool) -> None:
        self._settings.removePersonalInformation_val = bool(value)

    @property
    def remove_date_and_time(self) -> bool:
        """True when Word should strip revision/comment timestamps on save.

        Backed by the ``w:removeDateAndTime`` element. Read/write. Pairs
        with :attr:`remove_personal_information` to produce an anonymised
        document.

        .. versionadded:: 2026.05.10
        """
        return self._settings.removeDateAndTime_val

    @remove_date_and_time.setter
    def remove_date_and_time(self, value: bool) -> None:
        self._settings.removeDateAndTime_val = bool(value)

    @property
    def characters_with_numbers_width(self) -> bool:
        """True when ``w:charactersWithNumbersWidth`` toggle is set.

        Some East-Asian layouts use this flag to indicate that each CJK
        character should occupy the width of a digit rather than its
        native glyph width. Read/write.

        .. versionadded:: 2026.05.10
        """
        return self._settings.charactersWithNumbersWidth_val

    @characters_with_numbers_width.setter
    def characters_with_numbers_width(self, value: bool) -> None:
        self._settings.charactersWithNumbersWidth_val = bool(value)

    @property
    def update_fields_on_open(self) -> bool:
        """True when Word should refresh all fields when the document is opened.

        Maps to the ``w:updateFields`` child of ``w:settings``. When |True|,
        Word evaluates every field (TOC, PAGEREF, DATE, ...) on open rather
        than displaying the cached ``result`` text. Read/write.

        Closes upstream#1403.

        .. versionadded:: 2026.05.0
        """
        return self._settings.updateFields_val

    @update_fields_on_open.setter
    def update_fields_on_open(self, value: bool):
        self._settings.updateFields_val = value

    @property
    def rsid_root(self) -> str | None:
        """The document's root revision-save ID (``w:rsids/w:rsidRoot/@w:val``).

        Read-only. Returns the 8-character hex string of the first RSID ever
        assigned to this document, or |None| when no ``w:rsids`` or
        ``w:rsidRoot`` element is present. Word generates these values; they
        are surfaced here for downstream diff/merge tooling.

        .. versionadded:: 2026.05.0
        """
        rsids = self._settings.rsids
        if rsids is None:
            return None
        return rsids.rsidRoot_val

    @property
    def rsids(self) -> "RsidList":
        """The document's revision-save IDs (``w:rsids`` table).

        Returns an :class:`RsidList` proxy — a live, list-like view over
        every ``w:rsids/w:rsid/@w:val`` value in document order. The proxy
        also exposes the first-save root rsid via :attr:`RsidList.root`
        (``w:rsidRoot/@w:val``) and the complete id set via
        :attr:`RsidList.ids`, and can mint a new editing-session id via
        :meth:`RsidList.new_session`.

        The return value compares equal to the list of id strings, so code
        written against the pre-2026.05.12 signature (``settings.rsids ==
        ['00A1B2C3', ...]``) keeps working.

        An empty ``RsidList`` is returned when no ``w:rsids`` element is
        present, or when it has no ``w:rsid`` children. Any mutation via
        :meth:`~RsidList.new_session` or :meth:`~RsidList.add` materialises
        the ``w:rsids`` container on demand.

        .. versionadded:: 2026.05.0
        .. versionchanged:: 2026.05.12
            Returns :class:`RsidList` (``list[str]`` subclass) instead of a
            plain list — gains ``.root`` / ``.ids`` / ``.new_session()``.
        """
        return RsidList(self._settings)

    def add_rsids(self, rsid_root: str, extra: "set[str] | None" = None) -> None:
        """Record ``rsid_root`` in the settings' ``<w:rsids>`` table.

        Creates the ``<w:rsids>`` block if absent, sets ``<w:rsidRoot>``
        if absent (first-save wins), and appends any RSIDs in ``extra``
        that are not already present. Called from
        :meth:`DocumentPart.before_marshal` to keep document and
        settings rsid data consistent.

        .. versionadded:: 2026.05.2
        """
        rsids = self._settings.get_or_add_rsids()
        if rsids.rsidRoot is None:
            rsidRoot = rsids.get_or_add_rsidRoot()
            rsidRoot.val = rsid_root
        existing = set(rsids.rsid_vals)
        for value in {rsid_root, *(extra or set())}:
            if value and value not in existing:
                new_rsid = rsids.add_rsid()
                new_rsid.val = value
                existing.add(value)

    @property
    def zoom_percent(self) -> int | None:
        """The zoom percentage for the document view (e.g. 100 for 100%).

        Read/write. None when no zoom is specified.

        .. versionadded:: 2026.05.0
        """
        return self._settings.zoom_percent

    @zoom_percent.setter
    def zoom_percent(self, value: int | None):
        self._settings.zoom_percent = value

    @property
    def footnote_properties(self) -> FootnoteProperties | None:
        """A |FootnoteProperties| object or |None| if no ``w:footnotePr`` is present.

        Provides document-level footnote configuration (number format, start number,
        restart rule, and position). Use :meth:`add_footnote_properties` to add a
        ``w:footnotePr`` element when none is present.

        .. versionadded:: 2026.05.0
        """
        from docx.footnotes import FootnoteProperties

        footnotePr = self._settings.footnotePr
        if footnotePr is None:
            return None
        return FootnoteProperties(footnotePr)

    def add_footnote_properties(self) -> FootnoteProperties:
        """Return a |FootnoteProperties| proxy, adding ``w:footnotePr`` if needed.

        If a ``w:footnotePr`` element is already present, the existing element is used.

        .. versionadded:: 2026.05.0
        """
        from docx.footnotes import FootnoteProperties

        footnotePr = self._settings.get_or_add_footnotePr()
        return FootnoteProperties(footnotePr)

    def remove_footnote_properties(self) -> None:
        """Remove the ``w:footnotePr`` child element if present.

        .. versionadded:: 2026.05.0
        """
        self._settings._remove_footnotePr()  # pyright: ignore[reportPrivateUsage]

    @property
    def endnote_properties(self) -> EndnoteProperties | None:
        """An |EndnoteProperties| object or |None| if no ``w:endnotePr`` is present.

        Provides document-level endnote configuration (number format, start number,
        restart rule, and position). Use :meth:`add_endnote_properties` to add a
        ``w:endnotePr`` element when none is present.

        .. versionadded:: 2026.05.0
        """
        from docx.endnotes import EndnoteProperties

        endnotePr = self._settings.endnotePr
        if endnotePr is None:
            return None
        return EndnoteProperties(endnotePr)

    def add_endnote_properties(self) -> EndnoteProperties:
        """Return an |EndnoteProperties| proxy, adding ``w:endnotePr`` if needed.

        If a ``w:endnotePr`` element is already present, the existing element is used.

        .. versionadded:: 2026.05.0
        """
        from docx.endnotes import EndnoteProperties

        endnotePr = self._settings.get_or_add_endnotePr()
        return EndnoteProperties(endnotePr)

    def remove_endnote_properties(self) -> None:
        """Remove the ``w:endnotePr`` child element if present.

        .. versionadded:: 2026.05.0
        """
        self._settings._remove_endnotePr()  # pyright: ignore[reportPrivateUsage]

    # -- theme-font language ------------------------------------------------

    @property
    def theme_font_language(self) -> tuple[str | None, str | None, str | None]:
        """Document-level theme-font language tags.

        Returns a 3-tuple ``(latin, east_asian, bidi)`` of language tags
        (e.g. ``("en-US", None, None)``) read from
        ``w:themeFontLang/@w:val``, ``@w:eastAsia``, and ``@w:bidi``. Each
        component is |None| when the corresponding attribute is missing or
        no ``w:themeFontLang`` element is present.

        Assigning a plain string sets the Latin (``@w:val``) tag and clears
        the East-Asian and bidi tags. Assigning a tuple of 1-3 strings sets
        them in ``(latin, east_asian, bidi)`` order; shorter tuples leave
        the trailing entries unset. Assigning |None| (or ``(None, None, None)``)
        removes the element entirely.

        .. versionadded:: 2026.05.0
        """
        lang = self._settings.themeFontLang
        if lang is None:
            return (None, None, None)
        return (lang.val, lang.eastAsia, lang.bidi)

    @theme_font_language.setter
    def theme_font_language(
        self,
        value: str | tuple[str | None, ...] | list[str | None] | None,
    ) -> None:
        if value is None:
            self._settings._remove_themeFontLang()  # pyright: ignore[reportPrivateUsage]
            return
        if isinstance(value, str):
            latin, east_asian, bidi = value, None, None
        else:
            parts = list(value) + [None, None, None]
            latin, east_asian, bidi = parts[0], parts[1], parts[2]
        if latin is None and east_asian is None and bidi is None:
            self._settings._remove_themeFontLang()  # pyright: ignore[reportPrivateUsage]
            return
        lang = self._settings.get_or_add_themeFontLang()
        lang.val = latin
        lang.eastAsia = east_asian
        lang.bidi = bidi

    # -- spell / grammar check toggles --------------------------------------

    @property
    def hide_spelling_errors(self) -> bool:
        """True when Word should hide red spell-check underlines in this document.

        Backed by the ``w:hideSpellingErrors`` element. Read/write.

        .. versionadded:: 2026.05.0
        """
        el = self._settings.hideSpellingErrors
        if el is None:
            return False
        return el.val

    @hide_spelling_errors.setter
    def hide_spelling_errors(self, value: bool) -> None:
        if not value:
            self._settings._remove_hideSpellingErrors()  # pyright: ignore[reportPrivateUsage]
            return
        self._settings.get_or_add_hideSpellingErrors().val = True

    @property
    def hide_grammatical_errors(self) -> bool:
        """True when Word should hide green grammar-check underlines.

        Backed by the ``w:hideGrammaticalErrors`` element. Read/write.

        .. versionadded:: 2026.05.0
        """
        el = self._settings.hideGrammaticalErrors
        if el is None:
            return False
        return el.val

    @hide_grammatical_errors.setter
    def hide_grammatical_errors(self, value: bool) -> None:
        if not value:
            self._settings._remove_hideGrammaticalErrors()  # pyright: ignore[reportPrivateUsage]
            return
        self._settings.get_or_add_hideGrammaticalErrors().val = True

    # -- auto-hyphenation ---------------------------------------------------

    @property
    def auto_hyphenation(self) -> bool:
        """True when automatic hyphenation is enabled for the document.

        Backed by the ``w:autoHyphenation`` element. Read/write.

        .. versionadded:: 2026.05.0
        """
        el = self._settings.autoHyphenation
        if el is None:
            return False
        return el.val

    @auto_hyphenation.setter
    def auto_hyphenation(self, value: bool) -> None:
        if not value:
            self._settings._remove_autoHyphenation()  # pyright: ignore[reportPrivateUsage]
            return
        self._settings.get_or_add_autoHyphenation().val = True

    @property
    def do_not_hyphenate_caps(self) -> bool:
        """True when fully-capitalised words are excluded from hyphenation.

        Backed by the ``w:doNotHyphenateCaps`` element. Read/write.

        .. versionadded:: 2026.05.0
        """
        el = self._settings.doNotHyphenateCaps
        if el is None:
            return False
        return el.val

    @do_not_hyphenate_caps.setter
    def do_not_hyphenate_caps(self, value: bool) -> None:
        if not value:
            self._settings._remove_doNotHyphenateCaps()  # pyright: ignore[reportPrivateUsage]
            return
        self._settings.get_or_add_doNotHyphenateCaps().val = True

    @property
    def consecutive_hyphen_limit(self) -> int | None:
        """Maximum number of consecutive lines that may end with a hyphen, or |None|.

        Backed by ``w:consecutiveHyphenLimit/@w:val``. Read/write. Assigning
        |None| (or a value ≤ 0) removes the element.

        .. versionadded:: 2026.05.0
        """
        el = self._settings.consecutiveHyphenLimit
        if el is None:
            return None
        return el.val

    @consecutive_hyphen_limit.setter
    def consecutive_hyphen_limit(self, value: int | None) -> None:
        if value is None or value <= 0:
            self._settings._remove_consecutiveHyphenLimit()  # pyright: ignore[reportPrivateUsage]
            return
        self._settings.get_or_add_consecutiveHyphenLimit().val = int(value)

    @property
    def hyphenation_zone(self) -> Length | None:
        """Hyphenation zone width as a |Length|, or |None| when not set.

        Backed by ``w:hyphenationZone/@w:val`` (a twips measure). Read/write.

        .. versionadded:: 2026.05.0
        """
        el = self._settings.hyphenationZone
        if el is None:
            return None
        return el.val

    @hyphenation_zone.setter
    def hyphenation_zone(self, value: int | Length | None) -> None:
        if value is None:
            self._settings._remove_hyphenationZone()  # pyright: ignore[reportPrivateUsage]
            return
        self._settings.get_or_add_hyphenationZone().val = value

    # -- document variables -------------------------------------------------

    @property
    def doc_vars(self) -> DocVars:
        """A |DocVars| mapping proxy over ``w:docVars/w:docVar`` entries.

        Keys are the ``@w:name`` strings; values are the ``@w:val`` strings.
        The returned object is a live view: assignments and deletions mutate
        the underlying XML immediately and create or remove the ``w:docVars``
        container element as needed.

        .. versionadded:: 2026.05.0
        """
        return DocVars(self._settings)

    # -- Microsoft-extension document identifiers ------------------------

    @property
    def doc_id(self) -> str | None:
        """The document's stable GUID identifier or |None| if not set.

        Word 2013+ stamps ``<w15:docId w15:val="{GUID}"/>`` inside
        ``settings.xml`` on every new document; it drives revision-
        tracking and "same document?" heuristics across sessions.

        Read/write. Assigning a GUID string (with or without
        surrounding braces) creates or updates the element; assigning
        |None| removes it. Reads prefer ``w15:docId`` over the legacy
        ``w14:docId``.

        .. versionadded:: 2026.05.3
        """
        w15 = self._settings.w15_docId
        if w15 is not None and w15.w15_val is not None:
            return w15.w15_val
        w14 = self._settings.w14_docId
        if w14 is not None and w14.w14_val is not None:
            return w14.w14_val
        return None

    @doc_id.setter
    def doc_id(self, value: str | None) -> None:
        if value is None:
            self._settings._remove_w15_docId()  # pyright: ignore[reportPrivateUsage]
            self._settings._remove_w14_docId()  # pyright: ignore[reportPrivateUsage]
            return
        normalised = value if value.startswith("{") else "{%s}" % value
        self._settings.get_or_add_w15_docId().w15_val = normalised

    @property
    def chart_tracking_ref_based(self) -> bool:
        """True when ``<w15:chartTrackingRefBased/>`` is present in settings.

        Word writes this flag on every new document to drive chart
        reference-tracking; Word reads the absence of the element as
        "disabled". Read/write.

        .. versionadded:: 2026.05.3
        """
        return self._settings.chartTrackingRefBased is not None

    @chart_tracking_ref_based.setter
    def chart_tracking_ref_based(self, value: bool) -> None:
        if value:
            self._settings.get_or_add_chartTrackingRefBased()
        else:
            self._settings._remove_chartTrackingRefBased()  # pyright: ignore[reportPrivateUsage]


# -- default algorithm metadata matching Word's rsaAES/SHA-1 password scheme --
_DEFAULT_CRYPT_PROVIDER_TYPE = "rsaAES"
_DEFAULT_CRYPT_ALGORITHM_CLASS = "hash"
_DEFAULT_CRYPT_ALGORITHM_TYPE = "typeAny"
_DEFAULT_CRYPT_ALGORITHM_SID = 4  # -- 4 == SHA-1 in Office's algorithm-id table --
_DEFAULT_SPIN_COUNT = 100000


def _hash_password(password: str, salt: bytes, spin_count: int) -> str:
    """Compute Word-compatible SHA-1 password hash.

    Word's algorithm (ISO/IEC 29500-1 §17.15.1.28) hashes the UTF-16LE encoding
    of the password prefixed by the salt, then re-hashes the previous digest
    concatenated with a 4-byte little-endian iteration counter `spin_count`
    times. Returns the base64-encoded final 20-byte digest.

    Note: Word's implementation has historically had subtle variations; callers
    for whom Word must accept the password at open time should verify against
    their target Word version. For detection-only uses this implementation is
    sufficient.
    """
    digest = hashlib.sha1(salt + password.encode("utf-16-le")).digest()
    for iteration in range(spin_count):
        digest = hashlib.sha1(digest + iteration.to_bytes(4, "little")).digest()
    return base64.b64encode(digest).decode("ascii")


class DocumentProtection:
    """Read/write access to document-protection settings.

    Wraps the ``w:documentProtection`` child of ``w:settings``. All attributes
    are live — writes are persisted to the underlying XML immediately and the
    ``w:documentProtection`` element is created on demand. Setting an attribute
    to |None| (or |False| for bools) clears the corresponding XML attribute.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, settings: CT_Settings):
        self._settings = settings

    # -- internal helpers ---------------------------------------------------

    def _dp_or_none(self):
        return self._settings.documentProtection

    def _dp_or_add(self):
        return self._settings.get_or_add_documentProtection()

    # -- mode / enforcement / formatting ------------------------------------

    @property
    def enforce(self) -> bool:
        """True when document protection is enforced (``@w:enforcement``).

        .. versionadded:: 2026.05.0
        """
        return self._settings.documentProtection_enforcement

    @enforce.setter
    def enforce(self, value: bool) -> None:
        self._settings.documentProtection_enforcement = bool(value)

    @property
    def mode(self) -> WD_PROTECTION | None:
        """The protection mode as a |WD_PROTECTION| member, or |None|.

        Corresponds to the ``@w:edit`` attribute. Assigning |None| clears the
        attribute; assigning a :class:`WD_PROTECTION` member maps to the
        corresponding XML string (e.g. ``WD_PROTECTION.COMMENTS`` → ``comments``).

        .. versionadded:: 2026.05.0
        """
        edit = self._settings.documentProtection_edit
        if edit is None:
            return None
        return WD_PROTECTION.from_xml(edit)

    @mode.setter
    def mode(self, value: WD_PROTECTION | None) -> None:
        if value is None:
            self._settings.documentProtection_edit = None
            return
        self._settings.documentProtection_edit = WD_PROTECTION.to_xml(value)

    @property
    def formatting_locked(self) -> bool:
        """True when formatting restrictions are enabled (``@w:formatting``).

        .. versionadded:: 2026.05.0
        """
        dp = self._dp_or_none()
        if dp is None:
            return False
        return dp.formatting

    @formatting_locked.setter
    def formatting_locked(self, value: bool) -> None:
        self._dp_or_add().formatting = bool(value)

    # -- password hash / salt -----------------------------------------------

    @property
    def password_hash(self) -> str | None:
        """Base64-encoded password hash (``@w:hash``) or |None|.

        .. versionadded:: 2026.05.0
        """
        dp = self._dp_or_none()
        if dp is None:
            return None
        return dp.hash

    @password_hash.setter
    def password_hash(self, value: str | None) -> None:
        dp = self._dp_or_none()
        if value is None:
            if dp is not None:
                dp.hash = None
            return
        self._dp_or_add().hash = value

    @property
    def password_salt(self) -> str | None:
        """Base64-encoded salt (``@w:salt``) used with the password hash, or |None|.

        .. versionadded:: 2026.05.0
        """
        dp = self._dp_or_none()
        if dp is None:
            return None
        return dp.salt

    @password_salt.setter
    def password_salt(self, value: str | None) -> None:
        dp = self._dp_or_none()
        if value is None:
            if dp is not None:
                dp.salt = None
            return
        self._dp_or_add().salt = value

    # -- algorithm metadata -------------------------------------------------

    @property
    def crypto_provider_type(self) -> str | None:
        """Value of ``@w:cryptProviderType`` (e.g. ``"rsaAES"``) or |None|.

        .. versionadded:: 2026.05.0
        """
        dp = self._dp_or_none()
        if dp is None:
            return None
        return dp.cryptProviderType

    @crypto_provider_type.setter
    def crypto_provider_type(self, value: str | None) -> None:
        dp = self._dp_or_none()
        if value is None:
            if dp is not None:
                dp.cryptProviderType = None
            return
        self._dp_or_add().cryptProviderType = value

    @property
    def crypto_algorithm_class(self) -> str | None:
        """Value of ``@w:cryptAlgorithmClass`` (e.g. ``"hash"``) or |None|.

        .. versionadded:: 2026.05.0
        """
        dp = self._dp_or_none()
        if dp is None:
            return None
        return dp.cryptAlgorithmClass

    @crypto_algorithm_class.setter
    def crypto_algorithm_class(self, value: str | None) -> None:
        dp = self._dp_or_none()
        if value is None:
            if dp is not None:
                dp.cryptAlgorithmClass = None
            return
        self._dp_or_add().cryptAlgorithmClass = value

    @property
    def crypto_algorithm_type(self) -> str | None:
        """Value of ``@w:cryptAlgorithmType`` (e.g. ``"typeAny"``) or |None|.

        .. versionadded:: 2026.05.0
        """
        dp = self._dp_or_none()
        if dp is None:
            return None
        return dp.cryptAlgorithmType

    @crypto_algorithm_type.setter
    def crypto_algorithm_type(self, value: str | None) -> None:
        dp = self._dp_or_none()
        if value is None:
            if dp is not None:
                dp.cryptAlgorithmType = None
            return
        self._dp_or_add().cryptAlgorithmType = value

    @property
    def crypto_algorithm_sid(self) -> int | None:
        """Value of ``@w:cryptAlgorithmSid`` (algorithm-id integer) or |None|.

        .. versionadded:: 2026.05.0
        """
        dp = self._dp_or_none()
        if dp is None:
            return None
        return dp.cryptAlgorithmSid

    @crypto_algorithm_sid.setter
    def crypto_algorithm_sid(self, value: int | None) -> None:
        dp = self._dp_or_none()
        if value is None:
            if dp is not None:
                dp.cryptAlgorithmSid = None
            return
        self._dp_or_add().cryptAlgorithmSid = int(value)

    @property
    def spin_count(self) -> int | None:
        """Value of ``@w:cryptSpinCount`` (iteration count) or |None|.

        .. versionadded:: 2026.05.0
        """
        dp = self._dp_or_none()
        if dp is None:
            return None
        return dp.cryptSpinCount

    @spin_count.setter
    def spin_count(self, value: int | None) -> None:
        dp = self._dp_or_none()
        if value is None:
            if dp is not None:
                dp.cryptSpinCount = None
            return
        self._dp_or_add().cryptSpinCount = int(value)

    # -- high-level helpers --------------------------------------------------

    def set_password(self, password: str) -> None:
        """Populate ``@w:hash``/``@w:salt`` and algorithm metadata from `password`.

        Generates a fresh 16-byte random salt, hashes the password using the
        Word-standard SHA-1 scheme with 100,000 iterations, and sets the
        ``@w:cryptProviderType=rsaAES``, ``@w:cryptAlgorithmClass=hash``,
        ``@w:cryptAlgorithmType=typeAny``, ``@w:cryptAlgorithmSid=4``,
        ``@w:cryptSpinCount=100000`` attributes accordingly.

        .. versionadded:: 2026.05.0
        """
        salt_bytes = os.urandom(16)
        digest = _hash_password(password, salt_bytes, _DEFAULT_SPIN_COUNT)
        dp = self._dp_or_add()
        dp.cryptProviderType = _DEFAULT_CRYPT_PROVIDER_TYPE
        dp.cryptAlgorithmClass = _DEFAULT_CRYPT_ALGORITHM_CLASS
        dp.cryptAlgorithmType = _DEFAULT_CRYPT_ALGORITHM_TYPE
        dp.cryptAlgorithmSid = _DEFAULT_CRYPT_ALGORITHM_SID
        dp.cryptSpinCount = _DEFAULT_SPIN_COUNT
        dp.salt = base64.b64encode(salt_bytes).decode("ascii")
        dp.hash = digest

    # -- backward-compat aliases --------------------------------------------

    @property
    def enabled(self) -> bool:
        """Alias for :attr:`enforce` (pre-existing API).

        .. versionadded:: 2026.05.0
        """
        return self.enforce

    @property
    def type(self) -> str | None:
        """Raw ``@w:edit`` string or |None| (pre-existing API).

        Prefer :attr:`mode`, which returns a |WD_PROTECTION| enum member.

        .. versionadded:: 2026.05.0
        """
        return self._settings.documentProtection_edit


# -- backward-compat: preserve private name referenced elsewhere --
_DocumentProtection = DocumentProtection


class WriteProtection:
    """Read/write access to the document's ``w:writeProtection`` marker.

    Wraps the ``w:writeProtection`` child of ``w:settings``. The element is
    created on demand when any attribute is first written. All mutations are
    persisted to the underlying XML immediately.

    Distinct from :class:`DocumentProtection`: write-protection guards *save*
    access (Word will refuse to overwrite the file without the password),
    whereas document-protection restricts edits *within* an opened document.

    .. versionadded:: 2026.05.10
    """

    def __init__(self, settings: CT_Settings):
        self._settings = settings

    # -- internal helpers ---------------------------------------------------

    def _wp_or_none(self) -> "CT_WriteProtection | None":
        return self._settings.writeProtection

    def _wp_or_add(self) -> "CT_WriteProtection":
        return self._settings.get_or_add_writeProtection()

    # -- presence -----------------------------------------------------------

    @property
    def present(self) -> bool:
        """True when a ``w:writeProtection`` element exists.

        Note that an empty ``w:writeProtection`` element (no attributes) is
        semantically equivalent to no element at all: it neither enforces
        recommended-read-only nor enables password-to-modify.

        .. versionadded:: 2026.05.10
        """
        return self._wp_or_none() is not None

    # -- recommended_read_only ----------------------------------------------

    @property
    def recommended_read_only(self) -> bool:
        """True when ``@w:recommended`` is set, i.e. the "open read-only" banner.

        Reads False when the ``w:writeProtection`` element is missing.
        Assigning False while the element carries password attributes clears
        only the recommended flag and leaves the password intact.

        .. versionadded:: 2026.05.10
        """
        wp = self._wp_or_none()
        if wp is None:
            return False
        return wp.recommended

    @recommended_read_only.setter
    def recommended_read_only(self, value: bool) -> None:
        self._wp_or_add().recommended = bool(value)

    # -- ECMA-style enforcement alias --------------------------------------

    @property
    def enforcement(self) -> bool:
        """Alias for :attr:`recommended_read_only`.

        ``w:writeProtection`` has no distinct ``@w:enforcement`` attribute in
        the schema — the presence of the element together with
        ``@w:recommended`` or a populated password is what Word treats as
        "enforced". This alias mirrors the equivalent API on
        :class:`DocumentProtection` for symmetric call sites.

        .. versionadded:: 2026.05.10
        """
        return self.recommended_read_only

    @enforcement.setter
    def enforcement(self, value: bool) -> None:
        self.recommended_read_only = value

    # -- password hash / salt -----------------------------------------------

    @property
    def password_hash(self) -> str | None:
        """Base64-encoded password hash (``@w:hash``) or |None|.

        .. versionadded:: 2026.05.10
        """
        wp = self._wp_or_none()
        if wp is None:
            return None
        return wp.hash

    @password_hash.setter
    def password_hash(self, value: str | None) -> None:
        wp = self._wp_or_none()
        if value is None:
            if wp is not None:
                wp.hash = None
            return
        self._wp_or_add().hash = value

    @property
    def password_salt(self) -> str | None:
        """Base64-encoded salt (``@w:salt``) used with the password hash, or |None|.

        .. versionadded:: 2026.05.10
        """
        wp = self._wp_or_none()
        if wp is None:
            return None
        return wp.salt

    @password_salt.setter
    def password_salt(self, value: str | None) -> None:
        wp = self._wp_or_none()
        if value is None:
            if wp is not None:
                wp.salt = None
            return
        self._wp_or_add().salt = value

    # -- algorithm metadata -------------------------------------------------

    @property
    def crypto_provider_type(self) -> str | None:
        """Value of ``@w:cryptProviderType`` or |None|.

        .. versionadded:: 2026.05.10
        """
        wp = self._wp_or_none()
        if wp is None:
            return None
        return wp.cryptProviderType

    @crypto_provider_type.setter
    def crypto_provider_type(self, value: str | None) -> None:
        wp = self._wp_or_none()
        if value is None:
            if wp is not None:
                wp.cryptProviderType = None
            return
        self._wp_or_add().cryptProviderType = value

    @property
    def crypto_algorithm_class(self) -> str | None:
        """Value of ``@w:cryptAlgorithmClass`` or |None|.

        .. versionadded:: 2026.05.10
        """
        wp = self._wp_or_none()
        if wp is None:
            return None
        return wp.cryptAlgorithmClass

    @crypto_algorithm_class.setter
    def crypto_algorithm_class(self, value: str | None) -> None:
        wp = self._wp_or_none()
        if value is None:
            if wp is not None:
                wp.cryptAlgorithmClass = None
            return
        self._wp_or_add().cryptAlgorithmClass = value

    @property
    def crypto_algorithm_type(self) -> str | None:
        """Value of ``@w:cryptAlgorithmType`` or |None|.

        .. versionadded:: 2026.05.10
        """
        wp = self._wp_or_none()
        if wp is None:
            return None
        return wp.cryptAlgorithmType

    @crypto_algorithm_type.setter
    def crypto_algorithm_type(self, value: str | None) -> None:
        wp = self._wp_or_none()
        if value is None:
            if wp is not None:
                wp.cryptAlgorithmType = None
            return
        self._wp_or_add().cryptAlgorithmType = value

    @property
    def crypto_algorithm_sid(self) -> int | None:
        """Value of ``@w:cryptAlgorithmSid`` or |None|.

        .. versionadded:: 2026.05.10
        """
        wp = self._wp_or_none()
        if wp is None:
            return None
        return wp.cryptAlgorithmSid

    @crypto_algorithm_sid.setter
    def crypto_algorithm_sid(self, value: int | None) -> None:
        wp = self._wp_or_none()
        if value is None:
            if wp is not None:
                wp.cryptAlgorithmSid = None
            return
        self._wp_or_add().cryptAlgorithmSid = int(value)

    @property
    def spin_count(self) -> int | None:
        """Value of ``@w:cryptSpinCount`` or |None|.

        .. versionadded:: 2026.05.10
        """
        wp = self._wp_or_none()
        if wp is None:
            return None
        return wp.cryptSpinCount

    @spin_count.setter
    def spin_count(self, value: int | None) -> None:
        wp = self._wp_or_none()
        if value is None:
            if wp is not None:
                wp.cryptSpinCount = None
            return
        self._wp_or_add().cryptSpinCount = int(value)

    # -- high-level helpers -------------------------------------------------

    def set_password(self, password: str) -> None:
        """Populate ``@w:hash``/``@w:salt`` and algorithm metadata from `password`.

        Uses the same Word-standard SHA-1 scheme (100,000 iterations, 16-byte
        random salt, ``rsaAES``/``hash``/``typeAny``/SID=4) as
        :meth:`DocumentProtection.set_password`.

        .. versionadded:: 2026.05.10
        """
        salt_bytes = os.urandom(16)
        digest = _hash_password(password, salt_bytes, _DEFAULT_SPIN_COUNT)
        wp = self._wp_or_add()
        wp.cryptProviderType = _DEFAULT_CRYPT_PROVIDER_TYPE
        wp.cryptAlgorithmClass = _DEFAULT_CRYPT_ALGORITHM_CLASS
        wp.cryptAlgorithmType = _DEFAULT_CRYPT_ALGORITHM_TYPE
        wp.cryptAlgorithmSid = _DEFAULT_CRYPT_ALGORITHM_SID
        wp.cryptSpinCount = _DEFAULT_SPIN_COUNT
        wp.salt = base64.b64encode(salt_bytes).decode("ascii")
        wp.hash = digest


class CompatSettings:
    """Dict-like view over ``w:compat/w:compatSetting`` entries.

    Obtained via :attr:`Settings.compat_settings`. Keys are the ``@w:name`` strings;
    values are the ``@w:val`` strings. The collection is a live view -- all
    mutations are persisted to the underlying XML immediately.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, settings: CT_Settings):
        self._settings = settings

    # -- internal helpers ---------------------------------------------------

    def _compat_or_none(self) -> CT_Compat | None:
        return self._settings.compat

    def _compat_or_add(self) -> CT_Compat:
        return self._settings.get_or_add_compat()

    def _prune_compat_if_empty(self) -> None:
        compat = self._settings.compat
        if compat is None:
            return
        if len(compat) == 0:
            self._settings._remove_compat()  # pyright: ignore[reportPrivateUsage]

    # -- Mapping-like protocol ---------------------------------------------

    def __getitem__(self, name: str) -> str:
        compat = self._compat_or_none()
        if compat is not None:
            val = compat.get_compat_setting(name)
            if val is not None:
                return val
        raise KeyError(name)

    def __setitem__(self, name: str, value: str) -> None:
        self._compat_or_add().set_compat_setting(name, str(value))

    def __delitem__(self, name: str) -> None:
        compat = self._compat_or_none()
        if compat is None or not compat.remove_compat_setting(name):
            raise KeyError(name)
        self._prune_compat_if_empty()

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        compat = self._compat_or_none()
        if compat is None:
            return False
        return compat.get_compat_setting(name) is not None

    def __iter__(self) -> Iterator[str]:
        compat = self._compat_or_none()
        if compat is None:
            return iter(())
        return iter(list(compat.iter_compat_setting_names()))

    def __len__(self) -> int:
        compat = self._compat_or_none()
        if compat is None:
            return 0
        return len(compat.compatSetting_lst)

    # -- convenience --------------------------------------------------------

    def get(self, name: str, default: str | None = None) -> str | None:
        """Return the value for ``name`` if present, else `default`.

        .. versionadded:: 2026.05.0
        """
        compat = self._compat_or_none()
        if compat is None:
            return default
        val = compat.get_compat_setting(name)
        return default if val is None else val

    def find(self, name: str, uri: str | None = None) -> str | None:
        """Return the ``@w:val`` of the compatSetting matching `name` (and optional `uri`).

        Distinct from :meth:`get`: returns |None| on miss (no default), and
        narrows the match by `uri` when supplied. Multiple compatSettings
        can share a name if they carry different URIs (e.g. a vendor's
        custom namespace), so callers round-tripping third-party flags
        should prefer this over the bare :meth:`__getitem__`.

        .. versionadded:: 2026.05.10
        """
        compat = self._compat_or_none()
        if compat is None:
            return None
        for setting in compat.compatSetting_lst:
            if setting.name != name:
                continue
            if uri is not None and setting.uri != uri:
                continue
            return setting.val
        return None

    def set(
        self,
        name: str,
        uri: str,
        val: str,
    ) -> None:
        """Create or update the compatSetting identified by ``(name, uri)``.

        When a compatSetting with matching ``@w:name`` and ``@w:uri`` already
        exists, its ``@w:val`` is updated in place. Otherwise a new
        ``w:compatSetting`` element is appended.

        Distinct from ``proxy[name] = val``: that form only matches by name
        and preserves the existing ``@w:uri``, which is appropriate for the
        common ``http://schemas.microsoft.com/office/word`` URI but not for
        third-party flags that share a name across namespaces.

        .. versionadded:: 2026.05.10
        """
        compat = self._compat_or_add()
        for setting in compat.compatSetting_lst:
            if setting.name == name and setting.uri == uri:
                setting.val = str(val)
                return
        compat._add_compatSetting(name=name, uri=uri, val=str(val))  # pyright: ignore[reportPrivateUsage]

    def as_dict(self) -> dict[str, str]:
        """Return a plain ``{name: val}`` dict snapshot of every compatSetting.

        When two compatSettings share a name but differ by URI, later
        entries overwrite earlier ones. Callers needing URI fidelity
        should iterate ``compatSetting_lst`` on the underlying element
        directly.

        .. versionadded:: 2026.05.10
        """
        compat = self._compat_or_none()
        if compat is None:
            return {}
        return {s.name: s.val for s in compat.compatSetting_lst}

    def remove(self, name: str) -> None:
        """Remove the compatSetting named `name`, raising |KeyError| if absent.

        .. versionadded:: 2026.05.0
        """
        del self[name]


class CompatFlags:
    """Dict-like view over direct-child flag elements under ``w:compat``.

    Obtained via :attr:`Settings.compat_flags`. Keys are the flag element's local
    name (without the ``w:`` prefix); values are booleans indicating the element's
    presence. Missing flags read as |False| rather than raising |KeyError| -- this
    matches how Word treats absent flag elements.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, settings: CT_Settings):
        self._settings = settings

    # -- internal helpers ---------------------------------------------------

    def _compat_or_none(self) -> CT_Compat | None:
        return self._settings.compat

    def _compat_or_add(self) -> CT_Compat:
        return self._settings.get_or_add_compat()

    def _prune_compat_if_empty(self) -> None:
        compat = self._settings.compat
        if compat is None:
            return
        if len(compat) == 0:
            self._settings._remove_compat()  # pyright: ignore[reportPrivateUsage]

    # -- Mapping-like protocol ---------------------------------------------

    def __getitem__(self, name: str) -> bool:
        compat = self._compat_or_none()
        if compat is None:
            return False
        return compat.has_flag(name)

    def __setitem__(self, name: str, value: bool) -> None:
        if value:
            self._compat_or_add().set_flag(name, True)
            return
        compat = self._compat_or_none()
        if compat is None:
            return
        compat.set_flag(name, False)
        self._prune_compat_if_empty()

    def __delitem__(self, name: str) -> None:
        compat = self._compat_or_none()
        if compat is None or not compat.has_flag(name):
            raise KeyError(name)
        compat.set_flag(name, False)
        self._prune_compat_if_empty()

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        compat = self._compat_or_none()
        if compat is None:
            return False
        return compat.has_flag(name)

    def __iter__(self) -> Iterator[str]:
        compat = self._compat_or_none()
        if compat is None:
            return iter(())
        return iter(list(compat.iter_flag_names()))

    def __len__(self) -> int:
        compat = self._compat_or_none()
        if compat is None:
            return 0
        return sum(1 for _ in compat.iter_flag_names())

    # -- convenience --------------------------------------------------------

    def clear(self) -> None:
        """Remove every non-``w:compatSetting`` child from ``w:compat``.

        .. versionadded:: 2026.05.0
        """
        compat = self._compat_or_none()
        if compat is None:
            return
        compat.clear_flags()
        self._prune_compat_if_empty()

    @staticmethod
    def names() -> tuple[str, ...]:
        """Return a tuple of well-known compatibility flag names.

        Useful for discoverability -- the returned names correspond to direct child
        elements commonly seen under ``w:compat`` in real-world Word documents.
        Setting a name not in this list still works.

        .. versionadded:: 2026.05.0
        """
        return _KNOWN_COMPAT_FLAG_NAMES


class DocVars:
    """Dict-like view over ``w:docVars/w:docVar`` entries.

    Obtained via :attr:`Settings.doc_vars`. Keys are the ``@w:name`` strings;
    values are the ``@w:val`` strings. The collection is live -- assignments
    and deletions persist to the underlying XML immediately, and the
    ``w:docVars`` container element is created / pruned on demand.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, settings: CT_Settings):
        self._settings = settings

    # -- internal helpers ---------------------------------------------------

    def _docVars_or_none(self) -> "CT_DocVars | None":
        return self._settings.docVars

    def _docVars_or_add(self) -> "CT_DocVars":
        return self._settings.get_or_add_docVars()

    def _prune_if_empty(self) -> None:
        container = self._settings.docVars
        if container is None:
            return
        if len(container.docVar_lst) == 0:
            self._settings._remove_docVars()  # pyright: ignore[reportPrivateUsage]

    # -- Mapping-like protocol ---------------------------------------------

    def __getitem__(self, name: str) -> str:
        container = self._docVars_or_none()
        if container is not None:
            val = container.get_var(name)
            if val is not None:
                return val
        raise KeyError(name)

    def __setitem__(self, name: str, value: str) -> None:
        self._docVars_or_add().set_var(name, str(value))

    def __delitem__(self, name: str) -> None:
        container = self._docVars_or_none()
        if container is None or not container.remove_var(name):
            raise KeyError(name)
        self._prune_if_empty()

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        container = self._docVars_or_none()
        if container is None:
            return False
        return container.get_var(name) is not None

    def __iter__(self) -> Iterator[str]:
        container = self._docVars_or_none()
        if container is None:
            return iter(())
        return iter([dv.name for dv in container.docVar_lst])

    def __len__(self) -> int:
        container = self._docVars_or_none()
        if container is None:
            return 0
        return len(container.docVar_lst)

    # -- convenience --------------------------------------------------------

    def get(self, name: str, default: str | None = None) -> str | None:
        """Return the value for ``name`` if present, else `default`.

        .. versionadded:: 2026.05.0
        """
        container = self._docVars_or_none()
        if container is None:
            return default
        val = container.get_var(name)
        return default if val is None else val

    def items(self) -> list[tuple[str, str]]:
        """Return a list of ``(name, value)`` pairs in document order.

        .. versionadded:: 2026.05.0
        """
        container = self._docVars_or_none()
        if container is None:
            return []
        return [(dv.name, dv.val) for dv in container.docVar_lst]


class RsidList(list):  # type: ignore[type-arg]
    """Live, list-like view over ``w:settings/w:rsids``.

    Obtained via :attr:`Settings.rsids`. Subclasses :class:`list` so that
    ``settings.rsids == ['00A1B2C3', ...]`` comparisons keep working while
    exposing rsid-specific helpers:

    - :attr:`root` — value of ``w:rsidRoot/@w:val`` (first-save rsid), or
      |None| when no ``w:rsidRoot`` is present.
    - :attr:`ids` — the complete ``set`` of rsids (``w:rsid`` values plus
      the root, when present). Constant-time containment checks.
    - :meth:`new_session` — mint a fresh 8-hex-digit rsid, add it to the
      ``w:rsids`` table, and return it. Matches Word's per-edit-session
      behaviour; the caller then uses the returned rsid to tag changed
      paragraphs/runs via :meth:`docx.document.Document.tag_revisions`.

    The list contents are snapshotted at construction time — mutating the
    underlying XML after the proxy was retrieved does not refresh the
    list. Call ``document.settings.rsids`` again to see newer rsids.

    .. versionadded:: 2026.05.12
    """

    def __init__(self, settings: "CT_Settings"):
        rsids_elm = settings.rsids
        super().__init__(rsids_elm.rsid_vals if rsids_elm is not None else [])
        self._settings = settings

    @property
    def root(self) -> str | None:
        """The first-save rsid (``w:rsidRoot/@w:val``) or |None| when absent."""
        rsids_elm = self._settings.rsids
        if rsids_elm is None:
            return None
        return rsids_elm.rsidRoot_val

    @property
    def ids(self) -> set[str]:
        """Every rsid referenced by this document as a :class:`set`.

        Includes ``w:rsidRoot/@w:val`` (when present) plus each
        ``w:rsid/@w:val``. Intended for constant-time containment checks
        ("has this rsid been seen before?").
        """
        rsids_elm = self._settings.rsids
        if rsids_elm is None:
            return set()
        values = set(rsids_elm.rsid_vals)
        root = rsids_elm.rsidRoot_val
        if root is not None:
            values.add(root)
        return values

    def add(self, rsid: str) -> None:
        """Append ``rsid`` to ``w:rsids`` when not already present.

        Creates the ``w:rsids`` container on demand. A no-op when
        ``rsid`` is already recorded. Does not touch ``w:rsidRoot``.
        """
        rsids_elm = self._settings.get_or_add_rsids()
        if rsid in set(rsids_elm.rsid_vals):
            return
        new_rsid = rsids_elm.add_rsid()
        new_rsid.val = rsid
        # keep the in-memory list snapshot consistent with the XML
        super().append(rsid)

    def new_session(self) -> str:
        """Mint and register a fresh editing-session rsid, returning it.

        Generates a random 8-character uppercase-hex string (Word's
        ``ST_LongHexNumber`` shape — Word in practice prefixes its rsids
        with ``00`` so the values fit in a signed 32-bit integer, and
        this implementation follows the same convention). The value is
        added to ``w:rsids`` so the returned token is immediately
        referenceable from any ``w:rsidR`` / ``w:rsidP`` / ``w:rsidRPr``
        attribute on downstream paragraphs, runs, or section properties.

        The first ``new_session()`` call on a document whose
        ``w:rsidRoot`` is not yet set also populates it — this matches
        Word's first-save behaviour where the initial editing-session
        rsid becomes the document's root rsid.
        """
        token = "00" + secrets.token_hex(3).upper()
        rsids_elm = self._settings.get_or_add_rsids()
        if rsids_elm.rsidRoot is None:
            rsidRoot = rsids_elm.get_or_add_rsidRoot()
            rsidRoot.val = token
        if token not in set(rsids_elm.rsid_vals):
            new_rsid = rsids_elm.add_rsid()
            new_rsid.val = token
            super().append(token)
        return token


class MailMerge:
    """Access to the mail-merge configuration stored in ``w:settings/w:mailMerge``.

    python-docx does not execute mail merges; this proxy exposes the stored
    configuration (main-document type, destination, data-source metadata,
    query, active record, etc.) so callers can inspect or modify the settings
    that Word will use when the merge is run.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, mailMerge: CT_MailMerge):
        self._mm = mailMerge

    # -- mainDocumentType ---------------------------------------------------

    @property
    def main_document_type(self) -> WD_MAIL_MERGE_TYPE | None:
        el = self._mm.mainDocumentType
        if el is None or el.val is None:
            return None
        return WD_MAIL_MERGE_TYPE.from_xml(el.val)

    @main_document_type.setter
    def main_document_type(self, value: WD_MAIL_MERGE_TYPE | None):
        if value is None:
            self._mm._remove_mainDocumentType()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_mainDocumentType().val = value.xml_value

    # -- destination --------------------------------------------------------

    @property
    def destination(self) -> WD_MAIL_MERGE_DESTINATION | None:
        el = self._mm.destination
        if el is None or el.val is None:
            return None
        return WD_MAIL_MERGE_DESTINATION.from_xml(el.val)

    @destination.setter
    def destination(self, value: WD_MAIL_MERGE_DESTINATION | None):
        if value is None:
            self._mm._remove_destination()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_destination().val = value.xml_value

    # -- dataType -----------------------------------------------------------

    @property
    def data_type(self) -> WD_MAIL_MERGE_DATA_TYPE | None:
        el = self._mm.dataType
        if el is None or el.val is None:
            return None
        try:
            return WD_MAIL_MERGE_DATA_TYPE.from_xml(el.val)
        except ValueError:
            return None

    @data_type.setter
    def data_type(self, value: WD_MAIL_MERGE_DATA_TYPE | None):
        if value is None:
            self._mm._remove_dataType()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_dataType().val = value.xml_value

    # -- connect_string -----------------------------------------------------

    @property
    def connect_string(self) -> str | None:
        el = self._mm.connectString
        return el.val if el is not None else None

    @connect_string.setter
    def connect_string(self, value: str | None):
        if value is None:
            self._mm._remove_connectString()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_connectString().val = value

    # -- query --------------------------------------------------------------

    @property
    def query(self) -> str | None:
        el = self._mm.query
        return el.val if el is not None else None

    @query.setter
    def query(self, value: str | None):
        if value is None:
            self._mm._remove_query()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_query().val = value

    # -- mail_subject -------------------------------------------------------

    @property
    def mail_subject(self) -> str | None:
        el = self._mm.mailSubject
        return el.val if el is not None else None

    @mail_subject.setter
    def mail_subject(self, value: str | None):
        if value is None:
            self._mm._remove_mailSubject()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_mailSubject().val = value

    # -- address_field_name -------------------------------------------------

    @property
    def address_field_name(self) -> str | None:
        el = self._mm.addressFieldName
        return el.val if el is not None else None

    @address_field_name.setter
    def address_field_name(self, value: str | None):
        if value is None:
            self._mm._remove_addressFieldName()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_addressFieldName().val = value

    # -- active_record ------------------------------------------------------

    @property
    def active_record(self) -> int | None:
        el = self._mm.activeRecord
        if el is None or el.val is None:
            return None
        try:
            return int(el.val)
        except (TypeError, ValueError):
            return None

    @active_record.setter
    def active_record(self, value: int | None):
        if value is None:
            self._mm._remove_activeRecord()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_activeRecord().val = str(int(value))

    # -- check_errors -------------------------------------------------------

    @property
    def check_errors(self) -> int | None:
        el = self._mm.checkErrors
        if el is None or el.val is None:
            return None
        try:
            return int(el.val)
        except (TypeError, ValueError):
            return None

    @check_errors.setter
    def check_errors(self, value: int | None):
        if value is None:
            self._mm._remove_checkErrors()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_checkErrors().val = str(int(value))

    # -- bool flags ---------------------------------------------------------

    def _get_bool(self, tag: str) -> bool:
        el = getattr(self._mm, tag)
        if el is None:
            return False
        val = el.val if hasattr(el, "val") else None
        # absent val on an ST_OnOff wrapper is "on"
        return True if val is None else bool(val)

    def _set_bool(self, tag: str, value: bool) -> None:
        if value:
            getattr(self._mm, f"get_or_add_{tag}")()
        else:
            getattr(self._mm, f"_remove_{tag}")()

    @property
    def link_to_query(self) -> bool:
        return self._get_bool("linkToQuery")

    @link_to_query.setter
    def link_to_query(self, value: bool):
        self._set_bool("linkToQuery", value)

    @property
    def do_not_suppress_blank_lines(self) -> bool:
        return self._get_bool("doNotSuppressBlankLines")

    @do_not_suppress_blank_lines.setter
    def do_not_suppress_blank_lines(self, value: bool):
        self._set_bool("doNotSuppressBlankLines", value)

    @property
    def mail_as_attachment(self) -> bool:
        return self._get_bool("mailAsAttachment")

    @mail_as_attachment.setter
    def mail_as_attachment(self, value: bool):
        self._set_bool("mailAsAttachment", value)

    @property
    def view_merged_data(self) -> bool:
        return self._get_bool("viewMergedData")

    @view_merged_data.setter
    def view_merged_data(self, value: bool):
        self._set_bool("viewMergedData", value)

    # -- odso ---------------------------------------------------------------

    @property
    def odso(self) -> "OdsoSettings | None":
        """The :class:`OdsoSettings` proxy or |None| when no ``w:odso`` child.

        .. versionadded:: 2026.05.10
        """
        odso = self._mm.odso
        if odso is None:
            return None
        return OdsoSettings(odso)

    def add_odso(self) -> "OdsoSettings":
        """Create ``w:mailMerge/w:odso`` (if missing) and return the proxy.

        .. versionadded:: 2026.05.10
        """
        odso = self._mm.get_or_add_odso()
        return OdsoSettings(odso)

    def remove_odso(self) -> None:
        """Remove the ``w:odso`` child element if present.

        .. versionadded:: 2026.05.10
        """
        self._mm._remove_odso()  # pyright: ignore[reportPrivateUsage]

    # -- data_source (`w:mailMerge/w:dataSource` — the rId-based relationship) ---

    @property
    def data_source(self) -> str | None:
        """The relationship id (``r:id``) referencing the merge data-source part.

        Corresponds to ``w:mailMerge/w:dataSource/@r:id``. Read/write. Assigning
        |None| removes the element.

        .. versionadded:: 2026.05.10
        """
        ds = self._mm.dataSource
        if ds is None:
            return None
        return ds.rId

    @data_source.setter
    def data_source(self, value: str | None) -> None:
        if value is None:
            self._mm._remove_dataSource()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_dataSource().rId = value

    # -- header_source (same shape as dataSource) ---------------------------

    @property
    def header_source(self) -> str | None:
        """The relationship id (``r:id``) referencing the header-source part.

        Corresponds to ``w:mailMerge/w:headerSource/@r:id``. Read/write.

        .. versionadded:: 2026.05.10
        """
        hs = self._mm.headerSource
        if hs is None:
            return None
        return hs.rId

    @header_source.setter
    def header_source(self, value: str | None) -> None:
        if value is None:
            self._mm._remove_headerSource()  # pyright: ignore[reportPrivateUsage]
            return
        self._mm.get_or_add_headerSource().rId = value


class OdsoSettings:
    """Read/write access to the ``w:mailMerge/w:odso`` (ODSO) data manifest.

    ODSO — *Office Data Source Object* — is Word's schema-tagged description
    of the external data source backing a mail merge. It records the UDL
    path, the table/view name, the optional relationship-referenced source
    file, the column delimiter, the source-type tag, whether the first row
    carries column names, and a dict-like mapping from external column names
    to merge-field names.

    python-docx does not execute mail-merge — this proxy is a faithful
    read/write view that preserves authored metadata across save.

    .. versionadded:: 2026.05.10
    """

    def __init__(self, odso: "CT_Odso"):
        self._odso = odso

    # -- udl / table (string val-wrappers) ---------------------------------

    @property
    def udl(self) -> str | None:
        """Value of ``w:udl/@w:val``, a UDL (Microsoft Data Link) file path.

        .. versionadded:: 2026.05.10
        """
        return self._odso._val_child_read("udl")  # pyright: ignore[reportPrivateUsage]

    @udl.setter
    def udl(self, value: str | None) -> None:
        self._odso._val_child_write("udl", value)  # pyright: ignore[reportPrivateUsage]

    @property
    def table(self) -> str | None:
        """Value of ``w:table/@w:val`` — the table / view / sheet name.

        .. versionadded:: 2026.05.10
        """
        return self._odso._val_child_read("table")  # pyright: ignore[reportPrivateUsage]

    @table.setter
    def table(self, value: str | None) -> None:
        self._odso._val_child_write("table", value)  # pyright: ignore[reportPrivateUsage]

    # -- src (relationship reference) ---------------------------------------

    @property
    def src(self) -> str | None:
        """The relationship id (``r:id``) of the ``w:src`` child, or |None|.

        Corresponds to ``w:odso/w:src/@r:id`` — a pointer to the external
        source-file relationship.

        .. versionadded:: 2026.05.10
        """
        el = self._odso.src
        if el is None:
            return None
        return el.rId

    @src.setter
    def src(self, value: str | None) -> None:
        if value is None:
            self._odso._remove_src()  # pyright: ignore[reportPrivateUsage]
            return
        self._odso.get_or_add_src().rId = value

    # -- colDelim (decimal val-wrapper) ------------------------------------

    @property
    def column_delimiter(self) -> int | None:
        """The ASCII code of the column delimiter character (``w:colDelim``).

        For example, ``44`` for comma-separated files or ``9`` for tab.
        Read/write. |None| when the element is absent or the value cannot be
        parsed as an integer.

        .. versionadded:: 2026.05.10
        """
        raw = self._odso._val_child_read("colDelim")  # pyright: ignore[reportPrivateUsage]
        if raw is None:
            return None
        try:
            return int(raw)
        except (TypeError, ValueError):
            return None

    @column_delimiter.setter
    def column_delimiter(self, value: int | None) -> None:
        if value is None:
            self._odso._val_child_write("colDelim", None)  # pyright: ignore[reportPrivateUsage]
            return
        self._odso._val_child_write("colDelim", int(value))  # pyright: ignore[reportPrivateUsage]

    # -- type (enum val-wrapper) -------------------------------------------

    @property
    def type(self) -> "WD_ODSO_TYPE | None":
        """The ODSO source-type as a :class:`WD_ODSO_TYPE` member, or |None|.

        Corresponds to ``w:odso/w:type/@w:val``. Returns |None| when the
        element is absent or the value isn't a recognised enum member.

        .. versionadded:: 2026.05.10
        """
        raw = self._odso._val_child_read("type")  # pyright: ignore[reportPrivateUsage]
        if raw is None:
            return None
        try:
            return WD_ODSO_TYPE.from_xml(raw)
        except ValueError:
            return None

    @type.setter
    def type(self, value: "WD_ODSO_TYPE | None") -> None:
        if value is None:
            self._odso._val_child_write("type", None)  # pyright: ignore[reportPrivateUsage]
            return
        self._odso._val_child_write(  # pyright: ignore[reportPrivateUsage]
            "type", value.xml_value
        )

    # -- fHdr (bool flag) ---------------------------------------------------

    @property
    def first_row_has_column_names(self) -> bool:
        """True when ``w:fHdr`` is present (first-row-is-header flag).

        An empty ``w:fHdr`` element reads as |True| (ST_OnOff default).
        Assigning |False| removes the element.

        .. versionadded:: 2026.05.10
        """
        el = self._odso.fHdr
        if el is None:
            return False
        val = el.get(qn_w("val"))
        if val is None:
            return True
        return val.lower() in ("1", "true", "on")

    @first_row_has_column_names.setter
    def first_row_has_column_names(self, value: bool) -> None:
        if value:
            el = self._odso.get_or_add_fHdr()
            # leave attribute absent so the element's presence-default (True) governs
            if el.get(qn_w("val")) is not None:
                del el.attrib[qn_w("val")]
            return
        self._odso._remove_fHdr()  # pyright: ignore[reportPrivateUsage]

    # -- field mapping -----------------------------------------------------

    @property
    def field_mapping(self) -> "dict[str, str]":
        """A ``{merge_field_name: external_column_name}`` dict.

        Built from the ``w:fieldMapData`` child records. Each record
        represents one mapping between a Word merge field and an ODSO data
        source column. Records missing ``w:name`` or ``w:mappedName`` are
        omitted.

        Assigning a dict replaces the entire ``w:fieldMapData`` list with
        fresh records in iteration order. Iteration order of the returned
        dict matches the on-disk ``w:fieldMapData`` order.

        .. versionadded:: 2026.05.10
        """
        return dict(self._odso.iter_field_map())

    @field_mapping.setter
    def field_mapping(self, value: "dict[str, str] | None") -> None:
        self._odso.set_field_map(value or {})
