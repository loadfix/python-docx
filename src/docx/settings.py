"""Settings object, providing access to document-level settings."""

from __future__ import annotations

import base64
import hashlib
import os
import warnings
from typing import TYPE_CHECKING, Iterator, cast

from docx.enum.text import WD_PROTECTION, WD_VIEW
from docx.shared import ElementProxy

if TYPE_CHECKING:
    import docx.types as t
    from docx.endnotes import EndnoteProperties
    from docx.footnotes import FootnoteProperties
    from docx.oxml.settings import CT_Compat, CT_Settings
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
        """
        return CompatFlags(self._settings)

    @property
    def default_tab_stop(self) -> Length | None:
        """The default tab-stop interval for the document as a |Length| value.

        Read/write. Assign a |Length| value (e.g. ``Twips(720)``) or |None| to remove.
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
        """Remove the ``w:documentProtection`` element entirely."""
        self._settings._remove_documentProtection()  # pyright: ignore[reportPrivateUsage]

    @property
    def even_and_odd_headers(self) -> bool:
        """True if this document has distinct odd and even page headers and footers.

        Read/write.
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

        Read/write.
        """
        return self._settings.trackRevisions_val

    @track_revisions.setter
    def track_revisions(self, value: bool):
        self._settings.trackRevisions_val = value

    @property
    def rsid_root(self) -> str | None:
        """The document's root revision-save ID (``w:rsids/w:rsidRoot/@w:val``).

        Read-only. Returns the 8-character hex string of the first RSID ever
        assigned to this document, or |None| when no ``w:rsids`` or
        ``w:rsidRoot`` element is present. Word generates these values; they
        are surfaced here for downstream diff/merge tooling.
        """
        rsids = self._settings.rsids
        if rsids is None:
            return None
        return rsids.rsidRoot_val

    @property
    def rsids(self) -> list[str]:
        """The document's revision-save IDs (``w:rsids/w:rsid/@w:val`` values).

        Read-only. Returns a list of 8-character hex strings in document order.
        An empty list is returned when no ``w:rsids`` element is present, or
        when it has no ``w:rsid`` children.
        """
        rsids = self._settings.rsids
        if rsids is None:
            return []
        return rsids.rsid_vals

    @property
    def zoom_percent(self) -> int | None:
        """The zoom percentage for the document view (e.g. 100 for 100%).

        Read/write. None when no zoom is specified.
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
        """
        from docx.footnotes import FootnoteProperties

        footnotePr = self._settings.footnotePr
        if footnotePr is None:
            return None
        return FootnoteProperties(footnotePr)

    def add_footnote_properties(self) -> FootnoteProperties:
        """Return a |FootnoteProperties| proxy, adding ``w:footnotePr`` if needed.

        If a ``w:footnotePr`` element is already present, the existing element is used.
        """
        from docx.footnotes import FootnoteProperties

        footnotePr = self._settings.get_or_add_footnotePr()
        return FootnoteProperties(footnotePr)

    def remove_footnote_properties(self) -> None:
        """Remove the ``w:footnotePr`` child element if present."""
        self._settings._remove_footnotePr()  # pyright: ignore[reportPrivateUsage]

    @property
    def endnote_properties(self) -> EndnoteProperties | None:
        """An |EndnoteProperties| object or |None| if no ``w:endnotePr`` is present.

        Provides document-level endnote configuration (number format, start number,
        restart rule, and position). Use :meth:`add_endnote_properties` to add a
        ``w:endnotePr`` element when none is present.
        """
        from docx.endnotes import EndnoteProperties

        endnotePr = self._settings.endnotePr
        if endnotePr is None:
            return None
        return EndnoteProperties(endnotePr)

    def add_endnote_properties(self) -> EndnoteProperties:
        """Return an |EndnoteProperties| proxy, adding ``w:endnotePr`` if needed.

        If a ``w:endnotePr`` element is already present, the existing element is used.
        """
        from docx.endnotes import EndnoteProperties

        endnotePr = self._settings.get_or_add_endnotePr()
        return EndnoteProperties(endnotePr)

    def remove_endnote_properties(self) -> None:
        """Remove the ``w:endnotePr`` child element if present."""
        self._settings._remove_endnotePr()  # pyright: ignore[reportPrivateUsage]


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
        """True when document protection is enforced (``@w:enforcement``)."""
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
        """True when formatting restrictions are enabled (``@w:formatting``)."""
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
        """Base64-encoded password hash (``@w:hash``) or |None|."""
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
        """Base64-encoded salt (``@w:salt``) used with the password hash, or |None|."""
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
        """Value of ``@w:cryptProviderType`` (e.g. ``"rsaAES"``) or |None|."""
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
        """Value of ``@w:cryptAlgorithmClass`` (e.g. ``"hash"``) or |None|."""
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
        """Value of ``@w:cryptAlgorithmType`` (e.g. ``"typeAny"``) or |None|."""
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
        """Value of ``@w:cryptAlgorithmSid`` (algorithm-id integer) or |None|."""
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
        """Value of ``@w:cryptSpinCount`` (iteration count) or |None|."""
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
        """Alias for :attr:`enforce` (pre-existing API)."""
        return self.enforce

    @property
    def type(self) -> str | None:
        """Raw ``@w:edit`` string or |None| (pre-existing API).

        Prefer :attr:`mode`, which returns a |WD_PROTECTION| enum member.
        """
        return self._settings.documentProtection_edit


# -- backward-compat: preserve private name referenced elsewhere --
_DocumentProtection = DocumentProtection


class CompatSettings:
    """Dict-like view over ``w:compat/w:compatSetting`` entries.

    Obtained via :attr:`Settings.compat_settings`. Keys are the ``@w:name`` strings;
    values are the ``@w:val`` strings. The collection is a live view -- all
    mutations are persisted to the underlying XML immediately.
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
        """Return the value for ``name`` if present, else `default`."""
        compat = self._compat_or_none()
        if compat is None:
            return default
        val = compat.get_compat_setting(name)
        return default if val is None else val

    def remove(self, name: str) -> None:
        """Remove the compatSetting named `name`, raising |KeyError| if absent."""
        del self[name]


class CompatFlags:
    """Dict-like view over direct-child flag elements under ``w:compat``.

    Obtained via :attr:`Settings.compat_flags`. Keys are the flag element's local
    name (without the ``w:`` prefix); values are booleans indicating the element's
    presence. Missing flags read as |False| rather than raising |KeyError| -- this
    matches how Word treats absent flag elements.
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
        """Remove every non-``w:compatSetting`` child from ``w:compat``."""
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
        """
        return _KNOWN_COMPAT_FLAG_NAMES
