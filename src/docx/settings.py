"""Settings object, providing access to document-level settings."""

from __future__ import annotations

import warnings
from typing import TYPE_CHECKING, cast

from docx.shared import ElementProxy

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.settings import CT_Settings
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.shared import Length


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
    def default_tab_stop(self) -> Length | None:
        """The default tab-stop interval for the document as a |Length| value.

        Read/write. Assign a |Length| value (e.g. ``Twips(720)``) or |None| to remove.
        """
        return self._settings.defaultTabStop_val

    @default_tab_stop.setter
    def default_tab_stop(self, value: int | Length | None):
        self._settings.defaultTabStop_val = value

    @property
    def document_protection(self) -> _DocumentProtection:
        """Read-only access to document protection settings.

        Provides `.type` (str or None) and `.enabled` (bool) properties.
        """
        return _DocumentProtection(self._settings)

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
    def track_revisions(self) -> bool:
        """True when revision tracking is enabled for this document.

        Read/write.
        """
        return self._settings.trackRevisions_val

    @track_revisions.setter
    def track_revisions(self, value: bool):
        self._settings.trackRevisions_val = value

    @property
    def zoom_percent(self) -> int | None:
        """The zoom percentage for the document view (e.g. 100 for 100%).

        Read/write. None when no zoom is specified.
        """
        return self._settings.zoom_percent

    @zoom_percent.setter
    def zoom_percent(self, value: int | None):
        self._settings.zoom_percent = value


class _DocumentProtection:
    """Read-only access to document-protection settings."""

    def __init__(self, settings: CT_Settings):
        self._settings = settings

    @property
    def enabled(self) -> bool:
        """True when document protection is enforced."""
        return self._settings.documentProtection_enforcement

    @property
    def type(self) -> str | None:
        """The protection type (e.g. "readOnly", "comments", "trackedChanges", "forms")
        or None if no protection is set."""
        return self._settings.documentProtection_edit
