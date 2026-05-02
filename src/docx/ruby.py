"""Read-only proxy for ruby annotations (`w:ruby`)."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.shared import ElementProxy

if TYPE_CHECKING:
    from docx.oxml.ruby import CT_Ruby


class RubyAnnotation(ElementProxy):
    """Pairing of base text with its phonetic annotation (e.g. Japanese furigana).

    Read-only. The underlying `w:ruby` element is accessible via
    ``self._element`` if needed.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, ruby: CT_Ruby):
        super().__init__(ruby)
        self._ruby = ruby

    @property
    def base_text(self) -> str:
        """Plain text of the base (the characters being annotated).

        .. versionadded:: 2026.05.0
        """
        return self._ruby.base_text

    @property
    def ruby_text(self) -> str:
        """Plain text of the ruby annotation (e.g. furigana).

        .. versionadded:: 2026.05.0
        """
        return self._ruby.ruby_text

    @property
    def alignment(self) -> str | None:
        """Value of `w:rubyPr/w:rubyAlign/@w:val` or |None| if absent.

        Typical values: ``distributeLetter``, ``distributeSpace``, ``center``,
        ``left``, ``right``, ``rightVertical``.

        .. versionadded:: 2026.05.0
        """
        if self._ruby.rubyPr is None or self._ruby.rubyPr.rubyAlign is None:
            return None
        return self._ruby.rubyPr.rubyAlign.val

    @property
    def language(self) -> str | None:
        """Value of `w:rubyPr/w:lid/@w:val` or |None| if absent.

        Typically a BCP-47 language tag like ``"ja-JP"``.

        .. versionadded:: 2026.05.0
        """
        if self._ruby.rubyPr is None or self._ruby.rubyPr.lid is None:
            return None
        return self._ruby.rubyPr.lid.val
