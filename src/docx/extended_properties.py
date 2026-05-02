"""Extended (application) document properties.

The extended-properties part (``docProps/app.xml``) stores metadata written by
the authoring application: ``Company``, ``Manager``, ``Application``,
``AppVersion``, ``TotalTime`` (edit time, in minutes), ``Template``, as well as
cached statistics (``Pages``, ``Words``, ``Characters``, ``CharactersWithSpaces``,
``Lines``, ``Paragraphs``). These properties are independent of the Dublin-Core
"core" properties exposed via :attr:`Document.core_properties` and the
user-defined typed name/value pairs exposed via :attr:`Document.custom_properties`.

.. versionadded:: 1.3.0.dev0
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx.oxml.extended_properties import CT_ExtendedProperties


# -- well-known scalar properties exposed as Python attributes. The mapping
# -- associates the attribute name used on :class:`ExtendedProperties` with the
# -- unqualified XML element name under ``<Properties>``. All values are
# -- serialised as strings in the XML; callers are expected to pass the right
# -- Python type for the named property (int for counts, str for text).
_STR_PROPS: dict[str, str] = {
    "application": "Application",
    "app_version": "AppVersion",
    "company": "Company",
    "manager": "Manager",
    "template": "Template",
    "hyperlink_base": "HyperlinkBase",
    "doc_security": "DocSecurity",
    "presentation_format": "PresentationFormat",
}

_INT_PROPS: dict[str, str] = {
    "total_time": "TotalTime",
    "pages": "Pages",
    "words": "Words",
    "characters": "Characters",
    "characters_with_spaces": "CharactersWithSpaces",
    "lines": "Lines",
    "paragraphs": "Paragraphs",
    "slides": "Slides",
    "notes": "Notes",
    "hidden_slides": "HiddenSlides",
    "mm_clips": "MMClips",
}


class ExtendedProperties:
    """Read/write access to the extended-properties part (``docProps/app.xml``).

    Provides typed accessors for common scalar children of the root
    ``<Properties>`` element (``Application``, ``Company``, ``Manager``,
    ``TotalTime``, cached ``Pages`` / ``Words`` / ``Characters`` statistics,
    etc.). Structured children (``HeadingPairs``, ``TitlesOfParts``) are left
    untouched by this API; use :meth:`get` / :meth:`set` for uncommon scalar
    fields.

    .. versionadded:: 1.3.0.dev0
    """

    def __init__(self, element: "CT_ExtendedProperties"):
        self._element = element

    # -- --- generic access --------------------------------------------------

    def get(self, name: str) -> str | None:
        """Return the text content of the child element named ``name``, or |None|.

        ``name`` is the unqualified XML element name (e.g. ``"Company"``).
        Returns |None| if no such child exists; returns the empty string if the
        child exists but has no text. Use this for uncommon fields that don't
        have dedicated attribute accessors.

        .. versionadded:: 1.3.0.dev0
        """
        return self._element.get_text(name)

    def set(self, name: str, value: str | int | float | bool | None) -> None:
        """Set the text of child element ``name`` to ``value``.

        Creates the child if it does not already exist. Passing |None| removes
        the child entirely. ``value`` is converted to a string for
        serialisation (booleans become ``"true"`` / ``"false"``).

        .. versionadded:: 1.3.0.dev0
        """
        self._element.set_text(name, value)

    def clear_all(self) -> None:
        """Clear every text-content scalar child of ``<Properties>``.

        Structured vector children (``HeadingPairs``, ``TitlesOfParts``) are
        preserved. Useful together with ``CoreProperties.clear_all()`` to
        strip identifying metadata from a new document.

        .. versionadded:: 1.3.0.dev0
        """
        for child in list(self._element):
            tag = child.tag
            # -- only remove children that have text content (scalar props) --
            if child.text is None and len(child) > 0:
                continue  # structured child (vector etc.)
            self._element.remove(child)

    # -- --- well-known string property helpers ------------------------------


def _make_str_property(xml_name: str):
    def _getter(self: ExtendedProperties) -> str | None:
        return self._element.get_text(xml_name)  # pyright: ignore[reportPrivateUsage]

    def _setter(self: ExtendedProperties, value: str | None) -> None:
        self._element.set_text(xml_name, value)  # pyright: ignore[reportPrivateUsage]

    return property(_getter, _setter)


def _make_int_property(xml_name: str):
    def _getter(self: ExtendedProperties) -> int | None:
        text = self._element.get_text(xml_name)  # pyright: ignore[reportPrivateUsage]
        if text is None or text == "":
            return None
        try:
            return int(text)
        except ValueError:
            return None

    def _setter(self: ExtendedProperties, value: int | None) -> None:
        self._element.set_text(xml_name, value)  # pyright: ignore[reportPrivateUsage]

    return property(_getter, _setter)


def _install_generated_properties() -> None:
    """Attach str/int property descriptors onto :class:`ExtendedProperties`.

    Extracted into a function (rather than free-standing loops at import time)
    so module-level temporaries don't leak and so static checkers don't flag
    possibly-unbound loop variables.
    """
    for attr, xml_name in _STR_PROPS.items():
        setattr(ExtendedProperties, attr, _make_str_property(xml_name))
    for attr, xml_name in _INT_PROPS.items():
        setattr(ExtendedProperties, attr, _make_int_property(xml_name))


_install_generated_properties()
