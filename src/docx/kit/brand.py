"""Brand asset manager — load brand colours / fonts / logos / spacing from YAML.

Closes #90.

A *brand* is a small, declarative bundle of corporate-identity assets that
multiple documents in an organisation re-use: a primary / secondary / accent
colour palette, a body / heading font pair, a few logo variants
(full-colour, monochrome, reverse), and conventional paragraph- and
section-spacing values. The :class:`BrandAssets` class loads such a
bundle from a YAML file and exposes it as a typed, attribute-accessible
object so kit helpers can compose against a single source of truth::

    from docx import Document
    from docx.kit.brand import BrandAssets
    from docx.kit.letterhead import set_letterhead

    brand = BrandAssets.load("aws-brand.yaml")
    doc = Document()

    doc.add_picture(brand.logos.full_color)
    para = doc.add_paragraph("AWS")
    para.runs[0].font.color.rgb = brand.colors.primary
    para.runs[0].font.name = brand.fonts.heading

    set_letterhead(doc, logo=brand.logos.full_color, return_address="...")

YAML schema (every block is optional — missing fields surface as ``None``)::

    name: AWS
    colors:
      primary: '#FF9900'      # AWS Smile Orange
      secondary: '#232F3E'    # Squid Ink
      accent: '#0073BB'       # Lightning Blue
      background: '#FAFAFA'
    fonts:
      heading: 'Amazon Ember Display'
      body: 'Amazon Ember'
    logos:
      full_color: 'logos/aws-full-color.png'
      monochrome: 'logos/aws-mono.png'
      reverse: 'logos/aws-reverse.png'
    spacing:
      paragraph: 12pt
      section: 24pt

Resolution rules:

* **Colours** parse as :class:`docx.shared.RGBColor` from any of
  ``"#FF9900"``, ``"FF9900"``, or a 3-character ``"#F90"`` shorthand.
* **Logos** resolve to absolute paths *relative to the YAML file's
  directory* — so a manifest checked in alongside its assets can be
  loaded from anywhere on disk and still hand kit helpers a working
  path. Pass an absolute path in the YAML to override.
* **Spacing** parses as :class:`docx.shared.Length` using the suffix
  in the value (``"12pt"`` → :class:`docx.shared.Pt`, ``"0.5in"`` →
  :class:`docx.shared.Inches`, etc.). Bare numbers are interpreted as
  points to match the most common authoring case.
* **Fonts** are font *name* strings — the kit doesn't embed fonts; the
  caller is responsible for ensuring the named family is available
  in the rendering environment.

The YAML loader uses :mod:`yaml` (PyYAML), which is widely installed
but not a hard dependency of python-docx. Callers who want to use
:meth:`BrandAssets.load` should install ``pip install pyyaml`` (or the
``[brand]`` extras flag if it is added later); :meth:`BrandAssets.from_dict`
works without PyYAML for callers who already have a parsed mapping in
memory.

.. versionadded:: 2026.05.29
"""

from __future__ import annotations

import os
from collections.abc import Mapping as _Mapping
from dataclasses import dataclass, field
from typing import Any, Mapping, Optional, Union

from docx.shared import Cm, Emu, Inches, Length, Mm, Pt, RGBColor, Twips


# -- Public exception types --------------------------------------------------


class BrandAssetsError(ValueError):
    """Raised when a brand-asset manifest cannot be parsed."""


# -- Internal helpers --------------------------------------------------------


def _parse_color(value: Any, *, where: str) -> RGBColor:
    """Parse *value* into an :class:`RGBColor`.

    Accepts a hex string (``"#RRGGBB"``, ``"RRGGBB"``, ``"#RGB"``,
    ``"RGB"``), an existing :class:`RGBColor` (returned as-is), or a
    3-tuple / 3-list of 0-255 ints.
    """

    if isinstance(value, RGBColor):
        return value
    if isinstance(value, (list, tuple)) and len(value) == 3:
        try:
            return RGBColor(int(value[0]), int(value[1]), int(value[2]))
        except (TypeError, ValueError) as exc:
            raise BrandAssetsError(
                f"{where}: invalid RGB triple {value!r}: {exc}"
            ) from exc
    if not isinstance(value, str):
        raise BrandAssetsError(
            f"{where}: expected a hex string or [r, g, b] triple, "
            f"got {type(value).__name__} ({value!r})"
        )
    s = value.strip()
    if s.startswith("#"):
        s = s[1:]
    if len(s) == 3:  # short form like "F90"
        s = "".join(ch * 2 for ch in s)
    if len(s) != 6:
        raise BrandAssetsError(
            f"{where}: expected a 3- or 6-character hex string, got {value!r}"
        )
    try:
        return RGBColor.from_string(s.upper())
    except ValueError as exc:
        raise BrandAssetsError(f"{where}: not a valid hex string ({value!r})") from exc


_LENGTH_SUFFIXES = {
    "pt": Pt,
    "in": Inches,
    "cm": Cm,
    "mm": Mm,
    "emu": Emu,
    "twip": Twips,
    "twips": Twips,
}


def _parse_length(value: Any, *, where: str) -> Length:
    """Parse *value* into a :class:`Length`.

    Accepts an existing :class:`Length` (returned as-is), an int / float
    (interpreted as points — the most common authoring unit), or a
    string with a unit suffix (``"12pt"``, ``"0.5in"``, ``"24mm"``,
    ``"914400emu"``, ``"720twips"``).
    """

    if isinstance(value, Length):
        return value
    if isinstance(value, bool):
        # bool is an int subclass; reject to avoid ``True == 1pt``
        raise BrandAssetsError(f"{where}: boolean is not a valid length ({value!r})")
    if isinstance(value, (int, float)):
        return Pt(float(value))
    if not isinstance(value, str):
        raise BrandAssetsError(
            f"{where}: expected a length string or number, "
            f"got {type(value).__name__} ({value!r})"
        )
    s = value.strip().lower()
    if not s:
        raise BrandAssetsError(f"{where}: empty length string")
    # Find the longest matching suffix. ``twips`` must beat ``twip``;
    # ``mm``/``cm`` must beat the bare number path. Search longest-first.
    matched_suffix: Optional[str] = None
    for suffix in sorted(_LENGTH_SUFFIXES, key=len, reverse=True):
        if s.endswith(suffix):
            matched_suffix = suffix
            break
    try:
        if matched_suffix is None:
            return Pt(float(s))
        magnitude = s[: -len(matched_suffix)].strip()
        if not magnitude:
            raise ValueError("missing magnitude")
        ctor = _LENGTH_SUFFIXES[matched_suffix]
        return ctor(float(magnitude))
    except (TypeError, ValueError) as exc:
        raise BrandAssetsError(
            f"{where}: cannot parse length {value!r}: {exc}"
        ) from exc


def _resolve_logo_path(value: Any, base_dir: Optional[str], *, where: str) -> str:
    """Resolve *value* to a logo path, anchored to *base_dir* if relative."""

    if not isinstance(value, str):
        raise BrandAssetsError(
            f"{where}: expected a path string, "
            f"got {type(value).__name__} ({value!r})"
        )
    if not value.strip():
        raise BrandAssetsError(f"{where}: empty path")
    if os.path.isabs(value) or base_dir is None:
        return value
    return os.path.normpath(os.path.join(base_dir, value))


# -- Public dataclass views --------------------------------------------------


@dataclass(frozen=True)
class BrandColors:
    """Brand colour palette. Each attribute is an :class:`RGBColor` or ``None``."""

    primary: Optional[RGBColor] = None
    secondary: Optional[RGBColor] = None
    accent: Optional[RGBColor] = None
    background: Optional[RGBColor] = None
    extras: Mapping[str, RGBColor] = field(default_factory=dict)

    def __getitem__(self, key: str) -> RGBColor:
        named = {"primary", "secondary", "accent", "background"}
        if key in named:
            value = getattr(self, key)
            if value is None:
                raise KeyError(key)
            return value
        return self.extras[key]


@dataclass(frozen=True)
class BrandFonts:
    """Brand font pair. Each attribute is a font-family name string or ``None``."""

    heading: Optional[str] = None
    body: Optional[str] = None
    extras: Mapping[str, str] = field(default_factory=dict)

    def __getitem__(self, key: str) -> str:
        named = {"heading", "body"}
        if key in named:
            value = getattr(self, key)
            if value is None:
                raise KeyError(key)
            return value
        return self.extras[key]


@dataclass(frozen=True)
class BrandLogos:
    """Brand logo paths.

    Each attribute is an absolute filesystem path resolved relative to the
    YAML file's directory at load time, or ``None`` if the manifest didn't
    declare it. Pass any of these values straight to
    :meth:`docx.document.Document.add_picture`.
    """

    full_color: Optional[str] = None
    monochrome: Optional[str] = None
    reverse: Optional[str] = None
    extras: Mapping[str, str] = field(default_factory=dict)

    def __getitem__(self, key: str) -> str:
        named = {"full_color", "monochrome", "reverse"}
        if key in named:
            value = getattr(self, key)
            if value is None:
                raise KeyError(key)
            return value
        return self.extras[key]


@dataclass(frozen=True)
class BrandSpacing:
    """Brand spacing values. Each attribute is a :class:`Length` or ``None``."""

    paragraph: Optional[Length] = None
    section: Optional[Length] = None
    extras: Mapping[str, Length] = field(default_factory=dict)

    def __getitem__(self, key: str) -> Length:
        named = {"paragraph", "section"}
        if key in named:
            value = getattr(self, key)
            if value is None:
                raise KeyError(key)
            return value
        return self.extras[key]


@dataclass(frozen=True)
class BrandAssets:
    """A complete brand-asset bundle loaded from a YAML manifest.

    Construct via :meth:`BrandAssets.load` (parses YAML from disk) or
    :meth:`BrandAssets.from_dict` (parses an already-loaded mapping).
    All four sub-views (:attr:`colors`, :attr:`fonts`, :attr:`logos`,
    :attr:`spacing`) are always populated; missing manifest blocks
    yield empty views with every field set to ``None``.
    """

    name: Optional[str] = None
    colors: BrandColors = field(default_factory=BrandColors)
    fonts: BrandFonts = field(default_factory=BrandFonts)
    logos: BrandLogos = field(default_factory=BrandLogos)
    spacing: BrandSpacing = field(default_factory=BrandSpacing)
    source_path: Optional[str] = None

    # -- Loaders -----------------------------------------------------------

    @classmethod
    def load(cls, yaml_path: Union[str, os.PathLike]) -> "BrandAssets":
        """Load a brand-asset manifest from a YAML file.

        Logo paths declared in the file are resolved relative to the
        YAML file's directory, so a brand kit that ships a manifest +
        ``logos/`` directory works regardless of the caller's CWD.

        Raises :class:`BrandAssetsError` if the file is malformed or
        :class:`ImportError` if PyYAML is not installed.
        """

        try:
            import yaml  # type: ignore[import-not-found]
        except ImportError as exc:  # pragma: no cover - exercised only without PyYAML
            raise ImportError(
                "BrandAssets.load() requires PyYAML. Install with "
                "`pip install pyyaml` or `pip install 'python-docx[brand]'`."
            ) from exc

        path = os.fspath(yaml_path)
        with open(path, "r", encoding="utf-8") as f:
            try:
                data = yaml.safe_load(f)
            except yaml.YAMLError as exc:
                raise BrandAssetsError(
                    f"{path}: invalid YAML ({exc})"
                ) from exc

        if data is None:
            data = {}
        if not isinstance(data, _Mapping):
            raise BrandAssetsError(
                f"{path}: top-level YAML node must be a mapping, "
                f"got {type(data).__name__}"
            )

        base_dir = os.path.dirname(os.path.abspath(path))
        return cls.from_dict(data, base_dir=base_dir, source_path=path)

    @classmethod
    def from_dict(
        cls,
        data: Mapping[str, Any],
        *,
        base_dir: Optional[str] = None,
        source_path: Optional[str] = None,
    ) -> "BrandAssets":
        """Build a :class:`BrandAssets` from an already-parsed mapping.

        Use this when the manifest comes from a non-YAML source (a Python
        dict, a TOML / JSON parser, an env-var harness, …). Logo paths
        are resolved relative to *base_dir* when supplied, otherwise
        passed through verbatim.

        Raises :class:`BrandAssetsError` if a block is malformed.
        """

        if not isinstance(data, _Mapping):
            raise BrandAssetsError(
                f"BrandAssets.from_dict() requires a mapping, "
                f"got {type(data).__name__}"
            )

        name = data.get("name")
        if name is not None and not isinstance(name, str):
            raise BrandAssetsError(
                f"name: expected a string, got {type(name).__name__}"
            )

        colors = _build_colors(data.get("colors") or {})
        fonts = _build_fonts(data.get("fonts") or {})
        logos = _build_logos(data.get("logos") or {}, base_dir=base_dir)
        spacing = _build_spacing(data.get("spacing") or {})

        return cls(
            name=name,
            colors=colors,
            fonts=fonts,
            logos=logos,
            spacing=spacing,
            source_path=source_path,
        )


# -- Sub-builders ------------------------------------------------------------


_COLOR_KEYS = ("primary", "secondary", "accent", "background")
_FONT_KEYS = ("heading", "body")
_LOGO_KEYS = ("full_color", "monochrome", "reverse")
_SPACING_KEYS = ("paragraph", "section")


def _ensure_mapping(value: Any, *, where: str) -> Mapping[str, Any]:
    if not isinstance(value, _Mapping):
        raise BrandAssetsError(
            f"{where}: expected a mapping, got {type(value).__name__}"
        )
    return value


def _build_colors(raw: Any) -> BrandColors:
    block = _ensure_mapping(raw, where="colors")
    named = {
        key: _parse_color(block[key], where=f"colors.{key}")
        for key in _COLOR_KEYS
        if key in block and block[key] is not None
    }
    extras = {
        key: _parse_color(block[key], where=f"colors.{key}")
        for key in block
        if key not in _COLOR_KEYS and block[key] is not None
    }
    return BrandColors(**named, extras=extras)


def _build_fonts(raw: Any) -> BrandFonts:
    block = _ensure_mapping(raw, where="fonts")
    named: dict[str, str] = {}
    for key in _FONT_KEYS:
        if key in block and block[key] is not None:
            value = block[key]
            if not isinstance(value, str):
                raise BrandAssetsError(
                    f"fonts.{key}: expected a string, got {type(value).__name__}"
                )
            named[key] = value
    extras: dict[str, str] = {}
    for key, value in block.items():
        if key in _FONT_KEYS or value is None:
            continue
        if not isinstance(value, str):
            raise BrandAssetsError(
                f"fonts.{key}: expected a string, got {type(value).__name__}"
            )
        extras[key] = value
    return BrandFonts(**named, extras=extras)


def _build_logos(raw: Any, *, base_dir: Optional[str]) -> BrandLogos:
    block = _ensure_mapping(raw, where="logos")
    named = {
        key: _resolve_logo_path(block[key], base_dir, where=f"logos.{key}")
        for key in _LOGO_KEYS
        if key in block and block[key] is not None
    }
    extras = {
        key: _resolve_logo_path(block[key], base_dir, where=f"logos.{key}")
        for key in block
        if key not in _LOGO_KEYS and block[key] is not None
    }
    return BrandLogos(**named, extras=extras)


def _build_spacing(raw: Any) -> BrandSpacing:
    block = _ensure_mapping(raw, where="spacing")
    named = {
        key: _parse_length(block[key], where=f"spacing.{key}")
        for key in _SPACING_KEYS
        if key in block and block[key] is not None
    }
    extras = {
        key: _parse_length(block[key], where=f"spacing.{key}")
        for key in block
        if key not in _SPACING_KEYS and block[key] is not None
    }
    return BrandSpacing(**named, extras=extras)


__all__ = [
    "BrandAssets",
    "BrandAssetsError",
    "BrandColors",
    "BrandFonts",
    "BrandLogos",
    "BrandSpacing",
]
