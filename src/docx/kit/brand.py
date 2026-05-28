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

This module also exposes :func:`validate_brand` — a read-only linter
that walks a :class:`Document` and reports drift from the brand palette
(font / colour / logo / heading-style / spacing). See the function
docstring for the rule catalogue.
"""

from __future__ import annotations

import os
from collections.abc import Mapping as _Mapping
from dataclasses import dataclass, field
from typing import (
    TYPE_CHECKING,
    Any,
    Dict,
    Iterable,
    List,
    Mapping,
    Optional,
    Tuple,
    Union,
)

from docx.shared import Cm, Emu, Inches, Length, Mm, Pt, RGBColor, Twips

if TYPE_CHECKING:
    from docx.document import Document
    from docx.shape import InlineShape
    from docx.styles.style import BaseStyle
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run



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

# -- Brand-guideline validator (#91) ----------------------------------------



# ---------------------------------------------------------------------------
# Public dataclass + module-level constants
# ---------------------------------------------------------------------------


SEVERITIES: Tuple[str, ...] = ("info", "warning", "error")
"""The three severity tiers, in ascending order of urgency."""

RULE_IDS: Tuple[str, ...] = (
    "font-not-on-brand",
    "color-not-on-brand",
    "wrong-logo",
    "heading-style-mismatch",
    "inconsistent-spacing",
)
"""The rule identifiers ``validate_brand`` may emit, in declaration order."""

# -- Word's default text colour. When a run's ``font.color.rgb`` is set
# -- explicitly to this value the validator surfaces it as the canonical
# -- "text-default" mismatch the issue example calls out.
_TEXT_DEFAULT_RGB = "#000000"

# -- Built-in heading style names — a paragraph whose style is one of
# -- these triggers ``heading-style-mismatch`` rather than the generic
# -- ``font-not-on-brand`` rule.
_HEADING_STYLE_NAMES = frozenset(
    "Heading %d" % n for n in range(1, 10)
) | {"Title", "Subtitle"}

# -- File-extension hint used by the ``wrong-logo`` heuristic. Common
# -- logo extensions a writer might paste from a corporate asset
# -- library.
_LOGO_FILE_EXTENSIONS = frozenset({".png", ".jpg", ".jpeg", ".svg", ".gif", ".webp"})

# -- Upper bound (bytes) for the "looks like a logo" heuristic. Real
# -- corporate logos are almost always under 200 kB; everything bigger
# -- is more likely a hero image or a screenshot.
_LOGO_SIZE_HINT = 200 * 1024


@dataclass(frozen=True)
class BrandFinding:
    """A single brand-guideline violation surfaced by :func:`validate_brand`.

    Attributes
    ----------
    severity
        One of ``"info"`` / ``"warning"`` / ``"error"``. The intended
        publication-blocking threshold is caller-defined.
    location
        Human-readable locator, e.g. ``"paragraph 5"``,
        ``"paragraph 12 run 0"``, or ``"image 'logo.jpg'"``. Designed
        to be printed verbatim — the validator avoids opaque XPath.
    rule
        One of :data:`RULE_IDS`.
    message
        Human-readable explanation of the drift, including (where
        relevant) the offending value and the closest brand suggestion.
    """

    severity: str
    location: str
    rule: str
    message: str


# ---------------------------------------------------------------------------
# Rules normalisation
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _Rules:
    """Normalised brand rules. Internal — callers construct via :func:`_load_rules`."""

    fonts: Tuple[str, ...]
    colors: Mapping[str, str]  # hex (uppercase, leading #) -> human name
    logos: Tuple[Mapping[str, Any], ...]
    spacing: Mapping[str, Any]


def _normalise_hex(value: str) -> str:
    """Return ``value`` as ``"#RRGGBB"`` (uppercase) for cross-source compares."""
    s = value.strip()
    if not s.startswith("#"):
        s = "#" + s
    return s.upper()


def _normalise_color_palette(
    raw,  # type: Any
):
    # type: (...) -> Dict[str, str]
    """Return the colour palette as ``{hex: name}`` regardless of input shape."""
    if raw is None:
        return {}
    if isinstance(raw, Mapping):
        return {_normalise_hex(k): str(v) for k, v in raw.items()}
    if isinstance(raw, Iterable) and not isinstance(raw, (str, bytes)):
        # -- list of hex strings; auto-name them by their position
        return {_normalise_hex(item): str(item) for item in raw}
    raise TypeError(
        "colors must be a mapping (hex -> name) or an iterable of hex strings"
    )


def _normalise_font_list(raw):
    # type: (Any) -> Tuple[str, ...]
    if raw is None:
        return ()
    if isinstance(raw, str):
        return (raw,)
    if isinstance(raw, Iterable):
        return tuple(str(item) for item in raw)
    raise TypeError("fonts must be a string or an iterable of strings")


def _normalise_logo_list(raw):
    # type: (Any) -> Tuple[Mapping[str, Any], ...]
    if raw is None:
        return ()
    out: List[Mapping[str, Any]] = []
    for item in raw:
        if isinstance(item, Mapping):
            out.append(dict(item))
        elif isinstance(item, str):
            # -- bare string treated as a sha1 OR a filename, not both
            if len(item) == 40 and all(c in "0123456789abcdefABCDEF" for c in item):
                out.append({"sha1": item.lower()})
            else:
                out.append({"path": item})
        else:
            raise TypeError(
                "each logo entry must be a string or a mapping; got %r" % (item,)
            )
    return tuple(out)


def _load_yaml(path):
    # type: (Union[str, os.PathLike]) -> Mapping[str, Any]
    """Read ``path`` and return the deserialised mapping.

    Tries :mod:`yaml` first; if PyYAML is not installed, falls back to
    the small built-in subset parser. Raises ``FileNotFoundError`` on a
    missing file (caller decides whether to swallow).
    """
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    try:
        import yaml  # type: ignore
    except ImportError:  # pragma: no cover - environment-specific
        return _yaml_subset_parse(text)
    data = yaml.safe_load(text)
    if data is None:
        return {}
    if not isinstance(data, Mapping):
        raise TypeError("brand rules YAML must deserialise to a mapping")
    return data


def _yaml_subset_parse(text):
    # type: (str) -> Mapping[str, Any]
    """A *very* small YAML subset parser — enough for brand-rule files.

    Supports two-level nesting: top-level mapping whose values are
    scalars, lists of scalars, or lists of single-line mappings.
    Anything more elaborate raises so the caller knows to install
    PyYAML for real coverage.
    """
    out: Dict[str, Any] = {}
    current_key: Optional[str] = None
    current_list: Optional[List[Any]] = None
    current_dict: Optional[Dict[str, Any]] = None
    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        if not line or line.lstrip().startswith("#"):
            continue
        if not line[0].isspace():  # top-level key
            if ":" not in line:
                raise ValueError("malformed YAML line: %r" % (raw_line,))
            current_dict = None
            key, _, rest = line.partition(":")
            key = key.strip()
            rest = rest.strip()
            if rest == "":
                current_key = key
                current_list = []
                out[key] = current_list
            else:
                current_key = None
                current_list = None
                out[key] = _coerce_scalar(rest)
        else:
            stripped = line.lstrip()
            if stripped.startswith("- "):
                if current_list is None:
                    raise ValueError("list item without parent key: %r" % (raw_line,))
                payload = stripped[2:].strip()
                if ":" in payload:
                    # -- list-of-mappings: open a new dict and seed it
                    k, _, v = payload.partition(":")
                    current_dict = {_strip_yaml_quotes(k.strip()): _coerce_scalar(v.strip())}
                    current_list.append(current_dict)
                else:
                    current_dict = None
                    current_list.append(_coerce_scalar(payload))
            elif current_dict is not None and ":" in stripped:
                k, _, v = stripped.partition(":")
                current_dict[_strip_yaml_quotes(k.strip())] = _coerce_scalar(v.strip())
            elif current_key is not None and ":" in stripped:
                # -- nested mapping (e.g. spacing: \n  line_spacing: 1.15)
                if not isinstance(out[current_key], dict):
                    out[current_key] = {}
                k, _, v = stripped.partition(":")
                out[current_key][_strip_yaml_quotes(k.strip())] = _coerce_scalar(v.strip())
            else:
                raise ValueError("unsupported YAML construct: %r" % (raw_line,))
    return out


def _strip_yaml_quotes(text):
    # type: (str) -> str
    """Strip a single matching pair of ``"`` or ``'`` around ``text``."""
    if len(text) >= 2 and text[0] == text[-1] and text[0] in ('"', "'"):
        return text[1:-1]
    return text


def _coerce_scalar(text):
    # type: (str) -> Any
    """Coerce a YAML scalar to int / float / bool / None / str."""
    s = text.strip()
    if (s.startswith('"') and s.endswith('"')) or (
        s.startswith("'") and s.endswith("'")
    ):
        return s[1:-1]
    lower = s.lower()
    if lower in ("null", "~", ""):
        return None
    if lower == "true":
        return True
    if lower == "false":
        return False
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        pass
    return s


def _coerce_to_mapping(rules):
    # type: (Any) -> Mapping[str, Any]
    """Pull a plain mapping out of a BrandAssets-shaped duck-typed object."""
    out: Dict[str, Any] = {}
    for key in ("fonts", "colors", "logos", "spacing"):
        if hasattr(rules, key):
            out[key] = getattr(rules, key)
    if not out:
        # -- not a BrandAssets-shaped instance — let the caller's
        # -- coercion helper raise so the error mentions the actual type.
        raise TypeError(
            "rules must be a path, a mapping, or a BrandAssets-like object "
            "exposing fonts / colors / logos / spacing; got %r" % (type(rules).__name__,)
        )
    return out


def _load_rules(rules):
    # type: (Any) -> _Rules
    """Coerce ``rules`` (path / dict / BrandAssets) into a :class:`_Rules`."""
    if isinstance(rules, (str, os.PathLike)):
        data = _load_yaml(rules)
    elif isinstance(rules, Mapping):
        data = rules
    else:
        data = _coerce_to_mapping(rules)
    return _Rules(
        fonts=_normalise_font_list(data.get("fonts")),
        colors=_normalise_color_palette(data.get("colors")),
        logos=_normalise_logo_list(data.get("logos")),
        spacing=dict(data.get("spacing") or {}),
    )


# ---------------------------------------------------------------------------
# Per-rule checkers
# ---------------------------------------------------------------------------


def _resolved_font_name(paragraph):
    # type: (Paragraph) -> Optional[str]
    """Return the paragraph's *effective* font name, walking the style chain."""
    style = paragraph.style
    return _style_font_name(style)


def _style_font_name(style):
    # type: (Optional[BaseStyle]) -> Optional[str]
    seen: List[int] = []
    while style is not None and id(style) not in seen:
        seen.append(id(style))
        font = getattr(style, "font", None)
        if font is not None and font.name:
            return font.name
        style = getattr(style, "base_style", None)
    return None


def _style_font_color(style):
    # type: (Optional[BaseStyle]) -> Optional[str]
    """Return the style chain's effective text colour as ``"#RRGGBB"`` or |None|."""
    seen: List[int] = []
    while style is not None and id(style) not in seen:
        seen.append(id(style))
        font = getattr(style, "font", None)
        color = getattr(font, "color", None) if font is not None else None
        rgb = getattr(color, "rgb", None)
        if rgb is not None:
            return "#%s" % (str(rgb).upper(),)
        style = getattr(style, "base_style", None)
    return None


def _is_heading_style(style):
    # type: (Optional[BaseStyle]) -> bool
    if style is None:
        return False
    name = getattr(style, "name", None)
    return name in _HEADING_STYLE_NAMES


def _check_paragraph_font(
    paragraph,  # type: Paragraph
    index,  # type: int
    rules,  # type: _Rules
    findings,  # type: List[BrandFinding]
):
    # type: (...) -> None
    """Surface ``font-not-on-brand`` and ``heading-style-mismatch`` findings."""
    if not rules.fonts:
        return
    is_heading = _is_heading_style(paragraph.style)
    style_font = _resolved_font_name(paragraph)
    if style_font is not None and style_font not in rules.fonts:
        rule = "heading-style-mismatch" if is_heading else "font-not-on-brand"
        findings.append(
            BrandFinding(
                severity="warning",
                location="paragraph %d" % (index,),
                rule=rule,
                message=(
                    "%s style %r uses font %r (allowed: %s)"
                    % (
                        "heading" if is_heading else "paragraph",
                        paragraph.style.name if paragraph.style else "Normal",
                        style_font,
                        ", ".join(rules.fonts),
                    )
                ),
            )
        )
    for run_index, run in enumerate(paragraph.runs):
        run_font = run.font.name
        if run_font is None or run_font in rules.fonts:
            continue
        rule = "heading-style-mismatch" if is_heading else "font-not-on-brand"
        findings.append(
            BrandFinding(
                severity="warning",
                location="paragraph %d run %d" % (index, run_index),
                rule=rule,
                message=(
                    "run uses font %r (allowed: %s)"
                    % (run_font, ", ".join(rules.fonts))
                ),
            )
        )


def _check_paragraph_color(
    paragraph,  # type: Paragraph
    index,  # type: int
    rules,  # type: _Rules
    findings,  # type: List[BrandFinding]
):
    # type: (...) -> None
    """Surface ``color-not-on-brand`` findings for run-level colour overrides."""
    if not rules.colors:
        return
    style_color = _style_font_color(paragraph.style)
    if (
        style_color is not None
        and style_color not in rules.colors
        and not _all_runs_override_color(paragraph)
    ):
        findings.append(
            BrandFinding(
                severity="warning",
                location="paragraph %d" % (index,),
                rule="color-not-on-brand",
                message=_color_message(style_color, rules),
            )
        )
    for run_index, run in enumerate(paragraph.runs):
        rgb = run.font.color.rgb
        if rgb is None:
            continue
        hex_value = "#%s" % (str(rgb).upper(),)
        if hex_value in rules.colors:
            continue
        findings.append(
            BrandFinding(
                severity="warning",
                location="paragraph %d run %d" % (index, run_index),
                rule="color-not-on-brand",
                message=_color_message(hex_value, rules),
            )
        )


def _all_runs_override_color(paragraph):
    # type: (Paragraph) -> bool
    """Return True when every run sets its own ``font.color.rgb`` (so the style colour is moot)."""
    runs = list(paragraph.runs)
    if not runs:
        return False
    return all(run.font.color.rgb is not None for run in runs)


def _color_message(hex_value, rules):
    # type: (str, _Rules) -> str
    """Build the ``color-not-on-brand`` message, mentioning the closest brand suggestion."""
    upper = hex_value.upper()
    if upper == _TEXT_DEFAULT_RGB and rules.colors:
        # -- preferred: pick the *first* declared brand colour as the
        # -- canonical "use this instead" suggestion. Mirrors the
        # -- example in the issue body (``Squid Ink``).
        first_hex, first_name = next(iter(rules.colors.items()))
        return (
            "%r is text-default but brand requires %r (%s)"
            % (upper, first_hex, first_name)
        )
    allowed = ", ".join(
        "%s (%s)" % (h, name) for h, name in rules.colors.items()
    )
    return "color %r is not on brand (allowed: %s)" % (upper, allowed)


def _check_paragraph_spacing(
    paragraph,  # type: Paragraph
    index,  # type: int
    rules,  # type: _Rules
    findings,  # type: List[BrandFinding]
):
    # type: (...) -> None
    """Surface ``inconsistent-spacing`` findings for paragraph-format drift."""
    if not rules.spacing:
        return
    fmt = paragraph.paragraph_format
    drifts: List[str] = []
    expected_ls = rules.spacing.get("line_spacing")
    if expected_ls is not None and fmt.line_spacing is not None:
        if not _approx_equal(fmt.line_spacing, expected_ls):
            drifts.append(
                "line_spacing %s != %s" % (fmt.line_spacing, expected_ls)
            )
    expected_before = _coerce_points(rules.spacing.get("space_before"))
    if expected_before is not None and fmt.space_before is not None:
        actual_before = fmt.space_before.pt
        if not _approx_equal(actual_before, expected_before):
            drifts.append(
                "space_before %.1fpt != %.1fpt" % (actual_before, expected_before)
            )
    expected_after = _coerce_points(rules.spacing.get("space_after"))
    if expected_after is not None and fmt.space_after is not None:
        actual_after = fmt.space_after.pt
        if not _approx_equal(actual_after, expected_after):
            drifts.append(
                "space_after %.1fpt != %.1fpt" % (actual_after, expected_after)
            )
    if drifts:
        findings.append(
            BrandFinding(
                severity="info",
                location="paragraph %d" % (index,),
                rule="inconsistent-spacing",
                message="paragraph spacing drifts from brand: " + "; ".join(drifts),
            )
        )


def _approx_equal(a, b, eps=0.05):
    # type: (float, float, float) -> bool
    return abs(float(a) - float(b)) <= eps


def _coerce_points(value):
    # type: (Any) -> Optional[float]
    """Coerce a brand-spacing scalar to a float number of points (or |None|)."""
    if value is None:
        return None
    if isinstance(value, bool):  # pragma: no cover - defensive
        return None
    return float(value)


def _check_inline_shape(
    shape,  # type: InlineShape
    rules,  # type: _Rules
    findings,  # type: List[BrandFinding]
):
    # type: (...) -> None
    """Surface ``wrong-logo`` findings for unauthorised logo-shaped images."""
    image = getattr(shape, "image", None)
    if image is None:
        return
    filename = getattr(image, "filename", None) or "(unnamed)"
    blob_size = len(getattr(image, "blob", b""))
    sha1 = getattr(image, "sha1", None)
    ext = os.path.splitext(filename)[1].lower()
    looks_like_logo = (
        ext in _LOGO_FILE_EXTENSIONS and 0 < blob_size <= _LOGO_SIZE_HINT
    )
    if not looks_like_logo:
        return
    if _is_known_brand_logo(filename, sha1, rules.logos):
        return
    findings.append(
        BrandFinding(
            severity="error",
            location="image %r" % (filename,),
            rule="wrong-logo",
            message=(
                "image looks like a logo (%s, %d bytes) but is not in "
                "brand.logos — possible competitor or unauthorised asset"
                % (ext, blob_size)
            ),
        )
    )


def _is_known_brand_logo(filename, sha1, registered):
    # type: (str, Optional[str], Tuple[Mapping[str, Any], ...]) -> bool
    """Return True when ``filename`` or ``sha1`` matches a registered brand logo."""
    base = os.path.basename(filename or "").lower()
    sha = (sha1 or "").lower()
    for entry in registered:
        registered_path = str(entry.get("path", "") or "").lower()
        registered_sha = str(entry.get("sha1", "") or "").lower()
        if registered_sha and sha and registered_sha == sha:
            return True
        if registered_path and base and os.path.basename(registered_path) == base:
            return True
    return False


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def validate_brand(document, rules):
    # type: (Document, Any) -> List[BrandFinding]
    """Return a list of :class:`BrandFinding` records describing brand drift.

    Walks ``document`` paragraph-by-paragraph applying the five rules
    described in the module docstring. The validator is read-only and
    never raises on a finding — callers decide what severity threshold
    to treat as publication-blocking.

    Parameters
    ----------
    document
        The :class:`Document` to validate.
    rules
        The brand palette. Accepts:

        - a path-like (``str`` / :class:`os.PathLike`) to a YAML file;
        - a pre-loaded :class:`dict` matching the YAML schema;
        - a :class:`BrandAssets` instance (issue #90; duck-typed —
          anything exposing ``fonts``, ``colors``, ``logos``,
          ``spacing`` works).

    Returns
    -------
    list[BrandFinding]
        Findings in document order, each carrying ``severity`` /
        ``location`` / ``rule`` / ``message``. An empty list means the
        document is fully on brand against the supplied rules.

    .. versionadded:: 2026.05.29
    """
    normalised = _load_rules(rules)
    findings: List[BrandFinding] = []

    for index, paragraph in enumerate(document.paragraphs):
        _check_paragraph_font(paragraph, index, normalised, findings)
        _check_paragraph_color(paragraph, index, normalised, findings)
        _check_paragraph_spacing(paragraph, index, normalised, findings)

    for shape in document.inline_shapes:
        _check_inline_shape(shape, normalised, findings)

    return findings

__all__ = [
    "BrandAssets",
    "BrandAssetsError",
    "BrandColors",
    "BrandFinding",
    "BrandFonts",
    "BrandLogos",
    "BrandSpacing",
    "RULE_IDS",
    "SEVERITIES",
    "validate_brand",
]
