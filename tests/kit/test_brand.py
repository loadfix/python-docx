"""Unit-test suite for ``docx.kit.brand`` (issue #90)."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from docx.kit.brand import (
    BrandAssets,
    BrandAssetsError,
    BrandColors,
    BrandFonts,
    BrandLogos,
    BrandSpacing,
)
from docx.shared import Cm, Inches, Length, Mm, Pt, RGBColor, Twips


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------


_AWS_YAML = """\
name: AWS
colors:
  primary: '#FF9900'
  secondary: '#232F3E'
  accent: '#0073BB'
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
"""


@pytest.fixture
def aws_brand_yaml(tmp_path: Path) -> Path:
    yaml_path = tmp_path / "aws-brand.yaml"
    yaml_path.write_text(_AWS_YAML, encoding="utf-8")
    return yaml_path


@pytest.fixture
def aws_brand(aws_brand_yaml: Path) -> BrandAssets:
    return BrandAssets.load(aws_brand_yaml)


# --------------------------------------------------------------------------
# load() -- end-to-end path
# --------------------------------------------------------------------------


class DescribeBrandAssetsLoad:
    """Acceptance-level tests for :meth:`BrandAssets.load`."""

    def it_returns_a_BrandAssets_instance(self, aws_brand: BrandAssets):
        assert isinstance(aws_brand, BrandAssets)

    def it_records_the_brand_name(self, aws_brand: BrandAssets):
        assert aws_brand.name == "AWS"

    def it_records_the_source_path(
        self, aws_brand: BrandAssets, aws_brand_yaml: Path
    ):
        assert aws_brand.source_path == str(aws_brand_yaml)

    def it_parses_colors_into_RGBColor_instances(self, aws_brand: BrandAssets):
        assert aws_brand.colors.primary == RGBColor(0xFF, 0x99, 0x00)
        assert aws_brand.colors.secondary == RGBColor(0x23, 0x2F, 0x3E)
        assert aws_brand.colors.accent == RGBColor(0x00, 0x73, 0xBB)
        assert aws_brand.colors.background == RGBColor(0xFA, 0xFA, 0xFA)

    def it_exposes_font_names_as_strings(self, aws_brand: BrandAssets):
        assert aws_brand.fonts.heading == "Amazon Ember Display"
        assert aws_brand.fonts.body == "Amazon Ember"

    def it_resolves_relative_logo_paths_against_the_yaml_directory(
        self, aws_brand: BrandAssets, aws_brand_yaml: Path
    ):
        expected_dir = str(aws_brand_yaml.parent)
        assert aws_brand.logos.full_color == os.path.normpath(
            os.path.join(expected_dir, "logos/aws-full-color.png")
        )
        assert os.path.isabs(aws_brand.logos.monochrome)
        assert os.path.isabs(aws_brand.logos.reverse)

    def it_parses_spacing_into_Length_instances(self, aws_brand: BrandAssets):
        assert isinstance(aws_brand.spacing.paragraph, Length)
        assert isinstance(aws_brand.spacing.section, Length)
        assert aws_brand.spacing.paragraph == Pt(12)
        assert aws_brand.spacing.section == Pt(24)

    def it_accepts_a_PathLike_argument(
        self, aws_brand_yaml: Path
    ):
        # Ensure load() accepts both ``str`` and ``pathlib.Path``.
        brand_str = BrandAssets.load(str(aws_brand_yaml))
        brand_path = BrandAssets.load(aws_brand_yaml)

        assert brand_str.name == brand_path.name == "AWS"

    def it_is_an_immutable_dataclass(self, aws_brand: BrandAssets):
        with pytest.raises((AttributeError, TypeError)):
            aws_brand.name = "Other"  # type: ignore[misc]


# --------------------------------------------------------------------------
# from_dict() -- programmatic construction
# --------------------------------------------------------------------------


class DescribeBrandAssetsFromDict:
    """Tests for :meth:`BrandAssets.from_dict`."""

    def it_builds_from_an_in_memory_mapping(self):
        data = {
            "name": "Acme",
            "colors": {"primary": "#FF0000"},
            "fonts": {"heading": "Helvetica"},
            "logos": {"full_color": "/abs/path/logo.png"},
            "spacing": {"paragraph": "10pt"},
        }

        brand = BrandAssets.from_dict(data)

        assert brand.name == "Acme"
        assert brand.colors.primary == RGBColor(0xFF, 0x00, 0x00)
        assert brand.fonts.heading == "Helvetica"
        assert brand.logos.full_color == "/abs/path/logo.png"
        assert brand.spacing.paragraph == Pt(10)

    def it_returns_empty_views_for_omitted_blocks(self):
        brand = BrandAssets.from_dict({})

        assert brand.name is None
        assert brand.colors == BrandColors()
        assert brand.fonts == BrandFonts()
        assert brand.logos == BrandLogos()
        assert brand.spacing == BrandSpacing()
        assert brand.colors.primary is None
        assert brand.fonts.heading is None
        assert brand.logos.full_color is None
        assert brand.spacing.paragraph is None

    def it_anchors_relative_logo_paths_to_base_dir_when_supplied(self):
        brand = BrandAssets.from_dict(
            {"logos": {"full_color": "logo.png"}},
            base_dir="/home/brand",
        )
        assert brand.logos.full_color == os.path.normpath(
            "/home/brand/logo.png"
        )

    def it_passes_relative_paths_through_when_base_dir_is_None(self):
        brand = BrandAssets.from_dict(
            {"logos": {"full_color": "logo.png"}}
        )
        assert brand.logos.full_color == "logo.png"

    def it_keeps_absolute_logo_paths_as_is(self):
        brand = BrandAssets.from_dict(
            {"logos": {"full_color": "/etc/logo.png"}},
            base_dir="/home/brand",
        )
        assert brand.logos.full_color == "/etc/logo.png"

    def it_collects_unknown_color_keys_into_extras(self):
        brand = BrandAssets.from_dict(
            {"colors": {"primary": "#FF0000", "warning": "#F39C12"}}
        )
        assert brand.colors.primary == RGBColor(0xFF, 0x00, 0x00)
        assert brand.colors.extras["warning"] == RGBColor(0xF3, 0x9C, 0x12)
        # Subscript also reaches both named and extras.
        assert brand.colors["primary"] == RGBColor(0xFF, 0x00, 0x00)
        assert brand.colors["warning"] == RGBColor(0xF3, 0x9C, 0x12)

    def it_collects_unknown_logo_keys_into_extras(self):
        brand = BrandAssets.from_dict(
            {"logos": {"full_color": "/a/full.png", "favicon": "/a/fav.ico"}}
        )
        assert brand.logos.full_color == "/a/full.png"
        assert brand.logos.extras["favicon"] == "/a/fav.ico"

    def it_collects_unknown_font_keys_into_extras(self):
        brand = BrandAssets.from_dict(
            {"fonts": {"heading": "Helvetica", "code": "Source Code Pro"}}
        )
        assert brand.fonts.heading == "Helvetica"
        assert brand.fonts.extras["code"] == "Source Code Pro"

    def it_collects_unknown_spacing_keys_into_extras(self):
        brand = BrandAssets.from_dict(
            {"spacing": {"paragraph": "10pt", "gutter": "0.25in"}}
        )
        assert brand.spacing.paragraph == Pt(10)
        assert brand.spacing.extras["gutter"] == Inches(0.25)


# --------------------------------------------------------------------------
# colour parsing
# --------------------------------------------------------------------------


class DescribeColorParsing:
    """Edge cases on the hex-string / RGB-triple colour parser."""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("#FF9900", RGBColor(0xFF, 0x99, 0x00)),
            ("FF9900", RGBColor(0xFF, 0x99, 0x00)),
            ("#ff9900", RGBColor(0xFF, 0x99, 0x00)),
            ("ff9900", RGBColor(0xFF, 0x99, 0x00)),
            ("#F90", RGBColor(0xFF, 0x99, 0x00)),
            ("F90", RGBColor(0xFF, 0x99, 0x00)),
        ],
    )
    def it_parses_hex_strings(self, raw: str, expected: RGBColor):
        brand = BrandAssets.from_dict({"colors": {"primary": raw}})
        assert brand.colors.primary == expected

    def it_accepts_a_3_int_list(self):
        brand = BrandAssets.from_dict({"colors": {"primary": [255, 153, 0]}})
        assert brand.colors.primary == RGBColor(0xFF, 0x99, 0x00)

    def it_passes_through_an_RGBColor_instance(self):
        existing = RGBColor(0x12, 0x34, 0x56)
        brand = BrandAssets.from_dict({"colors": {"primary": existing}})
        assert brand.colors.primary == existing

    @pytest.mark.parametrize(
        "bad",
        [
            "#GGHHII",  # not hex
            "#1234",     # neither 3 nor 6 chars
            "12345",     # neither 3 nor 6 chars
            42,          # not a string / list
        ],
    )
    def it_raises_on_malformed_color(self, bad: object):
        with pytest.raises(BrandAssetsError):
            BrandAssets.from_dict({"colors": {"primary": bad}})


# --------------------------------------------------------------------------
# length parsing
# --------------------------------------------------------------------------


class DescribeLengthParsing:
    """Edge cases on the spacing/length string parser."""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("12pt", Pt(12)),
            ("12 pt", Pt(12)),
            ("0.5in", Inches(0.5)),
            ("24mm", Mm(24)),
            ("3cm", Cm(3)),
            ("914400emu", Length(914400)),
            ("720twips", Twips(720)),
            ("720twip", Twips(720)),
        ],
    )
    def it_parses_suffixed_strings(self, raw: str, expected: Length):
        brand = BrandAssets.from_dict({"spacing": {"paragraph": raw}})
        assert brand.spacing.paragraph == expected

    def it_treats_bare_numbers_as_points(self):
        brand = BrandAssets.from_dict({"spacing": {"paragraph": 18}})
        assert brand.spacing.paragraph == Pt(18)

    def it_treats_bare_numeric_strings_as_points(self):
        brand = BrandAssets.from_dict({"spacing": {"paragraph": "18"}})
        assert brand.spacing.paragraph == Pt(18)

    def it_passes_a_Length_through_unchanged(self):
        brand = BrandAssets.from_dict({"spacing": {"paragraph": Pt(7)}})
        assert brand.spacing.paragraph == Pt(7)

    @pytest.mark.parametrize(
        "bad",
        ["", "pt", "abc", True, [12]],
    )
    def it_raises_on_malformed_length(self, bad: object):
        with pytest.raises(BrandAssetsError):
            BrandAssets.from_dict({"spacing": {"paragraph": bad}})


# --------------------------------------------------------------------------
# error paths
# --------------------------------------------------------------------------


class DescribeBrandAssetsErrors:
    """Negative cases — malformed manifests should surface clean errors."""

    def it_raises_on_a_non_mapping_top_level(self, tmp_path: Path):
        bad_yaml = tmp_path / "bad.yaml"
        bad_yaml.write_text("- just\n- a\n- list\n", encoding="utf-8")
        with pytest.raises(BrandAssetsError):
            BrandAssets.load(bad_yaml)

    def it_raises_on_invalid_yaml(self, tmp_path: Path):
        bad_yaml = tmp_path / "bad.yaml"
        bad_yaml.write_text(": : :\n  -broken\n", encoding="utf-8")
        with pytest.raises(BrandAssetsError):
            BrandAssets.load(bad_yaml)

    def it_treats_an_empty_yaml_file_as_an_empty_brand(self, tmp_path: Path):
        empty_yaml = tmp_path / "empty.yaml"
        empty_yaml.write_text("", encoding="utf-8")
        brand = BrandAssets.load(empty_yaml)
        assert brand.name is None
        assert brand.colors.primary is None

    def it_raises_on_non_string_name(self):
        with pytest.raises(BrandAssetsError):
            BrandAssets.from_dict({"name": 42})

    def it_raises_on_non_mapping_colors_block(self):
        with pytest.raises(BrandAssetsError):
            BrandAssets.from_dict({"colors": "not-a-mapping"})

    def it_raises_on_non_string_font_value(self):
        with pytest.raises(BrandAssetsError):
            BrandAssets.from_dict({"fonts": {"heading": 12}})

    def it_raises_on_non_string_logo_value(self):
        with pytest.raises(BrandAssetsError):
            BrandAssets.from_dict({"logos": {"full_color": 12}})

    def it_raises_on_empty_logo_path(self):
        with pytest.raises(BrandAssetsError):
            BrandAssets.from_dict({"logos": {"full_color": "  "}})

    def it_raises_on_non_mapping_argument_to_from_dict(self):
        with pytest.raises(BrandAssetsError):
            BrandAssets.from_dict("not-a-dict")  # type: ignore[arg-type]


# --------------------------------------------------------------------------
# subscript access
# --------------------------------------------------------------------------


class DescribeViewSubscript:
    """The four sub-views support ``view[key]`` for both named + extras."""

    def it_raises_KeyError_for_unset_named_field(self):
        brand = BrandAssets.from_dict({"colors": {"secondary": "#000000"}})
        with pytest.raises(KeyError):
            _ = brand.colors["primary"]

    def it_raises_KeyError_for_unknown_extras(self):
        brand = BrandAssets.from_dict({})
        with pytest.raises(KeyError):
            _ = brand.colors["nonsense"]
