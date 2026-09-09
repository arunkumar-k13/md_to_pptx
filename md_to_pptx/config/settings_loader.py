"""Configuration Loader Module.

Parses external settings.yaml files into strongly-typed dataclasses. Ensures zero
hardcoded thresholds or preferences across the application.
"""

from __future__ import annotations
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Union

import yaml

logger = logging.getLogger(__name__)


@dataclass(slots=True)
class OverflowSettings:
    """Overflow and text budgeting policies.

    Attributes:
        max_bullets_per_slide: Soft cap for bullet list items per slide.
        max_paragraph_words_per_slide: Soft cap for total paragraph words per slide.
        min_font_size_pt: Minimum allowable font size in points during auto-scaling.
        font_scaling_enabled: Allow font size scaling down to min_font_size_pt.
        enable_automatic_continuation: Automatically split overflow onto continuation slides.
    """

    max_bullets_per_slide: int = 10
    max_paragraph_words_per_slide: int = 200
    min_font_size_pt: int = 14
    font_scaling_enabled: bool = True
    enable_automatic_continuation: bool = True


@dataclass(slots=True)
class PresentationSettings:
    """Presentation structural and textual defaults.

    Attributes:
        continuation_suffix: Suffix appended to continuation slide titles.
        default_thank_you_title: Title text for auto-generated end slides.
        auto_detect_section_dividers: Create section divider slides for major H1 sections.
    """

    continuation_suffix: str = "(Continued)"
    default_thank_you_title: str = "Thank You"
    auto_detect_section_dividers: bool = True


@dataclass(slots=True)
class ImageSettings:
    """Image rendering and scaling options.

    Attributes:
        keep_aspect_ratio: Maintain original image width/height aspect ratio.
        max_width_percentage: Maximum allowable slide width fraction (0.0 to 1.0).
        max_height_percentage: Maximum allowable slide height fraction (0.0 to 1.0).
    """

    keep_aspect_ratio: bool = True
    max_width_percentage: float = 0.90
    max_height_percentage: float = 0.85


@dataclass(slots=True)
class ExportSettings:
    """PDF export preferences and environment paths.

    Attributes:
        preferred_pdf_exporter: Target exporter plugin ('com' or 'libreoffice').
        libreoffice_binary_path: Executable path for headless LibreOffice conversion.
    """

    preferred_pdf_exporter: str = "com"
    libreoffice_binary_path: str = "soffice"


@dataclass(slots=True)
class TemplateDefaultsSettings:
    """Fallback layout names when dynamic template resolution matches multiple targets.

    Attributes:
        fallback_title_layout_name: Default title slide layout name.
        fallback_content_layout_name: Default body content layout name.
        header_top_margin_inches: Fallback header top margin in inches.
        footer_bottom_margin_inches: Fallback footer bottom margin in inches.
        content_padding_inches: Content safety padding in inches.
    """

    fallback_title_layout_name: str = "Title Slide"
    fallback_content_layout_name: str = "Title and Content"
    header_top_margin_inches: float = 1.0
    footer_bottom_margin_inches: float = 0.6
    content_padding_inches: float = 0.2


@dataclass(slots=True)
class BrandSlideSettings:
    """Brand Slide Preservation Configuration.

    Attributes:
        enabled: Enable Brand Slide Manager workflow.
        cover: Preserve and populate BRAND_COVER slide.
        disclaimer: Preserve and populate BRAND_DISCLAIMER slide.
        about_company: Preserve and populate BRAND_ABOUT_COMPANY slide.
        case_study: Preserve and populate BRAND_CASE_STUDY slide.
        thank_you: Preserve and populate BRAND_THANK_YOU slide.
    """

    enabled: bool = True
    cover: bool = True
    disclaimer: bool = True
    about_company: bool = True
    case_study: bool = True
    thank_you: bool = True


@dataclass(slots=True)
class Settings:
    """Root Application Configuration Container."""

    overflow: OverflowSettings = field(default_factory=OverflowSettings)
    presentation: PresentationSettings = field(default_factory=PresentationSettings)
    images: ImageSettings = field(default_factory=ImageSettings)
    export: ExportSettings = field(default_factory=ExportSettings)
    template_defaults: TemplateDefaultsSettings = field(default_factory=TemplateDefaultsSettings)
    brand_slides: BrandSlideSettings = field(default_factory=BrandSlideSettings)


def get_default_config_path() -> Path:
    """Get default absolute path to config/settings.yaml.

    Returns:
        Path instance pointing to settings.yaml.
    """
    return Path(__file__).parent / "settings.yaml"


def load_settings(config_path: Optional[Union[str, Path]] = None) -> Settings:
    """Load configuration settings from a YAML file or fallback to defaults.

    Args:
        config_path: Optional path to settings.yaml file.

    Returns:
        Populated Settings instance.
    """
    target_path = Path(config_path) if config_path else get_default_config_path()

    if not target_path.is_file():
        logger.warning("Config file not found at %s. Returning default settings.", target_path)
        return Settings()

    try:
        raw_content = target_path.read_text(encoding="utf-8")
        data = yaml.safe_load(raw_content) or {}

        overflow_data = data.get("overflow", {})
        pres_data = data.get("presentation", {})
        img_data = data.get("images", {})
        export_data = data.get("export", {})
        tmpl_data = data.get("template_defaults", {})
        brand_data = data.get("brand_slides", {})

        return Settings(
            overflow=OverflowSettings(**overflow_data),
            presentation=PresentationSettings(**pres_data),
            images=ImageSettings(**img_data),
            export=ExportSettings(**export_data),
            template_defaults=TemplateDefaultsSettings(**tmpl_data),
            brand_slides=BrandSlideSettings(**brand_data),
        )
    except Exception as err:
        logger.error("Failed to load settings from %s: %s. Using default settings.", target_path, err, exc_info=True)
        return Settings()
