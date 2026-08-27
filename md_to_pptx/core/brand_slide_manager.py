"""Brand Slide Manager Module.

Recognizes, structures, and injects corporate brand slides (BRAND_COVER, BRAND_DISCLAIMER,
BRAND_ABOUT_COMPANY, BRAND_CASE_STUDY, BRAND_THANK_YOU) from template introspection into the
Presentation domain model.
"""

from __future__ import annotations
import logging
from typing import List, Optional

from md_to_pptx.config.settings_loader import Settings
from md_to_pptx.core.presentation_model import Presentation, Slide, SlideIntent, TitleBlock
from md_to_pptx.core.template.template_analyzer import TemplateMetadata

logger = logging.getLogger(__name__)


class BrandSlideManager:
    """Manager responsible for inserting and structuring brand slides in presentations."""

    def process_presentation(
        self,
        presentation: Presentation,
        template_meta: TemplateMetadata,
        settings: Optional[Settings] = None,
    ) -> Presentation:
        """Process Presentation IR model to inject or designate brand slides.

        Args:
            presentation: Input planned Presentation instance.
            template_meta: Introspected TemplateMetadata instance.
            settings: Optional Settings configuration instance.

        Returns:
            Updated Presentation instance with brand slides designated or injected.
        """
        cfg = settings or Settings()
        b_cfg = cfg.brand_slides

        # If brand slides workflow is disabled or no brand slides detected, return original
        if not b_cfg.enabled or not template_meta.brand_slides:
            return presentation

        brand_map = template_meta.brand_slides
        new_slides: List[Slide] = []

        # 1. BRAND_COVER
        if b_cfg.cover and "BRAND_COVER" in brand_map:
            cover_info = brand_map["BRAND_COVER"]
            cover_slide = Slide(
                intent=SlideIntent.TITLE_SLIDE,
                title=TitleBlock(text=presentation.title or "Presentation Title"),
                subtitle=presentation.subtitle,
                brand_role="BRAND_COVER",
                brand_slide_index=cover_info.get("slide_index"),
                brand_layout_index=cover_info.get("layout_index", 0),
            )
            new_slides.append(cover_slide)

        # 2. BRAND_DISCLAIMER
        if b_cfg.disclaimer and "BRAND_DISCLAIMER" in brand_map:
            disc_info = brand_map["BRAND_DISCLAIMER"]
            disc_slide = Slide(
                intent=SlideIntent.CONTENT_SLIDE,
                title=TitleBlock(text=disc_info.get("title", "Disclaimer")),
                brand_role="BRAND_DISCLAIMER",
                brand_slide_index=disc_info.get("slide_index"),
                brand_layout_index=disc_info.get("layout_index", 1),
            )
            new_slides.append(disc_slide)

        # 3. BRAND_ABOUT_COMPANY
        if b_cfg.about_company and "BRAND_ABOUT_COMPANY" in brand_map:
            about_info = brand_map["BRAND_ABOUT_COMPANY"]
            about_slide = Slide(
                intent=SlideIntent.CONTENT_SLIDE,
                title=TitleBlock(text=about_info.get("title", "About Company")),
                brand_role="BRAND_ABOUT_COMPANY",
                brand_slide_index=about_info.get("slide_index"),
                brand_layout_index=about_info.get("layout_index", 1),
            )
            new_slides.append(about_slide)

        # 4. BRAND_CASE_STUDY
        if b_cfg.case_study and "BRAND_CASE_STUDY" in brand_map:
            cs_info = brand_map["BRAND_CASE_STUDY"]
            cs_slide = Slide(
                intent=SlideIntent.CONTENT_SLIDE,
                title=TitleBlock(text=cs_info.get("title", "Case Study")),
                brand_role="BRAND_CASE_STUDY",
                brand_slide_index=cs_info.get("slide_index"),
                brand_layout_index=cs_info.get("layout_index", 1),
            )
            new_slides.append(cs_slide)

        # 5. Dynamic AI Content Slides
        for s in presentation.slides:
            if s.intent != SlideIntent.TITLE_SLIDE and s.intent != SlideIntent.THANK_YOU_SLIDE:
                new_slides.append(s)

        # 6. BRAND_THANK_YOU
        if b_cfg.thank_you and "BRAND_THANK_YOU" in brand_map:
            ty_info = brand_map["BRAND_THANK_YOU"]
            ty_slide = Slide(
                intent=SlideIntent.THANK_YOU_SLIDE,
                title=TitleBlock(text=ty_info.get("title", "Thank You!")),
                brand_role="BRAND_THANK_YOU",
                brand_slide_index=ty_info.get("slide_index"),
                brand_layout_index=ty_info.get("layout_index", 8),
            )
            new_slides.append(ty_slide)
        else:
            ty_match = [s for s in presentation.slides if s.intent == SlideIntent.THANK_YOU_SLIDE]
            if ty_match:
                new_slides.append(ty_match[0])

        return Presentation(
            title=presentation.title,
            subtitle=presentation.subtitle,
            slides=new_slides,
            metadata=presentation.metadata,
        )
