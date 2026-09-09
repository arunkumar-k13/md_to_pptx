"""Slide Layout Resolver Module.

Multi-signal weighted scoring engine matching logical Slide intents and content
composition to optimal layout metadata within TemplateMetadata. Contains zero direct
PowerPoint shape manipulation.
"""

from __future__ import annotations
import logging
from typing import Dict, List, Optional, Tuple

from md_to_pptx.core.presentation_model import (
    CodeBlock,
    ContentBlock,
    ImageBlock,
    ParagraphBlock,
    QuoteBlock,
    Slide,
    SlideIntent,
    TableBlock,
)
from md_to_pptx.core.template.template_analyzer import LayoutMetadata, TemplateMetadata

logger = logging.getLogger(__name__)


class SlideLayoutResolver:
    """Multi-signal weighted layout resolver scoring candidate layouts by semantic intent and placeholder signature."""

    def resolve_layout(self, slide: Slide, template_meta: TemplateMetadata) -> LayoutMetadata:
        """Select the best matching LayoutMetadata for a given Slide using multi-signal scoring.

        Args:
            slide: Target Slide instance.
            template_meta: Introspected TemplateMetadata object.

        Returns:
            Matched LayoutMetadata object.
        """
        best_layout, _, _ = self.resolve_layout_with_score(slide, template_meta)
        return best_layout

    def resolve_layout_with_score(
        self, slide: Slide, template_meta: TemplateMetadata
    ) -> Tuple[LayoutMetadata, float, str]:
        """Select best LayoutMetadata returning confidence score and decision explanation.

        Args:
            slide: Target Slide instance.
            template_meta: Introspected TemplateMetadata object.

        Returns:
            Tuple of (best_layout, confidence_score, explanation_reason).
        """
        if not template_meta.layouts:
            raise ValueError("TemplateMetadata contains no layouts.")

        if slide.brand_role:
            if template_meta.brand_slides and slide.brand_role in template_meta.brand_slides:
                b_info = template_meta.brand_slides[slide.brand_role]
                l_idx = b_info.get("layout_index", 0)
                if 0 <= l_idx < len(template_meta.layouts):
                    return template_meta.layouts[l_idx], 1.0, f"Preserved Brand Slide ({slide.brand_role})"
            for layout in template_meta.layouts:
                if layout.brand_role == slide.brand_role:
                    return layout, 1.0, f"Preserved Brand Slide ({slide.brand_role})"

        scored_layouts: List[Tuple[LayoutMetadata, float, str]] = []

        for layout in template_meta.layouts:
            score, reason = self._score_layout_for_slide(slide, layout)
            scored_layouts.append((layout, score, reason))

        # Sort descending by confidence score
        scored_layouts.sort(key=lambda x: x[1], reverse=True)

        best_layout, best_score, best_reason = scored_layouts[0]
        logger.info(
            "Slide '%s' (Intent: %s) mapped to layout '%s' with confidence %.2f (%s)",
            slide.title.text if slide.title else "Untitled",
            slide.intent.value,
            best_layout.name,
            best_score,
            best_reason,
        )

        return best_layout, best_score, best_reason

    def _score_layout_for_slide(self, slide: Slide, layout: LayoutMetadata) -> Tuple[float, str]:
        """Calculate weighted multi-signal confidence score for a candidate layout."""
        score = 0.0
        reasons = []
        l_name = layout.name.lower()
        intent = slide.intent
        block_count = len(slide.blocks)

        has_image = any(isinstance(b, ImageBlock) for b in slide.blocks)
        has_table = any(isinstance(b, TableBlock) for b in slide.blocks)

        # 1. Slide Intent Signal (Max +4.0)
        if intent == SlideIntent.TITLE_SLIDE:
            if "title slide" in l_name or (layout.has_title and not layout.has_body):
                score += 4.0
                reasons.append("Title Slide Layout Match")
            elif "title" in l_name:
                score += 2.0

        elif intent == SlideIntent.SECTION_DIVIDER:
            if any(k in l_name for k in ("section", "divider", "header")):
                score += 4.0
                reasons.append("Section Divider Layout Match")

        elif intent == SlideIntent.IMAGE_SLIDE:
            if layout.has_picture:
                score += 4.0
                reasons.append("Picture Placeholder Signature Match")
            elif layout.has_body:
                score += 3.0
                reasons.append("Image Card rendered in Content Placeholder")

        elif intent == SlideIntent.TABLE_SLIDE or has_table:
            if layout.has_table:
                score += 4.0
                reasons.append("Table Placeholder Signature Match")
            elif "table" in l_name:
                score += 3.5
                reasons.append("Table Layout Name Match")
            elif layout.has_body:
                score += 2.0
                reasons.append("Table rendered in Content Placeholder")

        elif intent in (SlideIntent.TWO_COLUMN, SlideIntent.COMPARISON):
            if layout.body_count >= 2:
                score += 3.5
                reasons.append("Multi-Column Capacity Match")
            if any(k in l_name for k in ("two", "comparison", "column")):
                score += 1.0

        elif intent in (SlideIntent.ROADMAP, SlideIntent.TIMELINE, SlideIntent.PROCESS):
            if layout.body_count >= 2:
                score += 3.0
                reasons.append("Sequential Flow Multi-Body Match")
            elif layout.has_body:
                score += 2.0

        elif intent in (SlideIntent.EXECUTIVE_SUMMARY, SlideIntent.OVERVIEW):
            if layout.has_title and layout.has_body:
                score += 3.0
                reasons.append("Executive Overview Content Match")

        else:
            if layout.has_title and layout.has_body:
                score += 2.5
                reasons.append("Standard Title + Body Content Match")

        # 2. Body Capacity & Density Matching (Max +2.5, Min -1.0)
        if block_count > 0:
            if intent in (SlideIntent.TWO_COLUMN, SlideIntent.COMPARISON):
                if layout.body_count == block_count:
                    score += 2.5
                    reasons.append(f"Exact Body Count Match ({block_count})")
                elif layout.body_count >= 2:
                    score += 2.0
                    reasons.append("Multi-Column Layout Match")
            else:
                # For standard single-column content slides, strongly prefer 1-body layouts (e.g. Title and Content)
                if layout.body_count == 1:
                    score += 3.0
                    reasons.append("Single Column Content Layout Match")
                elif layout.body_count == block_count:
                    score += 1.0
                    reasons.append(f"Body Count Match ({block_count})")
                elif layout.body_count > 0:
                    score += 0.5

        # 3. Penalty for Title-Only layouts on body content slides
        if block_count > 0 and layout.body_count == 0 and not layout.has_picture and not layout.has_table:
            if intent not in (SlideIntent.TITLE_SLIDE, SlideIntent.SECTION_DIVIDER):
                score -= 3.0

        # Normalize score to confidence range [0.0, 1.0]
        max_possible = 9.0
        confidence = round(max(0.05, min(0.99, score / max_possible)), 2)
        reason_str = ", ".join(reasons) if reasons else "Default Candidate Match"

        return confidence, reason_str
