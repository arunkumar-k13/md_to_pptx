"""Overflow Analyzer Module.

Evaluates presentation slides and individual content blocks against spatial height
budgets and configured item caps to detect visual overflow conditions.
"""

from __future__ import annotations
import logging
from dataclasses import dataclass, field
from typing import List, Optional

from md_to_pptx.config.settings_loader import Settings
from md_to_pptx.core.presentation_model import BulletListBlock, ContentBlock, ParagraphBlock, Slide

logger = logging.getLogger(__name__)


@dataclass(slots=True)
class OverflowAnalysisResult:
    """Diagnostic result produced by OverflowAnalyzer for a single slide.

    Attributes:
        is_overflowing: True if total estimated height or item counts exceed thresholds.
        total_estimated_height: Sum of block estimated heights in inches.
        max_allowed_height: Maximum allowed height budget in inches.
        split_block_index: Recommended index of the block at which to split the slide.
        split_item_index: Recommended internal item split index if a single block overflows.
        reason: Explanatory diagnostic message.
    """

    is_overflowing: bool = False
    total_estimated_height: float = 0.0
    max_allowed_height: float = 5.0
    split_block_index: Optional[int] = None
    split_item_index: Optional[int] = None
    reason: str = ""


class OverflowAnalyzer:
    """Analyzer evaluating slides against spatial and item count overflow constraints."""

    def __init__(self, max_body_height_inches: float = 4.2, body_width_inches: float = 8.0) -> None:
        """Initialize OverflowAnalyzer with body placeholder dimensions.

        Args:
            max_body_height_inches: Target body placeholder height in inches (default 4.2).
            body_width_inches: Target body placeholder width in inches (default 8.0).
        """
        self.max_body_height_inches = max_body_height_inches
        self.body_width_inches = body_width_inches

    def analyze_slide(
        self,
        slide: Slide,
        settings: Optional[Settings] = None,
        template_meta: Optional[Any] = None,
        resolver: Optional[Any] = None,
    ) -> OverflowAnalysisResult:
        """Analyze a Slide instance for height or item count overflow.

        Args:
            slide: Slide instance to evaluate.
            settings: Settings configuration instance.
            template_meta: Optional TemplateMetadata instance for template-aware bounds.
            resolver: Optional SlideLayoutResolver instance for layout resolution.

        Returns:
            OverflowAnalysisResult instance.
        """
        cfg = settings or Settings()
        max_bullets = cfg.overflow.max_bullets_per_slide
        max_words = cfg.overflow.max_paragraph_words_per_slide

        target_max_height = self.max_body_height_inches
        target_width = self.body_width_inches

        if template_meta and resolver:
            try:
                layout_meta = resolver.resolve_layout(slide, template_meta)
                h_margin = getattr(cfg.template_defaults, "header_top_margin_inches", 1.0)
                f_margin = getattr(cfg.template_defaults, "footer_bottom_margin_inches", 0.6)
                pad = getattr(cfg.template_defaults, "content_padding_inches", 0.2)

                bounds = layout_meta.get_usable_content_bounds(
                    template_meta,
                    header_top_margin=h_margin,
                    footer_bottom_margin=f_margin,
                    padding=pad,
                )
                target_max_height = bounds.height_inches
                target_width = bounds.width_inches
            except Exception as err:
                logger.warning("Failed to resolve layout bounds for slide overflow evaluation: %s", err)

        total_height = 0.0
        total_bullets = 0
        total_words = 0

        for i, block in enumerate(slide.blocks):
            block_h = block.estimate_height(target_width)

            # Check individual BulletListBlock overflow
            if isinstance(block, BulletListBlock):
                list_count = len(block.items)
                if list_count > max_bullets or (total_bullets + list_count > max_bullets):
                    split_item_idx = max(1, max_bullets - total_bullets)
                    return OverflowAnalysisResult(
                        is_overflowing=True,
                        total_estimated_height=total_height + block_h,
                        max_allowed_height=target_max_height,
                        split_block_index=i,
                        split_item_index=split_item_idx,
                        reason=f"Bullet item count ({total_bullets + list_count}) exceeds cap of {max_bullets}",
                    )
                total_bullets += list_count

            # Check individual ParagraphBlock word count overflow
            elif isinstance(block, ParagraphBlock):
                p_words = len(block.text.split())
                if p_words > max_words or (total_words + p_words > max_words):
                    return OverflowAnalysisResult(
                        is_overflowing=True,
                        total_estimated_height=total_height + block_h,
                        max_allowed_height=target_max_height,
                        split_block_index=i,
                        reason=f"Paragraph word count ({total_words + p_words}) exceeds cap of {max_words}",
                    )
                total_words += p_words

            # Accumulate height & check spatial height overflow budget
            if total_height + block_h > target_max_height:
                if i > 0:
                    return OverflowAnalysisResult(
                        is_overflowing=True,
                        total_estimated_height=total_height + block_h,
                        max_allowed_height=target_max_height,
                        split_block_index=i,
                        reason=f"Estimated block height ({total_height + block_h:.2f} in) exceeds height budget ({target_max_height:.2f} in)",
                    )
                elif isinstance(block, BulletListBlock) and len(block.items) > 3:
                    split_item_idx = max(2, int(len(block.items) * (target_max_height / block_h)))
                    if split_item_idx < len(block.items):
                        return OverflowAnalysisResult(
                            is_overflowing=True,
                            total_estimated_height=block_h,
                            max_allowed_height=target_max_height,
                            split_block_index=i,
                            split_item_index=split_item_idx,
                            reason=f"Single list height ({block_h:.2f} in) exceeds height budget ({target_max_height:.2f} in)",
                        )

            total_height += block_h

        return OverflowAnalysisResult(
            is_overflowing=False,
            total_estimated_height=total_height,
            max_allowed_height=target_max_height,
        )
