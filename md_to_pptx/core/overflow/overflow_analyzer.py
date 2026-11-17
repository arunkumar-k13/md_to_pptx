"""Overflow Analyzer Module.

Decides whether a slide's content blocks fit inside the template's body text box and,
if not, where to split them. Heights come from the same typography the renderer uses
(`md_to_pptx.core.typography`), measured against the real body box of the resolved
layout, so the decision matches what is actually drawn on the slide.
"""

from __future__ import annotations
import logging
from dataclasses import dataclass
from typing import Any, Optional, Tuple

from md_to_pptx.config.settings_loader import Settings
from md_to_pptx.core.presentation_model import BulletListBlock, ContentBlock, ParagraphBlock, Slide
from md_to_pptx.core.typography import DEFAULT_TEXT_INSET_TOP_BOTTOM_IN, Typography, typography_for

logger = logging.getLogger(__name__)

# Headroom kept free at the bottom of the body box to absorb small differences
# between PowerPoint's and LibreOffice's text layout.
SAFETY_FRACTION = 0.03
# A paragraph is only split across slides if at least this share of it stays on the
# current slide; otherwise the whole paragraph moves to the next slide.
MIN_PARAGRAPH_HEAD_FRACTION = 0.35
# Never leave fewer than this many top-level list items on either side of a split when avoidable.
MIN_LIST_ITEMS_PER_SIDE = 2


@dataclass(slots=True)
class OverflowAnalysisResult:
    """Diagnostic result produced by OverflowAnalyzer for a single slide.

    Attributes:
        is_overflowing: True if the slide must be split.
        total_estimated_height: Estimated content height in inches (up to the split point).
        max_allowed_height: Usable body height budget in inches.
        split_block_index: Index of the block where the slide is split.
        split_item_index: For a list block, index of the first item moved to the next slide.
        split_parts: For a paragraph block, the (stays, moves) halves split at a sentence boundary.
        reason: Explanatory diagnostic message.
    """

    is_overflowing: bool = False
    total_estimated_height: float = 0.0
    max_allowed_height: float = 5.0
    split_block_index: Optional[int] = None
    split_item_index: Optional[int] = None
    split_parts: Optional[Tuple[ContentBlock, ContentBlock]] = None
    reason: str = ""


class OverflowAnalyzer:
    """Analyzer evaluating slides against the template's body box and the bullet cap."""

    def __init__(self, max_body_height_inches: float = 3.8, body_width_inches: float = 8.0) -> None:
        """Initialize with fallback body dimensions used when no template layout is available."""
        self.max_body_height_inches = max_body_height_inches
        self.body_width_inches = body_width_inches

    def _content_box(
        self, slide: Slide, cfg: Settings, template_meta: Optional[Any], resolver: Optional[Any]
    ) -> Tuple[float, float]:
        """Width and height (inches) of the body text box this slide will be rendered into.

        Uses the template's own body placeholder, stopping above any footer, date or slide-number
        placeholder that starts inside it. Falls back to the generic usable-bounds calculation
        when the layout has no body placeholder.
        """
        width, height = self.body_width_inches, self.max_body_height_inches
        if not (template_meta and resolver):
            return width, height
        try:
            layout_meta = resolver.resolve_layout(slide, template_meta)
            phs = list(getattr(layout_meta, "placeholders", []) or [])
            bodies = [p for p in phs if p.ph_type in ("BODY", "OBJECT", "CONTENT", "TEXT") and p.height_inches > 0]
            if bodies:
                body = bodies[0]
                top = body.top_inches
                bottom = min(body.top_inches + body.height_inches, template_meta.slide_height_inches)
                for p in phs:
                    if p.ph_type in ("FOOTER", "DATE", "SLIDE_NUMBER") and top < p.top_inches < bottom:
                        bottom = p.top_inches
                right_edge = template_meta.slide_width_inches - 0.35  # same clamp the generator applies
                width = min(body.width_inches, right_edge - body.left_inches)
                return width, max(0.5, bottom - top)
            bounds = layout_meta.get_usable_content_bounds(
                template_meta,
                header_top_margin=getattr(cfg.template_defaults, "header_top_margin_inches", 1.0),
                footer_bottom_margin=getattr(cfg.template_defaults, "footer_bottom_margin_inches", 0.6),
                padding=getattr(cfg.template_defaults, "content_padding_inches", 0.2),
            )
            return bounds.width_inches, bounds.height_inches
        except Exception as err:
            logger.warning("Failed to resolve layout bounds for slide overflow evaluation: %s", err)
            return width, height

    def analyze_slide(
        self,
        slide: Slide,
        settings: Optional[Settings] = None,
        template_meta: Optional[Any] = None,
        resolver: Optional[Any] = None,
    ) -> OverflowAnalysisResult:
        cfg = settings or Settings()
        slide_h = getattr(template_meta, "slide_height_inches", 7.5) if template_meta else 7.5
        typo: Typography = typography_for(slide_h)

        width, box_h = self._content_box(slide, cfg, template_meta, resolver)
        budget = max(0.5, (box_h - 2 * DEFAULT_TEXT_INSET_TOP_BOTTOM_IN) * (1.0 - SAFETY_FRACTION))
        max_items = max(1, int(getattr(cfg.overflow, "max_bullets_per_slide", 12)))

        used = 0.0
        items_used = 0

        def overflow(idx: int, reason: str, item_idx: Optional[int] = None, parts=None) -> OverflowAnalysisResult:
            return OverflowAnalysisResult(
                is_overflowing=True,
                total_estimated_height=used,
                max_allowed_height=budget,
                split_block_index=idx,
                split_item_index=item_idx,
                split_parts=parts,
                reason=reason,
            )

        for i, block in enumerate(slide.blocks):
            remaining = budget - used
            block_h = block.estimate_height(width, typo)

            if isinstance(block, BulletListBlock):
                n_items = len(block.items)
                allowed = max_items - items_used
                if block_h <= remaining and n_items <= allowed:
                    used += block_h
                    items_used += n_items
                    continue

                heights = block.item_heights(width, typo)
                prefix = [0.0]
                for h in heights:
                    prefix.append(prefix[-1] + h)
                split_points = block.split_indices()
                fitting = [k for k in split_points if prefix[k] <= remaining and k <= allowed]
                # Balance is judged on top-level items, so a parent and its nested items count as one
                n_groups = len(split_points) + 1 if n_items else 0
                balanced = [
                    k for k in fitting
                    if split_points.index(k) + 1 >= MIN_LIST_ITEMS_PER_SIDE
                    and n_groups - (split_points.index(k) + 1) >= MIN_LIST_ITEMS_PER_SIDE
                ]
                if balanced:
                    k = max(balanced)
                    return overflow(i, f"List exceeds remaining {remaining:.2f} in; keeping {k} of {n_items} items", item_idx=k)
                if i == 0:
                    # First block on the slide: must make progress, even if unbalanced
                    k = max(fitting) if fitting else (split_points[0] if split_points else None)
                    if k is not None:
                        return overflow(i, f"List taller than the whole body box; keeping {k} of {n_items} items", item_idx=k)
                    used += block_h  # a single unsplittable item: accept it
                    continue
                return overflow(i, f"Only {remaining:.2f} in left; moving the list to the next slide")

            if isinstance(block, ParagraphBlock):
                if block_h <= remaining:
                    used += block_h
                    continue
                parts = block.split_to_fit(width, remaining, typo)
                if parts is not None:
                    head_h = parts[0].estimate_height(width, typo)
                    if i == 0 or head_h >= MIN_PARAGRAPH_HEAD_FRACTION * block_h:
                        return overflow(i, f"Paragraph split at a sentence boundary ({head_h:.2f} in kept)", parts=parts)
                if i > 0:
                    return overflow(i, f"Paragraph ({block_h:.2f} in) does not fit in remaining {remaining:.2f} in")
                used += block_h  # cannot be split further: accept rather than loop
                continue

            # Tables, code, images, quotes: moved whole, never split here
            if block_h <= remaining or i == 0:
                used += block_h
                continue
            return overflow(i, f"{block.__class__.__name__} ({block_h:.2f} in) does not fit in remaining {remaining:.2f} in")

        return OverflowAnalysisResult(
            is_overflowing=False,
            total_estimated_height=used,
            max_allowed_height=budget,
        )
