"""Overflow Strategy Module.

Defines rules for resolving detected slide overflow conditions (e.g. splitting
bullet lists across blocks, splitting paragraph blocks, moving blocks to continuation slides).
"""

from __future__ import annotations
import logging
from typing import List, Tuple

from md_to_pptx.core.overflow.overflow_analyzer import OverflowAnalysisResult
from md_to_pptx.core.presentation_model import BulletListBlock, ContentBlock, ParagraphBlock, Slide

logger = logging.getLogger(__name__)


class OverflowStrategy:
    """Strategy handler deciding how to split blocks and resolve overflow."""

    def resolve_overflow(
        self, slide: Slide, analysis: OverflowAnalysisResult
    ) -> Tuple[List[ContentBlock], List[ContentBlock]]:
        """Split a slide's content blocks into current slide blocks and continuation slide blocks.

        Args:
            slide: Slide instance experiencing overflow.
            analysis: OverflowAnalysisResult instance from OverflowAnalyzer.

        Returns:
            Tuple of (current_slide_kept_blocks, continuation_slide_moved_blocks).
        """
        if not analysis.is_overflowing or analysis.split_block_index is None:
            return slide.blocks, []

        split_idx = analysis.split_block_index
        kept_blocks = list(slide.blocks[:split_idx])
        target_block = slide.blocks[split_idx]
        remaining_blocks = list(slide.blocks[split_idx + 1 :])

        # If a single BulletListBlock overflows internally, split it into two list blocks
        if isinstance(target_block, BulletListBlock) and analysis.split_item_index:
            item_split_idx = analysis.split_item_index
            first_items = target_block.items[:item_split_idx]
            second_items = target_block.items[item_split_idx:]

            cur_start = getattr(target_block, "start_index", 1)
            if first_items:
                kept_blocks.append(BulletListBlock(items=first_items, is_ordered=target_block.is_ordered, start_index=cur_start))
            
            moved_blocks = []
            if second_items:
                moved_blocks.append(BulletListBlock(items=second_items, is_ordered=target_block.is_ordered, start_index=cur_start + len(first_items)))
            moved_blocks.extend(remaining_blocks)
            
            return kept_blocks, moved_blocks

        # Default strategy: move overflowing block and all subsequent blocks to continuation slide
        moved_blocks = [target_block] + remaining_blocks

        # Ensure section header paragraph is not left orphaned at bottom of kept_blocks
        if kept_blocks and isinstance(kept_blocks[-1], ParagraphBlock):
            header_txt = kept_blocks[-1].text.strip()
            if any(header_txt.startswith(prefix) for prefix in ("I.", "II.", "III.", "IV.", "V.", "VI.", "VII.", "VIII.", "IX.", "X.", "**I", "**II", "**III")) or "(" in header_txt:
                moved_header = kept_blocks.pop()
                moved_blocks.insert(0, moved_header)

        return kept_blocks, moved_blocks
