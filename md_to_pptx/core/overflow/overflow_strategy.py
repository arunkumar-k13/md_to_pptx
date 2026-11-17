"""Overflow Strategy Module.

Applies the split decided by OverflowAnalyzer: keeps the blocks that fit on the current
slide and moves the rest to a continuation slide. Lists split between top-level items
(nested items stay with their parent), paragraphs split at sentence boundaries, and a
heading is never left as the last line of a slide.
"""

from __future__ import annotations
import logging
from typing import List, Tuple

from md_to_pptx.core.overflow.overflow_analyzer import OverflowAnalysisResult
from md_to_pptx.core.presentation_model import (
    BulletListBlock,
    ContentBlock,
    ParagraphBlock,
    Slide,
    heading_level_of,
)

logger = logging.getLogger(__name__)


def _is_heading_block(block: ContentBlock) -> bool:
    """A Markdown heading or an all-bold label paragraph; it introduces the content after it."""
    return isinstance(block, ParagraphBlock) and heading_level_of(block) is not None


def _top_level_count(items: list) -> int:
    """Number of items at the list's outermost level (used to continue numbering)."""
    if not items:
        return 0
    min_level = min(getattr(i, "level", 0) for i in items)
    return sum(1 for i in items if getattr(i, "level", 0) == min_level)


class OverflowStrategy:
    """Strategy handler applying split decisions to a slide's blocks."""

    def resolve_overflow(
        self, slide: Slide, analysis: OverflowAnalysisResult
    ) -> Tuple[List[ContentBlock], List[ContentBlock]]:
        """Split a slide's blocks into (kept on this slide, moved to the continuation slide)."""
        if not analysis.is_overflowing or analysis.split_block_index is None:
            return slide.blocks, []

        split_idx = analysis.split_block_index
        kept_blocks: List[ContentBlock] = list(slide.blocks[:split_idx])
        target_block = slide.blocks[split_idx]
        remaining_blocks = list(slide.blocks[split_idx + 1 :])

        if analysis.split_parts is not None:
            head, tail = analysis.split_parts
            kept_blocks.append(head)
            moved_blocks = [tail] + remaining_blocks
            return kept_blocks, moved_blocks

        if isinstance(target_block, BulletListBlock) and analysis.split_item_index:
            k = analysis.split_item_index
            first_items = target_block.items[:k]
            second_items = target_block.items[k:]
            start = getattr(target_block, "start_index", 1)
            if first_items:
                kept_blocks.append(
                    BulletListBlock(items=first_items, is_ordered=target_block.is_ordered, start_index=start)
                )
            moved_blocks: List[ContentBlock] = []
            if second_items:
                moved_blocks.append(
                    BulletListBlock(
                        items=second_items,
                        is_ordered=target_block.is_ordered,
                        start_index=start + _top_level_count(first_items),
                    )
                )
            moved_blocks.extend(remaining_blocks)
            return kept_blocks, moved_blocks

        moved_blocks = [target_block] + remaining_blocks

        # Keep headings with the content they introduce (but never empty the slide)
        while len(kept_blocks) > 1 and _is_heading_block(kept_blocks[-1]):
            moved_blocks.insert(0, kept_blocks.pop())

        return kept_blocks, moved_blocks
