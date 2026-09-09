"""Paginator Module.

Iterates over an initial Presentation IR model, invoking OverflowAnalyzer and
OverflowStrategy to spawn continuation slides ('Title (Continued)') until all slides
fit within spatial bounds.
"""

from __future__ import annotations
import logging
from typing import List, Optional

from md_to_pptx.config.settings_loader import Settings
from md_to_pptx.core.overflow.overflow_analyzer import OverflowAnalyzer
from md_to_pptx.core.overflow.overflow_strategy import OverflowStrategy
from md_to_pptx.core.presentation_model import Presentation, Slide, TitleBlock, ParagraphBlock

logger = logging.getLogger(__name__)


class Paginator:
    """Paginator orchestrator applying overflow strategies to generate a fully paginated Presentation."""

    def __init__(
        self,
        analyzer: Optional[OverflowAnalyzer] = None,
        strategy: Optional[OverflowStrategy] = None,
        resolver: Optional[Any] = None,
    ) -> None:
        """Initialize Paginator with analyzer, strategy, and resolver instances.

        Args:
            analyzer: Optional OverflowAnalyzer instance.
            strategy: Optional OverflowStrategy instance.
            resolver: Optional SlideLayoutResolver instance.
        """
        self.analyzer = analyzer or OverflowAnalyzer()
        self.strategy = strategy or OverflowStrategy()
        self.resolver = resolver

    def paginate(
        self,
        presentation: Presentation,
        settings: Optional[Settings] = None,
        context: Optional[Any] = None,
        template_meta: Optional[Any] = None,
    ) -> Presentation:
        """Paginate a Presentation model, spawning continuation slides where necessary.

        Args:
            presentation: Input unpaginated Presentation instance.
            settings: Settings configuration instance.
            context: Optional DiagnosticContext telemetry collector.
            template_meta: Optional TemplateMetadata instance for template-aware bounds.

        Returns:
            Fully paginated Presentation instance.
        """
        cfg = settings or Settings()
        suffix = cfg.presentation.continuation_suffix

        paginated_slides: List[Slide] = []

        for idx, slide in enumerate(presentation.slides, start=1):
            processed_slides = self._paginate_slide(slide, suffix, cfg, context, idx, template_meta)
            paginated_slides.extend(processed_slides)

        return Presentation(
            title=presentation.title,
            subtitle=presentation.subtitle,
            slides=paginated_slides,
            metadata=presentation.metadata,
        )

    def _paginate_slide(
        self,
        slide: Slide,
        continuation_suffix: str,
        settings: Settings,
        context: Optional[Any] = None,
        slide_idx: int = 1,
        template_meta: Optional[Any] = None,
    ) -> List[Slide]:
        """Recursively paginate a single slide if overflow occurs.

        Args:
            slide: Target Slide instance.
            continuation_suffix: Title suffix for continuation slides.
            settings: Settings configuration instance.
            context: Optional DiagnosticContext telemetry collector.
            slide_idx: 1-indexed slide number.
            template_meta: Optional TemplateMetadata instance.

        Returns:
            List of non-overflowing Slide instances.
        """
        results: List[Slide] = []
        current_slide = slide
        continuation_count = 0

        while True:
            analysis = self.analyzer.analyze_slide(
                current_slide,
                settings=settings,
                template_meta=template_meta,
                resolver=self.resolver,
            )

            if not analysis.is_overflowing or not current_slide.blocks:
                if current_slide.blocks or not results:
                    results.append(current_slide)
                break

            kept_blocks, moved_blocks = self.strategy.resolve_overflow(current_slide, analysis)

            if not moved_blocks or kept_blocks == current_slide.blocks:
                # Unable to split further
                results.append(current_slide)
                break

            # Record overflow event in telemetry context
            if context and hasattr(context, "record_overflow"):
                target_block = current_slide.blocks[analysis.split_block_index] if analysis.split_block_index is not None else None
                b_type_name = target_block.__class__.__name__ if target_block else "Block"
                context.record_overflow(
                    slide_index=slide_idx,
                    block_type=b_type_name,
                    action_taken="split_to_continuation_slide",
                    items_moved=len(moved_blocks),
                    reason=analysis.reason,
                )

            # Update current slide with kept blocks
            current_slide.blocks = kept_blocks
            results.append(current_slide)

            continuation_count += 1

            # Prepare next continuation slide
            first_moved = moved_blocks[0] if moved_blocks else None
            new_section_title = None
            if isinstance(first_moved, ParagraphBlock):
                txt = first_moved.text.strip()
                import re
                if re.match(r"^\s*(\*\*|\#\#\s*)?(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+[A-Z]", txt):
                    new_section_title = txt

            if new_section_title:
                continuation_title = TitleBlock(
                    text=new_section_title,
                    level=slide.title.level if slide.title else 2,
                )
            else:
                base_title_text = (
                    slide.title.text if slide.title else "Content"
                )
                continuation_title = TitleBlock(
                    text=f"{base_title_text} {continuation_suffix}",
                    level=slide.title.level if slide.title else 2,
                )

            current_slide = Slide(
                intent=slide.intent,
                title=continuation_title,
                blocks=moved_blocks,
                is_continuation=True,
                continuation_index=continuation_count,
            )

        return results
