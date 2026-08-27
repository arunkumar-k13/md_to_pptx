"""Slide Planner Module.

Presentation Planning Layer that converts an AnalyzedDocument into a structured,
unpaginated Presentation model. Decides slide boundaries, content grouping, and slide
intents (Title, Section Divider, Content, Image, Table, Code).
"""

from __future__ import annotations
import logging
from typing import List, Optional

from md_to_pptx.config.settings_loader import Settings
from md_to_pptx.core.ast_nodes import (
    BulletListNode,
    CodeBlockNode,
    HeaderNode,
    ImageNode,
    ParagraphNode,
    QuoteNode,
    TableNode,
)
from md_to_pptx.core.content_analyzer import (
    AnalyzedBlock,
    AnalyzedDocument,
    ContentBlockType,
)
from md_to_pptx.core.design_engine import PresentationDesignEngine, PresentationDesignPlan, PresentationIntent
from md_to_pptx.core.presentation_model import (
    BulletListBlock,
    CodeBlock,
    ContentBlock,
    ImageBlock,
    ParagraphBlock,
    Presentation,
    QuoteBlock,
    Slide,
    SlideIntent,
    TableBlock,
    TitleBlock,
)

logger = logging.getLogger(__name__)


class SlidePlanner:
    """Planner converting AnalyzedDocument structures into initial Presentation IR objects using PresentationDesignEngine."""

    def plan(
        self,
        analyzed_doc: AnalyzedDocument,
        settings: Optional[Settings] = None,
        design_plan: Optional[PresentationDesignPlan] = None,
    ) -> Presentation:
        """Plan a Presentation model from an AnalyzedDocument instance and PresentationDesignPlan.

        Args:
            analyzed_doc: Classified document structure from ContentAnalyzer.
            settings: Optional Settings configuration instance.
            design_plan: Optional PresentationDesignPlan from PresentationDesignEngine.

        Returns:
            Populated unpaginated Presentation model.
        """
        cfg = settings or Settings()
        engine = PresentationDesignEngine()
        plan = design_plan or engine.create_design_plan(analyzed_doc)

        INTENT_MAP = {
            PresentationIntent.TITLE_COVER: SlideIntent.TITLE_SLIDE,
            PresentationIntent.SECTION_DIVIDER: SlideIntent.SECTION_DIVIDER,
            PresentationIntent.EXECUTIVE_SUMMARY: SlideIntent.EXECUTIVE_SUMMARY,
            PresentationIntent.AGENDA: SlideIntent.EXECUTIVE_SUMMARY,
            PresentationIntent.BULLET_CONTENT: SlideIntent.BULLET_CONTENT,
            PresentationIntent.CARD_GRID: SlideIntent.TWO_COLUMN,
            PresentationIntent.TWO_COLUMN: SlideIntent.TWO_COLUMN,
            PresentationIntent.COMPARISON: SlideIntent.COMPARISON,
            PresentationIntent.TABLE: SlideIntent.TABLE_SLIDE,
            PresentationIntent.IMAGE: SlideIntent.IMAGE_SLIDE,
            PresentationIntent.CODE: SlideIntent.CODE_SLIDE,
            PresentationIntent.QUOTE: SlideIntent.QUOTE_SLIDE,
            PresentationIntent.ROADMAP: SlideIntent.ROADMAP,
            PresentationIntent.TIMELINE: SlideIntent.TIMELINE,
            PresentationIntent.PROCESS: SlideIntent.PROCESS,
            PresentationIntent.METRICS: SlideIntent.METRICS,
            PresentationIntent.KPI_DASHBOARD: SlideIntent.METRICS,
            PresentationIntent.THANK_YOU: SlideIntent.THANK_YOU_SLIDE,
            PresentationIntent.CONCLUSION: SlideIntent.CONCLUSION,
            PresentationIntent.RECOMMENDATIONS: SlideIntent.CONCLUSION,
            PresentationIntent.KEY_TAKEAWAYS: SlideIntent.EXECUTIVE_SUMMARY,
            PresentationIntent.INFOGRAPHIC: SlideIntent.IMAGE_SLIDE,
        }

        slides: List[Slide] = []

        for dp in plan.slides:
            s_intent = INTENT_MAP.get(dp.intent, SlideIntent.CONTENT_SLIDE)

            cblocks: List[ContentBlock] = []
            for raw_b in dp.raw_blocks:
                cb = self._map_ast_to_content_block(raw_b)
                if cb:
                    cblocks.append(cb)

            # Auto-split single multi-item bullet list into 2 balanced columns for TWO_COLUMN intent
            if s_intent == SlideIntent.TWO_COLUMN and len(cblocks) == 1 and isinstance(cblocks[0], BulletListBlock):
                blist = cblocks[0]
                if len(blist.items) >= 4:
                    mid = (len(blist.items) + 1) // 2
                    left_b = BulletListBlock(items=blist.items[:mid], is_ordered=blist.is_ordered)
                    right_b = BulletListBlock(items=blist.items[mid:], is_ordered=blist.is_ordered)
                    cblocks = [left_b, right_b]

            title_block = TitleBlock(text=dp.title_text, level=1 if dp.intent == PresentationIntent.TITLE_COVER else 2)
            slide = Slide(
                intent=s_intent,
                title=title_block,
                subtitle=dp.subtitle_text,
                blocks=cblocks,
            )
            slides.append(slide)

        # Prune empty slides with no content blocks (except Title, Thank You, and Brand slides)
        pruned_slides = []
        for s in slides:
            if s.intent in (SlideIntent.TITLE_SLIDE, SlideIntent.THANK_YOU_SLIDE, SlideIntent.SECTION_DIVIDER) or getattr(s, "brand_role", None):
                pruned_slides.append(s)
            elif s.blocks and len(s.blocks) > 0:
                pruned_slides.append(s)
            else:
                logger.info("Pruned empty content slide titled '%s'", s.title.text if s.title else "Untitled")

        return Presentation(title=analyzed_doc.title, slides=pruned_slides if pruned_slides else slides)

    def _map_ast_to_content_block(self, analyzed_block: AnalyzedBlock) -> Optional[ContentBlock]:
        """Convert an AnalyzedBlock into a presentation ContentBlock instance.

        Args:
            analyzed_block: Input AnalyzedBlock node.

        Returns:
            Mapped ContentBlock instance or None.
        """
        node = analyzed_block.ast_node

        if isinstance(node, ParagraphNode):
            return ParagraphBlock(text=node.text, runs=node.runs)

        elif isinstance(node, BulletListNode):
            return BulletListBlock(items=node.items, is_ordered=node.is_ordered)

        elif isinstance(node, ImageNode):
            return ImageBlock(src=node.src, alt=node.alt, title=node.title)

        elif isinstance(node, TableNode):
            return TableBlock(headers=node.headers, rows=node.rows)

        elif isinstance(node, CodeBlockNode):
            return CodeBlock(code=node.code, language=node.language)

        elif isinstance(node, QuoteNode):
            return QuoteBlock(text=node.text)

        return None
