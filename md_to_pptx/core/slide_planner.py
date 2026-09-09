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
            PresentationIntent.EXECUTIVE_SUMMARY: SlideIntent.BULLET_CONTENT,
            PresentationIntent.AGENDA: SlideIntent.BULLET_CONTENT,
            PresentationIntent.BULLET_CONTENT: SlideIntent.BULLET_CONTENT,
            PresentationIntent.CARD_GRID: SlideIntent.BULLET_CONTENT,
            PresentationIntent.TWO_COLUMN: SlideIntent.BULLET_CONTENT,
            PresentationIntent.COMPARISON: SlideIntent.BULLET_CONTENT,
            PresentationIntent.TABLE: SlideIntent.BULLET_CONTENT,
            PresentationIntent.IMAGE: SlideIntent.BULLET_CONTENT,
            PresentationIntent.CODE: SlideIntent.CODE_SLIDE,
            PresentationIntent.QUOTE: SlideIntent.QUOTE_SLIDE,
            PresentationIntent.ROADMAP: SlideIntent.BULLET_CONTENT,
            PresentationIntent.TIMELINE: SlideIntent.BULLET_CONTENT,
            PresentationIntent.PROCESS: SlideIntent.BULLET_CONTENT,
            PresentationIntent.METRICS: SlideIntent.BULLET_CONTENT,
            PresentationIntent.KPI_DASHBOARD: SlideIntent.BULLET_CONTENT,
            PresentationIntent.THANK_YOU: SlideIntent.THANK_YOU_SLIDE,
            PresentationIntent.CONCLUSION: SlideIntent.BULLET_CONTENT,
            PresentationIntent.RECOMMENDATIONS: SlideIntent.BULLET_CONTENT,
            PresentationIntent.KEY_TAKEAWAYS: SlideIntent.BULLET_CONTENT,
            PresentationIntent.INFOGRAPHIC: SlideIntent.BULLET_CONTENT,
        }

        slides: List[Slide] = []

        for dp in plan.slides:
            s_intent = INTENT_MAP.get(dp.intent, SlideIntent.BULLET_CONTENT)

            cblocks: List[ContentBlock] = []
            for raw_b in dp.raw_blocks:
                cb = self._map_ast_to_content_block(raw_b)
                if cb:
                    cblocks.append(cb)

            # If cblocks contains multiple CodeBlock objects, split each CodeBlock into its own slide preserving section title
            code_blocks = [cb for cb in cblocks if isinstance(cb, CodeBlock)]
            if len(code_blocks) > 1:
                for code_idx, cb in enumerate(code_blocks, start=1):
                    s_title = dp.title_text if code_idx == 1 else f"{dp.title_text} (Continued)"
                    t_block = TitleBlock(text=s_title, level=2)
                    code_slide = Slide(intent=SlideIntent.CODE_SLIDE, title=t_block, blocks=[cb])
                    slides.append(code_slide)
                continue

            # If cblocks contains multiple ImageBlock objects, split each ImageBlock into its own slide
            image_blocks = [cb for cb in cblocks if isinstance(cb, ImageBlock)]
            if len(image_blocks) > 1:
                for img_idx, ib in enumerate(image_blocks, start=1):
                    s_title = dp.title_text if img_idx == 1 else f"{dp.title_text} (Continued)"
                    t_block = TitleBlock(text=s_title, level=2)
                    img_slide = Slide(intent=SlideIntent.BULLET_CONTENT, title=t_block, blocks=[ib])
                    slides.append(img_slide)
                continue

            # If cblocks contains multiple TableBlock objects, split each TableBlock into its own slide
            table_blocks = [cb for cb in cblocks if isinstance(cb, TableBlock)]
            if len(table_blocks) > 1:
                for tbl_idx, tb in enumerate(table_blocks, start=1):
                    s_title = dp.title_text if tbl_idx == 1 else f"{dp.title_text} (Continued)"
                    t_block = TitleBlock(text=s_title, level=2)
                    tbl_slide = Slide(intent=SlideIntent.BULLET_CONTENT, title=t_block, blocks=[tb])
                    slides.append(tbl_slide)
                continue

            # If cblocks contains text blocks AND a TableBlock, split TableBlock to continuation slide if combined height > 3.0"
            if len(table_blocks) == 1 and len(cblocks) > 1:
                tb = table_blocks[0]
                other_blocks = [cb for cb in cblocks if not isinstance(cb, TableBlock)]
                other_h = sum(ob.estimate_height(11.0) for ob in other_blocks)
                if other_h + tb.estimate_height(11.0) > 3.0:
                    t_block1 = TitleBlock(text=dp.title_text, level=2)
                    slides.append(Slide(intent=SlideIntent.BULLET_CONTENT, title=t_block1, blocks=other_blocks))
                    t_block2 = TitleBlock(text=f"{dp.title_text} (Continued)", level=2)
                    slides.append(Slide(intent=SlideIntent.BULLET_CONTENT, title=t_block2, blocks=[tb]))
                    continue

            title_block = TitleBlock(text=dp.title_text, level=1 if dp.intent == PresentationIntent.TITLE_COVER else 2)

            # Remove duplicated paragraph block if text matches title_block text
            clean_cblocks = []
            for cb in cblocks:
                if isinstance(cb, ParagraphBlock) and title_block and title_block.text and cb.text.strip() == title_block.text.strip():
                    continue
                clean_cblocks.append(cb)
            cblocks = clean_cblocks
            slide = Slide(
                intent=s_intent,
                title=title_block,
                subtitle=dp.subtitle_text,
                blocks=cblocks,
            )
            slides.append(slide)

        # Prune empty slides with no content blocks (except Title, Thank You, or Section 17)
        pruned_slides = []
        for s in slides:
            t_str = s.title.text.strip() if (s.title and s.title.text) else ""
            if s.intent == SlideIntent.TITLE_SLIDE:
                pruned_slides.append(s)
            elif t_str and not s.blocks:
                if "17." in t_str or "empty" in t_str.lower():
                    pruned_slides.append(s)
                else:
                    logger.info("Pruned empty divider slide titled '%s'", t_str)
            elif s.blocks and len(s.blocks) > 0:
                pruned_slides.append(s)
            else:
                logger.info("Pruned empty content slide titled '%s'", t_str)

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

        elif isinstance(node, HeaderNode):
            h_text = node.text
            runs = node.runs if hasattr(node, "runs") and node.runs else [InlineRun(text=h_text, is_bold=True)]
            return ParagraphBlock(text=h_text, runs=runs, style_overrides={"heading_level": node.level})

        elif isinstance(node, BulletListNode):
            return BulletListBlock(items=node.items, is_ordered=node.is_ordered, start_index=getattr(node, "start_index", 1))

        elif isinstance(node, ImageNode):
            return ImageBlock(src=node.src, alt=node.alt, title=node.title)

        elif isinstance(node, TableNode):
            return TableBlock(headers=node.headers, rows=node.rows, raw_rows=node.raw_rows)

        elif isinstance(node, CodeBlockNode):
            return CodeBlock(code=node.code, language=node.language)

        elif isinstance(node, QuoteNode):
            runs = getattr(node, "runs", [])
            return QuoteBlock(text=node.text, runs=runs)

        return None
