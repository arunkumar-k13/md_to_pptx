"""Presentation Design Engine Module.

Core intelligence engine that analyzes AnalyzedDocument domain models,
infers presentation patterns, and produces an abstract PresentationDesignPlan
containing PresentationIntents, visual priorities, and suggested layout traits.
"""

from __future__ import annotations
import enum
import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

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

logger = logging.getLogger(__name__)


class PresentationIntent(enum.Enum):
    """Abstract presentation intents representing executive visual communication goals."""

    TITLE_COVER = "title_cover"
    EXECUTIVE_SUMMARY = "executive_summary"
    AGENDA = "agenda"
    SECTION_DIVIDER = "section_divider"
    BULLET_CONTENT = "bullet_content"
    CARD_GRID = "card_grid"
    TWO_COLUMN = "two_column"
    COMPARISON = "comparison"
    TABLE = "table"
    IMAGE = "image"
    CODE = "code"
    QUOTE = "quote"
    INFOGRAPHIC = "infographic"
    TIMELINE = "timeline"
    ROADMAP = "roadmap"
    PROCESS = "process"
    METRICS = "metrics"
    KPI_DASHBOARD = "kpi_dashboard"
    KEY_TAKEAWAYS = "key_takeaways"
    RECOMMENDATIONS = "recommendations"
    THANK_YOU = "thank_you"
    CONCLUSION = "conclusion"


@dataclass(slots=True)
class SlideDesignPlan:
    """Design plan for an individual presentation slide.

    Attributes:
        slide_index: 1-based slide sequence index.
        title_text: Target slide heading text.
        subtitle_text: Target slide subtitle text.
        intent: Classified PresentationIntent enum.
        why_chosen: Executive design rationale string.
        visual_priority: Importance level ('High', 'Medium', 'Low').
        density: Estimated visual density ('High', 'Medium', 'Low').
        estimated_occupancy_percent: Estimated vertical canvas occupancy percentage.
        suggested_layout_traits: List of recommended layout characteristics.
        raw_blocks: List of associated AnalyzedBlock instances.
    """

    slide_index: int
    title_text: str = ""
    subtitle_text: str = ""
    intent: PresentationIntent = PresentationIntent.BULLET_CONTENT
    why_chosen: str = ""
    visual_priority: str = "Medium"
    density: str = "Medium"
    estimated_occupancy_percent: float = 50.0
    suggested_layout_traits: List[str] = field(default_factory=list)
    raw_blocks: List[AnalyzedBlock] = field(default_factory=list)


@dataclass(slots=True)
class PresentationDesignPlan:
    """Complete presentation design plan container.

    Attributes:
        doc_title: Presentation main document title.
        slides: List of SlideDesignPlan objects.
        metadata: Dictionary of design metadata.
    """

    doc_title: str = ""
    slides: List[SlideDesignPlan] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)


class PresentationDesignEngine:
    """Intelligence engine inferring PresentationDesignPlans from AnalyzedDocument structures."""

    def create_design_plan(self, doc: AnalyzedDocument) -> PresentationDesignPlan:
        """Analyze an AnalyzedDocument and generate a PresentationDesignPlan.

        Args:
            doc: Parsed and classified AnalyzedDocument.

        Returns:
            Populated PresentationDesignPlan instance.
        """
        plans: List[SlideDesignPlan] = []
        slide_idx = 1

        # 1. Document Cover
        if doc.title:
            plans.append(
                SlideDesignPlan(
                    slide_index=slide_idx,
                    title_text=doc.title,
                    subtitle_text="",
                    intent=PresentationIntent.TITLE_COVER,
                    why_chosen="Document Title H1 Header -> Cover Slide",
                    visual_priority="High",
                    density="Low",
                    estimated_occupancy_percent=35.0,
                    suggested_layout_traits=["Title Slide", "Centered Title", "Subtitle Frame"],
                    raw_blocks=[],
                )
            )
            slide_idx += 1

        # 2. Process Sections
        for sec in doc.sections:
            sec_title = sec.title
            blocks = sec.blocks

            # Skip duplicate section slide if title matches main document cover title and section has no body blocks (or only H1 title)
            if doc.title and sec_title == doc.title and not [b for b in blocks if b.block_type != ContentBlockType.SECTION_HEADER]:
                continue

            title = sec_title
            subtitle = ""
            body_blocks = blocks

            intent, why, priority, traits = self._infer_intent_and_traits(title, subtitle, body_blocks)
            density_str, occupancy = self._estimate_density_and_occupancy(body_blocks)

            plans.append(
                SlideDesignPlan(
                    slide_index=slide_idx,
                    title_text=title,
                    subtitle_text=subtitle,
                    intent=intent,
                    why_chosen=why,
                    visual_priority=priority,
                    density=density_str,
                    estimated_occupancy_percent=occupancy,
                    suggested_layout_traits=traits,
                    raw_blocks=body_blocks,
                )
            )
            slide_idx += 1

        # Handle flat document blocks if no sections were generated
        if not plans and doc.flat_blocks:
            intent, why, priority, traits = self._infer_intent_and_traits(doc.title or "Overview", "", doc.flat_blocks)
            density_str, occupancy = self._estimate_density_and_occupancy(doc.flat_blocks)
            plans.append(
                SlideDesignPlan(
                    slide_index=1,
                    title_text=doc.title or "Overview",
                    subtitle_text="",
                    intent=intent,
                    why_chosen=why,
                    visual_priority=priority,
                    density=density_str,
                    estimated_occupancy_percent=occupancy,
                    suggested_layout_traits=traits,
                    raw_blocks=doc.flat_blocks,
                )
            )

        return PresentationDesignPlan(doc_title=doc.title, slides=plans)

    def _infer_intent_and_traits(
        self, title: str, subtitle: str, blocks: List[AnalyzedBlock]
    ) -> tuple[PresentationIntent, str, str, List[str]]:
        """Infer PresentationIntent, executive rationale, visual priority, and layout traits."""
        t_combined = f"{title} {subtitle}".lower()

        # Check structural block types
        has_table = any(b.block_type == ContentBlockType.TABLE for b in blocks)
        has_image = any(b.block_type == ContentBlockType.IMAGE for b in blocks)
        has_code = any(b.block_type == ContentBlockType.CODE_BLOCK for b in blocks)
        has_quote = any(b.block_type == ContentBlockType.QUOTE for b in blocks)

        if has_table:
            return (
                PresentationIntent.TABLE,
                "Tabular data block detected -> Standalone Table Layout",
                "High",
                ["Full Width Table", "Title Only", "Header Styling"],
            )
        if has_image:
            return (
                PresentationIntent.IMAGE,
                "Embedded image block detected -> Picture + Caption Layout",
                "High",
                ["Picture Frame", "Side Caption", "Widescreen Aspect Ratio"],
            )
        if has_code:
            return (
                PresentationIntent.CODE,
                "Source code block detected -> Monospace Code Layout",
                "Medium",
                ["Consolas Monospace", "Syntax Frame", "High Contrast"],
            )
        if has_quote:
            return (
                PresentationIntent.QUOTE,
                "Blockquote element detected -> Emphasis Quote Layout",
                "Medium",
                ["Centered Quote", "Italic Styling", "Author Citation"],
            )

        # Keyword heuristics
        if any(k in t_combined for k in ("executive summary", "key takeaways", "highlights", "at a glance")):
            return (
                PresentationIntent.EXECUTIVE_SUMMARY,
                "Executive Summary keywords detected -> High-priority Summary Card",
                "High",
                ["Executive Header", "Bullet List", "Highlighted Points"],
            )

        if any(k in t_combined for k in ("roadmap", "implementation plan", "phase", "next steps")):
            return (
                PresentationIntent.ROADMAP,
                "Sequential roadmap keywords detected -> Multi-phase Horizontal Flow",
                "High",
                ["Horizontal Phases", "Sequential Flow", "Multi-Block Columns"],
            )

        if any(k in t_combined for k in ("timeline", "milestone", "schedule")):
            return (
                PresentationIntent.TIMELINE,
                "Timeline/milestone keywords detected -> Sequential Timeline Layout",
                "High",
                ["Timeline Axis", "Milestone Cards", "Date Callouts"],
            )

        if any(k in t_combined for k in ("process", "workflow", "pipeline", "methodology")):
            return (
                PresentationIntent.PROCESS,
                "Process/workflow keywords detected -> Step-by-Step Process Flow",
                "Medium",
                ["Numbered Steps", "Process Arrows", "Balanced Cards"],
            )

        if any(k in t_combined for k in ("vs", "versus", "comparison", "compare", "pros", "cons", "before", "after")):
            return (
                PresentationIntent.COMPARISON,
                "Comparison keywords detected -> Side-by-side Dual Column Layout",
                "High",
                ["Two Content", "Side-by-Side Comparison", "Category Headers"],
            )

        if any(k in t_combined for k in ("metrics", "kpi", "performance", "roi", "results")):
            return (
                PresentationIntent.KPI_DASHBOARD,
                "KPI/Metrics keywords detected -> Executive Metric Callouts",
                "High",
                ["Large Numbers", "Metric Cards", "KPI Highlights"],
            )

        if any(k in t_combined for k in ("conclusion", "closing", "recommendations")):
            return (
                PresentationIntent.CONCLUSION,
                "Conclusion/Recommendations keywords detected -> Closing Action Slide",
                "High",
                ["Summary Bullets", "Action Items", "Emphasis Cards"],
            )

        # Multi-block card grid / 2 column heuristic (filtering out structural rules)
        substantive_blocks = [
            b for b in blocks
            if b.block_type not in (ContentBlockType.HORIZONTAL_RULE, ContentBlockType.UNKNOWN)
        ]

        if len(substantive_blocks) >= 4:
            return (
                PresentationIntent.CARD_GRID,
                "4+ distinct content blocks detected -> Executive 4-Card Grid",
                "High",
                ["4 Cards", "2x2 Grid", "Short Bullets"],
            )

        if len(substantive_blocks) >= 2:
            # Check if block 0 is a short intro paragraph preceding a bullet list
            has_short_intro = (
                len(substantive_blocks) == 2
                and substantive_blocks[0].block_type == ContentBlockType.PARAGRAPH_GROUP
                and substantive_blocks[0].word_count <= 20
                and substantive_blocks[1].block_type in (ContentBlockType.BULLET_LIST, ContentBlockType.ORDERED_LIST)
            )
            if not has_short_intro:
                return (
                    PresentationIntent.TWO_COLUMN,
                    "Multiple substantial content blocks detected -> Balanced 2-Column Layout",
                    "Medium",
                    ["Two Content", "Dual Columns", "Balanced Spacing"],
                )

        return (
            PresentationIntent.BULLET_CONTENT,
            "Standard paragraph/list block -> Single Column Executive Bullets",
            "Medium",
            ["Title and Content", "Level-0 Bullets", "Standard Margins"],
        )

    def _estimate_density_and_occupancy(self, blocks: List[AnalyzedBlock]) -> tuple[str, float]:
        """Estimate visual density label and canvas occupancy percentage."""
        total_items = 0
        for b in blocks:
            node = b.ast_node
            if isinstance(node, BulletListNode):
                total_items += len(node.items)
            elif isinstance(node, TableNode):
                total_items += len(node.rows) + 2
            else:
                total_items += 1

        occupancy = float(min(90.0, max(25.0, total_items * 12.5)))
        density = "Low" if total_items <= 3 else ("High" if total_items >= 7 else "Medium")
        return density, occupancy
