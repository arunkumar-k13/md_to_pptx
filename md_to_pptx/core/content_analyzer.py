"""Content Analyzer Module.

Consumes the Document AST and produces an AnalyzedDocument structure containing
classified content blocks, spatial text metrics, section hierarchies, and candidate
slide hints. Decouples raw syntax parsing from downstream presentation planning.
"""

from __future__ import annotations
import enum
import logging
from dataclasses import dataclass, field
from typing import List, Set

from md_to_pptx.core.ast_nodes import (
    ASTNode,
    BulletListNode,
    CodeBlockNode,
    Document,
    HeaderNode,
    HorizontalRuleNode,
    ImageNode,
    ParagraphNode,
    QuoteNode,
    Section,
    TableNode,
)

logger = logging.getLogger(__name__)


class ContentBlockType(enum.Enum):
    """Semantic classifications for document content blocks."""

    DOCUMENT_TITLE = "document_title"
    SECTION_HEADER = "section_header"
    SUBSECTION_HEADER = "subsection_header"
    PARAGRAPH_GROUP = "paragraph_group"
    BULLET_LIST = "bullet_list"
    ORDERED_LIST = "ordered_list"
    TABLE = "table"
    IMAGE = "image"
    CODE_BLOCK = "code_block"
    QUOTE = "quote"
    HORIZONTAL_RULE = "horizontal_rule"
    UNKNOWN = "unknown"


class SlideCandidateHint(enum.Enum):
    """Non-binding structural hints for downstream slide layout planning."""

    POTENTIAL_TITLE_SLIDE = "potential_title_slide"
    POTENTIAL_SECTION_DIVIDER = "potential_section_divider"
    POTENTIAL_CONTENT_SLIDE = "potential_content_slide"
    POTENTIAL_IMAGE_SLIDE = "potential_image_slide"
    POTENTIAL_TABLE_SLIDE = "potential_table_slide"
    POTENTIAL_CODE_SLIDE = "potential_code_slide"


@dataclass(slots=True)
class AnalyzedBlock:
    """Classified content block containing spatial metrics and candidate hints.

    Attributes:
        block_type: Semantic classification of the block.
        ast_node: Original underlying AST node.
        word_count: Total word count within this block.
        item_count: Count of sub-items (e.g. list item count or table row count).
        candidate_hints: Set of potential slide layout hints for planning.
    """

    block_type: ContentBlockType
    ast_node: ASTNode
    word_count: int = 0
    item_count: int = 0
    candidate_hints: Set[SlideCandidateHint] = field(default_factory=set)


@dataclass(slots=True)
class AnalyzedSection:
    """Logical document section containing grouped analyzed content blocks.

    Attributes:
        title: Section header title.
        level: Section heading level (e.g., 1 for major section, 2 for subsection).
        blocks: List of AnalyzedBlocks directly belonging to this section.
        subsections: Nested subsections.
        is_major_divider_candidate: True if section header is suitable for a section divider slide.
    """

    title: str
    level: int
    blocks: List[AnalyzedBlock] = field(default_factory=list)
    subsections: List[AnalyzedSection] = field(default_factory=list)
    is_major_divider_candidate: bool = False


@dataclass(slots=True)
class AnalyzedDocument:
    """Comprehensive content model produced by ContentAnalyzer.

    Attributes:
        title: Overall document title.
        sections: Hierarchical list of analyzed sections.
        flat_blocks: Linear sequence of all analyzed blocks.
        has_images: True if document contains at least one image.
        has_tables: True if document contains at least one table.
        has_code_blocks: True if document contains code blocks.
        total_word_count: Cumulative word count of entire document.
    """

    title: str = ""
    sections: List[AnalyzedSection] = field(default_factory=list)
    flat_blocks: List[AnalyzedBlock] = field(default_factory=list)
    has_images: bool = False
    has_tables: bool = False
    has_code_blocks: bool = False
    total_word_count: int = 0


class ContentAnalyzer:
    """Analyzer classifying Document AST nodes into structured AnalyzedDocument representations."""

    def analyze(self, document: Document) -> AnalyzedDocument:
        """Analyze a Document AST and return an AnalyzedDocument structure.

        Args:
            document: Document AST root instance.

        Returns:
            AnalyzedDocument instance populated with classified blocks and candidate hints.
        """
        if not document or not document.nodes:
            logger.info("Empty document analyzed.")
            return AnalyzedDocument(title=document.title or "Untitled Presentation")

        flat_blocks: List[AnalyzedBlock] = []
        has_images = False
        has_tables = False
        has_code_blocks = False
        total_word_count = 0

        for node in document.nodes:
            analyzed = self._analyze_node(node)
            flat_blocks.append(analyzed)
            total_word_count += analyzed.word_count

            if analyzed.block_type == ContentBlockType.IMAGE:
                has_images = True
            elif analyzed.block_type == ContentBlockType.TABLE:
                has_tables = True
            elif analyzed.block_type == ContentBlockType.CODE_BLOCK:
                has_code_blocks = True

        sections = self._analyze_sections(document.sections, doc_title=document.title or "")

        return AnalyzedDocument(
            title=document.title,
            sections=sections,
            flat_blocks=flat_blocks,
            has_images=has_images,
            has_tables=has_tables,
            has_code_blocks=has_code_blocks,
            total_word_count=total_word_count,
        )

    def _analyze_node(self, node: ASTNode) -> AnalyzedBlock:
        """Classify an individual ASTNode and compute metrics/hints.

        Args:
            node: Input ASTNode instance.

        Returns:
            AnalyzedBlock populated with type, metrics, and candidate hints.
        """
        if isinstance(node, HeaderNode):
            words = len(node.text.split())
            if node.level == 1:
                btype = ContentBlockType.SECTION_HEADER
                hints = {SlideCandidateHint.POTENTIAL_TITLE_SLIDE, SlideCandidateHint.POTENTIAL_SECTION_DIVIDER}
            else:
                btype = ContentBlockType.SUBSECTION_HEADER
                hints = {SlideCandidateHint.POTENTIAL_CONTENT_SLIDE}
            return AnalyzedBlock(block_type=btype, ast_node=node, word_count=words, candidate_hints=hints)

        elif isinstance(node, ParagraphNode):
            words = len(node.text.split())
            hints = {SlideCandidateHint.POTENTIAL_CONTENT_SLIDE}
            return AnalyzedBlock(
                block_type=ContentBlockType.PARAGRAPH_GROUP, ast_node=node, word_count=words, candidate_hints=hints
            )

        elif isinstance(node, BulletListNode):
            btype = ContentBlockType.ORDERED_LIST if node.is_ordered else ContentBlockType.BULLET_LIST
            items_count = len(node.items)
            words = sum(len(item.text.split()) for item in node.items)
            hints = {SlideCandidateHint.POTENTIAL_CONTENT_SLIDE}
            return AnalyzedBlock(
                block_type=btype, ast_node=node, word_count=words, item_count=items_count, candidate_hints=hints
            )

        elif isinstance(node, TableNode):
            row_count = len(node.rows)
            col_count = len(node.headers) if node.headers else (len(node.rows[0]) if node.rows else 0)
            cell_words = sum(len(h.split()) for h in node.headers) + sum(
                len(cell.split()) for row in node.rows for cell in row
            )
            hints = {SlideCandidateHint.POTENTIAL_TABLE_SLIDE}
            return AnalyzedBlock(
                block_type=ContentBlockType.TABLE,
                ast_node=node,
                word_count=cell_words,
                item_count=row_count * col_count,
                candidate_hints=hints,
            )

        elif isinstance(node, ImageNode):
            hints = {SlideCandidateHint.POTENTIAL_IMAGE_SLIDE}
            words = len(node.alt.split()) + len(node.title.split())
            return AnalyzedBlock(
                block_type=ContentBlockType.IMAGE, ast_node=node, word_count=words, item_count=1, candidate_hints=hints
            )

        elif isinstance(node, CodeBlockNode):
            line_count = len(node.code.splitlines()) if node.code else 0
            words = len(node.code.split())
            hints = {SlideCandidateHint.POTENTIAL_CODE_SLIDE}
            return AnalyzedBlock(
                block_type=ContentBlockType.CODE_BLOCK,
                ast_node=node,
                word_count=words,
                item_count=line_count,
                candidate_hints=hints,
            )

        elif isinstance(node, QuoteNode):
            words = len(node.text.split())
            hints = {SlideCandidateHint.POTENTIAL_CONTENT_SLIDE}
            return AnalyzedBlock(
                block_type=ContentBlockType.QUOTE, ast_node=node, word_count=words, candidate_hints=hints
            )

        elif isinstance(node, HorizontalRuleNode):
            hints = {SlideCandidateHint.POTENTIAL_SECTION_DIVIDER}
            return AnalyzedBlock(
                block_type=ContentBlockType.HORIZONTAL_RULE, ast_node=node, word_count=0, candidate_hints=hints
            )

        else:
            logger.warning("Unknown ASTNode encountered during content analysis: %s", type(node))
            return AnalyzedBlock(
                block_type=ContentBlockType.UNKNOWN, ast_node=node, candidate_hints={SlideCandidateHint.POTENTIAL_CONTENT_SLIDE}
            )

    def _is_major_section_header_text(self, text: str) -> bool:
        t = text.strip()
        import re
        pattern = r"^\s*(\*\*|\#\#\s*)?(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+[A-Z]"
        return bool(re.match(pattern, t))

    def _analyze_sections(self, sections: List[Section], doc_title: str = "") -> List[AnalyzedSection]:
        """Convert AST Section structures into AnalyzedSection objects.

        Args:
            sections: List of AST Section instances.
            doc_title: Main document title string.

        Returns:
            List of AnalyzedSection objects.
        """
        # Merge preamble section without header preceding initial slide section
        if len(sections) > 1 and sections[0].header is None:
            sections[1].nodes = sections[0].nodes + sections[1].nodes
            sections = sections[1:]

        analyzed_sections: List[AnalyzedSection] = []

        for sec in sections:
            title = sec.header.text if sec.header else "General Content"
            level = sec.header.level if sec.header else 2
            is_major = (level == 1)

            doc_title_clean = doc_title.strip()
            sec_blocks = [
                self._analyze_node(n) for n in sec.nodes
                if n != sec.header and (not hasattr(n, "text") or getattr(n, "text", "").strip() != doc_title_clean)
            ]

            current_title = title
            current_blocks: List[AnalyzedBlock] = []

            for b in sec_blocks:
                b_text = getattr(b.ast_node, "text", "").strip() if hasattr(b.ast_node, "text") else ""

                if b.block_type == ContentBlockType.PARAGRAPH_GROUP and self._is_major_section_header_text(b_text):
                    if current_blocks:
                        analyzed_sections.append(
                            AnalyzedSection(
                                title=current_title,
                                level=level,
                                blocks=current_blocks,
                                is_major_divider_candidate=is_major,
                            )
                        )
                        current_blocks = []
                    current_title = b_text
                    current_blocks.append(b)

                elif b.block_type == ContentBlockType.BULLET_LIST and isinstance(b.ast_node, BulletListNode):
                    # Check if list contains any major Roman Numeral headers inside list items
                    sub_items: List[ListItem] = []
                    for item in b.ast_node.items:
                        if self._is_major_section_header_text(item.text):
                            if current_blocks or sub_items:
                                if sub_items:
                                    sub_node = BulletListNode(items=sub_items, is_ordered=b.ast_node.is_ordered)
                                    current_blocks.append(self._analyze_node(sub_node))
                                analyzed_sections.append(
                                    AnalyzedSection(
                                        title=current_title,
                                        level=level,
                                        blocks=current_blocks,
                                        is_major_divider_candidate=is_major,
                                    )
                                )
                                current_blocks = []
                                sub_items = []
                            current_title = item.text
                        else:
                            sub_items.append(item)

                    if sub_items:
                        sub_node = BulletListNode(items=sub_items, is_ordered=b.ast_node.is_ordered)
                        current_blocks.append(self._analyze_node(sub_node))

                else:
                    current_blocks.append(b)

            if current_blocks or (sec.header and sec.header.text):
                analyzed_sections.append(
                    AnalyzedSection(
                        title=current_title,
                        level=level,
                        blocks=current_blocks,
                        is_major_divider_candidate=is_major,
                    )
                )

        return analyzed_sections
