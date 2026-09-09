"""Presentation Intermediate Representation (IR) Model.

Defines target-agnostic presentation domain dataclasses based on an abstract
ContentBlock hierarchy. Completely decouples Markdown AST parsing from specific
presentation rendering engines (PowerPoint, Google Slides, HTML, PDF).
"""

from __future__ import annotations
import enum
import uuid
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from md_to_pptx.core.ast_nodes import InlineRun, ListItem


class SlideIntent(enum.Enum):
    """Semantic intent classification for logical presentation slides."""

    TITLE_SLIDE = "title_slide"
    SECTION_DIVIDER = "section_divider"
    EXECUTIVE_SUMMARY = "executive_summary"
    OVERVIEW = "overview"
    BULLET_CONTENT = "bullet_content"
    COMPARISON = "comparison"
    TWO_COLUMN = "two_column"
    TABLE_SLIDE = "table_slide"
    IMAGE_SLIDE = "image_slide"
    QUOTE_SLIDE = "quote_slide"
    CODE_SLIDE = "code_slide"
    ROADMAP = "roadmap"
    TIMELINE = "timeline"
    PROCESS = "process"
    METRICS = "metrics"
    THANK_YOU_SLIDE = "thank_you_slide"
    CONCLUSION = "conclusion"
    CONTENT_SLIDE = "content_slide"


@dataclass(slots=True)
class ContentBlock:
    """Abstract base class for all target-agnostic presentation content blocks.

    Attributes:
        block_id: Unique string identifier for this block instance.
        style_overrides: Dictionary of optional visual styling overrides.
    """

    block_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    style_overrides: Dict[str, Any] = field(default_factory=dict)

    def estimate_height(self, available_width_inches: float = 8.0) -> float:
        """Estimate visual rendering height in inches based on block metrics.

        Args:
            available_width_inches: Width of the target placeholder frame in inches.

        Returns:
            Estimated height in inches.
        """
        return 0.5


@dataclass(slots=True)
class TitleBlock(ContentBlock):
    """Represents a slide heading or major title block.

    Attributes:
        text: Heading text content.
        level: Title hierarchy level (1 for document title, 2 for slide title, 3+ for subtitles).
    """

    text: str = ""
    level: int = 1

    def estimate_height(self, available_width_inches: float = 8.0) -> float:
        chars_per_line = max(1, int(available_width_inches * 12))  # ~12 chars/inch for titles
        lines = max(1, (len(self.text) + chars_per_line - 1) // chars_per_line)
        return float(lines * 0.45)


@dataclass(slots=True)
class ParagraphBlock(ContentBlock):
    """Represents a body paragraph block.

    Attributes:
        text: Plain text content.
        runs: Formatted inline runs.
    """

    text: str = ""
    runs: List[InlineRun] = field(default_factory=list)

    def estimate_height(self, available_width_inches: float = 8.0) -> float:
        font_size_pt = 16.0
        line_spacing = 1.15
        char_width_in = (font_size_pt * 0.52) / 72.0
        line_height_in = (font_size_pt * line_spacing) / 72.0
        eff_width = max(2.0, available_width_inches - 0.2)
        chars_per_line = max(15, int(eff_width / char_width_in))
        lines = max(1, (len(self.text) + chars_per_line - 1) // chars_per_line) if self.text else 1
        return float(lines * line_height_in + 0.12)


@dataclass(slots=True)
class BulletListBlock(ContentBlock):
    """Represents a bullet or numbered list block.

    Attributes:
        items: List of item instances.
        is_ordered: True if numbered list, False if bullet list.
    """

    items: List[ListItem] = field(default_factory=list)
    is_ordered: bool = False
    start_index: int = 1

    def estimate_height(self, available_width_inches: float = 8.0) -> float:
        font_size_pt = 20.0
        line_spacing = 1.20
        char_width_in = (font_size_pt * 0.56) / 72.0
        line_height_in = (font_size_pt * line_spacing) / 72.0
        para_space_in = (4.0 + 10.0) / 72.0

        def _calc_item_height(item: ListItem) -> float:
            level = max(0, getattr(item, "level", 0))
            level_indent = 0.35 * (level + 1)
            eff_width = max(2.0, available_width_inches - 0.2 - level_indent)
            chars_per_line = max(10, int(eff_width / char_width_in))

            text_str = item.text.strip() if item.text else ""
            lines = max(1, (len(text_str) + chars_per_line - 1) // chars_per_line) if text_str else 1
            h = (lines * line_height_in) + para_space_in

            if hasattr(item, "children") and item.children:
                for c in item.children:
                    if isinstance(c, ListItem):
                        h += _calc_item_height(c)
                    elif hasattr(c, "items"):
                        for sub_i in getattr(c, "items", []):
                            if isinstance(sub_i, ListItem):
                                h += _calc_item_height(sub_i)
            return h

        return sum(_calc_item_height(item) for item in self.items) + 0.20


@dataclass(slots=True)
class ImageBlock(ContentBlock):
    """Represents an embedded image block.

    Attributes:
        src: File path or URL of the image.
        alt: Alternate text description.
        title: Optional title attribute.
        aspect_ratio: Width-to-height aspect ratio (default 1.33).
    """

    src: str = ""
    alt: str = ""
    title: str = ""
    aspect_ratio: float = 1.33

    def estimate_height(self, available_width_inches: float = 8.0) -> float:
        return float(min(4.5, available_width_inches / self.aspect_ratio))


@dataclass(slots=True)
class TableBlock(ContentBlock):
    """Represents a tabular data block.

    Attributes:
        headers: List of column title strings.
        rows: List of row cell values.
    """

    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    raw_rows: List[Any] = field(default_factory=list)

    def estimate_height(self, available_width_inches: float = 8.0) -> float:
        all_rows = []
        if self.headers:
            all_rows.append(self.headers)
        all_rows.extend(self.rows)
        if not all_rows:
            return 1.0

        num_cols = max(len(r) for r in all_rows)
        col_w = available_width_inches / max(1, num_cols)
        chars_per_col_line = max(8, int(col_w * 12))

        total_height = 0.0
        for row in all_rows:
            row_lines = 1
            for cell in row:
                cell_len = len(str(cell))
                lines = max(1, (cell_len + chars_per_col_line - 1) // chars_per_col_line)
                if lines > row_lines:
                    row_lines = lines
            total_height += row_lines * 0.32 + 0.12
        return float(max(1.0, total_height))


@dataclass(slots=True)
class CodeBlock(ContentBlock):
    """Represents a code snippet block.

    Attributes:
        code: Source code text string.
        language: Programming or markup language tag.
    """

    code: str = ""
    language: str = ""

    def estimate_height(self, available_width_inches: float = 8.0) -> float:
        line_count = max(1, len(self.code.splitlines()))
        return float(min(4.5, line_count * 0.25 + 0.2))


@dataclass(slots=True)
class QuoteBlock(ContentBlock):
    """Represents a blockquote element.

    Attributes:
        text: Quote text content.
        author: Optional author attribution.
    """

    text: str = ""
    author: str = ""
    runs: List[InlineRun] = field(default_factory=list)

    def estimate_height(self, available_width_inches: float = 8.0) -> float:
        chars_per_line = max(1, int(available_width_inches * 15))
        lines = max(1, (len(self.text) + chars_per_line - 1) // chars_per_line)
        return float(lines * 0.32 + 0.2)


@dataclass(slots=True)
class Slide:
    """Represents an individual logical presentation slide.

    Attributes:
        slide_id: Unique slide identifier string.
        intent: Logical slide intent (e.g. TITLE_SLIDE, CONTENT_SLIDE).
        title: Optional title block for this slide.
        blocks: Sequence of content blocks rendering on this slide.
        is_continuation: True if slide was spawned as an overflow continuation.
        continuation_index: Continuation sequence index (0 for original, 1 for first continuation).
        brand_role: Abstract brand role identifier ('BRAND_COVER', 'BRAND_DISCLAIMER', etc.).
        brand_slide_index: Index of pre-existing slide in template presentation slides array.
        brand_layout_index: Index of slide layout in template presentation slide_layouts array.
    """

    slide_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    intent: SlideIntent = SlideIntent.CONTENT_SLIDE
    title: Optional[TitleBlock] = None
    subtitle: str = ""
    blocks: List[ContentBlock] = field(default_factory=list)
    is_continuation: bool = False
    continuation_index: int = 0
    brand_role: Optional[str] = None
    brand_slide_index: Optional[int] = None
    brand_layout_index: Optional[int] = None


@dataclass(slots=True)
class Presentation:
    """Root presentation domain model container.

    Attributes:
        title: Presentation title string.
        subtitle: Optional subtitle string.
        slides: List of logical Slide instances.
        metadata: Arbitrary metadata dictionary.
    """

    title: str = ""
    subtitle: str = ""
    slides: List[Slide] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
