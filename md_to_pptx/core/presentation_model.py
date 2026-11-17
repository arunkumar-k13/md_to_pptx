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
from md_to_pptx.core.typography import (
    DEFAULT_TEXT_INSET_LEFT_RIGHT_IN,
    STANDARD,
    Typography,
    paragraph_height_in,
    split_sentences,
)


def heading_level_of(block: Any) -> Optional[int]:
    """Heading level of a paragraph block as the renderer styles it.

    Returns the Markdown heading level (1-6), 0 for an all-bold label paragraph
    (rendered at heading size), or None for ordinary body text.
    """
    overrides = getattr(block, "style_overrides", None) or {}
    level = overrides.get("heading_level", 0)
    if level and level > 0:
        return int(level)
    runs = getattr(block, "runs", None)
    if runs and all(getattr(r, "is_bold", False) for r in runs if getattr(r, "text", "").strip()):
        return 0
    return None


def _text_width(available_width_inches: float) -> float:
    """Usable line width inside a text box after the default left/right insets."""
    return max(1.0, available_width_inches - 2 * DEFAULT_TEXT_INSET_LEFT_RIGHT_IN)


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

    def estimate_height(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> float:
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

    def estimate_height(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> float:
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

    def estimate_height(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> float:
        typo = typography or STANDARD
        level = heading_level_of(self)
        if level is None:
            size, before, after, bold = typo.body_pt, typo.para_before_pt, typo.para_after_pt, False
        else:
            size, before, after, bold = typo.heading_size(level), typo.heading_before_pt, typo.heading_after_pt, True
        # Body paragraphs sit at the master's first bullet level, so they share its left margin
        width = max(1.0, _text_width(available_width_inches) - typo.list_indent(0))
        return float(paragraph_height_in(
            self.text, width, size, typo.line_spacing, before, after, bold=bold, runs=self.runs,
        ))

    def split_to_fit(
        self, available_width_inches: float, max_height_inches: float, typography: Optional[Typography] = None
    ) -> Optional[tuple]:
        """Split into (fits, rest) at the last sentence (or word) boundary that fits.

        Returns None when not even one sentence/word fits or no split is possible.
        """
        typo = typography or STANDARD
        units = split_sentences(self.text)
        joiner = " "
        if len(units) < 2:
            units = self.text.split(" ")
        if len(units) < 2:
            return None
        best = 0
        for k in range(1, len(units)):
            head = ParagraphBlock(text=joiner.join(units[:k]), style_overrides=dict(self.style_overrides))
            if head.estimate_height(available_width_inches, typo) <= max_height_inches:
                best = k
            else:
                break
        if best == 0:
            return None
        head_text = joiner.join(units[:best])
        tail_text = joiner.join(units[best:])
        # Formatting runs cannot be cut reliably mid-run; continuation pieces keep plain text styling
        plain = [InlineRun(text=head_text)], [InlineRun(text=tail_text)]
        return (
            ParagraphBlock(text=head_text, runs=plain[0], style_overrides=dict(self.style_overrides)),
            ParagraphBlock(text=tail_text, runs=plain[1], style_overrides=dict(self.style_overrides)),
        )


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

    def item_heights(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> List[float]:
        """Rendered height of each list item (items are flattened, nesting is in `level`)."""
        typo = typography or STANDARD
        width = _text_width(available_width_inches)
        heights: List[float] = []
        counters: Dict[int, int] = {}
        for item in self.items:
            depth = max(0, min(4, getattr(item, "level", 0)))
            ordered = bool(getattr(item, "is_ordered", self.is_ordered))
            counters[depth] = counters.get(depth, 0) + 1
            for lvl in [k for k in counters if k > depth]:
                counters[lvl] = 0
            text = (item.text or "").strip()
            size = typo.list_size(depth)
            if ordered:
                text = f"{counters[depth]}. {text}"
            line_w = max(1.0, width - typo.list_indent(depth))
            runs = list(getattr(item, "runs", None) or [])
            if ordered and runs:
                runs = [InlineRun(text=f"{counters[depth]}. ")] + runs
            h = paragraph_height_in(
                text, line_w, size, typo.line_spacing, typo.list_before_pt, typo.list_after_pt, runs=runs or None
            )
            for child in getattr(item, "children", None) or []:
                if isinstance(child, ListItem):
                    h += BulletListBlock(items=[child], is_ordered=ordered).estimate_height(available_width_inches, typo)
            heights.append(h)
        return heights

    def split_indices(self) -> List[int]:
        """Indices where the list may be split: before a top-level item, never inside a nested group."""
        min_level = min((getattr(i, "level", 0) for i in self.items), default=0)
        return [k for k in range(1, len(self.items)) if getattr(self.items[k], "level", 0) == min_level]

    def estimate_height(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> float:
        return float(sum(self.item_heights(available_width_inches, typography)))


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

    def estimate_height(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> float:
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

    def estimate_height(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> float:
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

    def estimate_height(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> float:
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

    def estimate_height(self, available_width_inches: float = 8.0, typography: Optional[Typography] = None) -> float:
        typo = typography or STANDARD
        quoted = f'"{self.text}"'
        width = max(1.0, _text_width(available_width_inches) - typo.list_indent(0))
        return float(paragraph_height_in(
            quoted, width, typo.quote_pt, 1.0, typo.quote_before_pt, typo.quote_after_pt,
        ))


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
