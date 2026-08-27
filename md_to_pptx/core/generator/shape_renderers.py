"""Shape Renderers Module.

Modular renderers responsible for populating python-pptx shapes (TextFrames,
Tables, Pictures) from target-agnostic ContentBlock IR objects while preserving
template master font formatting, line spacing, margins, and theme colors.
"""

from __future__ import annotations
import logging
from pathlib import Path
from typing import Any

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from md_to_pptx.core.ast_nodes import BulletListNode, InlineRun, ListItem
from md_to_pptx.core.presentation_model import (
    BulletListBlock,
    CodeBlock,
    ImageBlock,
    ParagraphBlock,
    QuoteBlock,
    TableBlock,
    TitleBlock,
)

logger = logging.getLogger(__name__)


def render_title_block(shape: Any, title_block: TitleBlock) -> None:
    """Render a TitleBlock into a title shape textframe with dynamic font scaling.

    Args:
        shape: PowerPoint title placeholder shape.
        title_block: Input TitleBlock IR instance.
    """
    if not hasattr(shape, "text_frame"):
        return

    tf = shape.text_frame
    tf.word_wrap = True
    if tf.paragraphs:
        p = tf.paragraphs[0]
        p.text = title_block.text

        # Auto-scale font size for long titles matching 'About Company' template slide title size (80pt base)
        t_len = len(title_block.text)
        if t_len > 110:
            target_size = Pt(36)
        elif t_len > 75:
            target_size = Pt(48)
        elif t_len > 40:
            target_size = Pt(60)
        else:
            target_size = Pt(80)

        p.font.name = "Arial"
        p.font.size = target_size

        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = target_size


def render_paragraph_block(text_frame: Any, para_block: ParagraphBlock, slide_height_inches: float = 7.5) -> None:
    """Render a ParagraphBlock into a TextFrame preserving master styles and spacing.

    Args:
        text_frame: PowerPoint TextFrame object.
        para_block: Input ParagraphBlock IR instance.
        slide_height_inches: Total presentation slide height for proportional font scaling.
    """
    scale = max(1.0, slide_height_inches / 7.5)
    base_font_pt = 18.0 * min(1.35, scale)

    p = text_frame.add_paragraph() if text_frame.paragraphs[0].text else text_frame.paragraphs[0]
    p.level = 0
    p.space_after = Pt(int(6 * scale))
    p.space_before = Pt(int(2 * scale))
    p.line_spacing = 1.15

    if not para_block.runs:
        p.text = para_block.text
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(int(base_font_pt))
        return

    p.text = ""
    for run in para_block.runs:
        r = p.add_run()
        r.text = run.text
        r.font.name = "Consolas" if run.is_code else "Arial"
        r.font.size = Pt(int(base_font_pt))
        if run.is_bold:
            r.font.bold = True
        if run.is_italic:
            r.font.italic = True
        if run.url:
            r.hyperlink.address = run.url


def render_bullet_list_block(text_frame: Any, list_block: BulletListBlock, slide_height_inches: float = 7.5) -> None:
    """Render a BulletListBlock into a TextFrame preserving nested list levels and template spacing.

    Args:
        text_frame: PowerPoint TextFrame object.
        list_block: Input BulletListBlock IR instance.
        slide_height_inches: Total presentation slide height for proportional font scaling.
    """
    scale = max(1.0, slide_height_inches / 7.5)

    def _render_items(items: list[ListItem], depth: int = 0):
        for item in items:
            p = text_frame.add_paragraph() if text_frame.paragraphs[0].text else text_frame.paragraphs[0]
            item_depth = max(depth, getattr(item, "level", depth))
            p.level = min(4, item_depth)
            p.space_after = Pt(int(4 * scale))
            p.space_before = Pt(int(2 * scale))
            p.line_spacing = 1.15

            base_pt = (20.0 if item_depth == 0 else 18.0 if item_depth == 1 else 16.0) * min(1.35, scale)
            target_font_size = Pt(int(base_pt))

            if not item.runs:
                p.text = item.text
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = target_font_size
            else:
                p.text = ""
                for run in item.runs:
                    r = p.add_run()
                    r.text = run.text
                    r.font.name = "Consolas" if run.is_code else "Arial"
                    r.font.size = target_font_size
                    if run.is_bold:
                        r.font.bold = True
                    if run.is_italic:
                        r.font.italic = True
                    if run.url:
                        r.hyperlink.address = run.url

            if item.children:
                for child_node in item.children:
                    if isinstance(child_node, ListItem):
                        _render_items([child_node], depth=item_depth + 1)
                    elif isinstance(child_node, BulletListNode):
                        _render_items(child_node.items, depth=item_depth + 1)

    _render_items(list_block.items, depth=0)


def render_table_block(
    slide: Any,
    table_block: TableBlock,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
) -> None:
    """Render a TableBlock as a corporate-styled table shape with column auto-balancing and padding.

    Args:
        slide: Target PowerPoint Slide object.
        table_block: Input TableBlock IR instance.
        left_in: X position in inches.
        top_in: Y position in inches.
        width_in: Table width in inches.
        height_in: Table height in inches.
    """
    all_rows = []
    if table_block.headers:
        all_rows.append(table_block.headers)
    all_rows.extend(table_block.rows)

    if not all_rows:
        return

    num_rows = len(all_rows)
    num_cols = max(len(r) for r in all_rows)

    table_shape = slide.shapes.add_table(
        num_rows, num_cols, Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
    )
    tbl = table_shape.table

    # Auto column width distribution based on content length and longest single word floor
    col_max_lens = [
        max(len(str(row[c])) if c < len(row) else 1 for row in all_rows)
        for c in range(num_cols)
    ]
    total_chars = sum(col_max_lens) or 1

    col_word_floors = []
    for c in range(num_cols):
        words = [
            w for row in all_rows if c < len(row)
            for w in str(row[c]).replace("-", " ").split()
        ]
        max_word_len = max([len(w) for w in words] + [5])
        col_word_floors.append(max(1.8, (max_word_len * 0.20) + 0.40))

    for c_idx in range(num_cols):
        fraction = col_max_lens[c_idx] / total_chars
        calc_w = width_in * fraction
        final_col_w = max(col_word_floors[c_idx], calc_w)
        tbl.columns[c_idx].width = Inches(final_col_w)

    # MC Corporate Dark Blue (#004065) & Alternating Shading (#F2F4F7)
    header_bg_color = RGBColor(0, 64, 101)
    alt_row_color = RGBColor(242, 244, 247)

    for r_idx, row in enumerate(all_rows):
        is_header = (r_idx == 0 and bool(table_block.headers))
        for c_idx, val in enumerate(row):
            if c_idx < num_cols:
                cell = tbl.cell(r_idx, c_idx)
                cell.text = str(val)
                cell.margin_left = Inches(0.12)
                cell.margin_right = Inches(0.12)
                cell.margin_top = Inches(0.08)
                cell.margin_bottom = Inches(0.08)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE

                if is_header:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = header_bg_color
                    for p in cell.text_frame.paragraphs:
                        p.alignment = PP_ALIGN.LEFT
                        for run in p.runs:
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(255, 255, 255)
                elif r_idx % 2 == 1:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = alt_row_color


def render_image_block(
    slide: Any,
    image_block: ImageBlock,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
) -> None:
    """Render an ImageBlock onto a slide centered inside placeholder bounds while preserving aspect ratio.

    Args:
        slide: Target PowerPoint Slide object.
        image_block: Input ImageBlock IR instance.
        left_in: X position in inches.
        top_in: Y position in inches.
        width_in: Placeholder width in inches.
        height_in: Placeholder height in inches.
    """
    img_path = Path(image_block.src)
    if not img_path.is_file():
        logger.warning("Image file not found at %s. Skipping picture rendering.", image_block.src)
        return

    aspect = getattr(image_block, "aspect_ratio", 1.33)
    final_w = min(width_in, height_in * aspect)
    final_h = min(height_in, width_in / aspect)
    offset_x = left_in + (width_in - final_w) / 2.0
    offset_y = top_in + (height_in - final_h) / 2.0

    try:
        slide.shapes.add_picture(
            str(img_path.resolve()),
            Inches(offset_x),
            Inches(offset_y),
            width=Inches(final_w),
            height=Inches(final_h),
        )
    except Exception as err:
        logger.error("Failed to render image %s: %s", image_block.src, err, exc_info=True)


from pptx.oxml import parse_xml

def render_code_block(text_frame: Any, code_block: CodeBlock) -> None:
    """Render a CodeBlock into a TextFrame formatted with monospace font.

    Args:
        text_frame: PowerPoint TextFrame object.
        code_block: Input CodeBlock IR instance.
    """
    lines = code_block.code.split("\n")
    first = True
    for line in lines:
        p = text_frame.add_paragraph() if (not first or text_frame.paragraphs[0].text) else text_frame.paragraphs[0]
        first = False
        p.text = line
        p.level = 0
        p.space_after = Pt(2)
        p.space_before = Pt(0)
        p.line_spacing = 1.0
        try:
            pPr = p._p.get_or_add_pPr()
            for child in list(pPr):
                if child.tag.endswith(("buChar", "buAutoNum", "buNone", "buBlip")):
                    pPr.remove(child)
            pPr.insert(0, parse_xml('<a:buNone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
        except Exception:
            pass
        for r in p.runs:
            r.font.name = "Consolas"
            r.font.size = Pt(11)
