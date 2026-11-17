"""Shape Renderers Module.

Modular renderers responsible for populating python-pptx shapes (TextFrames,
Tables, Pictures) from target-agnostic ContentBlock IR objects while preserving
template master font formatting, line spacing, margins, and theme colors.
"""

from __future__ import annotations
import logging
import re
from pathlib import Path
from typing import Any, Optional

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt

from md_to_pptx.core.ast_nodes import BulletListNode, ListItem
from md_to_pptx.core.presentation_model import (
    BulletListBlock,
    CodeBlock,
    ImageBlock,
    ParagraphBlock,
    TableBlock,
    TitleBlock,
    heading_level_of,
)
from md_to_pptx.core.typography import typography_for

logger = logging.getLogger(__name__)


def apply_font_family(run_obj: Any, text_str: str, is_code: bool = False, default_size_pt: Optional[float] = None) -> None:
    """Apply Segoe UI Emoji font for emoji characters or Arial/Consolas for standard text."""
    if not text_str:
        return
    has_emoji = any(
        (0x1F300 <= ord(c) <= 0x1F9FF) or 
        (0x2600 <= ord(c) <= 0x27BF) or 
        (0x1F000 <= ord(c) <= 0x1FFFF) or
        (0x2000 <= ord(c) <= 0x32FF)
        for c in text_str
    )
    if is_code:
        run_obj.font.name = "Consolas"
    elif has_emoji:
        run_obj.font.name = "Segoe UI Emoji"
    else:
        run_obj.font.name = "Arial"

    if default_size_pt is not None:
        run_obj.font.size = Pt(int(default_size_pt))


def render_title_block(shape: Any, title_block: TitleBlock, slide_height_inches: float = 7.5) -> None:
    """Render a TitleBlock into a title shape textframe with dynamic font scaling.

    Calibrates font size according to placeholder dimensions and slide hierarchy so titles
    never overflow their bounding box or collide with body content.

    Args:
        shape: PowerPoint title placeholder shape.
        title_block: Input TitleBlock IR instance.
        slide_height_inches: Total presentation slide height for proportional font scaling.
    """
    if not hasattr(shape, "text_frame"):
        return

    tf = shape.text_frame
    tf.word_wrap = True
    try:
        tf.margin_left = Inches(0.0)
        tf.margin_right = Inches(0.0)
        tf.margin_top = Inches(0.0)
        tf.margin_bottom = Inches(0.0)
    except Exception:
        pass

    if tf.paragraphs:
        p = tf.paragraphs[0]
        p.text = title_block.text
        p.line_spacing = 1.0
        p.space_after = Pt(0)
        p.space_before = Pt(0)

        t_len = len(title_block.text)
        box_h = shape.height.inches if hasattr(shape, "height") and shape.height else 1.25

        # Standard slide title placeholder (height < 1.8 inches, typical for 7.5" standard slides)
        if box_h < 1.8:
            if title_block.level == 1:
                # Cover / Title slide
                if t_len > 80:
                    target_size = Pt(24)
                elif t_len > 45:
                    target_size = Pt(28)
                else:
                    target_size = Pt(34)
            else:
                # Content slide title: 20pt - 26pt ensures single or clean 2-line title within ~1 inch height
                if t_len > 75:
                    target_size = Pt(20)
                elif t_len > 45:
                    target_size = Pt(22)
                elif t_len > 25:
                    target_size = Pt(24)
                else:
                    target_size = Pt(26)
        else:
            # Oversized canvas (like MC template's 2.9" title box on 15" canvas)
            scale = max(1.0, slide_height_inches / 7.5)
            if title_block.level == 1:
                if t_len > 80:
                    target_size = Pt(int(32 * min(1.35, scale)))
                elif t_len > 45:
                    target_size = Pt(int(40 * min(1.35, scale)))
                else:
                    target_size = Pt(int(48 * min(1.35, scale)))
            else:
                if t_len > 110:
                    target_size = Pt(24)
                elif t_len > 75:
                    target_size = Pt(30)
                elif t_len > 40:
                    target_size = Pt(36)
                else:
                    target_size = Pt(40)

        p.font.name = "Arial"
        p.font.size = target_size

        for r in p.runs:
            apply_font_family(r, r.text, is_code=False, default_size_pt=target_size.pt)


def render_paragraph_block(text_frame: Any, para_block: ParagraphBlock, slide_height_inches: float = 7.5) -> None:
    """Render a ParagraphBlock into a TextFrame preserving master styles and spacing.

    Args:
        text_frame: PowerPoint TextFrame object.
        para_block: Input ParagraphBlock IR instance.
        slide_height_inches: Total presentation slide height for proportional font scaling.
    """
    p = text_frame.add_paragraph() if text_frame.paragraphs[0].text else text_frame.paragraphs[0]

    typo = typography_for(slide_height_inches)
    p.level = 0
    p.line_spacing = typo.line_spacing

    # Heading level comes from the Markdown AST (or 0 for an all-bold label paragraph);
    # the same rule drives the overflow height estimate, so sizes always agree.
    h_level = heading_level_of(para_block)
    is_heading_block = h_level is not None
    is_h5 = h_level == 5
    is_h6 = h_level == 6

    if is_heading_block:
        base_font_pt = typo.heading_size(h_level)
        p.space_before = Pt(typo.heading_before_pt)
        p.space_after = Pt(typo.heading_after_pt)
    else:
        base_font_pt = typo.body_pt
        p.space_before = Pt(typo.para_before_pt)
        p.space_after = Pt(typo.para_after_pt)

    if not para_block.runs:
        p.text = para_block.text
        for r in p.runs:
            apply_font_family(r, r.text, is_code=False, default_size_pt=base_font_pt)
            if is_heading_block:
                r.font.bold = True
                r.font.color.rgb = RGBColor(11, 79, 108)
                if is_h5 or is_h6:
                    r.font.italic = True
        return

    p.text = ""
    for run in para_block.runs:
        r = p.add_run()
        r.text = run.text
        apply_font_family(r, run.text, is_code=run.is_code, default_size_pt=base_font_pt)
        if run.is_bold or is_heading_block:
            r.font.bold = True
            if is_heading_block:
                r.font.color.rgb = RGBColor(11, 79, 108)
        if run.is_italic or is_h5 or is_h6:
            r.font.italic = True
        if getattr(run, "is_strikethrough", False):
            try:
                r.font.strikethrough = True
            except Exception:
                pass
            rPr = r._r.get_or_add_rPr()
            rPr.set("strike", "sngStrike")
        if run.url:
            r.hyperlink.address = run.url


def render_bullet_list_block(text_frame: Any, list_block: BulletListBlock, slide_height_inches: float = 7.5) -> None:
    """Render a BulletListBlock into a TextFrame preserving nested list levels and template spacing.

    Args:
        text_frame: PowerPoint TextFrame object.
        list_block: Input BulletListBlock IR instance.
        slide_height_inches: Total presentation slide height for proportional font scaling.
    """
    typo = typography_for(slide_height_inches)
    sp_after = Pt(typo.list_after_pt)
    sp_before = Pt(typo.list_before_pt)

    def _render_items(items: list[ListItem], depth: int = 0):
        # Numbering belongs to this list only. Continuation slides carry the
        # correct offset in list_block.start_index (set by OverflowStrategy).
        level_counters: dict[int, int] = {0: getattr(list_block, "start_index", 1) - 1}
        level_kinds: dict[int, bool] = {}

        for item in items:
            p = text_frame.add_paragraph() if text_frame.paragraphs[0].text else text_frame.paragraphs[0]
            item_depth = max(depth, getattr(item, "level", depth))
            p.level = min(4, item_depth)

            # Reset counters for deeper levels when returning to a shallower level
            for lvl in list(level_counters.keys()):
                if lvl > item_depth:
                    level_counters[lvl] = 0

            item_kind = bool(getattr(item, "is_ordered", list_block.is_ordered))
            # A switch between bulleted and numbered items at the same depth starts a new sequence
            if item_depth > 0 and level_kinds.get(item_depth) not in (None, item_kind):
                level_counters[item_depth] = 0
            level_kinds[item_depth] = item_kind
            level_counters[item_depth] = level_counters.get(item_depth, 0) + 1
            item_idx = level_counters[item_depth]

            pPr = p._p.get_or_add_pPr()
            if item_depth > 0:
                pPr.set("lvl", str(min(4, item_depth)))

            item_is_ordered = getattr(item, "is_ordered", list_block.is_ordered)

            # Clean existing bullet elements to prevent duplicate bullet/number glyphs
            for child_xml in list(pPr):
                if child_xml.tag.endswith(("buChar", "buAutoNum", "buNone", "buBlip")):
                    pPr.remove(child_xml)

            if item_is_ordered:
                try:
                    pPr.insert(0, parse_xml('<a:buNone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
                except Exception:
                    pass
            else:
                try:
                    pPr.insert(0, parse_xml('<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="•"/>'))
                except Exception:
                    pass

            p.space_after = sp_after
            p.space_before = sp_before
            p.line_spacing = typo.line_spacing

            base_pt = typo.list_size(item_depth)
            target_font_size = Pt(int(base_pt))

            item_text = item.text
            if item_is_ordered and not re.match(r"^\d+[\.\)]\s+", item_text):
                item_text = f"{item_idx}. {item_text}"

            if not item.runs:
                p.text = item_text
                for r in p.runs:
                    apply_font_family(r, r.text, is_code=False, default_size_pt=target_font_size.pt)
            else:
                p.text = ""
                # Add prefix for first run if ordered
                first_run = True
                for run in item.runs:
                    r = p.add_run()
                    r_text = run.text
                    if first_run and item_is_ordered and not re.match(r"^\d+[\.\)]\s+", item.text):
                        r_text = f"{item_idx}. {r_text}"
                        first_run = False
                    r.text = r_text
                    apply_font_family(r, r_text, is_code=run.is_code, default_size_pt=target_font_size.pt)
                    if run.is_bold:
                        r.font.bold = True
                    if run.is_italic:
                        r.font.italic = True
                    if getattr(run, "is_strikethrough", False):
                        try:
                            r.font.strikethrough = True
                        except Exception:
                            pass
                        rPr = r._r.get_or_add_rPr()
                        rPr.set("strike", "sngStrike")
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
    """Render a TableBlock as a corporate-styled table shape with canvas-proportional font scaling and padding.

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

    # Detect slide width scale relative to standard 13.33" widescreen
    slide_w_in = width_in + (left_in * 2.0)
    scale = max(1.0, min(2.5, slide_w_in / 13.33))

    header_font_size_pt = Pt(int(12.0 * scale)) if scale < 1.5 else Pt(22)
    body_font_size_pt = Pt(int(11.0 * scale)) if scale < 1.5 else Pt(20)

    calc_table_h = max(0.6 * scale, min(height_in, (num_rows * 0.45 * scale) + (0.15 * scale)))
    table_shape = slide.shapes.add_table(
        num_rows, num_cols, Inches(left_in), Inches(top_in), Inches(width_in), Inches(calc_table_h)
    )
    tbl = table_shape.table

    # Auto column width distribution based on content length and longest single word floor scaled to canvas
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
        max_word_len = max([len(w) for w in words] + [4])
        min_w = (max_word_len * 0.16 * scale) + (0.35 * scale)
        col_word_floors.append(min_w)

    calc_widths = []
    for c_idx in range(num_cols):
        fraction = col_max_lens[c_idx] / total_chars
        calc_w = width_in * fraction
        final_col_w = max(col_word_floors[c_idx], calc_w)
        calc_widths.append(final_col_w)

    usable_width = width_in
    sum_widths = sum(calc_widths)
    if sum_widths > usable_width and sum_widths > 0:
        extra_space = usable_width - sum(col_word_floors)
        if extra_space > 0:
            flex_widths = [max(0.0, w - col_word_floors[i]) for i, w in enumerate(calc_widths)]
            flex_sum = sum(flex_widths) or 1.0
            calc_widths = [col_word_floors[i] + (flex_widths[i] / flex_sum * extra_space) for i in range(num_cols)]
        else:
            scale_factor = usable_width / sum_widths
            calc_widths = [w * scale_factor for w in calc_widths]

    for c_idx in range(num_cols):
        tbl.columns[c_idx].width = Inches(calc_widths[c_idx])

    # MC Corporate Dark Blue (#004065) & Alternating Shading (#F2F4F7)
    header_bg_color = RGBColor(0, 64, 101)
    alt_row_color = RGBColor(242, 244, 247)

    for r_idx, row in enumerate(all_rows):
        is_header = (r_idx == 0 and bool(table_block.headers))
        for c_idx, val in enumerate(row):
            if c_idx < num_cols:
                cell = tbl.cell(r_idx, c_idx)
                cell.margin_left = Inches(0.12 * scale)
                cell.margin_right = Inches(0.12 * scale)
                cell.margin_top = Inches(0.08 * scale)
                cell.margin_bottom = Inches(0.08 * scale)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE

                tf = cell.text_frame
                tf.word_wrap = True

                cell_runs = []
                if table_block.raw_rows and r_idx < len(table_block.raw_rows):
                    raw_row = table_block.raw_rows[r_idx]
                    if c_idx < len(raw_row.cells):
                        cell_runs = getattr(raw_row.cells[c_idx], "runs", [])

                if cell_runs:
                    p0 = tf.paragraphs[0]
                    p0.text = ""
                    for run_obj in cell_runs:
                        r = p0.add_run()
                        r.text = run_obj.text
                        if run_obj.url:
                            r.hyperlink.address = run_obj.url
                else:
                    cell.text = str(val)

                if is_header:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = header_bg_color
                    for p in tf.paragraphs:
                        p.alignment = PP_ALIGN.LEFT
                        for run in p.runs:
                            run.font.name = "Arial"
                            run.font.size = header_font_size_pt
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    if r_idx % 2 == 1:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = alt_row_color
                    for p in tf.paragraphs:
                        p.alignment = PP_ALIGN.LEFT
                        for run in p.runs:
                            run.font.name = "Arial"
                            run.font.size = body_font_size_pt
                            run.font.color.rgb = RGBColor(40, 40, 40)


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
    img_path = Path(image_block.src) if image_block.src else None
    if img_path and img_path.is_file():
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
            return
        except Exception as err:
            logger.error("Failed to render image %s: %s", image_block.src, err, exc_info=True)

    # Fallback Image Card Container for Remote / Placeholder URLs or missing files
    # Proportional card sizing: max height 2.6 inches, max width 8.5 inches
    card_h = min(2.6, max(1.8, height_in)) if height_in > 0 else 2.2
    card_w = min(width_in if width_in > 0 else 8.0, 8.5)
    offset_x = left_in + (width_in - card_w) / 2.0 if width_in > card_w else left_in

    from pptx.enum.shapes import MSO_SHAPE
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(offset_x),
        Inches(top_in),
        Inches(card_w),
        Inches(card_h),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(240, 244, 248)
    card.line.color.rgb = RGBColor(11, 79, 108)
    card.line.width = Pt(1.5)

    tf = card.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = f"🖼️ Image: {image_block.alt or 'Sample Image'}"
    p1.font.name = "Arial"
    p1.font.bold = True
    p1.font.size = Pt(18)
    p1.font.color.rgb = RGBColor(11, 79, 108)

    p2 = tf.add_paragraph()
    p2.text = f"Source: {image_block.src}"
    p2.font.name = "Arial"
    p2.font.size = Pt(12)
    p2.font.italic = True
    p2.font.color.rgb = RGBColor(100, 100, 100)
    p2.space_before = Pt(8)


def render_code_block(
    slide_or_tf: Any,
    code_block: CodeBlock,
    left_in: float = 1.83,
    top_in: float = 2.0,
    width_in: float = 11.34,
    height_in: float = 4.5,
) -> None:
    """Render a CodeBlock into a dark slate monospaced card container with dynamic height and font scaling."""
    from pptx.enum.shapes import MSO_SHAPE

    lines = code_block.code.split("\n")
    line_count = len(lines) + (1 if code_block.language else 0)

    # Proportional monospaced font scaling for high presentation legibility
    if line_count <= 6:
        font_size_pt = 20.0
        line_step = 0.52
    elif line_count <= 12:
        font_size_pt = 17.0
        line_step = 0.44
    else:
        font_size_pt = 14.0
        line_step = 0.36

    # Dynamic card height fitting code content
    calc_height = max(1.8, min(4.8, line_count * line_step + 0.50))

    if hasattr(slide_or_tf, "shapes"):
        card = slide_or_tf.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left_in),
            Inches(top_in),
            Inches(width_in),
            Inches(calc_height),
        )
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(30, 30, 30)
        card.line.color.rgb = RGBColor(60, 60, 60)
        card.line.width = Pt(1.0)
        tf = card.text_frame
    elif hasattr(slide_or_tf, "text_frame"):
        tf = slide_or_tf.text_frame
    elif hasattr(slide_or_tf, "paragraphs"):
        tf = slide_or_tf
    else:
        return

    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.20)
    tf.margin_bottom = Inches(0.20)

    if code_block.language:
        p_hdr = tf.paragraphs[0] if (tf.paragraphs and not tf.paragraphs[0].text) else tf.add_paragraph()
        p_hdr.text = f"// Code Language: {code_block.language.upper()}"
        p_hdr.font.name = "Consolas"
        p_hdr.font.size = Pt(int(font_size_pt))
        p_hdr.font.bold = True
        p_hdr.font.color.rgb = RGBColor(86, 156, 214)
        p_hdr.space_after = Pt(6)

    for line in lines:
        p = tf.add_paragraph()
        p.text = line
        p.level = 0
        p.space_after = Pt(2)
        p.space_before = Pt(0)
        p.line_spacing = 1.15

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
            r.font.size = Pt(int(font_size_pt))
            r.font.color.rgb = RGBColor(220, 220, 220)
