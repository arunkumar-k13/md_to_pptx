"""PowerPoint Generator Module.

Consumes the planned and paginated Presentation IR model and populates PPTX template
placeholders using SlideLayoutResolver and shape_renderers. Preserves corporate layout formatting.
"""

from __future__ import annotations
import logging
import math
from pathlib import Path
from typing import Optional, Union

from pptx import Presentation as PPTXPresentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.packuri import PackURI
from pptx.util import Inches, Pt

from md_to_pptx.core.generator.shape_renderers import (
    render_bullet_list_block,
    render_code_block,
    render_image_block,
    render_paragraph_block,
    render_table_block,
    render_title_block,
)
from md_to_pptx.core.presentation_model import (
    BulletListBlock,
    CodeBlock,
    ImageBlock,
    ParagraphBlock,
    Presentation,
    QuoteBlock,
    Slide,
    SlideIntent,
    TableBlock,
    TitleBlock,
)
from md_to_pptx.core.template.layout_resolver import SlideLayoutResolver
from md_to_pptx.core.template.template_analyzer import TemplateMetadata

logger = logging.getLogger(__name__)


CORPORATE_FOOTER_TEXT = "Copyright © Molecular Connections Pvt. Ltd.  ISO/IEC 27001:2022 Certified"


def extract_title_metadata(blocks: list[ContentBlock], default_title: str = "", default_subtitle: str = "") -> tuple[str, str, list[ContentBlock]]:
    """Inspect content blocks for presentation metadata key-value patterns (Topic, Focus, Target Audience).

    Extracts clean Title and Subtitle text while filtering out raw metadata blocks.
    """
    clean_blocks = []
    extracted_topic = ""
    extracted_focus = ""
    extracted_audience = ""

    for b in blocks:
        if isinstance(b, BulletListBlock):
            retained_items = []
            for item in b.items:
                item_text = item.text.strip()
                item_lower = item_text.lower()
                if item_lower.startswith("topic:"):
                    extracted_topic = item_text[item_text.find(":") + 1:].strip()
                elif item_lower.startswith("focus:"):
                    extracted_focus = item_text[item_text.find(":") + 1:].strip()
                elif item_lower.startswith("target audience:") or item_lower.startswith("audience:"):
                    extracted_audience = item_text[item_text.find(":") + 1:].strip()
                else:
                    retained_items.append(item)
            if retained_items:
                clean_blocks.append(BulletListBlock(items=retained_items, is_ordered=b.is_ordered))
        elif isinstance(b, ParagraphBlock):
            p_text = b.text.strip()
            p_lower = p_text.lower()
            if p_lower.startswith("topic:"):
                extracted_topic = p_text[p_text.find(":") + 1:].strip()
            elif p_lower.startswith("focus:"):
                extracted_focus = p_text[p_text.find(":") + 1:].strip()
            elif p_lower.startswith("target audience:") or p_lower.startswith("audience:"):
                extracted_audience = p_text[p_text.find(":") + 1:].strip()
            else:
                clean_blocks.append(b)
        else:
            clean_blocks.append(b)

    final_title = extracted_topic or default_title
    subtitle_parts = []
    if extracted_focus:
        subtitle_parts.append(extracted_focus)
    if extracted_audience:
        aud_prefix = "" if extracted_audience.lower().startswith("for ") else "For "
        subtitle_parts.append(f"{aud_prefix}{extracted_audience}")

    final_subtitle = " | ".join(subtitle_parts) if subtitle_parts else default_subtitle
    return final_title, final_subtitle, clean_blocks


def get_all_slide_shapes(shapes: Any) -> list[Any]:
    """Recursively traverse slide shapes including sub-shapes inside GroupShapes."""
    result = []
    for shp in shapes:
        if getattr(shp, "shape_type", None) == 6:  # MSO_SHAPE_TYPE.GROUP
            result.extend(get_all_slide_shapes(shp.shapes))
        else:
            result.append(shp)
    return result


class PPTXGenerator:
    """Core rendering engine populating PowerPoint presentation templates."""

    def __init__(self, resolver: Optional[SlideLayoutResolver] = None) -> None:
        """Initialize PPTXGenerator with a SlideLayoutResolver instance.

        Args:
            resolver: Optional SlideLayoutResolver instance.
        """
        self.resolver = resolver or SlideLayoutResolver()
        self.slides_cloned_count = 0
        self.slides_rendered_count = 0
        self.slides_skipped_count = 0
        self.placeholders_populated_count = 0
        self.placeholders_skipped_count = 0
        self.protected_shapes_count = 0

    def generate(
        self,
        presentation: Presentation,
        template_path: Union[str, Path],
        template_meta: TemplateMetadata,
        output_path: Union[str, Path],
    ) -> Path:
        """Render a Presentation IR model into a PowerPoint file (.pptx).

        Args:
            presentation: Input paginated Presentation IR model.
            template_path: Path to target corporate PPTX template.
            template_meta: Introspected TemplateMetadata object.
            output_path: Target path for generated PPTX file.

        Returns:
            Resolved Path instance to generated PPTX file.
        """
        path = Path(template_path)
        out_path = Path(output_path)
        if not path.is_file():
            raise FileNotFoundError(f"Template PPTX not found: {template_path}")

        out_path.parent.mkdir(parents=True, exist_ok=True)
        self.tmpl_meta = template_meta

        prs = PPTXPresentation(str(path))

        # Sanitize master layouts and slide master to replace leftover "Standard Business Presentation"
        for layout in prs.slide_layouts:
            for shp in layout.shapes:
                if shp.has_text_frame and shp.text_frame:
                    if shp.text_frame.text and "Standard Business Presentation" in shp.text_frame.text:
                        shp.text_frame.text = CORPORATE_FOOTER_TEXT
                        for p in shp.text_frame.paragraphs:
                            p.font.name = "Arial"
                            p.font.size = Pt(9)
        for shp in prs.slide_master.shapes:
            if shp.has_text_frame and shp.text_frame:
                if shp.text_frame.text and "Standard Business Presentation" in shp.text_frame.text:
                    shp.text_frame.text = CORPORATE_FOOTER_TEXT
                    for p in shp.text_frame.paragraphs:
                        p.font.name = "Arial"
                        p.font.size = Pt(9)

        has_brand_slides = any(getattr(s, "brand_role", None) is not None for s in presentation.slides)

        if not has_brand_slides or len(prs.slides) == 0:
            # Clear starter slides if presentation is purely dynamic
            while len(prs.slides) > 0:
                rId = prs.slides._sldIdLst[0].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[0]

            for slide_ir in presentation.slides:
                matched_layout_meta = self.resolver.resolve_layout(slide_ir, template_meta)
                pptx_layout = prs.slide_layouts[matched_layout_meta.layout_index]
                slide_shape = prs.slides.add_slide(pptx_layout)
                self._render_dynamic_slide(slide_shape, slide_ir, presentation)
                self.slides_rendered_count += 1
        else:
            # Template Slide Retention & Cloning Engine
            kept_slide_map = {}
            kept_indices = set()

            for slide_ir in presentation.slides:
                if slide_ir.brand_role and slide_ir.brand_slide_index is not None:
                    if 0 <= slide_ir.brand_slide_index < len(prs.slides):
                        kept_indices.add(slide_ir.brand_slide_index)
                        kept_slide_map[id(slide_ir)] = prs.slides[slide_ir.brand_slide_index]

            # Drop non-brand pre-existing template slides cleanly
            remove_indices = [i for i in range(len(prs.slides)) if i not in kept_indices]
            self.slides_skipped_count = len(remove_indices)
            for idx in reversed(remove_indices):
                r_id = prs.slides._sldIdLst[idx].rId
                prs.part.drop_rel(r_id)
                del prs.slides._sldIdLst[idx]

            # Re-number all retained slide parts to prevent PackURI collisions
            for idx, slide_obj in enumerate(prs.slides):
                setattr(slide_obj.part, "_partname", PackURI(f"/ppt/slides/slide{idx + 1}.xml"))

            # Render Brand slides and add Dynamic slides
            target_slide_elements = []

            for slide_ir in presentation.slides:
                if slide_ir.brand_role:
                    slide_shape = kept_slide_map.get(id(slide_ir))
                    if not slide_shape:
                        matched_meta = self.resolver.resolve_layout(slide_ir, template_meta)
                        slide_shape = prs.slides.add_slide(prs.slide_layouts[matched_meta.layout_index])
                    self._render_brand_slide(slide_shape, slide_ir, presentation)
                    self.slides_cloned_count += 1

                    for elem in prs.slides._sldIdLst:
                        if prs.part.related_part(elem.rId) == slide_shape.part:
                            target_slide_elements.append(elem)
                            break
                else:
                    matched_layout_meta = self.resolver.resolve_layout(slide_ir, template_meta)
                    pptx_layout = prs.slide_layouts[matched_layout_meta.layout_index]
                    dyn_shape = prs.slides.add_slide(pptx_layout)
                    self._render_dynamic_slide(dyn_shape, slide_ir, presentation)
                    self.slides_rendered_count += 1

                    for elem in prs.slides._sldIdLst:
                        if prs.part.related_part(elem.rId) == dyn_shape.part:
                            target_slide_elements.append(elem)
                            break

            # Re-order slides in presentation XML to match exact planned sequence
            if len(target_slide_elements) == len(prs.slides):
                prs.slides._sldIdLst.clear()
                prs.slides._sldIdLst.extend(target_slide_elements)

        # Enforce Corporate Footer & Universal Font Sanitization (Arial) across all slides
        for s in prs.slides:
            all_shps = get_all_slide_shapes(s.shapes)
            for shp in all_shps:
                if shp.has_text_frame and shp.text_frame:
                    if shp.text_frame.text and "Standard Business Presentation" in shp.text_frame.text:
                        shp.text_frame.text = CORPORATE_FOOTER_TEXT
                    for p in shp.text_frame.paragraphs:
                        p.font.name = "Arial"
                        for r in p.runs:
                            r.font.name = "Arial"

        target_file_path = out_path.resolve()
        try:
            prs.save(str(target_file_path))
        except PermissionError:
            import time
            target_file_path = out_path.parent / f"{out_path.stem}_{int(time.time())}{out_path.suffix}"
            logger.warning("Original output file locked. Saving presentation to timestamped path: %s", target_file_path.name)
            prs.save(str(target_file_path))

        logger.info("Successfully generated PowerPoint presentation at %s", target_file_path)
        return target_file_path

    def _render_brand_slide(self, slide_shape: Any, slide_ir: Slide, presentation: Presentation) -> None:
        """Render a brand slide by populating only designated text placeholders without altering graphics."""
        role = slide_ir.brand_role

        if role == "BRAND_COVER":
            t_def = slide_ir.title.text if slide_ir.title else (presentation.title or "Presentation Title")
            s_def = slide_ir.subtitle or presentation.subtitle
            t_text, s_text, clean_blocks = extract_title_metadata(slide_ir.blocks, default_title=t_def, default_subtitle=s_def)

            title_updated = False
            title_ph = self._find_placeholder_by_type(slide_shape, ["TITLE", "CENTER_TITLE"])
            if title_ph:
                render_title_block(title_ph, TitleBlock(text=t_text, level=1))
                title_updated = True
                self.placeholders_populated_count += 1

            if not title_updated:
                for shp in slide_shape.shapes:
                    if shp.has_text_frame and (shp.name.lower().startswith("title") or "title" in shp.text_frame.text.lower()):
                        shp.text_frame.text = t_text
                        title_updated = True
                        self.placeholders_populated_count += 1
                        break

            if s_text:
                sub_ph = self._find_placeholder_by_type(slide_shape, ["SUBTITLE"])
                if sub_ph:
                    render_title_block(sub_ph, TitleBlock(text=s_text, level=2))
                    self.placeholders_populated_count += 1
                else:
                    for shp in slide_shape.shapes:
                        if shp.has_text_frame and shp.name.lower().startswith("sub"):
                            shp.text_frame.text = s_text
                            self.placeholders_populated_count += 1
                            break

        elif role == "BRAND_CASE_STUDY":
            if slide_ir.title:
                title_ph = self._find_placeholder_by_type(slide_shape, ["TITLE", "CENTER_TITLE"])
                if title_ph:
                    render_title_block(title_ph, slide_ir.title)
                    self.placeholders_populated_count += 1
                else:
                    for shp in slide_shape.shapes:
                        if shp.has_text_frame and shp.name.lower().startswith("title"):
                            shp.text_frame.text = slide_ir.title.text
                            self.placeholders_populated_count += 1
                            break

            # Smart semantic placeholder population for Case Study text boxes
            all_shps = get_all_slide_shapes(slide_shape.shapes)
            body_lorem_shapes = [
                shp for shp in all_shps
                if shp.has_text_frame and "lorem ipsum" in shp.text_frame.text.lower()
            ]

            if slide_ir.blocks:
                for i, block in enumerate(slide_ir.blocks):
                    if i < len(body_lorem_shapes):
                        tf = body_lorem_shapes[i].text_frame
                        if isinstance(block, ParagraphBlock):
                            render_paragraph_block(tf, block)
                            self.placeholders_populated_count += 1
                        elif isinstance(block, BulletListBlock):
                            render_bullet_list_block(tf, block)
                            self.placeholders_populated_count += 1
            else:
                # If no markdown content provided, preserve template or clear placeholder text
                pass

            # Dynamically adjust label textbox height and body textbox position to prevent text overlap
            self._adjust_case_study_label_and_body_overlap(slide_shape)

        elif role == "BRAND_THANK_YOU":
            if slide_ir.title:
                title_ph = self._find_placeholder_by_type(slide_shape, ["TITLE", "CENTER_TITLE"])
                if title_ph:
                    render_title_block(title_ph, slide_ir.title)
                    self.placeholders_populated_count += 1

        self.protected_shapes_count += max(0, len(get_all_slide_shapes(slide_shape.shapes)) - 2)

    def _render_dynamic_slide(self, slide_shape: Any, slide_ir: Slide, presentation: Presentation) -> None:
        """Render a dynamic slide populating content blocks and clearing unpopulated body placeholders."""
        # Handle Dynamic Title Slide
        if slide_ir.intent == SlideIntent.TITLE_SLIDE:
            t_def = slide_ir.title.text if slide_ir.title else (presentation.title or "Presentation Title")
            s_def = slide_ir.subtitle or presentation.subtitle
            t_text, s_text, clean_blocks = extract_title_metadata(slide_ir.blocks, default_title=t_def, default_subtitle=s_def)

            title_ph = self._find_placeholder_by_type(slide_shape, ["TITLE", "CENTER_TITLE"])
            if title_ph:
                render_title_block(title_ph, TitleBlock(text=t_text, level=1))
                self.placeholders_populated_count += 1

            if s_text:
                subtitle_ph = self._find_placeholder_by_type(slide_shape, ["SUBTITLE"])
                if subtitle_ph:
                    render_title_block(subtitle_ph, TitleBlock(text=s_text, level=2))
                    self.placeholders_populated_count += 1
            slide_blocks = clean_blocks
        else:
            if slide_ir.title:
                title_ph = self._find_placeholder_by_type(slide_shape, ["TITLE", "CENTER_TITLE"])
                if title_ph:
                    render_title_block(title_ph, slide_ir.title)
                    self.placeholders_populated_count += 1

            sub_text = slide_ir.subtitle
            if sub_text:
                subtitle_ph = self._find_placeholder_by_type(slide_shape, ["SUBTITLE"])
                if subtitle_ph:
                    render_title_block(subtitle_ph, TitleBlock(text=sub_text, level=2))
                    self.placeholders_populated_count += 1
            text_blocks = [b for b in slide_ir.blocks if isinstance(b, (ParagraphBlock, BulletListBlock, QuoteBlock))]
            visual_blocks = [b for b in slide_ir.blocks if not isinstance(b, (ParagraphBlock, BulletListBlock, QuoteBlock))]
            slide_blocks = text_blocks + visual_blocks

        body_phs = self._find_all_placeholders_by_type(slide_shape, ["BODY", "OBJECT", "CONTENT", "TEXT"])
        pic_phs = self._find_all_placeholders_by_type(slide_shape, ["PICTURE"])
        tbl_phs = self._find_all_placeholders_by_type(slide_shape, ["TABLE"])

        pic_idx = 0
        tbl_idx = 0

        used_ph_ids = set()
        for i, block in enumerate(slide_blocks):
            target_ph = None

            if isinstance(block, ImageBlock) and pic_phs:
                target_ph = pic_phs[pic_idx % len(pic_phs)]
                pic_idx += 1
            elif isinstance(block, TableBlock) and tbl_phs:
                target_ph = tbl_phs[tbl_idx % len(tbl_phs)]
                tbl_idx += 1

            if not target_ph and body_phs:
                if slide_ir.intent in (SlideIntent.TWO_COLUMN, SlideIntent.COMPARISON) and len(slide_blocks) == 2 and len(body_phs) >= 2:
                    target_ph = body_phs[i % len(body_phs)]
                else:
                    target_ph = body_phs[0]

            if target_ph:
                used_ph_ids.add(getattr(target_ph, "shape_id", id(target_ph)))

            tf = target_ph.text_frame if target_ph and hasattr(target_ph, "text_frame") else None

            slide_h = self.tmpl_meta.slide_height_inches if getattr(self, "tmpl_meta", None) and hasattr(self.tmpl_meta, "slide_height_inches") else 7.5
            # Track running Y-offset across ALL preceding blocks on this slide to prevent shape overlap
            base_top = target_ph.top.inches if target_ph else 3.99
            calc_top = base_top
            if i > 0 and target_ph:
                prev_h_total = 0.0
                for prev_idx in range(i):
                    prev_b = slide_blocks[prev_idx]
                    ph_w = target_ph.width.inches if hasattr(target_ph, "width") else 11.34
                    prev_h_total += prev_b.estimate_height(ph_w) + 0.35
                if prev_h_total > 0:
                    calc_top = min(10.2, base_top + prev_h_total)

            if isinstance(block, ParagraphBlock) and tf:
                render_paragraph_block(tf, block, slide_height_inches=slide_h)
                self.placeholders_populated_count += 1
            elif isinstance(block, BulletListBlock) and tf:
                render_bullet_list_block(tf, block, slide_height_inches=slide_h)
                self.placeholders_populated_count += 1
            elif isinstance(block, TableBlock):
                left = target_ph.left.inches if target_ph else 1.83
                width = target_ph.width.inches if target_ph else 11.34
                height = target_ph.height.inches if target_ph else 4.0
                render_table_block(slide_shape, block, left, calc_top, width, height)
                self.placeholders_populated_count += 1
            elif isinstance(block, ImageBlock):
                left = target_ph.left.inches if target_ph else 1.83
                width = target_ph.width.inches if target_ph else 11.34
                height = target_ph.height.inches if target_ph else 4.5
                render_image_block(slide_shape, block, left, calc_top, width, height)
                self.placeholders_populated_count += 1
            elif isinstance(block, CodeBlock):
                left = target_ph.left.inches if target_ph else 1.83
                width = target_ph.width.inches if target_ph else 11.34
                height = target_ph.height.inches if target_ph else 4.5
                render_code_block(slide_shape, block, left, calc_top, width, height)
                self.placeholders_populated_count += 1
            elif isinstance(block, QuoteBlock) and tf:
                qp = tf.add_paragraph() if (tf.paragraphs and not tf.paragraphs[0].text) else tf.add_paragraph()
                qp.level = 0
                qp.space_before = Pt(6)
                qp.space_after = Pt(6)
                try:
                    from pptx.oxml import parse_xml
                    q_pPr = qp._p.get_or_add_pPr()
                    for child in list(q_pPr):
                        if child.tag.endswith(("buChar", "buAutoNum", "buNone", "buBlip")):
                            q_pPr.remove(child)
                    q_pPr.insert(0, parse_xml('<a:buNone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
                except Exception:
                    pass

                runs_to_render = block.runs if block.runs else [InlineRun(text=block.text, is_italic=True)]
                qp.text = ""
                r_start = qp.add_run()
                r_start.text = '"'
                r_start.font.name = "Arial"
                r_start.font.italic = True
                r_start.font.size = Pt(18)
                r_start.font.color.rgb = RGBColor(11, 79, 108)

                for run in runs_to_render:
                    r_txt = run.text.strip('"')
                    if not r_txt:
                        continue
                    r = qp.add_run()
                    r.text = r_txt
                    r.font.name = "Consolas" if run.is_code else "Arial"
                    r.font.size = Pt(18)
                    r.font.bold = run.is_bold
                    r.font.italic = run.is_italic or (not run.is_bold and not run.is_code)
                    r.font.color.rgb = RGBColor(86, 156, 214) if run.is_code else RGBColor(11, 79, 108)

                r_end = qp.add_run()
                r_end.text = '"'
                r_end.font.name = "Arial"
                r_end.font.italic = True
                r_end.font.size = Pt(18)
                r_end.font.color.rgb = RGBColor(11, 79, 108)
                self.placeholders_populated_count += 1

        # Remove unpopulated placeholder shapes from slide DOM to prevent ghost edit boxes and stray prompt characters
        for ph in body_phs + pic_phs + tbl_phs:
            if hasattr(ph, "text_frame") and ph.text_frame:
                if not ph.text_frame.text.strip():
                    try:
                        ph._element.getparent().remove(ph._element)
                        self.placeholders_skipped_count += 1
                    except Exception:
                        ph.text_frame.text = ""

        # Apply corporate visual branding (vertical blue accent bar, footer, page badge)
        self._apply_corporate_branding_to_slide(slide_shape, slide_ir)

    def _apply_corporate_branding_to_slide(self, slide_shape: Any, slide_ir: Slide) -> None:
        """Apply corporate visual branding (vertical blue accent bar and corporate footer) to dynamic slides."""
        all_shps = get_all_slide_shapes(slide_shape.shapes)

        # 1. Vertical Blue Accent Bar (Rectangle 12)
        has_accent_bar = False
        for shp in all_shps:
            if getattr(shp, "left", None) and getattr(shp, "top", None) and getattr(shp, "width", None):
                if shp.left.inches < 0.4 and shp.top.inches < 0.5 and 0.15 <= shp.width.inches <= 0.45:
                    has_accent_bar = True
                    break

        if not has_accent_bar and hasattr(slide_shape, "slide_layout") and slide_shape.slide_layout:
            for l_shp in slide_shape.slide_layout.shapes:
                if getattr(l_shp, "left", None) and getattr(l_shp, "top", None) and getattr(l_shp, "width", None):
                    if l_shp.left.inches < 0.4 and l_shp.top.inches < 0.5 and 0.15 <= l_shp.width.inches <= 0.45:
                        has_accent_bar = True
                        break

        if not has_accent_bar and slide_ir.intent != SlideIntent.TITLE_SLIDE:
            accent_bar = slide_shape.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left=Inches(0.0),
                top=Inches(0.0),
                width=Inches(0.25),
                height=Inches(3.70),
            )
            accent_bar.fill.solid()
            accent_bar.fill.fore_color.rgb = RGBColor(11, 79, 108)  # #0B4F6C Corporate Blue
            accent_bar.line.fill.background()

        # 2. Corporate Footer & Leftover Text Replacement
        for shp in all_shps:
            if shp.has_text_frame and shp.text_frame:
                txt = shp.text_frame.text.strip()
                if "Standard Business Presentation" in txt:
                    shp.text_frame.text = CORPORATE_FOOTER_TEXT
                    for p in shp.text_frame.paragraphs:
                        p.font.name = "Arial"
                        p.font.size = Pt(9)
                        p.font.color.rgb = RGBColor(166, 166, 166)

    def _find_placeholder_by_type(self, slide_shape: Any, types: list[str]) -> Optional[Any]:
        """Find first placeholder shape matching type string names."""
        for shape in slide_shape.placeholders:
            ph_type = str(shape.placeholder_format.type).split(".")[-1].split(" ")[0].upper()
            if ph_type in types:
                return shape
        return None

    def _find_all_placeholders_by_type(self, slide_shape: Any, types: list[str]) -> list[Any]:
        """Find all placeholder shapes matching type string names."""
        matched = []
        for shape in slide_shape.placeholders:
            ph_type = str(shape.placeholder_format.type).split(".")[-1].split(" ")[0].upper()
            if ph_type in types:
                matched.append(shape)
        return matched

    def _adjust_case_study_label_and_body_overlap(self, slide_shape: Any) -> None:
        """Measure label text wrapping and dynamically adjust label height, width, and body top coordinates.

        Ensures label textboxes and body textboxes never vertically overlap regardless of label length.
        """
        all_shps = get_all_slide_shapes(slide_shape.shapes)
        body_lorem_shapes = [
            shp for shp in all_shps
            if shp.has_text_frame and "lorem ipsum" in shp.text_frame.text.lower()
        ]
        body_lorem_shapes.sort(key=lambda s: s.top)

        other_text_shapes = [
            shp for shp in all_shps
            if shp.has_text_frame
            and getattr(shp, "left", None)
            and shp.left.inches > 2.0
            and "lorem ipsum" not in shp.text_frame.text.lower()
            and not shp.name.lower().startswith("title")
        ]

        paired_groups = []
        for b_shp in body_lorem_shapes:
            candidate_labels = [l for l in other_text_shapes if l.top < b_shp.top]
            if candidate_labels:
                best_label = min(
                    candidate_labels,
                    key=lambda l: (b_shp.top.inches - l.top.inches, abs(b_shp.left.inches - l.left.inches))
                )
                paired_groups.append((best_label, b_shp))

        for idx, (label_shp, body_shp) in enumerate(paired_groups):
            label_text = label_shp.text_frame.text.strip() if label_shp.has_text_frame else ""
            if not label_text:
                continue

            font_size_pt = 28.0
            if label_shp.has_text_frame:
                for p in label_shp.text_frame.paragraphs:
                    if p.font and p.font.size and p.font.size.pt:
                        font_size_pt = float(p.font.size.pt)
                        break
                    for r in p.runs:
                        if r.font and r.font.size and r.font.size.pt:
                            font_size_pt = float(r.font.size.pt)
                            break
                    if font_size_pt != 28.0:
                        break

            tf = label_shp.text_frame
            margin_l = tf.margin_left.inches if tf and hasattr(tf, "margin_left") and tf.margin_left else 0.05
            margin_r = tf.margin_right.inches if tf and hasattr(tf, "margin_right") and tf.margin_right else 0.05
            margin_t = tf.margin_top.inches if tf and hasattr(tf, "margin_top") and tf.margin_top else 0.05
            margin_b = tf.margin_bottom.inches if tf and hasattr(tf, "margin_bottom") and tf.margin_bottom else 0.05

            char_width_in = (font_size_pt * 0.55) / 72.0
            single_line_w = (len(label_text) * char_width_in) + margin_l + margin_r

            target_w = max(label_shp.width.inches, min(6.0, single_line_w + 0.2))
            label_shp.width = Inches(target_w)

            if single_line_w <= 6.0:
                tf.word_wrap = False
                total_lines = 1
                actual_label_height = max(label_shp.height.inches, 0.28)
            else:
                tf.word_wrap = True
                usable_w = max(0.2, target_w - (margin_l + margin_r))
                chars_per_line = max(1, int(usable_w / char_width_in))
                total_lines = math.ceil(len(label_text) / chars_per_line)
                line_height_in = (font_size_pt * 1.0) / 72.0
                calc_height_in = (total_lines * line_height_in) + margin_t + margin_b
                actual_label_height = max(label_shp.height.inches, calc_height_in)

            label_shp.height = Inches(actual_label_height)

            # Enforce a mandatory vertical gap of at least 0.04 inches between label bottom and body top
            min_body_top = label_shp.top.inches + actual_label_height + 0.04
            if body_shp.top.inches < min_body_top:
                body_shp.top = Inches(min_body_top)

            # Re-check vertical spacing against subsequent label shape to prevent group-to-group collisions
            if idx + 1 < len(paired_groups):
                next_label, next_body = paired_groups[idx + 1]
                body_bottom = body_shp.top.inches + body_shp.height.inches
                min_next_label_top = body_bottom + 0.06
                if next_label.top.inches < min_next_label_top:
                    shift = min_next_label_top - next_label.top.inches
                    next_label.top = Inches(min_next_label_top)
                    next_body.top = Inches(next_body.top.inches + shift)
