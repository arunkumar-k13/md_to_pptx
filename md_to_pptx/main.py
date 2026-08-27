"""Main CLI Entrypoint and Presentation Pipeline Orchestrator.

Orchestrates the entire Markdown -> Presentation -> PDF pipeline:
MarkdownParser -> ContentAnalyzer -> SlidePlanner -> TemplateCache -> TemplateValidator -> Paginator -> PPTXGenerator -> ExporterFactory -> GenerationReport.
"""

from __future__ import annotations
import argparse
import logging
import sys
from pathlib import Path
from typing import Optional, Union

from md_to_pptx.config.settings_loader import Settings, load_settings
from md_to_pptx.core.content_analyzer import ContentAnalyzer
from md_to_pptx.core.design_engine import PresentationDesignEngine
from md_to_pptx.core.exporters.environment import PdfEnvironmentDetector
from md_to_pptx.core.exporters.exporter_factory import ExporterFactory
from md_to_pptx.core.generator.pptx_generator import PPTXGenerator
from md_to_pptx.core.markdown_parser import MarkdownParser
from md_to_pptx.core.overflow.paginator import Paginator
from md_to_pptx.core.presentation_model import BulletListBlock, CodeBlock, ImageBlock, SlideIntent, TableBlock
from md_to_pptx.core.slide_planner import SlidePlanner
from md_to_pptx.core.template.layout_resolver import SlideLayoutResolver
from md_to_pptx.core.template.template_cache import TemplateCache
from md_to_pptx.core.template.template_validator import TemplateValidator
from md_to_pptx.reporting.diagnostic_context import DiagnosticContext
from md_to_pptx.reporting.generation_report import GenerationReport
from md_to_pptx.utils.logger import setup_logger

from md_to_pptx.core.brand_slide_manager import BrandSlideManager
from md_to_pptx.core.qa.package_validator import PackageValidator
from md_to_pptx.core.qa.visual_qa import VisualQAEngine
from md_to_pptx.utils.path_utils import get_unique_filepath

logger = setup_logger(__name__)


def generate_presentation(
    markdown_path: Union[str, Path],
    template_path: Union[str, Path],
    output_dir: Union[str, Path] = "output",
    export_pdf: bool = False,
    preferred_exporter: str = "libreoffice",
    settings_path: Optional[Union[str, Path]] = None,
) -> GenerationReport:
    """Orchestrate end-to-end presentation generation pipeline.

    Args:
        markdown_path: Path to input Markdown (.md) file.
        template_path: Path to corporate PPTX template (.pptx) file.
        output_dir: Target output directory path.
        export_pdf: True to generate PDF output in addition to PPTX.
        preferred_exporter: Target PDF exporter engine ('libreoffice').
        settings_path: Optional path to settings.yaml configuration file.

    Returns:
        GenerationReport instance detailing execution diagnostics.
    """
    ctx = DiagnosticContext()
    ctx.start_timer()

    md_p = Path(markdown_path).resolve()
    tmpl_p = Path(template_path).resolve()
    out_dir = Path(output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not md_p.is_file():
        raise FileNotFoundError(f"Markdown file not found: {markdown_path}")
    if not tmpl_p.is_file():
        raise FileNotFoundError(f"Template file not found: {template_path}")

    # 1. Load Configuration
    cfg = load_settings(settings_path)

    # 2. Parse Markdown to Document AST
    logger.info("Parsing Markdown document: %s", md_p.name)
    parser = MarkdownParser()
    doc_ast = parser.parse_file(md_p)

    # 3. Analyze Content Structure
    logger.info("Analyzing content structure...")
    analyzer = ContentAnalyzer()
    analyzed_doc = analyzer.analyze(doc_ast)

    # 4. Infer Presentation Design Plan via PresentationDesignEngine
    logger.info("Inferring Presentation Design Plan...")
    design_engine = PresentationDesignEngine()
    design_plan = design_engine.create_design_plan(analyzed_doc)

    # 5. Plan Initial Presentation Model
    logger.info("Planning presentation slides...")
    planner = SlidePlanner()
    unpaginated_pres = planner.plan(analyzed_doc, settings=cfg, design_plan=design_plan)

    # 6. Introspect & Validate Template
    logger.info("Introspecting template metadata...")
    cache = TemplateCache()
    tmpl_meta = cache.get_or_analyze(tmpl_p)
    ctx.template_name = tmpl_p.name

    validator = TemplateValidator()
    val_report = validator.validate(tmpl_meta)
    ctx.validation_passed = val_report.is_valid

    for issue in val_report.issues:
        if issue.severity == "ERROR":
            ctx.record_warning(f"Template Error: {issue.message}")
        else:
            ctx.record_warning(f"Template Warning: {issue.message}")

    if not val_report.is_valid:
        logger.error("Template validation failed. Aborting generation.")
        ctx.stop_timer()
        return ctx.build_report()

    # 6.5 Apply Brand Slide Manager (Recognize and Preserve Corporate Brand Slides)
    logger.info("Applying Brand Slide Manager...")
    brand_manager = BrandSlideManager()
    unpaginated_pres = brand_manager.process_presentation(unpaginated_pres, tmpl_meta, settings=cfg)

    # Record Brand Slide Verification status
    brand_map = tmpl_meta.brand_slides
    used_roles = {s.brand_role for s in unpaginated_pres.slides if s.brand_role}

    brand_verification = {"total_template_slides": tmpl_meta.total_template_slides}
    for role_name in ["Cover", "Disclaimer", "About Company", "Case Study", "Thank You"]:
        b_key = f"BRAND_{role_name.upper().replace(' ', '_')}"
        detected_info = brand_map.get(b_key)
        is_detected = detected_info is not None
        status_str = "Inserted" if b_key in used_roles else "Skipped"
        layout_str = detected_info.get("layout_name", "Brand Layout") if detected_info else "N/A"
        s_idx = detected_info.get("slide_index", 0) if detected_info else None
        
        brand_verification[role_name] = {
            "detected": is_detected,
            "layout_name": layout_str,
            "slide_index": s_idx,
            "status": status_str,
            "shapes_count": detected_info.get("shapes_count", 0) if detected_info else 0,
            "images_count": detected_info.get("images_count", 0) if detected_info else 0,
            "smartart_detected": detected_info.get("smartart_detected", False) if detected_info else False,
            "relationships_preserved": detected_info.get("relationships_preserved", True) if detected_info else True,
            "theme_preserved": detected_info.get("theme_preserved", True) if detected_info else True,
            "editable_placeholders": detected_info.get("editable_placeholders", []) if detected_info else [],
        }
    ctx.brand_slides_verification = brand_verification

    # 7. Apply Spatial Overflow & Pagination Engine
    logger.info("Applying overflow handling and pagination...")
    resolver = SlideLayoutResolver()
    paginator = Paginator(resolver=resolver)
    paginated_pres = paginator.paginate(unpaginated_pres, settings=cfg, context=ctx, template_meta=tmpl_meta)

    ctx.total_slides = len(paginated_pres.slides)
    ctx.continuation_slides = sum(1 for s in paginated_pres.slides if s.is_continuation)

    # Record element counts and layout mapping decisions
    for s_idx, s in enumerate(paginated_pres.slides, start=1):
        layout_meta, confidence_score, reason = resolver.resolve_layout_with_score(s, tmpl_meta)
        ctx.record_layout(layout_meta.name)
        
        s_title_text = s.title.text if s.title else "Untitled Slide"
        explanation = f"Intent: {s.intent.value.upper()} | Confidence: {confidence_score:.2f} ({reason})"

        ctx.record_layout_mapping(
            slide_number=s_idx,
            markdown_summary=s_title_text,
            selected_layout=layout_meta.name,
            reason=explanation,
        )

        dp_match = design_plan.slides[s_idx - 1] if (design_plan and s_idx <= len(design_plan.slides)) else None
        if dp_match:
            ctx.record_design_report(
                slide_number=s_idx,
                content_summary=s_title_text,
                presentation_intent=dp_match.intent.name,
                why_chosen=dp_match.why_chosen,
                visual_priority=dp_match.visual_priority,
                density=dp_match.density,
                candidate_layouts=dp_match.suggested_layout_traits,
                chosen_layout=layout_meta.name,
                confidence=confidence_score,
            )

        for b in s.blocks:
            if isinstance(b, TableBlock):
                ctx.tables_rendered += 1
            elif isinstance(b, ImageBlock):
                ctx.images_rendered += 1
            elif isinstance(b, CodeBlock):
                ctx.code_blocks_rendered += 1

    # 7. Render PowerPoint File (.pptx)
    raw_pptx_path = out_dir / f"{md_p.stem}_presentation.pptx"
    out_pptx_path = get_unique_filepath(raw_pptx_path)
    logger.info("Rendering PowerPoint presentation to %s", out_pptx_path.name)
    generator = PPTXGenerator(resolver=resolver)
    final_pptx_path = generator.generate(paginated_pres, tmpl_p, tmpl_meta, out_pptx_path)
    ctx.pptx_path = str(final_pptx_path)
    ctx.slides_cloned = generator.slides_cloned_count
    ctx.slides_rendered = generator.slides_rendered_count
    ctx.slides_skipped = generator.slides_skipped_count
    ctx.placeholders_populated = generator.placeholders_populated_count
    ctx.placeholders_skipped = generator.placeholders_skipped_count
    ctx.protected_shapes_count = generator.protected_shapes_count

    # 7.5 Run Visual QA Inspection Engine
    qa_engine = VisualQAEngine()
    ctx.qa_report = qa_engine.inspect(final_pptx_path, tmpl_meta)

    # 7.6 Run Open XML Package Validator Engine
    logger.info("Validating Open XML package structure...")
    pkg_validator = PackageValidator()
    pkg_val_report = pkg_validator.validate(final_pptx_path, test_com=False)
    ctx.package_validation_report = pkg_val_report

    if not pkg_val_report.is_valid:
        logger.error("PPTX Package Validation failed: %s", ", ".join(pkg_val_report.issues))
        ctx.record_warning(f"PPTX Package Validation failed: {', '.join(pkg_val_report.issues)}")
        if export_pdf:
            logger.warning("Aborting PDF export due to failed PPTX package validation.")
            ctx.export_status = "ABORTED (Failed Package Validation)"
            ctx.export_engine_used = "N/A"
            export_pdf = False

    # 8. Pre-flight PDF Environment Detection & Export (if requested)
    env_report = PdfEnvironmentDetector.detect(preferred_exporter=preferred_exporter, settings=cfg)
    ctx.export_environment = env_report.to_dict()

    if export_pdf:
        logger.info("Checking PDF export environment...")
        print("[INFO] Checking PDF export environment...")
        lo_str = "Yes" if env_report.libreoffice_available else "No"
        print(f"Detected: LibreOffice: {lo_str}")

        if not env_report.pdf_available:
            print("Skipping PDF export. Presentation generation will continue.")
            logger.info("PDF export skipped: No supported PDF engine detected.")
            ctx.export_status = "UNAVAILABLE"
            ctx.export_engine_used = "N/A"
            ctx.record_warning("PDF export unavailable: No supported PDF engine detected on host machine.")
        else:
            exporter = ExporterFactory().get_exporter(preferred_exporter=preferred_exporter, settings=cfg)
            if exporter:
                raw_pdf_path = out_dir / f"{final_pptx_path.stem}.pdf"
                out_pdf_path = get_unique_filepath(raw_pdf_path)
                logger.info("Exporting presentation to PDF via %s: %s", exporter.__class__.__name__, out_pdf_path.name)
                try:
                    engine_name = exporter.__class__.__name__
                    final_pdf_path = exporter.export_pdf(final_pptx_path, out_pdf_path)
                    ctx.pdf_path = str(final_pdf_path)
                    ctx.export_status = "SUCCESS"
                    ctx.export_engine_used = engine_name
                except Exception as err:
                    logger.error("PDF export failed unexpectedly: %s", err, exc_info=True)
                    ctx.export_status = f"FAILED ({err})"
                    ctx.record_warning(f"PDF export failed unexpectedly: {err}")

    ctx.stop_timer()
    report = ctx.build_report()
    logger.info("Presentation generation completed in %.3fs.", report.generation_time_seconds)
    return report


def main() -> None:
    """CLI entrypoint function for batch operations."""
    parser = argparse.ArgumentParser(description="Markdown to PowerPoint Generator CLI")
    parser.add_argument("-m", "--markdown", required=True, help="Path to input Markdown (.md) file")
    parser.add_argument("-t", "--template", required=True, help="Path to PowerPoint template (.pptx) file")
    parser.add_argument("-o", "--output-dir", default="output", help="Target output directory")
    parser.add_argument("-p", "--pdf", action="store_true", help="Also export presentation to PDF")
    parser.add_argument("-c", "--config", help="Optional path to settings.yaml configuration file")

    args = parser.parse_args()

    try:
        report = generate_presentation(
            markdown_path=args.markdown,
            template_path=args.template,
            output_dir=args.output_dir,
            export_pdf=args.pdf,
            settings_path=args.config,
        )
        print("\n" + report.to_markdown())
    except Exception as err:
        logger.error("Execution failed: %s", err, exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
