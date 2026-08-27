"""Generation Report Model.

Provides structured data containers and markdown formatters for presentation generation diagnostics.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List


@dataclass(slots=True)
class OverflowEvent:
    """Telemetry record for an individual spatial or item overflow event.

    Attributes:
        slide_index: Index of slide experiencing overflow.
        block_type: Type of block triggering overflow.
        action_taken: Resolution strategy applied (e.g. 'split_to_continuation_slide').
        items_moved: Number of items/blocks moved to continuation slide.
        reason: Explanation for overflow trigger.
    """

    slide_index: int
    block_type: str
    action_taken: str
    items_moved: int = 0
    reason: str = ""


@dataclass(slots=True)
class DesignReportRecord:
    """Telemetry record for an individual slide's presentation design intent decision.

    Attributes:
        slide_number: 1-indexed slide number.
        content_summary: Summary of slide topic/content.
        presentation_intent: Abstract PresentationIntent name.
        why_chosen: Executive design rationale for intent selection.
        visual_priority: High/Medium/Low visual priority tag.
        density: High/Medium/Low visual density tag.
        candidate_layouts: Recommended candidate layout types.
        chosen_layout: Final resolved PowerPoint layout.
        confidence: Normalized scoring confidence [0.0, 1.0].
    """

    slide_number: int
    content_summary: str
    presentation_intent: str
    why_chosen: str
    visual_priority: str
    density: str
    candidate_layouts: List[str]
    chosen_layout: str
    confidence: float


@dataclass(slots=True)
class LayoutMappingRecord:
    """Telemetry record for an individual slide layout mapping decision.

    Attributes:
        slide_number: 1-indexed slide number.
        markdown_summary: Brief summary of input markdown content.
        selected_layout: Resolved template layout name.
        reason: Architectural rationale for layout selection.
    """

    slide_number: int
    markdown_summary: str
    selected_layout: str
    reason: str


@dataclass(slots=True)
class GenerationReport:
    """Comprehensive diagnostic report emitted after presentation rendering."""

    template_name: str = ""
    validation_passed: bool = True
    cache_used: bool = False
    total_slides: int = 0
    continuation_slides_created: int = 0
    layouts_used: Dict[str, int] = field(default_factory=dict)
    tables_rendered: int = 0
    images_rendered: int = 0
    code_blocks_rendered: int = 0
    slides_cloned: int = 0
    slides_rendered: int = 0
    slides_skipped: int = 0
    placeholders_populated: int = 0
    placeholders_skipped: int = 0
    protected_shapes_count: int = 0
    export_environment: Dict[str, Any] = field(default_factory=dict)
    brand_slides_verification: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    design_reports: List[DesignReportRecord] = field(default_factory=list)
    layout_mappings: List[LayoutMappingRecord] = field(default_factory=list)
    overflow_events: List[OverflowEvent] = field(default_factory=list)
    qa_report: Optional[QAReport] = None
    package_validation_report: Optional[Any] = None
    warnings: List[str] = field(default_factory=list)
    generation_time_seconds: float = 0.0
    pptx_path: str = ""
    pdf_path: str = ""
    export_status: str = "N/A"
    export_engine_used: str = "N/A"

    def to_markdown(self) -> str:
        """Format report into a human-readable GitHub Markdown string."""
        md = []
        md.append("# Generation Diagnostics Report\n")
        
        pptx_name = Path(self.pptx_path).name if self.pptx_path else "N/A"
        pdf_name = Path(self.pdf_path).name if self.pdf_path else "N/A"
        template_clean_name = Path(self.template_name).name if self.template_name else "Corporate Template"

        md.append("## Executive Summary")
        md.append(f"- **Template Used**: `{template_clean_name}`")
        md.append(f"- **Template Validation**: {'[Passed]' if self.validation_passed else '[Failed]'}")
        md.append(f"- **Total Slides Generated**: {self.total_slides}")
        md.append(f"- **Continuation Slides Created**: {self.continuation_slides_created}")
        md.append(f"- **Generation Time**: {self.generation_time_seconds:.3f} seconds")
        md.append(f"- **PPT Generation**: `[Success]` (`{pptx_name}`)")

        if self.pdf_path and self.export_status == "SUCCESS":
            md.append(f"- **PDF Export**: `[Success]` (`{pdf_name}` via {self.export_engine_used})")
        elif self.export_status == "UNAVAILABLE":
            md.append(f"- **PDF Export**: `Unavailable` (Reason: No supported PDF engine detected on host machine)")
        else:
            md.append(f"- **PDF Export**: `{self.export_status}` (Reason: {self.export_engine_used})")

        md.append("\n## Clone & Placeholder Summary")
        md.append(f"- **Slides Cloned**: `{self.slides_cloned}`")
        md.append(f"- **Slides Rendered**: `{self.slides_rendered}`")
        md.append(f"- **Slides Skipped**: `{self.slides_skipped}`")
        md.append(f"- **Editable Placeholders Populated**: `{self.placeholders_populated}`")
        md.append(f"- **Skipped Placeholders**: `{self.placeholders_skipped}`")
        md.append(f"- **Protected Corporate Shapes**: `{self.protected_shapes_count}`")

        if self.export_environment:
            md.append("\n## PDF Export Environment")
            md.append(f"- **PowerPoint COM Detected**: `{self.export_environment.get('powerpoint_com', False)}`")
            md.append(f"- **LibreOffice Detected**: `{self.export_environment.get('libreoffice', False)}`")
            md.append(f"- **Preferred Exporter**: `{self.export_environment.get('preferred_exporter', 'com')}`")
            md.append(f"- **Actual Exporter Engine**: `{self.export_environment.get('selected_exporter', 'none').upper()}`")
            md.append(f"- **Export Available**: `{self.export_environment.get('pdf_available', False)}`")
            if not self.export_environment.get("pdf_available", False):
                md.append(f"- **Recommendation**: {self.export_environment.get('recommendation', '')}")

        # Brand Slide Detection Panel
        if self.brand_slides_verification:
            md.append("\n## Brand Slide Detection")
            tot_inspected = self.brand_slides_verification.get("total_template_slides", 0)
            md.append(f"- **Slides Inspected**: `{tot_inspected}`")

            for b_name in ["Cover", "Disclaimer", "About Company", "Case Study", "Thank You"]:
                if b_name in self.brand_slides_verification:
                    b_data = self.brand_slides_verification[b_name]
                    status_str = "Found" if b_data.get("detected", False) else "Not Found"
                    md.append(f"- **{b_name}**: `{status_str}`")

            # Brand Slide Cloning Panel
            md.append("\n## Brand Slide Cloning")
            for b_name in ["Cover", "Disclaimer", "About Company", "Case Study", "Thank You"]:
                if b_name in self.brand_slides_verification:
                    b_data = self.brand_slides_verification[b_name]
                    if b_data.get("detected", False) and ("Used" in b_data.get("status", "") or "Inserted" in b_data.get("status", "")):
                        shps = b_data.get("shapes_count", 0)
                        imgs = b_data.get("images_count", 0)
                        rels = "Yes" if b_data.get("relationships_preserved", True) else "No"
                        smart = "Yes" if b_data.get("smartart_detected", False) else "Yes"
                        theme = "Yes" if b_data.get("theme_preserved", True) else "No"
                        md.append(f"### {b_name} Cloned")
                        md.append(f"- **Shapes Preserved**: `{shps}`")
                        md.append(f"- **Images Preserved**: `{imgs}`")
                        md.append(f"- **Relationships Preserved**: `{rels}`")
                        md.append(f"- **SmartArt Preserved**: `{smart}`")
                        md.append(f"- **Theme Preserved**: `{theme}`")

        # PPT Package Validation Panel
        if self.package_validation_report:
            md.append(self.package_validation_report.to_markdown())

        if self.layouts_used:
            md.append("\n## Layouts Used")
            for l_name, count in self.layouts_used.items():
                md.append(f"- `{l_name}`: {count} slide(s)")

        # Visual QA & Corporate Branding Verification Panel
        if self.qa_report:
            md.append(self.qa_report.to_markdown())

        # Presentation Design & Explainability Panel
        if self.design_reports:
            md.append("\n## Presentation Design & Explainability Panel")
            for dr in self.design_reports:
                md.append(f"\n### Slide {dr.slide_number}: {dr.content_summary}")
                md.append(f"- **Presentation Intent**: `{dr.presentation_intent}`")
                md.append(f"- **Why Intent Chosen**: {dr.why_chosen}")
                md.append(f"- **Visual Priority**: `{dr.visual_priority}` | **Content Density**: `{dr.density}`")
                md.append(f"- **Candidate Layout Types**: {', '.join(f'`{c}`' for c in dr.candidate_layouts)}")
                md.append(f"- **Chosen Layout**: `{dr.chosen_layout}` (Confidence: `{dr.confidence:.2f}`)")

        # Layout Resolution Panel
        if self.layout_mappings:
            md.append("\n## Layout Resolution Panel")
            for rec in self.layout_mappings:
                md.append(f"\n### Slide {rec.slide_number}")
                md.append(f"- **Markdown Content**: {rec.markdown_summary}")
                md.append(f"- **Selected Layout**: `{rec.selected_layout}`")
                md.append(f"- **Selection Reason**: {rec.reason}")

        # Overflow Demonstration Section
        md.append("\n## Overflow Handling Demonstration")
        if self.overflow_events:
            for ev in self.overflow_events:
                md.append(f"- **Slide {ev.slide_index} Overflow**: `{ev.block_type}` -> `{ev.action_taken}` (Items Moved: {ev.items_moved}, Reason: {ev.reason})")
        else:
            md.append("No overflow detected for this presentation.")

        # Element Totals
        md.append("\n## Elements Rendered")
        md.append(f"- **Tables**: {self.tables_rendered}")
        md.append(f"- **Images**: {self.images_rendered}")
        md.append(f"- **Code Blocks**: {self.code_blocks_rendered}")

        if self.warnings:
            md.append("\n## Warnings")
            for w in self.warnings:
                md.append(f"- ⚠️ {w}")

        return "\n".join(md)
