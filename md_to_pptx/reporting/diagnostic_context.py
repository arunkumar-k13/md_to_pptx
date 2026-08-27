"""Diagnostic Context Module.

Runtime telemetry collector tracking generation duration, layout usage, overflow events, layout mapping decisions, and warnings.
"""

from __future__ import annotations
import time
from typing import Dict, List, Optional

from md_to_pptx.reporting.generation_report import DesignReportRecord, GenerationReport, LayoutMappingRecord, OverflowEvent


class DiagnosticContext:
    """Telemetry collector compiling runtime metrics into a GenerationReport."""

    def __init__(self) -> None:
        self._start_time: Optional[float] = None
        self._end_time: Optional[float] = None
        self._layouts_used: Dict[str, int] = {}
        self._design_reports: List[DesignReportRecord] = []
        self._layout_mappings: List[LayoutMappingRecord] = []
        self._overflow_events: List[OverflowEvent] = []
        self._warnings: List[str] = []
        self.template_name: str = ""
        self.validation_passed: bool = True
        self.cache_used: bool = False
        self.tables_rendered: int = 0
        self.images_rendered: int = 0
        self.code_blocks_rendered: int = 0
        self.slides_cloned: int = 0
        self.slides_rendered: int = 0
        self.slides_skipped: int = 0
        self.placeholders_populated: int = 0
        self.placeholders_skipped: int = 0
        self.protected_shapes_count: int = 0
        self.total_slides: int = 0
        self.continuation_slides: int = 0
        self.pptx_path: str = ""
        self.pdf_path: str = ""
        self.export_status: str = "N/A"
        self.export_engine_used: str = "N/A"
        self.export_environment: Dict[str, Any] = {}
        self.brand_slides_verification: Dict[str, Dict[str, Any]] = {}
        self.qa_report: Optional[Any] = None
        self.package_validation_report: Optional[Any] = None

    def start_timer(self) -> None:
        """Start the generation timer."""
        self._start_time = time.perf_counter()

    def stop_timer(self) -> None:
        """Stop the generation timer."""
        self._end_time = time.perf_counter()

    def record_layout(self, layout_name: str) -> None:
        """Record usage of a slide layout."""
        self._layouts_used[layout_name] = self._layouts_used.get(layout_name, 0) + 1

    def record_design_report(
        self,
        slide_number: int,
        content_summary: str,
        presentation_intent: str,
        why_chosen: str,
        visual_priority: str,
        density: str,
        candidate_layouts: List[str],
        chosen_layout: str,
        confidence: float,
    ) -> None:
        """Record an explicit presentation design intent decision."""
        self._design_reports.append(
            DesignReportRecord(
                slide_number=slide_number,
                content_summary=content_summary,
                presentation_intent=presentation_intent,
                why_chosen=why_chosen,
                visual_priority=visual_priority,
                density=density,
                candidate_layouts=candidate_layouts,
                chosen_layout=chosen_layout,
                confidence=confidence,
            )
        )

    def record_layout_mapping(
        self, slide_number: int, markdown_summary: str, selected_layout: str, reason: str
    ) -> None:
        """Record an explicit slide layout mapping decision with architectural rationale."""
        self._layout_mappings.append(
            LayoutMappingRecord(
                slide_number=slide_number,
                markdown_summary=markdown_summary,
                selected_layout=selected_layout,
                reason=reason,
            )
        )

    def record_overflow(
        self, slide_index: int, block_type: str, action_taken: str, items_moved: int = 0, reason: str = ""
    ) -> None:
        """Record an overflow resolution event."""
        self._overflow_events.append(
            OverflowEvent(
                slide_index=slide_index,
                block_type=block_type,
                action_taken=action_taken,
                items_moved=items_moved,
                reason=reason,
            )
        )

    def record_warning(self, message: str) -> None:
        """Record a diagnostic warning message."""
        self._warnings.append(message)

    def build_report(self) -> GenerationReport:
        """Compile recorded metrics into a final GenerationReport."""
        elapsed = 0.0
        if self._start_time and self._end_time:
            elapsed = self._end_time - self._start_time
        elif self._start_time:
            elapsed = time.perf_counter() - self._start_time

        return GenerationReport(
            template_name=self.template_name,
            validation_passed=self.validation_passed,
            cache_used=self.cache_used,
            total_slides=self.total_slides,
            continuation_slides_created=self.continuation_slides,
            layouts_used=self._layouts_used,
            tables_rendered=self.tables_rendered,
            images_rendered=self.images_rendered,
            code_blocks_rendered=self.code_blocks_rendered,
            slides_cloned=self.slides_cloned,
            slides_rendered=self.slides_rendered,
            slides_skipped=self.slides_skipped,
            placeholders_populated=self.placeholders_populated,
            placeholders_skipped=self.placeholders_skipped,
            protected_shapes_count=self.protected_shapes_count,
            export_environment=self.export_environment,
            brand_slides_verification=self.brand_slides_verification,
            design_reports=self._design_reports,
            layout_mappings=self._layout_mappings,
            overflow_events=self._overflow_events,
            qa_report=self.qa_report,
            package_validation_report=self.package_validation_report,
            warnings=self._warnings,
            generation_time_seconds=elapsed,
            pptx_path=self.pptx_path,
            pdf_path=self.pdf_path,
            export_status=self.export_status,
            export_engine_used=self.export_engine_used,
        )
