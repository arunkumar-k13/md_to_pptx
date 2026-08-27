"""Template Validator Module.

Performs pre-flight compatibility verification on TemplateMetadata, verifying that
required slide layout types and placeholders exist before generation begins.
Emits a comprehensive diagnostic ValidationReport.
"""

from __future__ import annotations
import logging
from dataclasses import dataclass, field
from typing import List

from md_to_pptx.core.template.template_analyzer import TemplateMetadata

logger = logging.getLogger(__name__)


@dataclass(slots=True)
class ValidationIssue:
    """Diagnostic issue entry within a ValidationReport.

    Attributes:
        severity: 'ERROR' for breaking incompatibilities, 'WARNING' for non-fatal hints.
        message: Detailed diagnostic description.
    """

    severity: str
    message: str


@dataclass(slots=True)
class ValidationReport:
    """Diagnostic report emitted by TemplateValidator.

    Attributes:
        is_valid: True if no critical ERROR severity issues were identified.
        issues: List of ValidationIssue instances.
    """

    is_valid: bool = True
    issues: List[ValidationIssue] = field(default_factory=list)


class TemplateValidator:
    """Pre-flight validator checking TemplateMetadata compatibility."""

    def validate(self, metadata: TemplateMetadata) -> ValidationReport:
        """Validate a TemplateMetadata instance.

        Args:
            metadata: TemplateMetadata object to evaluate.

        Returns:
            ValidationReport instance.
        """
        issues: List[ValidationIssue] = []
        is_valid = True

        if not metadata.layouts:
            issues.append(ValidationIssue("ERROR", "Template contains no slide layouts."))
            return ValidationReport(is_valid=False, issues=issues)

        # 1. Verify Title placeholder layout exists
        has_title_layout = any(layout.has_title for layout in metadata.layouts)
        if not has_title_layout:
            issues.append(ValidationIssue("ERROR", "No layout with a TITLE placeholder found in template."))
            is_valid = False

        # 2. Verify Body placeholder layout exists
        has_body_layout = any(layout.has_body for layout in metadata.layouts)
        if not has_body_layout:
            issues.append(ValidationIssue("ERROR", "No layout with a BODY/CONTENT placeholder found in template."))
            is_valid = False

        # 3. Check for multi-column / two-content layout capabilities (warning only)
        has_two_body = any(layout.body_count >= 2 for layout in metadata.layouts)
        if not has_two_body:
            issues.append(
                ValidationIssue("WARNING", "Template has no multi-body (Two Column) layouts. Content will fall back to single body layouts.")
            )

        # 4. Check for picture placeholder support (warning only)
        has_picture = any(layout.has_picture for layout in metadata.layouts)
        if not has_picture:
            issues.append(
                ValidationIssue("WARNING", "Template has no native PICTURE placeholders. Images will be inserted as positioned shapes.")
            )

        return ValidationReport(is_valid=is_valid, issues=issues)
