"""PowerPoint COM PDF Exporter Module (Disabled / Removed Dependency).

Microsoft PowerPoint COM automation has been disabled to eliminate the dependency
on Microsoft PowerPoint software. All PDF exports are handled via LibreOffice.
"""

from __future__ import annotations
import logging
from pathlib import Path
from typing import Union

from md_to_pptx.core.exporters.base_exporter import BaseExporter

logger = logging.getLogger(__name__)


class COMExporter(BaseExporter):
    """Legacy COM Exporter stub. Disabled to remove Microsoft PowerPoint dependency."""

    def is_available(self) -> bool:
        """Check if Microsoft PowerPoint COM automation is available. Always returns False."""
        return False

    def export_pdf(self, pptx_path: Union[str, Path], output_pdf_path: Union[str, Path]) -> Path:
        """Disabled COM export method.

        Raises:
            RuntimeError: Always raised to indicate MS PowerPoint dependency is removed.
        """
        raise RuntimeError(
            "Microsoft PowerPoint COM dependency has been removed from this project. "
            "Please use LibreOffice ('soffice') for PDF export."
        )
