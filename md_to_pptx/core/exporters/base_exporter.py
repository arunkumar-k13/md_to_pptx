"""Base Exporter Plugin Interface Module.

Defines the Abstract Base Class interface for all presentation PDF exporters.
"""

from __future__ import annotations
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Union


class BaseExporter(ABC):
    """Abstract Base Class for presentation PDF export plugins."""

    @abstractmethod
    def export_pdf(self, pptx_path: Union[str, Path], output_pdf_path: Union[str, Path]) -> Path:
        """Export a PowerPoint file (.pptx) to PDF (.pdf).

        Args:
            pptx_path: Path to target PPTX presentation.
            output_pdf_path: Target path for output PDF document.

        Returns:
            Resolved Path instance to generated PDF file.
        """
        pass

    @abstractmethod
    def is_available(self) -> bool:
        """Check if this export engine is available on the current operating system/environment.

        Returns:
            True if engine can be invoked, False otherwise.
        """
        pass
