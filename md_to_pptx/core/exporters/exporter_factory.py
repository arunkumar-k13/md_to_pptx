"""Exporter Factory Module.

Factory resolving LibreOfficeExporter implementation based on runtime operating system capabilities.
"""

from __future__ import annotations
import logging
from typing import Optional

from md_to_pptx.config.settings_loader import Settings
from md_to_pptx.core.exporters.base_exporter import BaseExporter
from md_to_pptx.core.exporters.environment import PdfEnvironmentDetector
from md_to_pptx.core.exporters.libreoffice_exporter import LibreOfficeExporter

logger = logging.getLogger(__name__)


class ExporterFactory:
    """Factory resolving available PDF exporter plugins."""

    def get_exporter(
        self, preferred_exporter: Optional[str] = None, settings: Optional[Settings] = None
    ) -> Optional[BaseExporter]:
        """Resolve and instantiate LibreOfficeExporter or None if unavailable.

        Args:
            preferred_exporter: Preferred exporter key ('libreoffice').
            settings: Settings configuration instance.

        Returns:
            Instantiated LibreOfficeExporter implementation or None.
        """
        cfg = settings or Settings()
        report = PdfEnvironmentDetector.detect(preferred_exporter=preferred_exporter, settings=cfg)

        if report.selected_exporter == "libreoffice":
            logger.info("Selected LibreOfficeExporter for PDF export.")
            return LibreOfficeExporter(binary_path=cfg.export.libreoffice_binary_path)

        logger.info("LibreOffice PDF exporter engine is not available on this system.")
        return None
