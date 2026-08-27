"""PDF Exporter Environment Detector Module.

Pre-flight environment detection utility checking host machine capabilities for LibreOffice executables.
"""

from __future__ import annotations
import logging
from dataclasses import dataclass
from typing import Any, Dict, Optional

from md_to_pptx.config.settings_loader import Settings
from md_to_pptx.core.exporters.libreoffice_exporter import LibreOfficeExporter

logger = logging.getLogger(__name__)


@dataclass(slots=True)
class PdfEnvironmentReport:
    """Pre-flight environment detection report for PDF export capabilities.

    Attributes:
        powerpoint_com_available: Always False (MS PowerPoint dependency removed).
        libreoffice_available: True if LibreOffice executable ('soffice') is in PATH.
        preferred_exporter: Exporter key ('libreoffice').
        selected_exporter: Actual exporter engine selected ('libreoffice' or 'none').
        pdf_available: True if LibreOffice is ready on host machine.
        recommendation: Human-readable setup recommendation if unavailable.
    """

    powerpoint_com_available: bool = False
    libreoffice_available: bool = False
    preferred_exporter: str = "libreoffice"
    selected_exporter: str = "none"
    pdf_available: bool = False
    recommendation: str = ""

    def to_dict(self) -> Dict[str, Any]:
        """Export report as dictionary for API integration."""
        return {
            "powerpoint_com": False,
            "libreoffice": self.libreoffice_available,
            "preferred_exporter": self.preferred_exporter,
            "selected_exporter": self.selected_exporter,
            "pdf_available": self.pdf_available,
            "recommendation": self.recommendation,
        }


class PdfEnvironmentDetector:
    """Detector for inspecting host PDF exporter capabilities."""

    @staticmethod
    def detect(preferred_exporter: str = "libreoffice", settings: Optional[Settings] = None) -> PdfEnvironmentReport:
        """Inspect host environment and return a PdfEnvironmentReport.

        Args:
            preferred_exporter: Preferred exporter key ('libreoffice').
            settings: Settings configuration instance.

        Returns:
            Populated PdfEnvironmentReport instance.
        """
        cfg = settings or Settings()
        pref = "libreoffice"

        lo_exp = LibreOfficeExporter(binary_path=cfg.export.libreoffice_binary_path)
        lo_ok = lo_exp.is_available()

        selected = "libreoffice" if lo_ok else "none"
        pdf_ok = lo_ok
        rec = "" if pdf_ok else "No supported PDF engine detected. Install LibreOffice ('soffice') and add it to PATH to enable PDF export."

        return PdfEnvironmentReport(
            powerpoint_com_available=False,
            libreoffice_available=lo_ok,
            preferred_exporter=pref,
            selected_exporter=selected,
            pdf_available=pdf_ok,
            recommendation=rec,
        )
