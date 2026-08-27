"""LibreOffice Headless PDF Exporter Module.

Cross-platform PDF exporter leveraging headless LibreOffice CLI ('soffice').
Automatically checks system PATH and standard installation paths (e.g. C:\\Program Files\\LibreOffice\\program\\soffice.exe).
"""

from __future__ import annotations
import logging
import os
import shutil
import subprocess
from pathlib import Path
from typing import Optional, Union

from md_to_pptx.core.exporters.base_exporter import BaseExporter

logger = logging.getLogger(__name__)


class LibreOfficeExporter(BaseExporter):
    """PDF Exporter using headless LibreOffice CLI command."""

    def __init__(self, binary_path: Optional[str] = None) -> None:
        """Initialize LibreOfficeExporter with optional binary path.

        Args:
            binary_path: Optional command executable path for LibreOffice ('soffice').
        """
        self.binary_path = self._resolve_binary_path(binary_path)

    @staticmethod
    def _resolve_binary_path(user_path: Optional[str]) -> str:
        """Resolve valid LibreOffice binary path from user settings, PATH, or standard installation directories.

        Args:
            user_path: Optional explicit binary path string.

        Returns:
            Resolved binary executable path.
        """
        if user_path and user_path != "soffice":
            if Path(user_path).is_file() or shutil.which(user_path):
                return user_path

        env_path = os.getenv("LIBREOFFICE_BINARY")
        if env_path:
            if Path(env_path).is_file() or shutil.which(env_path):
                return env_path

        # 1. Check system PATH
        in_path = shutil.which("soffice")
        if in_path:
            return in_path

        # 2. Check standard Windows installation paths
        win_candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for cand in win_candidates:
            if Path(cand).is_file():
                return cand

        # 3. Check standard Linux / macOS installation paths
        nix_candidates = [
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
        for cand in nix_candidates:
            if Path(cand).is_file():
                return cand

        return "soffice"

    def is_available(self) -> bool:
        """Check if LibreOffice executable is accessible."""
        if Path(self.binary_path).is_file():
            return True
        return shutil.which(self.binary_path) is not None

    def export_pdf(self, pptx_path: Union[str, Path], output_pdf_path: Union[str, Path]) -> Path:
        """Export PowerPoint file to PDF using headless LibreOffice.

        Args:
            pptx_path: Target PPTX file path.
            output_pdf_path: Target output PDF file path.

        Returns:
            Resolved Path instance to generated PDF file.
        """
        if not self.is_available():
            raise RuntimeError(f"LibreOffice binary ('{self.binary_path}') is not available.")

        src_path = Path(pptx_path).resolve()
        out_path = Path(output_pdf_path).resolve()
        out_dir = out_path.parent
        out_dir.mkdir(parents=True, exist_ok=True)

        if not src_path.is_file():
            raise FileNotFoundError(f"PPTX file not found: {pptx_path}")

        cmd = [
            self.binary_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(src_path),
        ]

        logger.info("Executing LibreOffice command: %s", " ".join(cmd))

        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if res.returncode != 0:
            logger.error("LibreOffice export failed (code %d): %s", res.returncode, res.stderr)
            raise RuntimeError(f"LibreOffice conversion failed: {res.stderr}")

        # LibreOffice outputs output_name.pdf in outdir
        expected_pdf = out_dir / f"{src_path.stem}.pdf"
        if expected_pdf.is_file() and expected_pdf != out_path:
            shutil.move(str(expected_pdf), str(out_path))

        if not out_path.is_file():
            raise FileNotFoundError(f"LibreOffice output PDF was not produced at expected path: {out_path}")

        logger.info("Successfully exported PDF via LibreOffice to %s", out_path)
        return out_path
