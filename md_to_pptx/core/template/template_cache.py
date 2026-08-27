"""Template Cache Module.

Serializes and caches TemplateMetadata to disk (JSON) based on SHA256 hashes and file modification timestamps to prevent redundant XML template analysis.
"""

from __future__ import annotations
import json
import logging
from pathlib import Path
from typing import Optional, Union

from md_to_pptx.core.template.template_analyzer import TemplateAnalyzer, TemplateMetadata

logger = logging.getLogger(__name__)


class TemplateCache:
    """Caching service for TemplateMetadata instances."""

    def __init__(self, cache_dir: Optional[Union[str, Path]] = None) -> None:
        """Initialize TemplateCache with a target cache directory.

        Args:
            cache_dir: Optional directory path for storing JSON cache files.
        """
        if cache_dir:
            self.cache_dir = Path(cache_dir)
        else:
            self.cache_dir = Path.home() / ".gemini" / "antigravity" / "cache" / "templates"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self._analyzer = TemplateAnalyzer()

    def get_or_analyze(self, pptx_path: Union[str, Path]) -> TemplateMetadata:
        """Get cached TemplateMetadata or analyze and cache if missing/outdated.

        Args:
            pptx_path: Path to target PowerPoint template file.

        Returns:
            TemplateMetadata instance.
        """
        path = Path(pptx_path)
        if not path.is_file():
            raise FileNotFoundError(f"Template file not found: {pptx_path}")

        current_mtime = path.stat().st_mtime
        cache_file = self.cache_dir / f"{path.stem}_{path.stat().st_size}.json"

        if cache_file.is_file():
            try:
                raw_data = json.loads(cache_file.read_text(encoding="utf-8"))
                cached_meta = TemplateMetadata.from_dict(raw_data)
                if cached_meta.file_mtime == current_mtime:
                    logger.info("Retrieved cached TemplateMetadata for %s", path.name)
                    return cached_meta
            except Exception as err:
                logger.warning("Failed to read cache file %s: %s. Re-analyzing.", cache_file, err)

        # Cache miss or invalid -> Analyze template
        meta = self._analyzer.analyze(path)
        try:
            cache_file.write_text(json.dumps(meta.to_dict(), indent=2), encoding="utf-8")
            logger.info("Saved fresh TemplateMetadata cache for %s", path.name)
        except Exception as err:
            logger.error("Failed to write TemplateMetadata cache to %s: %s", cache_file, err)

        return meta
