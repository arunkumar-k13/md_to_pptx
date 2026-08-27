"""Path Utilities Module.

Utility functions for generating unique non-conflicting output file paths.
"""

from __future__ import annotations
from datetime import datetime
from pathlib import Path
from typing import Union


def get_unique_filepath(target_path: Union[str, Path]) -> Path:
    """Resolve a unique non-conflicting file path with a timestamp or index suffix if target_path exists.

    Args:
        target_path: Desired target file path.

    Returns:
        Resolved Path instance guaranteed not to overwrite existing files.
    """
    path = Path(target_path).resolve()
    if not path.exists():
        return path

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = path.parent / f"{path.stem}_{timestamp}{path.suffix}"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = path.parent / f"{path.stem}_{timestamp}_{counter}{path.suffix}"
        if not candidate.exists():
            return candidate
        counter += 1
