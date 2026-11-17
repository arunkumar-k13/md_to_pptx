"""Shared Typography and Text Measurement.

Single source of truth for the font sizes and spacing that the PPTX renderer applies
to body text, and for measuring how tall that text will be. The renderer
(`generator/shape_renderers.py`) and the overflow estimator
(`presentation_model.py` / `overflow/overflow_analyzer.py`) both read from here,
so the paginator's idea of "how much fits" always matches what is actually drawn.
"""

from __future__ import annotations
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from md_to_pptx.core.font_metrics import BOLD_WIDTHS, REGULAR_WIDTHS

# Single-spaced line height of Arial as a fraction of the font size (PowerPoint and
# LibreOffice both lay out Arial at 1.2 em per line before the line-spacing multiple).
ARIAL_LINE_HEIGHT_EM = 1.20
# Text frame insets used by PowerPoint when a template does not override them.
DEFAULT_TEXT_INSET_TOP_BOTTOM_IN = 0.05
DEFAULT_TEXT_INSET_LEFT_RIGHT_IN = 0.10

_DEFAULT_CHAR_EM = 0.556
_EMOJI_EM = 1.10
_CJK_EM = 1.00
_MONO_EM = 0.60


@dataclass(frozen=True)
class Typography:
    """Body text sizes (pt) and spacing (pt) for one slide canvas class."""

    name: str
    body_pt: int
    heading_pt: Dict[int, int]
    list_pt: Dict[int, int]
    quote_pt: int
    line_spacing: float
    para_before_pt: float
    para_after_pt: float
    heading_before_pt: float
    heading_after_pt: float
    list_before_pt: float
    list_after_pt: float
    quote_before_pt: float
    quote_after_pt: float
    # Left margin (marL) of each bullet level, in inches, from the slide master body style
    list_indent_in: Tuple[float, ...] = field(default=(0.375, 0.81, 1.25))

    def heading_size(self, level: int) -> int:
        """Font size for a heading level; level 0 means a bold label (rendered like H3)."""
        if level <= 0:
            return self.heading_pt[3]
        return self.heading_pt.get(min(level, 6), self.heading_pt[6])

    def list_size(self, depth: int) -> int:
        return self.list_pt.get(depth, self.list_pt[2])

    def list_indent(self, depth: int) -> float:
        return self.list_indent_in[min(depth, len(self.list_indent_in) - 1)]


_MC_SCALE = 1.35  # MC corporate canvas (15" tall) scales the 7.5" sizes by at most 1.35

WIDESCREEN = Typography(
    name="widescreen_16x9",
    body_pt=15,
    heading_pt={1: 22, 2: 20, 3: 18, 4: 16, 5: 15, 6: 14},
    list_pt={0: 15, 1: 13, 2: 12},
    quote_pt=18,
    line_spacing=1.15,
    para_before_pt=2, para_after_pt=3,
    heading_before_pt=6, heading_after_pt=2,
    list_before_pt=1, list_after_pt=3,
    quote_before_pt=6, quote_after_pt=6,
)

STANDARD = Typography(
    name="standard_4x3",
    body_pt=16,
    heading_pt={1: 24, 2: 22, 3: 19, 4: 17, 5: 16, 6: 15},
    list_pt={0: 16, 1: 15, 2: 13},
    quote_pt=18,
    line_spacing=1.15,
    para_before_pt=2, para_after_pt=4,
    heading_before_pt=8, heading_after_pt=3,
    list_before_pt=2, list_after_pt=4,
    quote_before_pt=6, quote_after_pt=6,
)

LARGE_CANVAS = Typography(
    name="large_canvas",
    body_pt=int(18 * _MC_SCALE),
    heading_pt={
        1: int(32 * _MC_SCALE), 2: int(28 * _MC_SCALE), 3: int(24 * _MC_SCALE),
        4: int(20 * _MC_SCALE), 5: int(18 * _MC_SCALE), 6: int(16 * _MC_SCALE),
    },
    list_pt={0: int(22 * _MC_SCALE), 1: int(19 * _MC_SCALE), 2: int(17 * _MC_SCALE)},
    quote_pt=18,
    line_spacing=1.15,
    para_before_pt=4, para_after_pt=12,
    heading_before_pt=20, heading_after_pt=8,
    list_before_pt=int(3 * _MC_SCALE), list_after_pt=int(6 * _MC_SCALE),
    quote_before_pt=6, quote_after_pt=6,
    list_indent_in=(0.5, 1.5, 2.5),
)


def typography_for(slide_height_inches: float) -> Typography:
    """Pick the typography class for a slide canvas height."""
    if slide_height_inches < 6.5:
        return WIDESCREEN
    if slide_height_inches > 10.0:
        return LARGE_CANVAS
    return STANDARD


# --------------------------------------------------------------------------------------
# Text measurement
# --------------------------------------------------------------------------------------

def _char_em(ch: str, bold: bool, mono: bool) -> float:
    if mono:
        return _MONO_EM
    cp = ord(ch)
    table = BOLD_WIDTHS if bold else REGULAR_WIDTHS
    w = table.get(cp)
    if w is not None:
        return w / 1000.0
    if 0x1F000 <= cp <= 0x1FAFF or 0x2600 <= cp <= 0x27BF:
        return _EMOJI_EM
    if 0xFE00 <= cp <= 0xFE0F or cp == 0x200D:  # variation selectors / joiners take no width
        return 0.0
    if 0x2E80 <= cp <= 0x9FFF or 0xAC00 <= cp <= 0xD7AF:
        return _CJK_EM
    return _DEFAULT_CHAR_EM


def text_width_in(text: str, size_pt: float, bold: bool = False, mono: bool = False) -> float:
    """Width of a single line of text in inches."""
    return sum(_char_em(c, bold, mono) for c in text) * size_pt / 72.0


def wrapped_line_count(text: str, width_in: float, size_pt: float, bold: bool = False, mono: bool = False) -> int:
    """Number of lines `text` occupies when word-wrapped into `width_in` inches."""
    return wrapped_line_count_styled([(text or "", bold, mono)], width_in, size_pt)


def wrapped_line_count_styled(segments: List[Tuple[str, bool, bool]], width_in: float, size_pt: float) -> int:
    """Line count for text made of styled segments (text, bold, monospace), word-wrapped.

    Inline code is set in a monospace font that is much wider than Arial, so mixed
    runs are measured character by character in their own style.
    """
    if width_in <= 0:
        return 1
    scale = size_pt / 72.0
    total_lines = 1
    cur = 0.0          # width of the current line
    word = 0.0         # width of the word being built
    space_w = _char_em(" ", False, False) * scale

    def place_word(word_w: float) -> None:
        nonlocal cur, total_lines
        if word_w <= 0:
            return
        if word_w > width_in:
            # A word longer than a line breaks across lines
            if cur > 0:
                total_lines += 1
            total_lines += int(word_w // width_in)
            cur = word_w % width_in
            return
        needed = word_w if cur == 0 else cur + space_w + word_w
        if needed <= width_in:
            cur = needed
        else:
            total_lines += 1
            cur = word_w

    for text, bold, mono in segments:
        for ch in text:
            if ch == "\n":
                place_word(word)
                word = 0.0
                total_lines += 1
                cur = 0.0
            elif ch == " ":
                place_word(word)
                word = 0.0
            else:
                word += _char_em(ch, bold, mono) * scale
    place_word(word)
    return max(1, total_lines)


def runs_to_segments(text: str, runs: Optional[list], bold: bool = False) -> List[Tuple[str, bool, bool]]:
    """Styled segments for measurement from inline runs, falling back to plain text."""
    if runs:
        segs = [(r.text, bold or bool(getattr(r, "is_bold", False)), bool(getattr(r, "is_code", False)))
                for r in runs if getattr(r, "text", "")]
        if segs:
            return segs
    return [(text or "", bold, False)]


def line_height_in(size_pt: float, line_spacing: float) -> float:
    return size_pt * ARIAL_LINE_HEIGHT_EM * line_spacing / 72.0


def paragraph_height_in(
    text: str, width_in: float, size_pt: float, line_spacing: float,
    before_pt: float, after_pt: float, bold: bool = False, mono: bool = False,
    runs: Optional[list] = None,
) -> float:
    """Height of one rendered paragraph including its spacing before and after."""
    if runs:
        lines = wrapped_line_count_styled(runs_to_segments(text, runs, bold), width_in, size_pt)
    else:
        lines = wrapped_line_count(text, width_in, size_pt, bold, mono)
    return lines * line_height_in(size_pt, line_spacing) + (before_pt + after_pt) / 72.0


_SENTENCE_END = re.compile(r"(?<=[.!?])\s+(?=\S)")


def split_sentences(text: str) -> List[str]:
    """Split text into sentences, keeping punctuation with each sentence."""
    parts = _SENTENCE_END.split(text.strip())
    return [p for p in parts if p]
