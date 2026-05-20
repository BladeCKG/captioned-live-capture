from __future__ import annotations

from chrome_live_caption_common import (
    LIVE_CAPTION_CLASS,
    LIVE_CAPTION_PROCESS_NAME,
    LIVE_CAPTION_WINDOW_NAME,
    is_live_caption_target,
    merge_live_caption_history,
    split_live_caption_paragraphs,
)
from chrome_live_caption_windows import find_live_caption_document
from chrome_live_caption_windows import (
    extract_live_caption_paragraphs as _extract_windows_live_caption_paragraphs,
)


def extract_live_caption_paragraphs(root_control, find_first_control, normalize_paragraph=None):
    # Compatibility wrapper for older callers; normalization now lives in the
    # Windows-specific implementation.
    return _extract_windows_live_caption_paragraphs(root_control, find_first_control)

__all__ = [
    "LIVE_CAPTION_CLASS",
    "LIVE_CAPTION_PROCESS_NAME",
    "LIVE_CAPTION_WINDOW_NAME",
    "extract_live_caption_paragraphs",
    "find_live_caption_document",
    "is_live_caption_target",
    "merge_live_caption_history",
    "split_live_caption_paragraphs",
]
