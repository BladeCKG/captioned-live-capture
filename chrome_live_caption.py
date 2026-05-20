from __future__ import annotations

import difflib
import re
import sys


LIVE_CAPTION_WINDOW_NAME = "Live Caption"
LIVE_CAPTION_CLASS = "Chrome_WidgetWin_1" if sys.platform.startswith("win") else ""
LIVE_CAPTION_PROCESS_NAME = "chrome.exe" if sys.platform.startswith("win") else "Google Chrome"
WORD_PATTERN = re.compile(r"\S+")
STRIP_PUNCTUATION = ".,!?;:'\"()[]{}<>"
LIVE_CAPTION_TAIL_WINDOW_WORDS = 140
LIVE_CAPTION_MIN_ANCHOR_WORDS = 6


def is_live_caption_target(target) -> bool:
    window_name = (target.window_name or "").strip()
    expected_class = (target.expected_class or "").strip()
    process_name = (target.process_name or "").strip().casefold()
    if window_name == LIVE_CAPTION_WINDOW_NAME and (
        not LIVE_CAPTION_CLASS or expected_class == LIVE_CAPTION_CLASS
    ):
        return True
    class_matches = not LIVE_CAPTION_CLASS or expected_class == LIVE_CAPTION_CLASS
    return class_matches and process_name in {"", LIVE_CAPTION_PROCESS_NAME.casefold()}


def find_live_caption_document(root_control, find_first_control):
    return find_first_control(
        root_control,
        lambda control: getattr(control, "ControlTypeName", "") == "DocumentControl"
        and getattr(control, "ClassName", "") == "CaptionBubbleLabel",
    )


def split_live_caption_paragraphs(raw_text: str) -> list[str]:
    text = re.sub(r"\s+", " ", raw_text or "").strip()
    if not text:
        return []
    return [text]


def normalize_live_caption_lines(lines: list[str], normalize_paragraph) -> list[str]:
    paragraph = normalize_paragraph(" ".join(lines))
    if not paragraph:
        return []
    return [paragraph]


def extract_live_caption_paragraphs(root_control, find_first_control, normalize_paragraph) -> list[str]:
    document = find_live_caption_document(root_control, find_first_control)
    if document is None:
        return []

    try:
        name_text = (document.Name or "").strip()
    except Exception:
        name_text = ""

    if name_text:
        return split_live_caption_paragraphs(name_text)

    lines: list[str] = []
    try:
        for child in document.GetChildren():
            if child.ControlTypeName != "TextControl":
                continue
            if child.ClassName != "AXVirtualView":
                continue
            line = (child.Name or "").strip()
            if line:
                lines.append(line)
    except Exception:
        return []

    if not lines:
        return []
    return normalize_live_caption_lines(lines, normalize_paragraph)


def merge_live_caption_history(existing_paragraphs: list[str], new_paragraphs: list[str]) -> list[str]:
    existing_text = " ".join(paragraph.strip() for paragraph in existing_paragraphs if paragraph.strip())
    candidate_text = " ".join(paragraph.strip() for paragraph in new_paragraphs if paragraph.strip())

    if not existing_text:
        return [candidate_text] if candidate_text else []
    if not candidate_text:
        return [existing_text]

    merged_text, changed = merge_live_caption_paragraph(existing_text, candidate_text)
    if not changed:
        return [existing_text]
    return [merged_text]


def merge_live_caption_paragraph(existing: str, candidate: str) -> tuple[str, bool]:
    existing_norm = normalize_caption_text(existing)
    candidate_norm = normalize_caption_text(candidate)
    if not existing_norm or not candidate_norm:
        return existing, False
    if existing_norm == candidate_norm:
        return existing, False
    if candidate_norm.startswith(existing_norm):
        return candidate, True
    if existing_norm.startswith(candidate_norm):
        return existing, False

    if candidate_norm in existing_norm:
        return existing, False
    if existing_norm in candidate_norm:
        return candidate, True

    existing_words = WORD_PATTERN.findall(existing)
    candidate_words = WORD_PATTERN.findall(candidate)

    revised_tail = merge_live_caption_tail(existing_words, candidate_words)
    if revised_tail is not None:
        return revised_tail, revised_tail != existing

    overlap = word_overlap_suffix_prefix(existing_words, candidate_words)
    if overlap >= 3 and overlap < len(candidate_words):
        merged_words = existing_words + candidate_words[overlap:]
        return " ".join(merged_words), True

    return existing, False


def merge_live_caption_tail(existing_words: list[str], candidate_words: list[str]) -> str | None:
    if not existing_words or not candidate_words:
        return None

    tail_start = max(0, len(existing_words) - LIVE_CAPTION_TAIL_WINDOW_WORDS)
    tail_words = existing_words[tail_start:]
    tail_normalized = [normalize_word(word) for word in tail_words]
    candidate_normalized = [normalize_word(word) for word in candidate_words]

    matcher = difflib.SequenceMatcher(None, tail_normalized, candidate_normalized, autojunk=False)
    match = matcher.find_longest_match(0, len(tail_normalized), 0, len(candidate_normalized))
    if match.size < LIVE_CAPTION_MIN_ANCHOR_WORDS:
        return None

    global_start = tail_start + match.a
    merged_words = existing_words[:global_start] + candidate_words[match.b:]
    if not merged_words:
        return None
    return " ".join(merged_words)


def word_overlap_suffix_prefix(existing_words: list[str], candidate_words: list[str]) -> int:
    max_overlap = min(len(existing_words), len(candidate_words))
    normalized_existing = [normalize_word(word) for word in existing_words]
    normalized_candidate = [normalize_word(word) for word in candidate_words]
    for overlap in range(max_overlap, 0, -1):
        if normalized_existing[-overlap:] == normalized_candidate[:overlap]:
            return overlap
    return 0


def normalize_caption_text(text: str) -> str:
    words = [normalize_word(word) for word in WORD_PATTERN.findall(text)]
    return " ".join(word for word in words if word)


def normalize_word(word: str) -> str:
    return word.strip(STRIP_PUNCTUATION).casefold()
