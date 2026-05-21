from __future__ import annotations

import re

from captioned_text import (
    TIMESTAMP_LINE_PATTERN,
    normalize_transcript_paragraph,
    parse_transcript_value,
)
from capture_types import TargetWindow
from macos_accessibility import (
    AccessibilityTextRecord,
    create_application,
    dependency_error as accessibility_dependency_error,
    describe_window,
    find_ax_window,
    find_window_info,
    kCGWindowOwnerPID,
    extract_text_records,
    normalize_line,
)


TEXT_ROLES = {"AXStaticText", "AXTextArea"}
WORD_TOKEN_PATTERN = re.compile(r"^[A-Za-z][A-Za-z'.,?!-]*$")


class MacCaptionedSession:
    def __init__(self, target: TargetWindow):
        self.target = target
        self.window_info: dict | None = None
        self.app_element = None
        self.window_element = None

    def refresh(self) -> str | None:
        window_info = find_window_info(self.target)
        if window_info is None:
            return (
                f"Window not found.\n"
                f"Process: {self.target.process_name or '(any)'}\n"
                f"Window: {self.target.window_name or '(any)'}"
            )

        pid = int(window_info.get(kCGWindowOwnerPID, 0) or 0)
        if pid <= 0:
            return "Target window did not expose a valid process id."

        if self.window_info == window_info and self.window_element is not None:
            return None

        self.window_info = window_info
        self.app_element = create_application(pid)
        self.window_element = find_ax_window(self.app_element, window_info)
        if self.window_element is None:
            return (
                "Transcript window was not found through macOS Accessibility.\n"
                "Make sure Accessibility permission is granted to this app."
            )
        return None

    def extract_text(self) -> str:
        refresh_error = self.refresh()
        if refresh_error:
            return refresh_error

        records = extract_text_records(self.window_element)
        lines = transcript_lines_from_records(records)
        if not lines:
            return "(No text detected yet.)"

        parsed = parse_transcript_value("\n".join(lines))
        paragraphs = parsed or lines
        if not paragraphs:
            return "(No text detected yet.)"
        return "\n\n".join(paragraphs)


def dependency_error(prompt: bool = False) -> str | None:
    return accessibility_dependency_error(prompt=prompt)


def describe_target(target: TargetWindow) -> str:
    window_info = find_window_info(target)
    if window_info is None:
        return (
            f"No matching target window found.\n"
            f"Process: {target.process_name or '(any)'}\n"
            f"Window: {target.window_name or '(any)'}"
        )
    return describe_window(window_info)


def capture_text(target: TargetWindow) -> str:
    return MacCaptionedSession(target).extract_text()


def transcript_lines_from_records(records: list[AccessibilityTextRecord]) -> list[str]:
    word_stream = build_captioned_word_stream(records)
    if word_stream:
        return [word_stream]
    return []


def build_captioned_word_stream(records: list[AccessibilityTextRecord]) -> str:
    segments = caption_word_segments(records)
    if not segments:
        return ""

    paragraphs: list[str] = []
    for segment in segments:
        paragraph = build_visual_word_stream(segment) or build_traversal_word_stream(segment)
        if paragraph:
            paragraphs.append(paragraph)
    paragraphs = dedupe_lines(paragraphs)
    if not paragraphs:
        return ""
    return normalize_transcript_paragraph(" ".join(paragraphs))


def caption_word_segments(records: list[AccessibilityTextRecord]) -> list[list[AccessibilityTextRecord]]:
    segments: list[list[AccessibilityTextRecord]] = []
    current: list[AccessibilityTextRecord] = []
    seen_transcript_time = False

    for record in records:
        text = normalize_line(record.text)
        if not text:
            continue
        if record.role not in TEXT_ROLES:
            continue
        if TIMESTAMP_LINE_PATTERN.match(text):
            if current:
                segments.append(current)
                current = []
            seen_transcript_time = True
            continue
        if not seen_transcript_time:
            continue
        if not is_caption_word_token(text):
            continue
        if not normalize_caption_word(text):
            continue
        current.append(record)
    if current:
        segments.append(current)
    return [segment for segment in segments if len(segment) >= 4]


def build_traversal_word_stream(records: list[AccessibilityTextRecord]) -> str:
    words: list[str] = []
    recent_tokens: list[str] = []
    for record in records:
        text = normalize_line(record.text)
        normalized = normalize_caption_word(text)
        if normalized in recent_tokens:
            continue
        words.append(text)
        recent_tokens.append(normalized)
        if len(recent_tokens) > 8:
            recent_tokens.pop(0)
    if len(words) < 4:
        return ""
    return normalize_transcript_paragraph(" ".join(words))


def build_visual_word_stream(records: list[AccessibilityTextRecord]) -> str:
    positioned = [record for record in records if record.x is not None and record.y is not None]
    if len(positioned) < 4:
        return ""

    rows: list[list[AccessibilityTextRecord]] = []
    for record in sorted(positioned, key=lambda item: (item.y or 0.0, item.x or 0.0)):
        for row in rows:
            row_y = sum(item.y or 0.0 for item in row) / len(row)
            tolerance = max(8.0, ((record.height or 0.0) + (row[0].height or 0.0)) / 2.0)
            if abs((record.y or 0.0) - row_y) <= tolerance:
                row.append(record)
                break
        else:
            rows.append([record])

    words: list[str] = []
    recent_rows: list[str] = []
    for row in rows:
        row_words: list[str] = []
        row_recent: list[str] = []
        last_x: float | None = None
        for record in sorted(row, key=lambda item: item.x or 0.0):
            text = normalize_line(record.text)
            normalized = normalize_caption_word(text)
            if not normalized:
                continue
            if normalized in row_recent:
                continue
            if last_x is not None and record.x is not None and abs(record.x - last_x) <= 2:
                continue
            row_words.append(text)
            row_recent.append(normalized)
            last_x = record.x
            if len(row_recent) > 12:
                row_recent.pop(0)

        row_text = normalized_compare_text(" ".join(row_words))
        if not row_words or row_text in recent_rows:
            continue
        words.extend(row_words)
        recent_rows.append(row_text)
        if len(recent_rows) > 12:
            recent_rows.pop(0)

    if len(words) < 4:
        return ""
    return normalize_transcript_paragraph(" ".join(words))


def is_caption_word_token(text: str) -> bool:
    if " " in text:
        return False
    if len(text) > 32:
        return False
    return bool(WORD_TOKEN_PATTERN.match(text))


def normalize_caption_word(text: str) -> str:
    return re.sub(r"[^a-z0-9']+", "", text.casefold())


def normalized_compare_text(text: str) -> str:
    return " ".join(text.casefold().split())


def dedupe_lines(lines: list[str]) -> list[str]:
    deduped: list[str] = []
    seen_recent: list[str] = []
    for line in lines:
        if not line:
            continue
        normalized = normalized_compare_text(line)
        if deduped and normalized == normalized_compare_text(deduped[-1]):
            continue
        if normalized in seen_recent:
            continue
        deduped.append(line)
        seen_recent.append(normalized)
        if len(seen_recent) > 40:
            seen_recent.pop(0)
    return deduped
