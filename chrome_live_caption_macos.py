from __future__ import annotations

from capture_types import TargetWindow
from chrome_live_caption_common import split_live_caption_paragraphs
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


class MacChromeLiveCaptionSession:
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
                "Live Caption window was not found through macOS Accessibility.\n"
                "Make sure Accessibility permission is granted to this app."
            )
        return None

    def extract_text(self) -> str:
        refresh_error = self.refresh()
        if refresh_error:
            return refresh_error
        records = extract_text_records(self.window_element)
        paragraph = live_caption_text_from_records(records)
        if not paragraph:
            return "(No text detected yet.)"
        return paragraph


def dependency_error() -> str | None:
    return accessibility_dependency_error()


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
    return MacChromeLiveCaptionSession(target).extract_text()


def live_caption_text_from_records(records: list[AccessibilityTextRecord]) -> str:
    lines: list[str] = []
    for record in records:
        text = normalize_line(record.text)
        if not text:
            continue
        if record.role not in {"AXStaticText", "AXTextArea"}:
            continue
        lines.append(text)
    paragraphs = split_live_caption_paragraphs(" ".join(dedupe_lines(lines)))
    return "\n\n".join(paragraphs)


def dedupe_lines(lines: list[str]) -> list[str]:
    deduped: list[str] = []
    for line in lines:
        if deduped and deduped[-1] == line:
            continue
        deduped.append(line)
    return deduped
