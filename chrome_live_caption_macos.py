from __future__ import annotations

from capture_types import TargetWindow
from chrome_live_caption_common import split_live_caption_paragraphs
from macos_accessibility import (
    AccessibilityTextRecord,
    create_application,
    dependency_error as accessibility_dependency_error,
    describe_window,
    find_ax_window,
    get_window_list,
    kCGWindowOwnerPID,
    kCGWindowOwnerName,
    window_area,
    extract_text_records,
    normalized_process_name,
    normalize_line,
)


class MacChromeLiveCaptionSession:
    def __init__(self, target: TargetWindow):
        self.target = target
        self.window_info: dict | None = None
        self.app_element = None
        self.window_element = None

    def refresh(self) -> str | None:
        window_info = find_live_caption_window_info(self.target)
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


def dependency_error(prompt: bool = False) -> str | None:
    return accessibility_dependency_error(prompt=prompt)


def describe_target(target: TargetWindow) -> str:
    window_info = find_live_caption_window_info(target)
    if window_info is None:
        return (
            f"No matching target window found.\n"
            f"Process: {target.process_name or '(any)'}\n"
            f"Window: {target.window_name or '(any)'}"
        )
    return describe_window(window_info)


def capture_text(target: TargetWindow) -> str:
    return MacChromeLiveCaptionSession(target).extract_text()


def find_live_caption_window_info(target: TargetWindow) -> dict | None:
    wanted_process = normalized_process_name(target.process_name)
    candidates: list[tuple[int, int, dict]] = []
    for window_info in get_window_list():
        owner_name = str(window_info.get(kCGWindowOwnerName, "") or "").strip()
        if wanted_process and normalized_process_name(owner_name) != wanted_process:
            continue
        area = window_area(window_info)
        if area <= 0:
            continue
        # Chrome Live Caption on macOS is exposed as a small untitled Chrome
        # floating window. Prefer smaller Chrome windows over full browser tabs.
        candidates.append((0 if area < 200_000 else 1, area, window_info))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def live_caption_text_from_records(records: list[AccessibilityTextRecord]) -> str:
    lines = visual_caption_lines_from_records(records)
    if not lines:
        lines = text_lines_from_records(records)
    paragraphs = split_live_caption_paragraphs(" ".join(lines))
    return "\n\n".join(paragraphs)


def visual_caption_lines_from_records(records: list[AccessibilityTextRecord]) -> list[str]:
    rows: dict[int, str] = {}
    for record in records:
        text = normalize_line(record.text)
        if not text:
            continue
        if record.role not in {"AXStaticText", "AXTextArea"}:
            continue
        if record.y is None:
            continue
        row_key = round(record.y)
        existing = rows.get(row_key, "")
        # Keep the most complete text exposed for a visual row. If equal length,
        # prefer the later record because Chrome often appends the freshest word
        # to a duplicate row near the end of the AX tree.
        if len(text) >= len(existing):
            rows[row_key] = text
    return [rows[key] for key in sorted(rows)]


def text_lines_from_records(records: list[AccessibilityTextRecord]) -> list[str]:
    lines: list[str] = []
    for record in records:
        text = normalize_line(record.text)
        if not text:
            continue
        if record.role not in {"AXStaticText", "AXTextArea"}:
            continue
        lines.append(text)
    return dedupe_lines(lines)


def dedupe_lines(lines: list[str]) -> list[str]:
    deduped: list[str] = []
    for line in lines:
        if deduped and deduped[-1] == line:
            continue
        deduped.append(line)
    return deduped
