from __future__ import annotations

import time

from capture_types import TargetWindow
from chrome_live_caption_common import split_live_caption_paragraphs
from macos_accessibility import (
    AccessibilityTextRecord,
    accessibility_children,
    ax_copy_attribute,
    create_application,
    dependency_error as accessibility_dependency_error,
    describe_window,
    find_ax_window,
    get_window_list,
    kAXDescriptionAttribute,
    kCGWindowOwnerPID,
    kCGWindowOwnerName,
    kAXRoleAttribute,
    kAXTitleAttribute,
    kAXValueAttribute,
    kAXWindowRole,
    MAX_ACCESSIBILITY_DEPTH,
    MAX_ACCESSIBILITY_NODES,
    node_rect,
    window_area,
    normalized_process_name,
    normalize_line,
)

TEXT_ROLES = {"AXStaticText", "AXTextArea"}
WINDOW_REFRESH_INTERVAL_SECONDS = 1.0
TEXT_NODE_RESCAN_INTERVAL_SECONDS = 0.75


class MacChromeLiveCaptionSession:
    def __init__(self, target: TargetWindow):
        self.target = target
        self.window_info: dict | None = None
        self.app_element = None
        self.window_element = None
        self.text_nodes: list = []
        self.last_window_refresh_at = 0.0
        self.last_text_scan_at = 0.0

    def refresh(self, force: bool = False) -> str | None:
        now = time.monotonic()
        if (
            not force
            and self.window_element is not None
            and (now - self.last_window_refresh_at) < WINDOW_REFRESH_INTERVAL_SECONDS
        ):
            return None

        window_info = find_live_caption_window_info(self.target)
        self.last_window_refresh_at = now
        if window_info is None:
            self.window_info = None
            self.window_element = None
            self.text_nodes = []
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
        self.text_nodes = []
        self.last_text_scan_at = 0.0
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

        now = time.monotonic()
        records = text_records_from_nodes(self.text_nodes)
        should_rescan = not records or (now - self.last_text_scan_at) >= TEXT_NODE_RESCAN_INTERVAL_SECONDS
        if should_rescan:
            records, self.text_nodes = scan_live_caption_records(self.window_element)
            self.last_text_scan_at = now
        if not records:
            refresh_error = self.refresh(force=True)
            if refresh_error:
                return refresh_error
            records, self.text_nodes = scan_live_caption_records(self.window_element)
            self.last_text_scan_at = time.monotonic()

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


def scan_live_caption_records(window_element) -> tuple[list[AccessibilityTextRecord], list]:
    records: list[AccessibilityTextRecord] = []
    text_nodes: list = []
    visited: set[str] = set()
    node_count = 0

    def walk(node, depth: int = 0) -> None:
        nonlocal node_count
        if node_count >= MAX_ACCESSIBILITY_NODES or depth > MAX_ACCESSIBILITY_DEPTH:
            return
        node_key = repr(node)
        if node_key in visited:
            return
        visited.add(node_key)
        node_count += 1

        role = str(ax_copy_attribute(node, kAXRoleAttribute) or "")
        if role in TEXT_ROLES:
            record = text_record_from_node(node, role)
            if record is not None:
                records.append(record)
                text_nodes.append(node)

        if role in TEXT_ROLES and depth >= 2:
            return
        if role == (kAXWindowRole or "AXWindow") and depth >= 1:
            return

        for child in accessibility_children(node):
            walk(child, depth + 1)

    walk(window_element)
    return records, text_nodes


def text_records_from_nodes(nodes: list) -> list[AccessibilityTextRecord]:
    records: list[AccessibilityTextRecord] = []
    alive_nodes: list = []
    for node in nodes:
        role = str(ax_copy_attribute(node, kAXRoleAttribute) or "")
        if role not in TEXT_ROLES:
            continue
        record = text_record_from_node(node, role)
        if record is None:
            continue
        records.append(record)
        alive_nodes.append(node)
    if len(alive_nodes) != len(nodes):
        nodes[:] = alive_nodes
    return records


def text_record_from_node(node, role: str) -> AccessibilityTextRecord | None:
    for attribute in (kAXValueAttribute, kAXTitleAttribute, kAXDescriptionAttribute):
        value = ax_copy_attribute(node, attribute)
        if not isinstance(value, str):
            continue
        text = normalize_line(value)
        if not text:
            continue
        x, y, width, height = node_rect(node)
        return AccessibilityTextRecord(text, role, str(attribute), x, y, width, height)
    return None
