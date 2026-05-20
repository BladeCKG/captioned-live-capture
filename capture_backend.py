from __future__ import annotations

import platform
import re
import threading
from collections.abc import Iterable
from dataclasses import dataclass
from chrome_live_caption import (
    LIVE_CAPTION_CLASS,
    LIVE_CAPTION_PROCESS_NAME,
    LIVE_CAPTION_WINDOW_NAME,
    extract_live_caption_paragraphs,
    is_live_caption_target,
    split_live_caption_paragraphs,
)

try:
    import uiautomation as auto
except ImportError:
    auto = None

try:
    import comtypes
except ImportError:
    comtypes = None

try:
    import win32con
    import win32gui
    import win32process
except ImportError:
    win32con = None
    win32gui = None
    win32process = None

try:
    from ApplicationServices import (
        AXIsProcessTrusted,
        AXUIElementCopyAttributeNames,
        AXUIElementCopyElementAtPosition,
        AXUIElementCopyAttributeValue,
        AXUIElementCreateApplication,
        AXUIElementCreateSystemWide,
        kAXChildrenAttribute,
        kAXDescriptionAttribute,
        kAXParentAttribute,
        kAXRoleAttribute,
        kAXStaticTextRole,
        kAXTitleAttribute,
        kAXValueAttribute,
        kAXWindowRole,
        kAXWindowsAttribute,
    )
except ImportError:
    AXIsProcessTrusted = None
    AXUIElementCopyAttributeNames = None
    AXUIElementCopyElementAtPosition = None
    AXUIElementCopyAttributeValue = None
    AXUIElementCreateApplication = None
    AXUIElementCreateSystemWide = None
    kAXChildrenAttribute = None
    kAXDescriptionAttribute = None
    kAXParentAttribute = None
    kAXRoleAttribute = None
    kAXStaticTextRole = None
    kAXTitleAttribute = None
    kAXValueAttribute = None
    kAXWindowRole = None
    kAXWindowsAttribute = None

MACOS_CHILD_ATTRIBUTES = tuple(
    attribute
    for attribute in (
        kAXChildrenAttribute,
        "AXChildrenInNavigationOrder",
        "AXVisibleChildren",
        "AXRows",
        "AXContents",
    )
    if attribute
)

try:
    from Quartz import (
        kCGNullWindowID,
        kCGWindowBounds,
        kCGWindowListExcludeDesktopElements,
        kCGWindowListOptionAll,
        kCGWindowListOptionOnScreenOnly,
        kCGWindowName,
        kCGWindowNumber,
        kCGWindowOwnerName,
        kCGWindowOwnerPID,
        CGWindowListCopyWindowInfo,
    )
except ImportError:
    kCGNullWindowID = None
    kCGWindowBounds = None
    kCGWindowListExcludeDesktopElements = None
    kCGWindowListOptionAll = None
    kCGWindowListOptionOnScreenOnly = None
    kCGWindowName = None
    kCGWindowNumber = None
    kCGWindowOwnerName = None
    kCGWindowOwnerPID = None
    CGWindowListCopyWindowInfo = None


IS_WINDOWS = platform.system() == "Windows"
IS_MACOS = platform.system() == "Darwin"

DEFAULT_CLASS = "Chrome_RenderWidgetHostHWND" if IS_WINDOWS else ""
DEFAULT_PROCESS_NAME = "Caption.Ed.exe" if IS_WINDOWS else "Caption.Ed"
DEFAULT_WINDOW_NAME = "Caption.Ed" if IS_WINDOWS else ""
MACOS_MAX_ACCESSIBILITY_NODES = 1000
MACOS_MAX_ACCESSIBILITY_DEPTH = 24
PARAGRAPH_ID_PREFIX = "paragraph-"
TIMESTAMP_PREFIX_PATTERN = re.compile(
    r"^\s*(?:(?:\d+\s+(?:hours?|minutes?|seconds?))+\s*)+",
    re.IGNORECASE,
)
SPEAKER_BLOCK_PATTERN = re.compile(
    r"(Speaker\s+\d+)\s+((?:\d+\s+(?:hours?|minutes?|seconds?))(?:\s+\d+\s+(?:hours?|minutes?|seconds?))*)\s+",
    re.IGNORECASE,
)
SPEAKER_LINE_PATTERN = re.compile(r"^speaker\s+\d+\s*$", re.IGNORECASE)
TIMESTAMP_LINE_PATTERN = re.compile(
    r"^(?:\d+\s+(?:hours?|minutes?|seconds?))(?:\s+\d+\s+(?:hours?|minutes?|seconds?))*\s*$",
    re.IGNORECASE,
)
THREAD_STATE = threading.local()


@dataclass(frozen=True)
class TargetWindow:
    hwnd: int | None = None
    expected_class: str = DEFAULT_CLASS
    process_name: str = DEFAULT_PROCESS_NAME
    window_name: str = DEFAULT_WINDOW_NAME


class TranscriptAutomationSession:
    def __init__(self, target: TargetWindow):
        self.target = target
        self.target_hwnd: int | None = None
        self.root_hwnd: int | None = None
        self.root_control = None
        self.document_control = None
        self.transcript_control = None

    def refresh(self) -> str | None:
        hwnd = resolve_hwnd(self.target)
        if hwnd is None:
            return f"Window not found.\nProcess: {self.target.process_name}\nClass: {self.target.expected_class}"

        actual_class = win32gui.GetClassName(hwnd)
        if self.target.expected_class and actual_class != self.target.expected_class:
            return f"Window class mismatch.\nExpected: {self.target.expected_class}\nActual: {actual_class}"

        root_hwnd = get_root_hwnd(hwnd)
        if self.target_hwnd == hwnd and self.root_hwnd == root_hwnd and self.transcript_control is not None:
            return None

        self.target_hwnd = hwnd
        self.root_hwnd = root_hwnd
        self.root_control = auto.ControlFromHandle(root_hwnd)
        self.document_control = self.root_control.DocumentControl()
        self.transcript_control = find_transcript_control(self.root_control)
        return None

    def extract_text(self) -> str:
        refresh_error = self.refresh()
        if refresh_error:
            return refresh_error

        if is_live_caption_target(self.target):
            paragraphs = extract_live_caption_paragraphs(
                self.root_control,
                find_first_control,
                normalize_transcript_paragraph,
            )
            if not paragraphs:
                return "(No text detected yet.)"
            return "\n\n".join(paragraphs)

        if not self.transcript_control.Exists(0, 0):
            self.root_control = auto.ControlFromHandle(self.root_hwnd)
            self.document_control = self.root_control.DocumentControl()
            self.transcript_control = find_transcript_control(self.root_control)
            if not self.transcript_control.Exists(0, 0):
                return "Transcript control was not found through UI Automation."

        paragraphs = self._read_transcript_paragraphs()
        if not paragraphs:
            return "(No text detected yet.)"
        return "\n\n".join(paragraphs)

    def _read_transcript_paragraphs(self) -> list[str]:
        try:
            if self.document_control is not None and self.document_control.Exists(0, 0):
                text_pattern = self.document_control.GetTextPattern()
                if text_pattern and text_pattern.DocumentRange:
                    document_text = text_pattern.DocumentRange.GetText(-1)
                    paragraph_metadata = extract_transcript_paragraph_metadata(self.transcript_control)
                    paragraphs = extract_transcript_paragraphs_from_document_text(
                        document_text,
                        paragraph_metadata=paragraph_metadata,
                    )
                    if paragraphs:
                        return paragraphs
        except Exception:
            pass

        try:
            value_pattern = self.transcript_control.GetValuePattern()
            if value_pattern and value_pattern.Value:
                return parse_transcript_value(value_pattern.Value)
        except Exception:
            pass

        try:
            text_pattern = self.transcript_control.GetTextPattern()
            if text_pattern and text_pattern.DocumentRange:
                return parse_transcript_value(text_pattern.DocumentRange.GetText(-1))
        except Exception:
            pass

        paragraphs: list[str] = []
        for child in self.transcript_control.GetChildren():
            automation_id = child.AutomationId or ""
            if not automation_id.startswith(PARAGRAPH_ID_PREFIX):
                continue
            paragraph = normalize_transcript_paragraph(extract_text_controls(child))
            if paragraph:
                paragraphs.append(paragraph)
        return paragraphs


class MacAccessibilitySession:
    def __init__(self, target: TargetWindow):
        self.target = target
        self.window_info: dict | None = None
        self.app_element = None
        self.window_element = None

    def refresh(self) -> str | None:
        window_info = find_macos_window_info(self.target)
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
        self.app_element = AXUIElementCreateApplication(pid)
        self.window_element = find_macos_ax_window(self.app_element, window_info, self.target)
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

        lines = extract_macos_accessibility_lines(self.window_element)
        if not lines:
            return "(No text detected yet.)"

        if is_live_caption_target(self.target):
            paragraphs = split_live_caption_paragraphs(" ".join(lines))
        else:
            parsed = parse_transcript_value("\n".join(lines))
            paragraphs = parsed or split_generic_macos_paragraphs(lines)

        if not paragraphs:
            return "(No text detected yet.)"
        return "\n\n".join(paragraphs)

def get_root_hwnd(hwnd: int) -> int:
    if win32gui is None:
        return hwnd
    return win32gui.GetAncestor(hwnd, win32con.GA_ROOT)


def get_macos_window_list() -> list[dict]:
    if CGWindowListCopyWindowInfo is None:
        return []
    options = (kCGWindowListOptionOnScreenOnly or 0) | (kCGWindowListExcludeDesktopElements or 0)
    try:
        windows = CGWindowListCopyWindowInfo(options, kCGNullWindowID or 0) or []
    except Exception:
        windows = []
    return list(windows)


def macos_window_area(window_info: dict) -> int:
    bounds = window_info.get(kCGWindowBounds, {}) or {}
    width = int(bounds.get("Width", 0) or 0)
    height = int(bounds.get("Height", 0) or 0)
    return max(0, width) * max(0, height)


def find_macos_window_info(target: TargetWindow) -> dict | None:
    matches: list[tuple[int, int, dict]] = []
    wanted_process = normalized_macos_process_name(target.process_name)
    wanted_window = (target.window_name or "").strip()

    for window_info in get_macos_window_list():
        owner_name = str(window_info.get(kCGWindowOwnerName, "") or "").strip()
        window_name = str(window_info.get(kCGWindowName, "") or "").strip()
        if wanted_process and normalized_macos_process_name(owner_name) != wanted_process:
            continue
        if wanted_window and window_name != wanted_window:
            continue
        area = macos_window_area(window_info)
        if area <= 0:
            continue
        title_score = 1 if window_name else 0
        matches.append((title_score, area, window_info))

    if not matches:
        return None
    matches.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return matches[0][2]


def normalized_macos_process_name(name: str | None) -> str:
    text = (name or "").strip().casefold()
    return re.sub(r"[^a-z0-9]+", "", text)


def ax_copy_attribute(element, attribute: str):
    if AXUIElementCopyAttributeValue is None:
        return None
    try:
        _, value = AXUIElementCopyAttributeValue(element, attribute, None)
        return value
    except TypeError:
        try:
            result = AXUIElementCopyAttributeValue(element, attribute)
            if isinstance(result, tuple) and len(result) >= 2:
                return result[1]
            return result
        except Exception:
            return None
    except Exception:
        return None


def ax_copy_attribute_names(element) -> list[str]:
    if AXUIElementCopyAttributeNames is None:
        return []
    try:
        _, names = AXUIElementCopyAttributeNames(element, None)
        return list(names or [])
    except TypeError:
        try:
            result = AXUIElementCopyAttributeNames(element)
            if isinstance(result, tuple) and len(result) >= 2:
                return list(result[1] or [])
            return list(result or [])
        except Exception:
            return []
    except Exception:
        return []


def find_macos_ax_window(app_element, window_info: dict, target: TargetWindow):
    windows = ax_copy_attribute(app_element, kAXWindowsAttribute) or []
    wanted_title = str(window_info.get(kCGWindowName, "") or "").strip()
    fallback = None
    for window in windows:
        title = str(ax_copy_attribute(window, kAXTitleAttribute) or "").strip()
        if fallback is None:
            fallback = window
        if wanted_title and title == wanted_title:
            return window
    if fallback is not None:
        return fallback

    return find_macos_ax_window_by_position(window_info)


def find_macos_ax_window_by_position(window_info: dict):
    if AXUIElementCreateSystemWide is None or AXUIElementCopyElementAtPosition is None:
        return None
    bounds = window_info.get(kCGWindowBounds, {}) or {}
    width = float(bounds.get("Width", 0) or 0)
    height = float(bounds.get("Height", 0) or 0)
    if width <= 0 or height <= 0:
        return None
    center_x = float(bounds.get("X", 0) or 0) + (width / 2.0)
    center_y = float(bounds.get("Y", 0) or 0) + (height / 2.0)
    system_wide = AXUIElementCreateSystemWide()
    try:
        result = AXUIElementCopyElementAtPosition(system_wide, center_x, center_y, None)
    except TypeError:
        try:
            result = AXUIElementCopyElementAtPosition(system_wide, center_x, center_y)
        except Exception:
            return None
    except Exception:
        return None

    element = unpack_ax_result(result)
    if element is None:
        return None
    return ascend_to_ax_window(element)


def unpack_ax_result(result):
    if result is None:
        return None
    if isinstance(result, tuple):
        if len(result) >= 2:
            return result[1]
        if len(result) == 1:
            return result[0]
    return result


def ascend_to_ax_window(element):
    node = element
    visited_ids: set[int] = set()
    while node is not None:
        node_id = id(node)
        if node_id in visited_ids:
            break
        visited_ids.add(node_id)
        role = str(ax_copy_attribute(node, kAXRoleAttribute) or "")
        if role == (kAXWindowRole or "AXWindow"):
            return node
        node = ax_copy_attribute(node, kAXParentAttribute)
    return element


def extract_macos_accessibility_lines(window_element) -> list[str]:
    lines: list[str] = []
    visited: set[str] = set()
    node_count = 0

    def walk(node, depth: int = 0) -> None:
        nonlocal node_count
        if node_count >= MACOS_MAX_ACCESSIBILITY_NODES or depth > MACOS_MAX_ACCESSIBILITY_DEPTH:
            return
        node_key = repr(node)
        if node_key in visited:
            return
        visited.add(node_key)
        node_count += 1
        role = str(ax_copy_attribute(node, kAXRoleAttribute) or "")
        children = macos_accessibility_children(node)

        if role == kAXStaticTextRole or (not children and role != kAXWindowRole):
            for attribute in (kAXValueAttribute, kAXTitleAttribute, kAXDescriptionAttribute):
                value = ax_copy_attribute(node, attribute)
                if isinstance(value, str):
                    text = normalize_macos_line(value)
                    if text and (not lines or lines[-1] != text):
                        lines.append(text)
                    break

        for child in children:
            walk(child, depth + 1)

    walk(window_element)
    return lines


def macos_accessibility_children(node) -> list:
    children: list = []
    seen_ids: set[int] = set()
    for attribute in MACOS_CHILD_ATTRIBUTES:
        value = ax_copy_attribute(node, attribute)
        if isinstance(value, (str, bytes)) or not isinstance(value, Iterable):
            continue
        for child in value:
            child_id = id(child)
            if child_id in seen_ids:
                continue
            seen_ids.add(child_id)
            children.append(child)
    return children


def normalize_macos_line(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def split_generic_macos_paragraphs(lines: list[str]) -> list[str]:
    paragraphs: list[str] = []
    current: list[str] = []
    for line in lines:
        if not line:
            continue
        if current and line == current[-1]:
            continue
        current.append(line)
    if current:
        paragraphs.append(normalize_transcript_paragraph(" ".join(current)))
    return [paragraph for paragraph in paragraphs if paragraph]


def get_process_path(hwnd: int) -> str:
    if win32process is None:
        return ""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        import psutil

        return psutil.Process(pid).exe()
    except Exception:
        return ""


def get_process_name(hwnd: int) -> str:
    if win32process is None:
        return ""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        import psutil

        return psutil.Process(pid).name()
    except Exception:
        return ""


def find_target_hwnd(
    process_name: str = DEFAULT_PROCESS_NAME,
    class_name: str = DEFAULT_CLASS,
    window_name: str = DEFAULT_WINDOW_NAME,
) -> int | None:
    if win32gui is None:
        return None

    matches: list[tuple[int, int, int]] = []
    wanted_process = process_name.casefold()
    wanted_window = window_name.strip()

    def consider(hwnd: int) -> None:
        if not win32gui.IsWindow(hwnd):
            return
        try:
            if win32gui.GetClassName(hwnd) != class_name:
                return
            title = win32gui.GetWindowText(hwnd)
            root_title = ""
            if wanted_window:
                try:
                    root_title = win32gui.GetWindowText(get_root_hwnd(hwnd))
                except Exception:
                    root_title = ""
                if title != wanted_window and root_title != wanted_window:
                    return
            if wanted_process and get_process_name(hwnd).casefold() != wanted_process:
                return
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            area = max(0, right - left) * max(0, bottom - top)
            visible_score = 1 if win32gui.IsWindowVisible(hwnd) else 0
            matches.append((visible_score, area, hwnd))
        except Exception:
            return

    def visit_child(hwnd: int, _: object) -> bool:
        consider(hwnd)
        return True

    def visit_top_level(hwnd: int, _: object) -> bool:
        try:
            consider(hwnd)
            if wanted_process and get_process_name(hwnd).casefold() != wanted_process:
                return True
            win32gui.EnumChildWindows(hwnd, visit_child, None)
        except Exception:
            return True
        return True

    win32gui.EnumWindows(visit_top_level, None)
    if not matches:
        return None
    matches.sort(reverse=True)
    return matches[0][2]


def resolve_hwnd(target: TargetWindow) -> int | None:
    if target.hwnd and win32gui is not None and win32gui.IsWindow(target.hwnd):
        return target.hwnd
    return find_target_hwnd(target.process_name, target.expected_class, target.window_name)


def describe_window(hwnd: int) -> str:
    if win32gui is None or win32process is None:
        return "pywin32 is not installed"
    if not win32gui.IsWindow(hwnd):
        return f"HWND {hwnd} was not found"

    root_hwnd = get_root_hwnd(hwnd)
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    class_name = win32gui.GetClassName(hwnd)
    title = win32gui.GetWindowText(hwnd)
    root_title = win32gui.GetWindowText(root_hwnd)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return (
        f"HWND: {hwnd}\n"
        f"Root HWND: {root_hwnd}\n"
        f"PID: {pid}\n"
        f"Class: {class_name}\n"
        f"Title: {title or '(no title)'}\n"
        f"Root title: {root_title or '(no title)'}\n"
        f"Process: {get_process_path(hwnd) or '(unknown)'}\n"
        f"Bounds: {left}, {top}, {right}, {bottom}"
    )


def describe_macos_window(window_info: dict) -> str:
    owner_name = str(window_info.get(kCGWindowOwnerName, "") or "(unknown)")
    window_name = str(window_info.get(kCGWindowName, "") or "(no title)")
    pid = int(window_info.get(kCGWindowOwnerPID, 0) or 0)
    bounds = window_info.get(kCGWindowBounds, {}) or {}
    return (
        f"Window ID: {window_info.get(kCGWindowNumber, '(unknown)')}\n"
        f"PID: {pid}\n"
        f"Owner: {owner_name}\n"
        f"Title: {window_name}\n"
        f"Bounds: {bounds.get('X', 0)}, {bounds.get('Y', 0)}, "
        f"{bounds.get('Width', 0)}, {bounds.get('Height', 0)}"
    )


def describe_target(target: TargetWindow) -> str:
    if IS_MACOS:
        window_info = find_macos_window_info(target)
        if window_info is None:
            return (
                f"No matching target window found.\n"
                f"Process: {target.process_name or '(any)'}\n"
                f"Window: {target.window_name or '(any)'}"
            )
        return describe_macos_window(window_info)
    if win32gui is None:
        return "pywin32 is not installed"
    hwnd = resolve_hwnd(target)
    if hwnd is None:
        return f"No matching target window found.\nProcess: {target.process_name}\nClass: {target.expected_class}"
    return describe_window(hwnd)


def dependency_error() -> str | None:
    missing = []
    if IS_MACOS:
        if CGWindowListCopyWindowInfo is None or AXUIElementCreateApplication is None:
            missing.append("pyobjc-framework-ApplicationServices")
        elif AXIsProcessTrusted is not None and not AXIsProcessTrusted():
            return (
                "Accessibility permission is required on macOS. "
                "Allow this app or Terminal in System Settings > Privacy & Security > Accessibility."
            )
    else:
        if win32gui is None:
            missing.append("pywin32")
        if auto is None:
            missing.append("uiautomation")
        if comtypes is None:
            missing.append("comtypes")
    if missing:
        return "Missing Python packages: " + ", ".join(missing)
    return None


def ensure_com_initialized() -> None:
    if not IS_WINDOWS:
        return
    if getattr(THREAD_STATE, "com_initialized", False):
        return
    comtypes.CoInitialize()
    THREAD_STATE.com_initialized = True


def release_com_if_initialized() -> None:
    if not IS_WINDOWS:
        return
    if not getattr(THREAD_STATE, "com_initialized", False):
        return
    try:
        comtypes.CoUninitialize()
    finally:
        THREAD_STATE.com_initialized = False


def extract_text_controls(control) -> str:
    parts: list[str] = []

    def walk(node) -> None:
        try:
            if node.ControlTypeName == "TextControl" and node.Name:
                parts.append(node.Name)
        except Exception:
            pass
        try:
            for child in node.GetChildren():
                walk(child)
        except Exception:
            return

    walk(control)
    return "".join(parts)


def iter_controls_depth_first(control):
    yield control
    try:
        for child in control.GetChildren():
            yield from iter_controls_depth_first(child)
    except Exception:
        return


def find_first_control(control, predicate):
    for node in iter_controls_depth_first(control):
        try:
            if predicate(node):
                return node
        except Exception:
            continue
    return None


def find_transcript_control(root_control):
    direct_group = root_control.GroupControl(Name="Transcript")
    if direct_group.Exists(0, 0):
        return direct_group

    direct_control = root_control.Control(Name="Transcript")
    if direct_control.Exists(0, 0):
        return direct_control

    transcription_container = root_control.Control(AutomationId="transcription-container")
    if transcription_container.Exists(0, 0):
        nested_group = transcription_container.GroupControl(Name="Transcript")
        if nested_group.Exists(0, 0):
            return nested_group
        return transcription_container

    for control in iter_controls_depth_first(root_control):
        try:
            automation_id = (control.AutomationId or "").strip()
        except Exception:
            continue
        if automation_id.startswith(PARAGRAPH_ID_PREFIX):
            parent = getattr(control, "GetParentControl", lambda: None)()
            if parent is not None:
                return parent

    return root_control.Control(Name="Transcript")


def normalize_transcript_paragraph(raw_paragraph: str) -> str:
    text = TIMESTAMP_PREFIX_PATTERN.sub("", raw_paragraph).strip()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_transcript_value(raw_text: str) -> list[str]:
    if not raw_text:
        return []

    lines = [line.strip() for line in raw_text.splitlines()]
    paragraphs: list[str] = []
    current_parts: list[str] = []
    skip_timestamp = False

    def flush_current() -> None:
        if not current_parts:
            return
        paragraph = normalize_transcript_paragraph(" ".join(current_parts))
        if paragraph:
            paragraphs.append(paragraph)
        current_parts.clear()

    for line in lines:
        if not line:
            continue
        if SPEAKER_LINE_PATTERN.match(line):
            flush_current()
            skip_timestamp = True
            continue
        if skip_timestamp and TIMESTAMP_LINE_PATTERN.match(line):
            skip_timestamp = False
            continue
        skip_timestamp = False
        current_parts.append(line)

    flush_current()
    return paragraphs


def extract_transcript_paragraph_metadata(transcript_control) -> list[tuple[str, str]]:
    metadata: list[tuple[str, str]] = []
    try:
        children = transcript_control.GetChildren()
    except Exception:
        return metadata

    for child in children:
        automation_id = (child.AutomationId or "").strip()
        if not automation_id.startswith(PARAGRAPH_ID_PREFIX):
            continue
        label = ""
        paragraph_text = ""
        try:
            outer_children = child.GetChildren()
            if outer_children:
                inner_children = outer_children[0].GetChildren()
                if inner_children and inner_children[0].ControlTypeName == "ButtonControl":
                    candidate = (inner_children[0].Name or "").strip()
                    if candidate.startswith("Speaker "):
                        label = candidate
                paragraph_text = normalize_transcript_paragraph(extract_text_controls(child))
        except Exception:
            label = ""
            paragraph_text = ""
        metadata.append((label, paragraph_text))
    return metadata


def extract_transcript_paragraphs_from_document_text(
    raw_text: str,
    paragraph_metadata: list[tuple[str, str]] | None = None,
) -> list[str]:
    if not raw_text:
        return []

    matches = list(SPEAKER_BLOCK_PATTERN.finditer(raw_text))
    if not matches:
        return []

    paragraphs: list[str] = []
    for index, match in enumerate(matches):
        content_start = match.end()
        content_end = matches[index + 1].start() if index + 1 < len(matches) else len(raw_text)
        paragraph = normalize_transcript_paragraph(raw_text[content_start:content_end])
        if paragraph:
            speaker = ""
            structural_text = ""
            if paragraph_metadata and index < len(paragraph_metadata):
                speaker, structural_text = paragraph_metadata[index]
                speaker = speaker.strip()
            if not speaker:
                speaker = match.group(1).strip()
            if structural_text:
                paragraph = choose_structural_paragraph_text(paragraph, structural_text)
            paragraphs.append(f"{speaker}\n{paragraph}")
    return paragraphs


def choose_structural_paragraph_text(document_text: str, structural_text: str) -> str:
    document_norm = normalized_compare_text(document_text)
    structural_norm = normalized_compare_text(structural_text)
    if not structural_norm:
        return document_text
    if document_norm.startswith(structural_norm):
        return structural_text
    if structural_norm in document_norm:
        return structural_text
    return document_text


def normalized_compare_text(text: str) -> str:
    return " ".join(text.casefold().split())


def extract_transcript_text(target: TargetWindow) -> str:
    error = dependency_error()
    if error:
        return error
    if IS_WINDOWS:
        ensure_com_initialized()
        session_type = TranscriptAutomationSession
    elif IS_MACOS:
        session_type = MacAccessibilitySession
    else:
        return f"Unsupported platform: {platform.system()}"
    session = getattr(THREAD_STATE, "transcript_session", None)
    if session is None or session.target != target or not isinstance(session, session_type):
        session = session_type(target)
        THREAD_STATE.transcript_session = session
    return session.extract_text()


def capture_window_text(target: TargetWindow, session=None) -> str:
    return extract_transcript_text(target)
