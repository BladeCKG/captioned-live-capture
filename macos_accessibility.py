from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass

from capture_types import TargetWindow

try:
    from ApplicationServices import (
        AXIsProcessTrusted,
        AXIsProcessTrustedWithOptions,
        AXUIElementCopyAttributeNames,
        AXUIElementCopyElementAtPosition,
        AXUIElementCopyAttributeValue,
        AXUIElementCreateApplication,
        AXUIElementCreateSystemWide,
        AXValueGetType,
        AXValueGetValue,
        kAXChildrenAttribute,
        kAXDescriptionAttribute,
        kAXParentAttribute,
        kAXRoleAttribute,
        kAXStaticTextRole,
        kAXTitleAttribute,
        kAXTrustedCheckOptionPrompt,
        kAXValueAttribute,
        kAXValueCGPointType,
        kAXValueCGSizeType,
        kAXWindowRole,
        kAXWindowsAttribute,
    )
except ImportError:
    AXIsProcessTrusted = None
    AXIsProcessTrustedWithOptions = None
    AXUIElementCopyAttributeNames = None
    AXUIElementCopyElementAtPosition = None
    AXUIElementCopyAttributeValue = None
    AXUIElementCreateApplication = None
    AXUIElementCreateSystemWide = None
    AXValueGetType = None
    AXValueGetValue = None
    kAXChildrenAttribute = None
    kAXDescriptionAttribute = None
    kAXParentAttribute = None
    kAXRoleAttribute = None
    kAXStaticTextRole = None
    kAXTitleAttribute = None
    kAXTrustedCheckOptionPrompt = None
    kAXValueAttribute = None
    kAXValueCGPointType = None
    kAXValueCGSizeType = None
    kAXWindowRole = None
    kAXWindowsAttribute = None

try:
    from Quartz import (
        kCGNullWindowID,
        kCGWindowBounds,
        kCGWindowListExcludeDesktopElements,
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
    kCGWindowListOptionOnScreenOnly = None
    kCGWindowName = None
    kCGWindowNumber = None
    kCGWindowOwnerName = None
    kCGWindowOwnerPID = None
    CGWindowListCopyWindowInfo = None


MAX_ACCESSIBILITY_NODES = 1000
MAX_ACCESSIBILITY_DEPTH = 24
ACCESSIBILITY_PERMISSION_MESSAGE = (
    "Accessibility permission is required on macOS. "
    "Allow this app or Terminal in System Settings > Privacy & Security > Accessibility, "
    "then click Capture Once again."
)
CHILD_ATTRIBUTES = tuple(
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


@dataclass(frozen=True)
class AccessibilityTextRecord:
    text: str
    role: str
    attribute: str
    x: float | None = None
    y: float | None = None
    width: float | None = None
    height: float | None = None


def dependency_error(prompt: bool = False) -> str | None:
    if CGWindowListCopyWindowInfo is None or AXUIElementCreateApplication is None:
        return "Missing Python packages: pyobjc-framework-ApplicationServices"
    if AXIsProcessTrusted is not None and not has_accessibility_permission():
        if prompt and request_accessibility_permission():
            return None
        return ACCESSIBILITY_PERMISSION_MESSAGE
    return None


def has_accessibility_permission() -> bool:
    if AXIsProcessTrusted is None:
        return False
    try:
        return bool(AXIsProcessTrusted())
    except Exception:
        return False


def request_accessibility_permission() -> bool:
    if has_accessibility_permission():
        return True

    if AXIsProcessTrustedWithOptions is None:
        return False

    prompt_key = kAXTrustedCheckOptionPrompt or "AXTrustedCheckOptionPrompt"
    try:
        return bool(AXIsProcessTrustedWithOptions({prompt_key: True}))
    except Exception:
        return False


def get_window_list() -> list[dict]:
    if CGWindowListCopyWindowInfo is None:
        return []
    options = (kCGWindowListOptionOnScreenOnly or 0) | (kCGWindowListExcludeDesktopElements or 0)
    try:
        windows = CGWindowListCopyWindowInfo(options, kCGNullWindowID or 0) or []
    except Exception:
        windows = []
    return list(windows)


def window_area(window_info: dict) -> int:
    bounds = window_info.get(kCGWindowBounds, {}) or {}
    width = int(bounds.get("Width", 0) or 0)
    height = int(bounds.get("Height", 0) or 0)
    return max(0, width) * max(0, height)


def find_window_info(target: TargetWindow) -> dict | None:
    matches: list[tuple[int, int, dict]] = []
    wanted_process = normalized_process_name(target.process_name)
    wanted_window = (target.window_name or "").strip()

    for window_info in get_window_list():
        owner_name = str(window_info.get(kCGWindowOwnerName, "") or "").strip()
        window_name = str(window_info.get(kCGWindowName, "") or "").strip()
        if wanted_process and normalized_process_name(owner_name) != wanted_process:
            continue
        if wanted_window and window_name != wanted_window:
            continue
        area = window_area(window_info)
        if area <= 0:
            continue
        title_score = 1 if window_name else 0
        matches.append((title_score, area, window_info))

    if not matches:
        return None
    matches.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return matches[0][2]


def normalized_process_name(name: str | None) -> str:
    text = (name or "").strip().casefold()
    return re.sub(r"[^a-z0-9]+", "", text)


def create_application(pid: int):
    return AXUIElementCreateApplication(pid)


def find_ax_window(app_element, window_info: dict):
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
    return find_ax_window_by_position(window_info)


def find_ax_window_by_position(window_info: dict):
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


def extract_text_records(window_element) -> list[AccessibilityTextRecord]:
    records: list[AccessibilityTextRecord] = []
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
        children = accessibility_children(node)

        if role == kAXStaticTextRole or (not children and role != kAXWindowRole):
            for attribute in (kAXValueAttribute, kAXTitleAttribute, kAXDescriptionAttribute):
                value = ax_copy_attribute(node, attribute)
                if isinstance(value, str):
                    text = normalize_line(value)
                    if text:
                        x, y, width, height = node_rect(node)
                        records.append(AccessibilityTextRecord(text, role, str(attribute), x, y, width, height))
                    break

        for child in children:
            walk(child, depth + 1)

    walk(window_element)
    return records


def accessibility_children(node) -> list:
    children: list = []
    seen_ids: set[int] = set()
    for attribute in CHILD_ATTRIBUTES:
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


def node_rect(node) -> tuple[float | None, float | None, float | None, float | None]:
    x, y = ax_point(ax_copy_attribute(node, "AXPosition"))
    width, height = ax_size(ax_copy_attribute(node, "AXSize"))
    return x, y, width, height


def ax_point(value) -> tuple[float | None, float | None]:
    point = ax_value(value, kAXValueCGPointType)
    if point is None:
        return None, None
    return float(getattr(point, "x", 0.0)), float(getattr(point, "y", 0.0))


def ax_size(value) -> tuple[float | None, float | None]:
    size = ax_value(value, kAXValueCGSizeType)
    if size is None:
        return None, None
    return float(getattr(size, "width", 0.0)), float(getattr(size, "height", 0.0))


def ax_value(value, expected_type):
    if AXValueGetValue is None or AXValueGetType is None or value is None or expected_type is None:
        return None
    try:
        value_type = AXValueGetType(value)
        if value_type != expected_type:
            return None
        success, unpacked = AXValueGetValue(value, value_type, None)
        if success:
            return unpacked
    except Exception:
        return None
    return None


def normalize_line(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def describe_window(window_info: dict) -> str:
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
