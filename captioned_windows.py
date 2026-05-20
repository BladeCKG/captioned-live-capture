from __future__ import annotations

from captioned_text import (
    PARAGRAPH_ID_PREFIX,
    extract_transcript_paragraphs_from_document_text,
    normalize_transcript_paragraph,
    parse_transcript_value,
)
from capture_types import TargetWindow
from chrome_live_caption_common import is_live_caption_target
from chrome_live_caption_windows import extract_live_caption_paragraphs

try:
    import comtypes
except ImportError:
    comtypes = None

try:
    import uiautomation as auto
except ImportError:
    auto = None

try:
    import win32con
    import win32gui
    import win32process
except ImportError:
    win32con = None
    win32gui = None
    win32process = None


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
            paragraphs = extract_live_caption_paragraphs(self.root_control, find_first_control)
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


def dependency_error() -> str | None:
    missing = []
    if win32gui is None:
        missing.append("pywin32")
    if auto is None:
        missing.append("uiautomation")
    if comtypes is None:
        missing.append("comtypes")
    if missing:
        return "Missing Python packages: " + ", ".join(missing)
    return None


def ensure_com_initialized(thread_state) -> None:
    if getattr(thread_state, "com_initialized", False):
        return
    comtypes.CoInitialize()
    thread_state.com_initialized = True


def release_com_if_initialized(thread_state) -> None:
    if not getattr(thread_state, "com_initialized", False):
        return
    try:
        comtypes.CoUninitialize()
    finally:
        thread_state.com_initialized = False


def get_root_hwnd(hwnd: int) -> int:
    if win32gui is None:
        return hwnd
    return win32gui.GetAncestor(hwnd, win32con.GA_ROOT)


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
    process_name: str,
    class_name: str,
    window_name: str,
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


def describe_target(target: TargetWindow) -> str:
    if win32gui is None or win32process is None:
        return "pywin32 is not installed"
    hwnd = resolve_hwnd(target)
    if hwnd is None:
        return f"No matching target window found.\nProcess: {target.process_name}\nClass: {target.expected_class}"
    return describe_window(hwnd)


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


def capture_text(target: TargetWindow) -> str:
    session = TranscriptAutomationSession(target)
    return session.extract_text()
