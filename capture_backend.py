import re
import threading
from dataclasses import dataclass

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


DEFAULT_CLASS = "Chrome_RenderWidgetHostHWND"
DEFAULT_PROCESS_NAME = "Caption.Ed.exe"
DEFAULT_WINDOW_NAME = "Caption.Ed"
PARAGRAPH_ID_PREFIX = "paragraph-"
TIMESTAMP_PREFIX_PATTERN = re.compile(
    r"^\s*(?:(?:\d+\s+(?:hours?|minutes?|seconds?))+\s*)+",
    re.IGNORECASE,
)
SPEAKER_BLOCK_PATTERN = re.compile(
    r"Speaker\s+\d+\s+(?:(?:\d+\s+(?:hours?|minutes?|seconds?))\s+)+",
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

        if not self.transcript_control.Exists(0, 0):
            self.root_control = auto.ControlFromHandle(self.root_hwnd)
            self.document_control = self.root_control.DocumentControl()
            self.transcript_control = find_transcript_control(self.root_control)
            if not self.transcript_control.Exists(0, 0):
                return "Transcript control was not found through UI Automation."

        raw_text = self._read_transcript_value()
        paragraphs = parse_transcript_value(raw_text)
        if not paragraphs:
            return "(No text detected yet.)"
        return "\n\n".join(paragraphs)

    def _read_transcript_value(self) -> str:
        try:
            if self.document_control is not None and self.document_control.Exists(0, 0):
                text_pattern = self.document_control.GetTextPattern()
                if text_pattern and text_pattern.DocumentRange:
                    document_text = text_pattern.DocumentRange.GetText(-1)
                    transcript_text = extract_transcript_from_document_text(document_text)
                    if transcript_text:
                        return transcript_text
        except Exception:
            pass

        try:
            value_pattern = self.transcript_control.GetValuePattern()
            if value_pattern and value_pattern.Value:
                return value_pattern.Value
        except Exception:
            pass

        try:
            text_pattern = self.transcript_control.GetTextPattern()
            if text_pattern and text_pattern.DocumentRange:
                return text_pattern.DocumentRange.GetText(-1)
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
        return "\n\n".join(paragraphs)


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


def find_target_hwnd(process_name: str = DEFAULT_PROCESS_NAME, class_name: str = DEFAULT_CLASS) -> int | None:
    if win32gui is None:
        return None

    matches: list[tuple[int, int, int]] = []
    wanted_process = process_name.casefold()

    def consider(hwnd: int) -> None:
        if not win32gui.IsWindow(hwnd):
            return
        try:
            if win32gui.GetClassName(hwnd) != class_name:
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
            if wanted_process and get_process_name(hwnd).casefold() != wanted_process:
                return True
            consider(hwnd)
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
    return find_target_hwnd(target.process_name, target.expected_class)


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


def describe_target(target: TargetWindow) -> str:
    if win32gui is None:
        return "pywin32 is not installed"
    hwnd = resolve_hwnd(target)
    if hwnd is None:
        return f"No matching target window found.\nProcess: {target.process_name}\nClass: {target.expected_class}"
    return describe_window(hwnd)


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


def ensure_com_initialized() -> None:
    if getattr(THREAD_STATE, "com_initialized", False):
        return
    comtypes.CoInitialize()
    THREAD_STATE.com_initialized = True


def release_com_if_initialized() -> None:
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


def extract_transcript_from_document_text(raw_text: str) -> str:
    if not raw_text:
        return ""

    matches = list(SPEAKER_BLOCK_PATTERN.finditer(raw_text))
    if not matches:
        return ""

    paragraphs: list[str] = []
    for index, match in enumerate(matches):
        content_start = match.end()
        content_end = matches[index + 1].start() if index + 1 < len(matches) else len(raw_text)
        paragraph = normalize_transcript_paragraph(raw_text[content_start:content_end])
        if paragraph:
            paragraphs.append(paragraph)
    return "\n\n".join(paragraphs)


def extract_transcript_text(target: TargetWindow) -> str:
    error = dependency_error()
    if error:
        return error
    ensure_com_initialized()
    session = getattr(THREAD_STATE, "transcript_session", None)
    if session is None or session.target != target:
        session = TranscriptAutomationSession(target)
        THREAD_STATE.transcript_session = session
    return session.extract_text()


def capture_window_text(target: TargetWindow, session=None) -> str:
    return extract_transcript_text(target)
