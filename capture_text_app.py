import argparse
import os
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import font as tkfont
from tkinter import messagebox, ttk

try:
    from PIL import Image, ImageGrab, ImageOps
except ImportError:  # pragma: no cover - shown in the UI at runtime
    Image = None
    ImageGrab = None
    ImageOps = None

try:
    import pytesseract
except ImportError:  # pragma: no cover - shown in the UI at runtime
    pytesseract = None

try:
    from windows_capture import Frame, InternalCaptureControl, WindowsCapture
except ImportError:  # pragma: no cover - normal screenshot fallback still works
    Frame = None
    InternalCaptureControl = None
    WindowsCapture = None

try:
    import win32con
    import win32gui
    import win32process
except ImportError:  # pragma: no cover - shown in the UI at runtime
    win32con = None
    win32gui = None
    win32process = None


DEFAULT_CLASS = "Chrome_RenderWidgetHostHWND"
DEFAULT_PROCESS_NAME = "Caption.Ed.exe"
DEFAULT_WINDOW_NAME = "Caption.Ed"
DEFAULT_INTERVAL_SECONDS = 0.5
COMMON_TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


if pytesseract is not None and os.path.exists(COMMON_TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = COMMON_TESSERACT_PATH


@dataclass(frozen=True)
class TargetWindow:
    hwnd: int | None = None
    expected_class: str = DEFAULT_CLASS
    process_name: str = DEFAULT_PROCESS_NAME
    window_name: str = DEFAULT_WINDOW_NAME


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
        return (
            "No matching target window found.\n"
            f"Process: {target.process_name}\n"
            f"Class: {target.expected_class}"
        )
    return describe_window(hwnd)


def dependency_error() -> str | None:
    missing = []
    if win32gui is None:
        missing.append("pywin32")
    if ImageGrab is None:
        missing.append("Pillow")
    if pytesseract is None:
        missing.append("pytesseract")
    if missing:
        return "Missing Python packages: " + ", ".join(missing)
    return None


def capture_visible_window(hwnd: int):
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    if right <= left or bottom <= top:
        return None
    return ImageGrab.grab(bbox=(left, top, right, bottom))


def capture_with_windows_graphics(target: TargetWindow):
    if WindowsCapture is None or ImageGrab is None:
        return None

    window_name = target.window_name.strip()
    if not window_name:
        return None

    done = threading.Event()
    result = {"image": None, "error": None}
    capture_control = None

    try:
        capture = WindowsCapture(cursor_capture=False, draw_border=None, window_name=window_name)

        @capture.event
        def on_frame_arrived(frame: Frame, capture_control: InternalCaptureControl) -> None:
            try:
                rgb = frame.frame_buffer[:, :, :3][:, :, ::-1].copy()
                result["image"] = Image.fromarray(rgb, "RGB")
            except Exception as exc:
                result["error"] = exc
            finally:
                done.set()
                capture_control.stop()

        @capture.event
        def on_closed() -> None:
            done.set()

        capture_control = capture.start_free_threaded()
        if not done.wait(3):
            capture_control.stop()
            done.wait(1)
    except Exception as exc:
        result["error"] = exc

    return result["image"]


def preprocess_for_ocr(image):
    scale = 2
    image = image.resize((image.width * scale, image.height * scale))
    image = ImageOps.grayscale(image)
    return ImageOps.autocontrast(image)


def clean_captured_text(text: str) -> str:
    lines = []
    previous_blank = False

    for line in text.splitlines():
        cleaned = "" if "Speaker " in line else line
        is_blank = not cleaned.strip()
        if is_blank and previous_blank:
            continue
        lines.append(cleaned)
        previous_blank = is_blank

    return "\n".join(lines).strip()


def capture_window_text(target: TargetWindow) -> str:
    error = dependency_error()
    if error:
        return error

    hwnd = resolve_hwnd(target)
    if hwnd is None:
        return (
            "Window not found.\n"
            f"Process: {target.process_name}\n"
            f"Class: {target.expected_class}"
        )

    actual_class = win32gui.GetClassName(hwnd)
    if target.expected_class and actual_class != target.expected_class:
        return (
            f"Window class mismatch.\n"
            f"Expected: {target.expected_class}\n"
            f"Actual: {actual_class}"
        )

    root_hwnd = get_root_hwnd(hwnd)
    if win32gui.IsIconic(root_hwnd):
        return "Caption.Ed is minimized. Restore it so the visible pixels can be OCR-captured."

    image = capture_with_windows_graphics(target) or capture_visible_window(hwnd)
    if image is None:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        return f"Could not capture window bounds: {left}, {top}, {right}, {bottom}"

    image = preprocess_for_ocr(image)

    try:
        text = pytesseract.image_to_string(image, config="--psm 6").strip()
    except pytesseract.TesseractNotFoundError:
        return (
            "Tesseract OCR is not installed or is not on PATH.\n\n"
            "Install it with:\n"
            "winget install UB-Mannheim.TesseractOCR\n\n"
            f"The app also auto-detects: {COMMON_TESSERACT_PATH}"
        )
    except Exception as exc:
        return f"OCR failed: {exc}"

    return clean_captured_text(text) or "(No text detected yet.)"


class CaptureApp(tk.Tk):
    def __init__(self, initial_target: TargetWindow, interval_seconds: float):
        super().__init__()
        self.title("Captioned Live Capture")
        self.geometry("1100x760")
        self.minsize(620, 420)

        self.target = initial_target
        self.interval_seconds = interval_seconds
        self.running = False
        self.worker: threading.Thread | None = None

        self.hwnd_var = tk.StringVar(value="" if initial_target.hwnd is None else str(initial_target.hwnd))
        self.class_var = tk.StringVar(value=initial_target.expected_class)
        self.process_var = tk.StringVar(value=initial_target.process_name)
        self.window_name_var = tk.StringVar(value=initial_target.window_name)
        self.interval_var = tk.StringVar(value=str(interval_seconds))
        self.status_var = tk.StringVar(value="Ready")
        self.text_font = tkfont.Font(family="Consolas", size=11)

        self._build_ui()
        self.after(0, self._maximize)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=8)
        root.pack(fill=tk.BOTH, expand=True)

        button_row = ttk.Frame(root)
        button_row.pack(fill=tk.X, pady=(0, 8))

        self.start_button = ttk.Button(button_row, text="Start", command=self._start)
        self.start_button.pack(side=tk.LEFT)

        self.stop_button = ttk.Button(button_row, text="Stop", command=self._stop, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(8, 0))

        ttk.Button(button_row, text="Capture Once", command=self._capture_once).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(button_row, text="Target Info", command=self._show_target_info).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(button_row, textvariable=self.status_var, anchor=tk.E).pack(side=tk.RIGHT)

        self.text = tk.Text(root, wrap=tk.WORD, font=self.text_font, undo=False)
        self.text.pack(fill=tk.BOTH, expand=True)

    def _maximize(self) -> None:
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

    def _read_target(self) -> TargetWindow | None:
        try:
            hwnd_text = self.hwnd_var.get().strip()
            hwnd = int(hwnd_text, 0) if hwnd_text else None
            interval = float(self.interval_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid Settings", "HWND override and interval must be numbers.")
            return None

        if interval < 0.25:
            messagebox.showerror("Invalid Settings", "Interval must be at least 0.25 seconds.")
            return None

        self.interval_seconds = interval
        return TargetWindow(
            hwnd=hwnd,
            expected_class=self.class_var.get().strip(),
            process_name=self.process_var.get().strip(),
            window_name=self.window_name_var.get().strip(),
        )

    def _start(self) -> None:
        target = self._read_target()
        if target is None:
            return

        self.target = target
        self.running = True
        self.start_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        self.status_var.set("Capturing...")
        self.worker = threading.Thread(target=self._capture_loop, daemon=True)
        self.worker.start()

    def _stop(self) -> None:
        self.running = False
        self.start_button.configure(state=tk.NORMAL)
        self.stop_button.configure(state=tk.DISABLED)
        self.status_var.set("Stopped")

    def _capture_once(self) -> None:
        target = self._read_target()
        if target is None:
            return
        self.target = target
        threading.Thread(target=self._capture_and_display, daemon=True).start()

    def _capture_loop(self) -> None:
        while self.running:
            self._capture_and_display()
            time.sleep(self.interval_seconds)

    def _capture_and_display(self) -> None:
        text = capture_window_text(self.target)
        self.after(0, self._set_text, text)

    def _set_text(self, value: str) -> None:
        if self.text.tag_ranges(tk.SEL):
            self.status_var.set(f"Selection active; update paused at {time.strftime('%H:%M:%S')}")
            return

        self._fit_text_font(value)
        self.text.delete("1.0", tk.END)
        self.text.insert(tk.END, value)
        self.status_var.set(f"Updated at {time.strftime('%H:%M:%S')}")

    def _fit_text_font(self, value: str) -> None:
        lines = value.splitlines() or [""]
        longest_line = max((len(line) for line in lines), default=1)
        height = max(self.text.winfo_height(), 1)
        width = max(self.text.winfo_width(), 1)

        for size in range(11, 7, -1):
            self.text_font.configure(size=size)
            line_height = self.text_font.metrics("linespace") or size
            char_width = max(self.text_font.measure("M"), 1)
            wrapped_lines = sum(max(1, (len(line) * char_width) // width + 1) for line in lines)
            if wrapped_lines * line_height <= height and longest_line:
                return

        self.text_font.configure(size=8)

    def _show_target_info(self) -> None:
        target = self._read_target()
        if target is None:
            return
        self.target = target
        messagebox.showinfo("Target Info", describe_target(target))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="OCR text from the Caption.Ed Chromium render window.")
    parser.add_argument("--hwnd", type=lambda value: int(value, 0), default=None)
    parser.add_argument("--class-name", default=DEFAULT_CLASS)
    parser.add_argument("--process-name", default=DEFAULT_PROCESS_NAME)
    parser.add_argument("--window-name", default=DEFAULT_WINDOW_NAME)
    parser.add_argument("--interval", type=float, default=DEFAULT_INTERVAL_SECONDS)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    app = CaptureApp(
        initial_target=TargetWindow(
            hwnd=args.hwnd,
            expected_class=args.class_name,
            process_name=args.process_name,
            window_name=args.window_name,
        ),
        interval_seconds=args.interval,
    )
    app.mainloop()


if __name__ == "__main__":
    main()
