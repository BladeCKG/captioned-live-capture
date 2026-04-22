import argparse
import difflib
import os
import re
import sys
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
DEFAULT_INTERVAL_SECONDS = 0.3
COMMON_TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
SPEAKER_LINE_PATTERN = re.compile(r"\bspe[a-z]*ker\s*\d*\b", re.IGNORECASE)
WORD_PATTERN = re.compile(r"[A-Za-z]+(?:['-][A-Za-z]+)?")
COMMON_TRANSCRIPT_WORDS = {
    "a",
    "about",
    "all",
    "am",
    "an",
    "and",
    "are",
    "as",
    "at",
    "be",
    "but",
    "can",
    "do",
    "for",
    "from",
    "go",
    "have",
    "he",
    "i",
    "if",
    "in",
    "is",
    "it",
    "me",
    "my",
    "no",
    "not",
    "of",
    "okay",
    "on",
    "or",
    "our",
    "right",
    "so",
    "that",
    "the",
    "then",
    "there",
    "this",
    "to",
    "uh",
    "um",
    "we",
    "well",
    "what",
    "with",
    "would",
    "yeah",
    "yes",
    "you",
    "your",
}
ALLOWED_UPPERCASE_WORDS = {"CEO", "SS"}


def app_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def bundled_tesseract_path() -> str:
    return os.path.join(app_base_dir(), "tesseract", "tesseract.exe")


def configure_tesseract() -> None:
    if pytesseract is None:
        return

    candidates = [bundled_tesseract_path(), COMMON_TESSERACT_PATH]
    for candidate in candidates:
        if os.path.exists(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            tessdata = os.path.join(os.path.dirname(candidate), "tessdata")
            if os.path.isdir(tessdata):
                os.environ.setdefault("TESSDATA_PREFIX", tessdata)
            return


configure_tesseract()


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

    for line in text.splitlines():
        if SPEAKER_LINE_PATTERN.search(line):
            if lines and lines[-1] != "":
                lines.append("")
            if lines and lines[-1] == "":
                lines.append("")
            continue
        cleaned = line.rstrip()
        is_blank = not cleaned.strip()
        if is_blank and lines and lines[-1] == "":
            continue
        if not is_blank and not looks_like_sentence_part(cleaned):
            continue
        lines.append(cleaned)

    return "\n".join(lines).strip()


def normalized_line(line: str) -> str:
    return " ".join(WORD_PATTERN.findall(line.casefold()))


def line_similarity(left: str, right: str) -> float:
    normalized_left = normalized_line(left)
    normalized_right = normalized_line(right)
    if not normalized_left or not normalized_right:
        return 1.0 if normalized_left == normalized_right else 0.0
    return difflib.SequenceMatcher(None, normalized_left, normalized_right).ratio()


def split_paragraphs(text: str) -> list[str]:
    return [paragraph.strip() for paragraph in re.split(r"\n\s*\n+", text.strip()) if paragraph.strip()]


def paragraph_similarity(left: str, right: str) -> float:
    return line_similarity(left.replace("\n", " "), right.replace("\n", " "))


def paragraphs_related(left: str, right: str) -> bool:
    left_norm = normalized_line(left)
    right_norm = normalized_line(right)
    if not left_norm or not right_norm:
        return False
    if left_norm in right_norm or right_norm in left_norm:
        return True
    return difflib.SequenceMatcher(None, left_norm, right_norm).ratio() >= 0.52


def merge_capture_text(existing: str, captured: str, preserve_tail: bool = False) -> tuple[str, bool]:
    captured = captured.strip()
    if not captured or captured == "(No text detected yet.)":
        return existing, False

    existing = existing.strip()
    if not existing:
        return captured, True

    existing_paragraphs = split_paragraphs(existing)
    captured_paragraphs = split_paragraphs(captured)
    if not captured_paragraphs:
        return existing, False

    locked_existing = existing_paragraphs[:-1]
    tail_existing = existing_paragraphs[-1] if existing_paragraphs else ""
    merged_tail = tail_existing
    replacement_candidates: list[str] = []
    additions: list[str] = []

    tail_match_index = -1
    tail_match_score = 0.0
    for index, paragraph in enumerate(captured_paragraphs):
        score = paragraph_similarity(tail_existing, paragraph)
        if score > tail_match_score:
            tail_match_score = score
            tail_match_index = index

    tail_matches = tail_match_index >= 0 and paragraphs_related(tail_existing, captured_paragraphs[tail_match_index])

    if tail_matches:
        candidate_tail = captured_paragraphs[tail_match_index]
        if not preserve_tail and should_replace_tail_paragraph(tail_existing, candidate_tail):
            merged_tail = candidate_tail
        replacement_candidates = captured_paragraphs[:tail_match_index]
        additions = captured_paragraphs[tail_match_index + 1 :]
    else:
        last_locked = locked_existing[-1] if locked_existing else ""
        start_index = 0
        if last_locked:
            for index, paragraph in enumerate(captured_paragraphs):
                if paragraph_similarity(last_locked, paragraph) >= 0.72:
                    start_index = index + 1
        additions = captured_paragraphs[start_index:]

    paragraphs = locked_existing + [merged_tail]
    for paragraph in replacement_candidates:
        replacement_index = find_replaceable_fragment_index(paragraphs, paragraph)
        if replacement_index is not None:
            paragraphs[replacement_index] = paragraph

    for paragraph in additions:
        if not paragraph:
            continue
        replacement_index = find_replaceable_fragment_index(paragraphs, paragraph)
        if replacement_index is not None:
            paragraphs[replacement_index] = paragraph
            continue
        if any(paragraph_similarity(paragraph, existing_paragraph) >= 0.88 for existing_paragraph in paragraphs[-6:]):
            continue
        paragraphs.append(paragraph)

    merged = "\n\n".join(paragraphs).strip()
    return merged, merged != existing


def should_replace_tail_paragraph(existing_tail: str, candidate_tail: str) -> bool:
    if not candidate_tail.strip():
        return False
    existing_words = WORD_PATTERN.findall(existing_tail)
    candidate_words = WORD_PATTERN.findall(candidate_tail)
    if len(candidate_words) > len(existing_words):
        return True
    if len(candidate_tail) > len(existing_tail) + 20:
        return True
    return False


def find_replaceable_fragment_index(paragraphs: list[str], candidate: str) -> int | None:
    for index in range(max(0, len(paragraphs) - 8), len(paragraphs)):
        existing = paragraphs[index]
        if is_finished_paragraph(existing):
            continue
        if candidate_extends_fragment(existing, candidate):
            return index
    return None


def is_finished_paragraph(paragraph: str) -> bool:
    words = WORD_PATTERN.findall(paragraph)
    stripped = paragraph.strip()
    if len(words) < 6:
        return False
    return bool(re.search(r"[.!?][\"')\]]*$", stripped))


def candidate_extends_fragment(fragment: str, candidate: str) -> bool:
    fragment_norm = normalized_line(fragment)
    candidate_norm = normalized_line(candidate)
    if not fragment_norm or not candidate_norm:
        return False
    if len(candidate_norm) <= len(fragment_norm) + 8:
        return False
    if fragment_norm in candidate_norm:
        return True

    fragment_words = fragment_norm.split()
    candidate_words = candidate_norm.split()
    if len(fragment_words) <= 8:
        prefix = " ".join(candidate_words[: max(len(fragment_words) + 3, 1)])
        if difflib.SequenceMatcher(None, fragment_norm, prefix).ratio() >= 0.58:
            return True

    window_size = min(max(len(fragment_words) + 3, 4), len(candidate_words))
    best_score = 0.0
    for start in range(0, max(1, len(candidate_words) - window_size + 1)):
        window = " ".join(candidate_words[start : start + window_size])
        best_score = max(best_score, difflib.SequenceMatcher(None, fragment_norm, window).ratio())
    return best_score >= 0.62


def is_ui_noise_line(line: str) -> bool:
    lowered = line.casefold()
    if "untitled recording" in lowered:
        return True
    if "feedback" in lowered and "share" in lowered:
        return True
    if "processing audio" in lowered:
        return True
    if re.search(r"\b\d{1,2}/\d{1,2}/\d{4}\b", lowered):
        return True
    if re.fullmatch(r"[\W_]*\d+x[\W_]*", lowered):
        return True
    if re.search(r"\b\d{1,2}:\d{2}\b", lowered) and not WORD_PATTERN.findall(lowered):
        return True
    return False


def looks_like_sentence_part(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return True
    if is_ui_noise_line(stripped):
        return False

    words = WORD_PATTERN.findall(stripped)
    letters = sum(ch.isalpha() for ch in stripped)
    alnum = sum(ch.isalnum() for ch in stripped)
    if letters < 4 or len(words) < 2:
        return False

    if alnum:
        symbol_ratio = 1 - (alnum / len(stripped))
        if symbol_ratio > 0.45:
            return False

    normalized_words = [word.casefold().strip("'-") for word in words]
    common_word_count = sum(word in COMMON_TRANSCRIPT_WORDS for word in normalized_words)
    short_word_count = sum(len(word) <= 2 for word in normalized_words)
    malformed_word_count = sum(looks_malformed_word(word) for word in words)
    if len(words) <= 6 and short_word_count / len(words) > 0.6:
        return False
    if len(words) >= 6 and malformed_word_count / len(words) > 0.25:
        return False
    if len(words) <= 5 and common_word_count == 0 and not re.search(r"[.!?]", stripped):
        return False

    average_word_length = sum(len(word) for word in words) / len(words)
    return average_word_length >= 2.2


def looks_malformed_word(word: str) -> bool:
    if len(word) < 4:
        return False
    if word.upper() in ALLOWED_UPPERCASE_WORDS:
        return False
    if word.casefold() in COMMON_TRANSCRIPT_WORDS:
        return False

    uppercase_count = sum(ch.isupper() for ch in word)
    if word.isupper():
        return True
    return uppercase_count >= 2


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
            "The app auto-detects bundled Tesseract at:\n"
            f"{bundled_tesseract_path()}\n\n"
            f"It also checks: {COMMON_TESSERACT_PATH}"
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
        self.autoscroll_var = tk.BooleanVar(value=True)
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
        ttk.Checkbutton(button_row, text="Auto-scroll", variable=self.autoscroll_var).pack(side=tk.LEFT, padx=(12, 0))

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
        existing = self.text.get("1.0", tk.END).strip()
        selection_active = bool(self.text.tag_ranges(tk.SEL))
        merged, changed = merge_capture_text(existing, value, preserve_tail=selection_active)
        if not changed:
            self.status_var.set(f"No new text at {time.strftime('%H:%M:%S')}")
            return

        self._fit_text_font(merged)
        if selection_active:
            if merged.startswith(existing):
                self.text.insert(tk.END, merged[len(existing) :])
                self._maybe_autoscroll()
                self.status_var.set(f"Appended at {time.strftime('%H:%M:%S')}")
            else:
                self.status_var.set(f"Selection active; tail update deferred at {time.strftime('%H:%M:%S')}")
            return

        self.text.delete("1.0", tk.END)
        self.text.insert(tk.END, merged)
        self._maybe_autoscroll()
        self.status_var.set(f"Updated at {time.strftime('%H:%M:%S')}")

    def _maybe_autoscroll(self) -> None:
        if self.autoscroll_var.get():
            self.text.see(tk.END)

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
