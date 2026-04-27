import argparse
import threading
import time
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox, ttk

from capture_backend import (
    DEFAULT_CLASS,
    DEFAULT_PROCESS_NAME,
    DEFAULT_WINDOW_NAME,
    TargetWindow,
    capture_window_text,
    describe_target,
)
from text_processing import split_paragraphs


DEFAULT_INTERVAL_SECONDS = 0.3
PARAGRAPH_SEPARATOR = "\n\n"


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
        self.always_on_top_var = tk.BooleanVar(value=False)
        self.auto_copy_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready")
        self.text_font = tkfont.Font(family="Consolas", size=11)
        self.pointer_active = False
        self.pointer_offset: int | None = None
        self.displayed_paragraphs: list[str] = []
        self.autoscroll_token = 0

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
        ttk.Checkbutton(
            button_row,
            text="Always top",
            variable=self.always_on_top_var,
            command=self._apply_always_on_top,
        ).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Checkbutton(button_row, text="Auto-copy", variable=self.auto_copy_var).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(button_row, textvariable=self.status_var, anchor=tk.E).pack(side=tk.RIGHT)

        text_frame = ttk.Frame(root)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.text = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=self.text_font,
            undo=False,
            exportselection=False,
            selectbackground="#2f6fed",
            selectforeground="#ffffff",
        )
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text.configure(yscrollcommand=scrollbar.set)
        self.text.tag_configure("persistent_selection", background="#2f6fed", foreground="#ffffff")
        self.text.bind("<ButtonRelease-1>", self._set_pointer_from_click, add="+")

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
        self._initialize_worker_thread()
        next_run = time.monotonic()
        try:
            while self.running:
                self._capture_and_display()
                next_run += self.interval_seconds
                remaining = next_run - time.monotonic()
                if remaining > 0:
                    time.sleep(remaining)
                else:
                    next_run = time.monotonic()
        finally:
            self._finalize_worker_thread()

    def _initialize_worker_thread(self) -> None:
        try:
            from capture_backend import ensure_com_initialized

            ensure_com_initialized()
        except Exception:
            pass

    def _finalize_worker_thread(self) -> None:
        try:
            from capture_backend import release_com_if_initialized

            release_com_if_initialized()
        except Exception:
            pass

    def _capture_and_display(self) -> None:
        text = capture_window_text(self.target)
        self.after(0, self._set_text, text)

    def _set_text(self, value: str) -> None:
        if value.startswith("Window not found.") or value.startswith("Window class mismatch.") or value.startswith("Transcript control was not found"):
            self.status_var.set(value.splitlines()[0])
            return

        new_paragraphs = split_paragraphs(value)
        if not new_paragraphs and value.strip() == "(No text detected yet.)":
            self.status_var.set(f"No new text at {time.strftime('%H:%M:%S')}")
            return

        changed = self._apply_transcript_update(new_paragraphs)
        if not changed:
            self._restore_pointer_selection()
            self.status_var.set(f"No new text at {time.strftime('%H:%M:%S')}")
            return

        self._restore_pointer_selection()
        self._maybe_autoscroll(force=True)
        self._maybe_auto_copy_selection()
        self.status_var.set(f"Updated at {time.strftime('%H:%M:%S')}")

    def _apply_transcript_update(self, new_paragraphs: list[str]) -> bool:
        if new_paragraphs == self.displayed_paragraphs:
            return False

        prefix_len = 0
        max_prefix = min(len(self.displayed_paragraphs), len(new_paragraphs))
        while prefix_len < max_prefix and self.displayed_paragraphs[prefix_len] == new_paragraphs[prefix_len]:
            prefix_len += 1

        if prefix_len == 0:
            self.text.delete("1.0", tk.END)
            if new_paragraphs:
                self.text.insert("1.0", PARAGRAPH_SEPARATOR.join(new_paragraphs))
        else:
            start_index = self._paragraph_start_index(prefix_len)
            self.text.delete(start_index, tk.END)
            suffix_text = PARAGRAPH_SEPARATOR.join(new_paragraphs[prefix_len:])
            if suffix_text:
                if prefix_len > 0:
                    self.text.insert(start_index, PARAGRAPH_SEPARATOR + suffix_text)
                else:
                    self.text.insert(start_index, suffix_text)

        self.displayed_paragraphs = list(new_paragraphs)
        return True

    def _paragraph_start_index(self, paragraph_index: int) -> str:
        char_offset = 0
        for index, paragraph in enumerate(self.displayed_paragraphs[:paragraph_index]):
            char_offset += len(paragraph)
            if index < paragraph_index - 1:
                char_offset += len(PARAGRAPH_SEPARATOR)
        return f"1.0+{char_offset}c"

    def _maybe_autoscroll(self, force: bool = False) -> None:
        if not self.autoscroll_var.get():
            return
        if self.pointer_active and not force:
            return
        self.autoscroll_token += 1
        token = self.autoscroll_token
        self.after_idle(lambda: self._scroll_to_end(token))

    def _scroll_to_end(self, token: int) -> None:
        if token != self.autoscroll_token:
            return
        self.text.update_idletasks()
        self.text.see("end-1c")
        self.text.yview_moveto(1.0)

    def _apply_always_on_top(self) -> None:
        self.attributes("-topmost", self.always_on_top_var.get())

    def _maybe_auto_copy_selection(self) -> None:
        if not self.auto_copy_var.get():
            return
        ranges = self.text.tag_ranges("persistent_selection") or self.text.tag_ranges(tk.SEL)
        if not ranges:
            return
        selected_text = self.text.get(ranges[0], ranges[-1]).strip()
        if not selected_text:
            return
        self.clipboard_clear()
        self.clipboard_append(selected_text)

    def _set_pointer_from_click(self, event: tk.Event) -> None:
        click_index = self.text.index(f"@{event.x},{event.y}")
        self.pointer_active = True
        self.pointer_offset = self._index_to_offset(click_index)
        self.text.tag_remove(tk.SEL, "1.0", tk.END)
        self.text.tag_remove("persistent_selection", "1.0", tk.END)
        self._select_pointer_to_end()
        self._maybe_auto_copy_selection()
        self.status_var.set("Pointer set from click")

    def _index_to_offset(self, index: str) -> int:
        try:
            return int(self.text.count("1.0", index, "chars")[0])
        except (tk.TclError, TypeError, ValueError):
            return 0

    def _restore_pointer_selection(self) -> None:
        if not self.pointer_active or self.pointer_offset is None:
            return
        self._select_pointer_to_end()

    def _select_pointer_to_end(self) -> None:
        if not self.pointer_active or self.pointer_offset is None:
            return
        try:
            text_length = int(self.text.count("1.0", "end-1c", "chars")[0])
        except (tk.TclError, TypeError, ValueError):
            return
        safe_offset = max(0, min(self.pointer_offset, text_length))
        start_index = f"1.0+{safe_offset}c"
        self.text.tag_remove(tk.SEL, "1.0", tk.END)
        self.text.tag_remove("persistent_selection", "1.0", tk.END)
        if self.text.compare(start_index, "<", "end-1c"):
            self.text.tag_add(tk.SEL, start_index, "end-1c")
            self.text.tag_add("persistent_selection", start_index, "end-1c")
        self.text.mark_set(tk.INSERT, "end-1c")
        self.text.see(start_index)

    def _show_target_info(self) -> None:
        target = self._read_target()
        if target is None:
            return
        self.target = target
        messagebox.showinfo("Target Info", describe_target(target))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Read transcript text from the Caption.Ed Chromium window.")
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
