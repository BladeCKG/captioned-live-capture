from __future__ import annotations

import platform
from dataclasses import dataclass


IS_WINDOWS = platform.system() == "Windows"
IS_MACOS = platform.system() == "Darwin"

DEFAULT_CLASS = "Chrome_RenderWidgetHostHWND" if IS_WINDOWS else ""
DEFAULT_PROCESS_NAME = "Caption.Ed.exe" if IS_WINDOWS else "Caption.Ed"
DEFAULT_WINDOW_NAME = "Caption.Ed" if IS_WINDOWS else ""


@dataclass(frozen=True)
class TargetWindow:
    hwnd: int | None = None
    expected_class: str = DEFAULT_CLASS
    process_name: str = DEFAULT_PROCESS_NAME
    window_name: str = DEFAULT_WINDOW_NAME
