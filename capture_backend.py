from __future__ import annotations

import platform
import threading

from captioned_text import (
    normalize_transcript_paragraph,
    parse_transcript_value,
)
from capture_types import (
    DEFAULT_CLASS,
    DEFAULT_PROCESS_NAME,
    DEFAULT_WINDOW_NAME,
    IS_MACOS,
    IS_WINDOWS,
    TargetWindow,
)
from chrome_live_caption_common import is_live_caption_target


THREAD_STATE = threading.local()


def dependency_error(target: TargetWindow | None = None) -> str | None:
    backend = backend_for_target(target or TargetWindow())
    if hasattr(backend, "dependency_error"):
        return backend.dependency_error()
    return None


def ensure_com_initialized() -> None:
    if not IS_WINDOWS:
        return
    from captioned_windows import ensure_com_initialized as windows_ensure_com_initialized

    windows_ensure_com_initialized(THREAD_STATE)


def release_com_if_initialized() -> None:
    if not IS_WINDOWS:
        return
    from captioned_windows import release_com_if_initialized as windows_release_com_if_initialized

    windows_release_com_if_initialized(THREAD_STATE)


def extract_transcript_text(target: TargetWindow) -> str:
    error = dependency_error(target)
    if error:
        return error

    if IS_WINDOWS:
        ensure_com_initialized()

    backend = backend_for_target(target)
    session_type = backend_session_type(backend)
    session = getattr(THREAD_STATE, "transcript_session", None)
    if session is None or session.target != target or not isinstance(session, session_type):
        session = session_type(target)
        THREAD_STATE.transcript_session = session
    return session.extract_text()


def capture_window_text(target: TargetWindow, session=None) -> str:
    return extract_transcript_text(target)


def describe_target(target: TargetWindow) -> str:
    backend = backend_for_target(target)
    if hasattr(backend, "describe_target"):
        return backend.describe_target(target)
    return f"Unsupported platform: {platform.system()}"


def backend_for_target(target: TargetWindow):
    if IS_WINDOWS:
        return windows_backend_for_target(target)
    if IS_MACOS:
        return macos_backend_for_target(target)
    raise RuntimeError(f"Unsupported platform: {platform.system()}")


def windows_backend_for_target(target: TargetWindow):
    # Chrome Live Caption on Windows is discovered through the same Win32/UIA
    # window path as Caption.Ed, but extraction is delegated inside that module.
    import captioned_windows

    return captioned_windows


def macos_backend_for_target(target: TargetWindow):
    if is_live_caption_target(target):
        import chrome_live_caption_macos

        return chrome_live_caption_macos
    import captioned_macos

    return captioned_macos


def backend_session_type(backend):
    if hasattr(backend, "TranscriptAutomationSession"):
        return backend.TranscriptAutomationSession
    if hasattr(backend, "MacCaptionedSession"):
        return backend.MacCaptionedSession
    if hasattr(backend, "MacChromeLiveCaptionSession"):
        return backend.MacChromeLiveCaptionSession
    raise RuntimeError(f"Backend does not expose a session type: {backend.__name__}")
