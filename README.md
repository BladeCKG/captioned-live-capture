# Captioned Live Capture

A small Windows Python UI that captures Caption.Ed with Windows Graphics Capture, OCRs the frame, and displays the detected text.

The UI is intentionally minimal: only the capture buttons stay visible, and the captured text fills the rest of the window.

Each capture is compared fuzzily with the existing text. Paragraphs that already have a blank line before and after them are treated as finalized and are not changed again; only the trailing in-progress paragraph can be improved, and new paragraphs are appended.

Use the `Auto-scroll` checkbox to choose whether the text view jumps to the latest captured text after updates.

Use `Always top` to keep the app above other windows.

Use `Auto-copy` to copy the currently selected text to the clipboard after capture updates.

Click inside the captured text to set the pointer. The pointer stays fixed at that clicked position, and each update selects text from that pointer to the latest captured text.

The default target is discovered automatically by:

- Class: `Chrome_RenderWidgetHostHWND`
- Process: `Caption.Ed.exe`

The HWND value is intentionally optional because it changes whenever the app/window is recreated.

## Setup

Install Python packages:

```powershell
py -m pip install -r requirements.txt
```

Install the Tesseract OCR engine if it is not already installed:

```powershell
winget install UB-Mannheim.TesseractOCR
```

The app auto-detects the common install path:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

## Run

```powershell
py capture_text_app.py
```

Or override the target window:

```powershell
py capture_text_app.py --process-name Caption.Ed.exe --class-name Chrome_RenderWidgetHostHWND --interval 0.3
```

If you really need to force one specific handle:

```powershell
py capture_text_app.py --hwnd 394966
```

## Portable Build

The portable Windows build is created at:

```text
release\CaptionedLiveCapture-portable.zip
```

To use it on another PC, extract the zip and run:

```text
CaptionedLiveCapture.exe
```

The portable folder includes the Python runtime, Python packages, and Tesseract OCR runtime, so Python and Tesseract do not need to be installed separately.

## Notes

`Chrome_RenderWidgetHostHWND` itself is only a render surface. Standard Win32 APIs usually return only `Chrome Legacy Window`, not the rendered caption text.

The app now tries Windows Graphics Capture first using the Caption.Ed window title. This is the correct Windows compositor capture path for windows that are covered by other windows or off-screen-but-not-minimized. If Windows Graphics Capture is unavailable or fails, it falls back to visible screenshot OCR of the Chromium render widget.
