# Captioned Live Capture

A small Windows Python UI that OCR-captures the visible Caption.Ed render window and displays the detected text.

The UI is intentionally minimal: only the capture buttons stay visible, and the captured text fills the rest of the window.

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
py capture_text_app.py --process-name Caption.Ed.exe --class-name Chrome_RenderWidgetHostHWND --interval 1
```

If you really need to force one specific handle:

```powershell
py capture_text_app.py --hwnd 394966
```

## Notes

`Chrome_RenderWidgetHostHWND` itself is only a render surface. Standard Win32 APIs usually return only `Chrome Legacy Window`, not the rendered caption text. This version uses screenshot OCR because Caption.Ed did not expose the text through UI Automation or a DevTools endpoint.
