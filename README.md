# Captioned Live Capture

A small Windows Python UI that reads transcript text from Caption.Ed through Windows UI Automation and displays it live.

The UI is intentionally minimal: only the capture buttons stay visible, and the captured text fills the rest of the window.

Each update reads the current transcript from Caption.Ed, splits it into paragraphs, and updates only the changed tail of the displayed text.

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

Create the portable Windows build with:

```powershell
.\build_release.ps1
```

The portable Windows build is created at:

```text
release\CaptionedLiveCapture-portable.zip
```

To use it on another PC, extract the zip and run:

```text
CaptionedLiveCapture.exe
```

The portable folder includes the Python runtime and Python packages, so Python does not need to be installed separately.

## Notes

`Chrome_RenderWidgetHostHWND` itself is only a render surface. Standard Win32 APIs usually return only `Chrome Legacy Window`, not the rendered caption text.

For Caption.Ed specifically, the transcript is exposed through the app's accessibility tree. The app reads the `Transcript` UI Automation control directly instead of using OCR.
